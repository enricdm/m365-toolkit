<#
.SYNOPSIS
    Creates Entra ID security groups from a definition list and populates their
    members. Preview by default; writes only with -Execute.

.DESCRIPTION
    Data-driven and idempotent:
      - If the group already exists (matched by displayName), it is reused.
      - Members already in the group are skipped.
      - UPNs that don't resolve to a user are reported, not fatal.
    Groups are created as plain security groups (mailEnabled=$false,
    securityEnabled=$true, no groupTypes), which is what tools like Microsoft
    Fabric expect for workspace access.

    Definitions come either from -DefinitionsPath (a .psd1 data file) or from the
    CONFIG block below. Either way, each entry needs Name, Description and Members.

    Output: a per-user result table on screen and a CSV log next to the script.

.PARAMETER DefinitionsPath
    Optional .psd1 file with a 'Groups' key, e.g.

        @{ Groups = @(
            @{ Name = 'App_Dev_ES'
               Description = 'Application developers - Spain'
               Members = @('user1@contoso.com', 'user2@contoso.com') }
        ) }

.PARAMETER Execute
    Actually create groups and add members. Without it the script runs in preview
    mode and only reports what it would do.

.EXAMPLE
    .\New-SecurityGroup.ps1 -DefinitionsPath .\groups.psd1
    Preview: shows every group and member that would be created.

.EXAMPLE
    .\New-SecurityGroup.ps1 -DefinitionsPath .\groups.psd1 -Execute
    Creates the groups and adds the members.

.NOTES
    When to use  : A new project or workspace needs several groups with their members created in one go.
    Why it exists: Data-driven from a .psd1 and idempotent: an existing group is reused, members already in it are skipped, and a UPN that does not resolve is reported rather than fatal. Preview is the default and writes need -Execute.
    Requires: Microsoft.Graph (Groups + Users sub-modules).
    Scopes  : Group.ReadWrite.All, User.Read.All, GroupMember.ReadWrite.All
    WRITES TO THE DIRECTORY when -Execute is passed (creates groups, adds members).
#>

[CmdletBinding()]
param(
    [string]$DefinitionsPath,
    [switch]$Execute
)

# ---------------------------------------------------------------------------
# 1. CONFIG — used only when -DefinitionsPath is not supplied.
#    Add one block per group. Kept commented out on purpose: an empty or
#    half-filled template here would create a junk group in the tenant.
# ---------------------------------------------------------------------------
$GroupDefinitions = @(
    # [pscustomobject]@{
    #     Name        = 'App_Dev_ES'
    #     Description = 'Application developers - Spain'
    #     Members     = @(
    #         'user1@contoso.com'
    #         'user2@contoso.com'
    #     )
    # }
)

if ($DefinitionsPath) {
    if (-not (Test-Path $DefinitionsPath)) { throw "Definitions file not found: $DefinitionsPath" }
    $data = Import-PowerShellDataFile -Path $DefinitionsPath
    $GroupDefinitions = @($data.Groups | ForEach-Object { [pscustomobject]$_ })
}

$GroupDefinitions = @($GroupDefinitions | Where-Object { $_.Name })
if ($GroupDefinitions.Count -eq 0) {
    Write-Warning "No group definitions. Pass -DefinitionsPath, or fill in the CONFIG block. Nothing to do."
    return
}

# Preview unless -Execute was passed.
$WhatIfMode = -not $Execute

# ---------------------------------------------------------------------------
# 2. CONNECT
# ---------------------------------------------------------------------------
$requiredScopes = @('Group.ReadWrite.All', 'User.Read.All', 'GroupMember.ReadWrite.All')

if (-not (Get-MgContext)) {
    Connect-MgGraph -Scopes $requiredScopes | Out-Null
}
Write-Host "Connected as: $((Get-MgContext).Account)" -ForegroundColor Cyan
Write-Host ("Mode: {0}`n" -f $(if ($WhatIfMode) { 'PREVIEW (no changes)' } else { 'LIVE' })) -ForegroundColor Cyan

$results = New-Object System.Collections.Generic.List[object]

# ---------------------------------------------------------------------------
# 3. PROCESS
# ---------------------------------------------------------------------------
foreach ($def in $GroupDefinitions) {

    Write-Host "=== $($def.Name) ===" -ForegroundColor Yellow

    # --- ensure the group exists ---------------------------------------
    $group = Get-MgGroup -Filter "displayName eq '$($def.Name)'" -ConsistencyLevel eventual -CountVariable c -All |
    Select-Object -First 1

    if ($group) {
        Write-Host "  Group exists (Id $($group.Id)) - reusing." -ForegroundColor DarkGray
    } else {
        $mailNick = ($def.Name -replace '[^a-zA-Z0-9._-]', '')
        if ($WhatIfMode) {
            Write-Host "  [PREVIEW] would create security group '$($def.Name)'." -ForegroundColor Magenta
            $group = $null
        } else {
            $group = New-MgGroup -DisplayName $def.Name `
                -Description $def.Description `
                -MailNickname $mailNick `
                -SecurityEnabled:$true `
                -MailEnabled:$false
            Write-Host "  Created group (Id $($group.Id))." -ForegroundColor Green
        }
    }

    # current members (for idempotency) ---------------------------------
    $existingMemberIds = @()
    if ($group) {
        $existingMemberIds = (Get-MgGroupMember -GroupId $group.Id -All).Id
    }

    # --- members --------------------------------------------------------
    foreach ($upn in $def.Members) {

        $row = [pscustomobject]@{
            Group  = $def.Name
            Member = $upn
            Action = $null
            Detail = $null
        }

        # resolve the user
        $user = $null
        try { $user = Get-MgUser -UserId $upn -ErrorAction Stop }
        catch { $user = $null }

        if (-not $user) {
            $row.Action = 'NOT FOUND'
            $row.Detail = 'UPN did not resolve to a user object'
            Write-Host "  [!] $upn - not found" -ForegroundColor Red
            $results.Add($row); continue
        }

        if ($existingMemberIds -contains $user.Id) {
            $row.Action = 'ALREADY MEMBER'
            Write-Host "  [=] $upn - already a member" -ForegroundColor DarkGray
            $results.Add($row); continue
        }

        if ($WhatIfMode -or -not $group) {
            $row.Action = 'WOULD ADD'
            Write-Host "  [PREVIEW] would add $upn" -ForegroundColor Magenta
            $results.Add($row); continue
        }

        try {
            $ref = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($user.Id)" }
            New-MgGroupMemberByRef -GroupId $group.Id -BodyParameter $ref -ErrorAction Stop
            $row.Action = 'ADDED'
            Write-Host "  [+] $upn - added" -ForegroundColor Green
        } catch {
            $row.Action = 'ERROR'
            $row.Detail = $_.Exception.Message
            Write-Host "  [x] $upn - $($_.Exception.Message)" -ForegroundColor Red
        }
        $results.Add($row)
    }
    Write-Host ''
}

# ---------------------------------------------------------------------------
# 4. SUMMARY + LOG
# ---------------------------------------------------------------------------
Write-Host "==== SUMMARY ====" -ForegroundColor Cyan
$results | Format-Table -AutoSize

$logPath = Join-Path $PSScriptRoot ("SecurityGroups_{0:yyyyMMdd_HHmmss}.csv" -f (Get-Date))
$results | Export-Csv -Path $logPath -NoTypeInformation -Encoding UTF8
Write-Host "Log written to: $logPath" -ForegroundColor Cyan
