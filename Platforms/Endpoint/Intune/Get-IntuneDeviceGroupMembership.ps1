<#
.SYNOPSIS
    Resolves the relationship between Intune devices and Entra ID groups, in either
    direction: which groups a device belongs to, or which devices are in a group -
    including devices reached indirectly through their primary user's groups.

.DESCRIPTION
    Answers the question you actually have when a policy lands somewhere unexpected:
    "why is this device getting that?"

    Note this is the complement of Get-IntuneGroupAssignment.ps1:
      Get-IntuneGroupAssignment      group -> what is assigned to it
      Get-IntuneDeviceGroupMembership device <-> which groups it is in
    Together they close the loop from device to policy.

    DEVICE GROUPS AND USER GROUPS ARE NOT THE SAME THING, and conflating them is the
    usual reason a policy appears to apply "for no reason". An Intune policy assigned
    to a USER group reaches a device through whoever is its primary user; assigned to
    a DEVICE group it reaches the device directly. -IncludeUserGroups reports both and
    labels which is which in the MembershipVia column.

    The Entra device object is what holds group membership, not the Intune managed
    device, so the two are matched on azureADDeviceId.

.PARAMETER DeviceName
    Report the groups these devices belong to. Wildcards accepted.

.PARAMETER GroupName
    Report the Intune devices in these groups. Matched exactly, not as a substring.

.PARAMETER IncludeUserGroups
    Also report groups reached through the device's primary user. Costs one extra
    call per distinct user.

.PARAMETER OutputPath
    Optional CSV path. Nothing is written unless supplied.

.EXAMPLE
    # Which groups is this device in, directly and via its user
    .\Get-IntuneDeviceGroupMembership.ps1 -DeviceName 'LAP-0042' -IncludeUserGroups

.EXAMPLE
    # All laptops matching a naming convention
    .\Get-IntuneDeviceGroupMembership.ps1 -DeviceName 'LAP-BCN-*' -OutputPath .\Exports\membership.csv

.EXAMPLE
    # Which Intune devices are in these groups
    .\Get-IntuneDeviceGroupMembership.ps1 -GroupName 'GRP-Devices-Kiosk','GRP-Devices-Classrooms'

.NOTES
    Requires : Microsoft.Graph.Authentication  (Install-Module Microsoft.Graph -Scope CurrentUser)
    Scopes   : DeviceManagementManagedDevices.Read.All, Device.Read.All,
               Group.Read.All, User.Read.All (with -IncludeUserGroups)
    Rights   : read-only.

    DYNAMIC GROUP MEMBERSHIP IS EVALUATED ASYNCHRONOUSLY BY ENTRA. A device that has
    just been enrolled, renamed or recategorised may not yet appear in a dynamic group
    whose rule it already satisfies. An absent row here means "not a member right now",
    not "the rule does not match it".

    Nested groups are reported one level deep by default; transitive membership is not
    expanded. If your assignment model relies on nested groups, treat this output as a
    starting point rather than the final answer.

    Replaces (merged): a device group-membership report that existed as two
    divergent copies, plus a third script that exported the devices belonging to
    the members of a user group. Both directions are parameters here.
#>

[CmdletBinding(DefaultParameterSetName = 'ByDevice')]
param(
    [Parameter(Mandatory, ParameterSetName = 'ByDevice')]
    [string[]]$DeviceName,

    [Parameter(Mandatory, ParameterSetName = 'ByGroup')]
    [string[]]$GroupName,

    [switch]$IncludeUserGroups,

    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot '..\..\_Shared\Modules\M365.Common.psm1') -Force -ErrorAction Stop

function Write-Step { param($m) Write-Host "`n[>] $m" -ForegroundColor Cyan }
function Write-OK   { param($m) Write-Host "[OK] $m"  -ForegroundColor Green }
function Write-Warn { param($m) Write-Host "[!!] $m"  -ForegroundColor Yellow }

$scopes = @('DeviceManagementManagedDevices.Read.All', 'Device.Read.All', 'Group.Read.All')
if ($IncludeUserGroups) { $scopes += 'User.Read.All' }
Connect-MgGraph -Scopes $scopes -NoWelcome

function Get-Prop {
    param($Obj, [string[]]$Names)
    if ($null -eq $Obj) { return $null }
    foreach ($n in $Names) {
        if ($Obj -is [System.Collections.IDictionary]) {
            if ($Obj.Contains($n)) { return $Obj[$n] }
        } else {
            $p = $Obj.PSObject.Properties[$n]
            if ($p) { return $p.Value }
        }
    }
    return $null
}

$results = New-Object System.Collections.Generic.List[object]

# ---------------------------------------------------------------------------
# Entra device object id, keyed by the deviceId GUID Intune reports
# ---------------------------------------------------------------------------
$entraByDeviceId = @{}
function Initialize-EntraDeviceIndex {
    if ($entraByDeviceId.Count -gt 0) { return }
    Write-Step 'Indexing Entra device objects...'
    foreach ($e in (Invoke-GraphPaged -Uri 'https://graph.microsoft.com/v1.0/devices?$select=id,deviceId,displayName')) {
        $k = [string](Get-Prop $e @('deviceId'))
        if ($k) { $entraByDeviceId[$k] = [string](Get-Prop $e @('id')) }
    }
    Write-OK "Entra device objects indexed: $($entraByDeviceId.Count)"
}

$userGroupCache = @{}
function Get-UserGroup {
    param([string]$UserId)
    if ($userGroupCache.ContainsKey($UserId)) { return $userGroupCache[$UserId] }
    try {
        $g = Invoke-GraphPaged -Uri "https://graph.microsoft.com/v1.0/users/$UserId/memberOf?`$select=id,displayName"
        $userGroupCache[$UserId] = $g
    }
    catch {
        # A deleted user, or one we cannot read, must not abort the whole report.
        Write-Warn "Could not read groups for user $UserId : $($_.Exception.Message)"
        $userGroupCache[$UserId] = @()
    }
    return $userGroupCache[$UserId]
}

if ($PSCmdlet.ParameterSetName -eq 'ByDevice') {

    Write-Step 'Loading managed devices...'
    $all = Invoke-GraphPaged -Uri 'https://graph.microsoft.com/beta/deviceManagement/managedDevices'
    Write-OK "Managed devices: $($all.Count)"

    $matched = @($all | Where-Object {
        $n = [string](Get-Prop $_ @('deviceName'))
        foreach ($pat in $DeviceName) { if ($n -like $pat) { return $true } }
        return $false
    })

    Write-OK "Devices matching: $($matched.Count)"
    if ($matched.Count -eq 0) {
        Write-Warn 'No device matched. Patterns accept wildcards, e.g. LAP-BCN-*'
    }
    else { Initialize-EntraDeviceIndex }

    foreach ($d in $matched) {
        $name  = [string](Get-Prop $d @('deviceName'))
        $azId  = [string](Get-Prop $d @('azureADDeviceId'))
        $upn   = [string](Get-Prop $d @('userPrincipalName'))
        $uid   = [string](Get-Prop $d @('userId'))
        $found = 0

        # --- direct device group membership ---
        if ($azId -and $entraByDeviceId.ContainsKey($azId)) {
            $oid = $entraByDeviceId[$azId]
            try {
                foreach ($g in (Invoke-GraphPaged -Uri "https://graph.microsoft.com/v1.0/devices/$oid/memberOf?`$select=id,displayName")) {
                    $found++
                    $results.Add([pscustomobject]@{
                        DeviceName = $name; DeviceId = [string](Get-Prop $d @('id'))
                        PrimaryUser = $upn
                        GroupName = [string](Get-Prop $g @('displayName'))
                        GroupId   = [string](Get-Prop $g @('id'))
                        MembershipVia = 'Device'
                    })
                }
            }
            catch { Write-Warn "$name : group lookup failed - $($_.Exception.Message)" }
        }
        else {
            Write-Warn "$name : no matching Entra device object (azureADDeviceId '$azId'). Device groups cannot be resolved."
        }

        # --- indirect, through the primary user ---
        if ($IncludeUserGroups -and $uid) {
            foreach ($g in (Get-UserGroup $uid)) {
                $found++
                $results.Add([pscustomobject]@{
                    DeviceName = $name; DeviceId = [string](Get-Prop $d @('id'))
                    PrimaryUser = $upn
                    GroupName = [string](Get-Prop $g @('displayName'))
                    GroupId   = [string](Get-Prop $g @('id'))
                    MembershipVia = 'PrimaryUser'
                })
            }
        }

        if ($found -eq 0) {
            $results.Add([pscustomobject]@{
                DeviceName = $name; DeviceId = [string](Get-Prop $d @('id'))
                PrimaryUser = $upn; GroupName = '(no group membership found)'
                GroupId = ''; MembershipVia = ''
            })
        }
    }
}
else {
    # ---------------------------------------------------------------------------
    # Group -> devices
    # ---------------------------------------------------------------------------
    Initialize-EntraDeviceIndex

    # reverse index: Entra object id -> Intune device
    Write-Step 'Loading managed devices...'
    $all = Invoke-GraphPaged -Uri 'https://graph.microsoft.com/beta/deviceManagement/managedDevices'
    $intuneByEntraOid = @{}
    foreach ($d in $all) {
        $azId = [string](Get-Prop $d @('azureADDeviceId'))
        if ($azId -and $entraByDeviceId.ContainsKey($azId)) { $intuneByEntraOid[$entraByDeviceId[$azId]] = $d }
    }
    Write-OK "Managed devices: $($all.Count)"

    foreach ($gn in $GroupName) {
        $safe  = $gn.Replace("'", "''")
        $found = Invoke-GraphPaged -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$safe'&`$select=id,displayName"

        if ($found.Count -eq 0) { Write-Warn "No group named '$gn' - skipped."; continue }
        if ($found.Count -gt 1) { Write-Warn "'$gn' matches $($found.Count) groups; all are reported." }

        foreach ($g in $found) {
            $gid   = [string](Get-Prop $g @('id'))
            $gname = [string](Get-Prop $g @('displayName'))
            Write-Step "Group: $gname"

            $members = Invoke-GraphPaged -Uri "https://graph.microsoft.com/v1.0/groups/$gid/members?`$select=id,displayName,userPrincipalName,deviceId"
            $deviceMembers = 0

            foreach ($m in $members) {
                $mid = [string](Get-Prop $m @('id'))
                if ($intuneByEntraOid.ContainsKey($mid)) {
                    $d = $intuneByEntraOid[$mid]
                    $deviceMembers++
                    $results.Add([pscustomobject]@{
                        DeviceName = [string](Get-Prop $d @('deviceName'))
                        DeviceId   = [string](Get-Prop $d @('id'))
                        PrimaryUser = [string](Get-Prop $d @('userPrincipalName'))
                        GroupName = $gname; GroupId = $gid; MembershipVia = 'Device'
                    })
                }
            }

            Write-OK "Members: $($members.Count)  of which Intune-managed devices: $deviceMembers"
            if ($members.Count -gt 0 -and $deviceMembers -eq 0) {
                Write-Warn 'Group has members but none are Intune-managed devices - it is probably a user group.'
            }
        }
    }
}

if ($OutputPath) {
    $dir = Split-Path -Parent $OutputPath
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $results | Sort-Object DeviceName, MembershipVia, GroupName |
        Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
    Write-OK "CSV written: $OutputPath"
}

Write-Step "Rows: $($results.Count)"
$results | Sort-Object DeviceName, MembershipVia, GroupName
