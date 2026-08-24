<#
.SYNOPSIS
    Reports every Intune object assigned to one or more Entra ID groups: compliance
    policies, configuration profiles, settings-catalog policies, ADMX/group-policy
    configurations, applications, app configuration policies, app protection policies,
    remediation (device health) scripts, platform scripts and Autopilot profiles.

.DESCRIPTION
    Answers "what actually lands on this group?" before you change or delete it.

    Distinguishes INCLUDE from EXCLUDE assignments. The earlier versions of this
    script treated every assignment as an include, so a group used purely as an
    exclusion looked like it was receiving policy. That inverts the meaning of the
    report, so the target type is now reported explicitly.

    Emits objects, so the result can be piped, sorted or exported. Use -OutputPath
    to also write a CSV.

.PARAMETER GroupName
    Display name of the Entra ID group. Matched exactly, not as a substring.
    Accepts several names. Mutually exclusive with -GroupId.

.PARAMETER GroupId
    Object ID (GUID) of the group. Use this when display names are ambiguous.

.PARAMETER OutputPath
    Optional CSV path. Nothing is written unless this is supplied.

.PARAMETER IncludeEmpty
    Also emit a row for object types where the group has no assignment, so the
    report shows what was checked rather than only what was found.

.EXAMPLE
    # What is assigned to one group
    .\Get-IntuneGroupAssignment.ps1 -GroupName 'GRP-Devices-Classrooms'

.EXAMPLE
    # Several groups, exported
    .\Get-IntuneGroupAssignment.ps1 -GroupName 'GRP-Staff','GRP-Kiosk' `
        -OutputPath .\Exports\assignments.csv

.EXAMPLE
    # By object ID, showing object types with no assignment too
    .\Get-IntuneGroupAssignment.ps1 -GroupId '00000000-0000-0000-0000-000000000000' -IncludeEmpty

.NOTES
    When to use  : Before changing or deleting a group: what actually lands on it.
    Why it exists: Distinguishes include from exclude assignments - earlier versions treated everything as an include, so a group used purely for exclusions looked like it was receiving policy, which inverts the meaning of the report. -IncludeEmpty also shows which object types were queried, so an absent row is not confused with an unchecked one.
    Requires : Microsoft.Graph.Authentication  (Install-Module Microsoft.Graph -Scope CurrentUser)
    Scopes   : Group.Read.All, DeviceManagementConfiguration.Read.All,
               DeviceManagementApps.Read.All, DeviceManagementManagedDevices.Read.All,
               DeviceManagementServiceConfig.Read.All
    Rights   : read-only. This script never writes to Intune.

    Graph version: most of these collections only expose $expand=assignments on /beta,
    so /beta is used deliberately for those. Beta contracts can change without notice;
    if a collection starts returning a different shape, this is the first place to look.

    Absence of a row is not proof that nothing is assigned. Assignments made to
    "All devices" / "All users" built-in targets are NOT group assignments and will
    not appear here by design - they still apply to the devices in your group.
    -IncludeEmpty shows which object types were actually queried.

    Replaces (merged): four scripts that reported what is assigned to an Entra ID
    group. Two were the same original at different revisions, kept side by side;
    the other two covered fewer object types than this one does.
#>

[CmdletBinding(DefaultParameterSetName = 'ByName')]
param(
    [Parameter(Mandatory, ParameterSetName = 'ByName')]
    [string[]]$GroupName,

    [Parameter(Mandatory, ParameterSetName = 'ById')]
    [string[]]$GroupId,

    [string]$OutputPath,

    [switch]$IncludeEmpty
)

$ErrorActionPreference = 'Stop'

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot '..\..\_Shared\Modules\M365.Common.psm1') -Force -ErrorAction Stop

function Write-Step { param($m) Write-Host "`n[>] $m" -ForegroundColor Cyan }
function Write-OK   { param($m) Write-Host "[OK] $m"  -ForegroundColor Green }
function Write-Warn { param($m) Write-Host "[!!] $m"  -ForegroundColor Yellow }

Connect-MgGraph -Scopes @(
    'Group.Read.All'
    'DeviceManagementConfiguration.Read.All'
    'DeviceManagementApps.Read.All'
    'DeviceManagementManagedDevices.Read.All'
    'DeviceManagementServiceConfig.Read.All'
) -NoWelcome

# ---------------------------------------------------------------------------
# Dictionary-safe property read. Invoke-MgGraphRequest returns hashtables, so
# $obj.Property silently misses on some shapes.
# ---------------------------------------------------------------------------
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

# ---------------------------------------------------------------------------
# Resolve the groups
# ---------------------------------------------------------------------------
$targets = New-Object System.Collections.Generic.List[object]

if ($PSCmdlet.ParameterSetName -eq 'ByName') {
    foreach ($n in $GroupName) {
        # escape single quotes for the OData filter
        $safe  = $n.Replace("'", "''")
        $found = Invoke-GraphPaged -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$safe'&`$select=id,displayName"

        if (-not $found -or $found.Count -eq 0) {
            Write-Warn "No group found with display name '$n' - skipped."
            continue
        }
        if ($found.Count -gt 1) {
            Write-Warn "'$n' matches $($found.Count) groups. All of them are reported; use -GroupId to disambiguate."
        }
        foreach ($g in $found) {
            $targets.Add([pscustomobject]@{
                Id   = [string](Get-Prop $g @('id'))
                Name = [string](Get-Prop $g @('displayName'))
            })
        }
    }
} else {
    foreach ($id in $GroupId) {
        $g = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$id`?`$select=id,displayName" -ErrorAction Stop
        $targets.Add([pscustomobject]@{
            Id   = [string](Get-Prop $g @('id'))
            Name = [string](Get-Prop $g @('displayName'))
        })
    }
}

if ($targets.Count -eq 0) { throw 'No group could be resolved. Nothing to report.' }
Write-OK "Groups resolved: $($targets.Count)"

# ---------------------------------------------------------------------------
# What to inspect.
#   Expandable  = $expand=assignments works, one call for the whole collection.
#   Otherwise   = assignments must be fetched per object (N+1, slower).
# ---------------------------------------------------------------------------
$collections = @(
    @{ Type='Compliance policy';        Uri='https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies';       Expandable=$true;  NameProp=@('displayName','name') }
    @{ Type='Device configuration';     Uri='https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations';           Expandable=$true;  NameProp=@('displayName','name') }
    @{ Type='Settings catalog policy';  Uri='https://graph.microsoft.com/beta/deviceManagement/configurationPolicies';          Expandable=$true;  NameProp=@('name','displayName') }
    @{ Type='ADMX / group policy';      Uri='https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations';      Expandable=$true;  NameProp=@('displayName','name') }
    @{ Type='Application';              Uri='https://graph.microsoft.com/beta/deviceAppManagement/mobileApps';                  Expandable=$true;  NameProp=@('displayName','name') }
    @{ Type='App configuration';        Uri='https://graph.microsoft.com/beta/deviceAppManagement/targetedManagedAppConfigurations'; Expandable=$true; NameProp=@('displayName','name') }
    @{ Type='App config (managed dev)'; Uri='https://graph.microsoft.com/beta/deviceAppManagement/mobileAppConfigurations';     Expandable=$true;  NameProp=@('displayName','name') }
    @{ Type='App protection (iOS)';     Uri='https://graph.microsoft.com/beta/deviceAppManagement/iosManagedAppProtections';    Expandable=$true;  NameProp=@('displayName','name') }
    @{ Type='App protection (Android)'; Uri='https://graph.microsoft.com/beta/deviceAppManagement/androidManagedAppProtections';Expandable=$true;  NameProp=@('displayName','name') }
    @{ Type='App protection (Windows)'; Uri='https://graph.microsoft.com/beta/deviceAppManagement/windowsManagedAppProtections';Expandable=$true;  NameProp=@('displayName','name') }
    @{ Type='WIP policy';               Uri='https://graph.microsoft.com/beta/deviceAppManagement/mdmWindowsInformationProtectionPolicies'; Expandable=$true; NameProp=@('displayName','name') }
    @{ Type='Remediation script';       Uri='https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts';            Expandable=$false; NameProp=@('displayName','name') }
    @{ Type='Platform script';          Uri='https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts';        Expandable=$false; NameProp=@('displayName','name') }
    @{ Type='Autopilot profile';        Uri='https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeploymentProfiles'; Expandable=$false; NameProp=@('displayName','name') }
)

# Map the assignment target @odata.type to a plain word
function Get-TargetKind {
    param($Target)
    $t = [string](Get-Prop $Target @('@odata.type'))
    switch -Wildcard ($t) {
        '*exclusionGroupAssignmentTarget*' { 'Exclude' }
        '*groupAssignmentTarget*'          { 'Include' }
        default                            { 'Include' }
    }
}

$results = New-Object System.Collections.Generic.List[object]

foreach ($grp in $targets) {
    Write-Step "Group: $($grp.Name)  [$($grp.Id)]"

    foreach ($c in $collections) {
        $matched = 0
        try {
            if ($c.Expandable) {
                $uri   = "$($c.Uri)?`$expand=assignments"
                $items = Invoke-GraphPaged -Uri $uri

                foreach ($item in $items) {
                    foreach ($a in @(Get-Prop $item @('assignments'))) {
                        if ($null -eq $a) { continue }
                        $tgt = Get-Prop $a @('target')
                        $gid = [string](Get-Prop $tgt @('groupId'))
                        if ($gid -ne $grp.Id) { continue }   # exact match, not -match

                        $matched++
                        $results.Add([pscustomobject]@{
                            GroupName  = $grp.Name
                            GroupId    = $grp.Id
                            ObjectType = $c.Type
                            ObjectName = [string](Get-Prop $item $c.NameProp)
                            ObjectId   = [string](Get-Prop $item @('id'))
                            Assignment = Get-TargetKind $tgt
                            Intent     = [string](Get-Prop $a @('intent'))
                        })
                    }
                }
            }
            else {
                # assignments not expandable here: one call per object
                $items = Invoke-GraphPaged -Uri $c.Uri
                foreach ($item in $items) {
                    $id = [string](Get-Prop $item @('id'))
                    if (-not $id) { continue }
                    $asg = Invoke-GraphPaged -Uri "$($c.Uri)/$id/assignments"
                    foreach ($a in $asg) {
                        $tgt = Get-Prop $a @('target')
                        $gid = [string](Get-Prop $tgt @('groupId'))
                        if ($gid -ne $grp.Id) { continue }

                        $matched++
                        $results.Add([pscustomobject]@{
                            GroupName  = $grp.Name
                            GroupId    = $grp.Id
                            ObjectType = $c.Type
                            ObjectName = [string](Get-Prop $item $c.NameProp)
                            ObjectId   = $id
                            Assignment = Get-TargetKind $tgt
                            Intent     = [string](Get-Prop $a @('intent'))
                        })
                    }
                }
            }

            if ($matched -gt 0) {
                Write-OK "$($c.Type): $matched"
            }
            elseif ($IncludeEmpty) {
                $results.Add([pscustomobject]@{
                    GroupName  = $grp.Name
                    GroupId    = $grp.Id
                    ObjectType = $c.Type
                    ObjectName = '(none assigned)'
                    ObjectId   = ''
                    Assignment = ''
                    Intent     = ''
                })
            }
        }
        catch {
            # One collection failing (missing licence, tenant not onboarded for that
            # workload, beta contract change) must not lose the rest of the report.
            Write-Warn "$($c.Type): query failed - $($_.Exception.Message)"
            $results.Add([pscustomobject]@{
                GroupName  = $grp.Name
                GroupId    = $grp.Id
                ObjectType = $c.Type
                ObjectName = '(QUERY FAILED - result incomplete)'
                ObjectId   = ''
                Assignment = 'ERROR'
                Intent     = $_.Exception.Message
            })
        }
    }
}

if ($OutputPath) {
    $dir = Split-Path -Parent $OutputPath
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $results | Sort-Object GroupName, ObjectType, ObjectName |
        Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
    Write-OK "CSV written: $OutputPath"
}

Write-Step "Total assignment rows: $($results.Count)"
$results | Sort-Object GroupName, ObjectType, ObjectName
