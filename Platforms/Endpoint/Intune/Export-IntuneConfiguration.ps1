<#
.SYNOPSIS
    Exports Intune configuration to JSON as a point-in-time snapshot: configuration
    profiles, settings-catalog policies, compliance policies, app protection policies,
    platform scripts, remediation scripts, Win32 app detection/requirement scripts and
    application definitions - optionally with their assignments.

.DESCRIPTION
    A readable, diffable record of how the tenant was configured. Two uses:

      1. Change tracking. Run it on a schedule into a git repo. Every run is a commit,
         and `git diff` then answers "what changed in Intune last week?", which the
         portal cannot.
      2. Pre-change safety net. Run before a big change so there is something to
         compare against and to rebuild from.

    THIS IS AN EXPORT, NOT A BACKUP PRODUCT. Read this before relying on it:

      - There is no restore. The JSON is a record, not a package. Rebuilding means
        re-creating objects from it, by hand or by a script you write.
      - Secrets are never returned by Graph. Certificate payloads, VPN pre-shared
        keys, Wi-Fi passwords, and the contents of .intunewin packages are absent
        from the export because the API omits them. A profile restored from this
        JSON will be missing them.
      - Assignments reference group object IDs. Those IDs are meaningless in a
        different tenant, so this is not a tenant-migration tool.
      - Apps export their DEFINITION, not their installer.

    Script bodies (platform, remediation, Win32 detection/requirement) are returned by
    Graph base64-encoded. They are decoded and written as .ps1 next to the JSON, so
    the actual code is diffable rather than an opaque blob.

.PARAMETER OutputRoot
    Folder for the snapshot. A timestamped subfolder is created inside it unless
    -NoTimestampFolder is used. Defaults to .\Exports\IntuneConfig

.PARAMETER Include
    Which object types to export. Default All.

.PARAMETER IncludeAssignments
    Also export each object's assignments. Costs one extra call per object for the
    types that do not support $expand.

.PARAMETER NoTimestampFolder
    Write straight into -OutputRoot instead of a timestamped subfolder. Use this when
    the target is a git working tree, so each run overwrites and the diff is the history.

.EXAMPLE
    # Full snapshot with assignments
    .\Export-IntuneConfiguration.ps1 -IncludeAssignments

.EXAMPLE
    # Into a git repo, same paths every run, so git diff shows the change
    .\Export-IntuneConfiguration.ps1 -OutputRoot C:\IntuneSnapshot -NoTimestampFolder -IncludeAssignments

.EXAMPLE
    # Just the scripts
    .\Export-IntuneConfiguration.ps1 -Include PlatformScripts,RemediationScripts

.NOTES
    When to use  : You need to answer "what changed in Intune last week?", which the console cannot, or you want a reviewable record of tenant configuration before a change.
    Why it exists: Run on a schedule into a git working tree, the diff IS the history. Script bodies come back from Graph base64-encoded and are decoded to .ps1 beside the JSON so the actual code is diffable rather than an opaque blob. It is an export, not a backup: there is no restore, secrets are never returned by Graph, and assignments reference group IDs meaningless in another tenant. manifest.json records which types were queried and which failed, so zero objects reads as "none found" rather than "not checked".
    Requires : Microsoft.Graph.Authentication  (Install-Module Microsoft.Graph -Scope CurrentUser)
    Scopes   : DeviceManagementConfiguration.Read.All, DeviceManagementApps.Read.All,
               DeviceManagementManagedDevices.Read.All, DeviceManagementServiceConfig.Read.All
    Rights   : read-only. This script never writes to Intune.

    Uses the /beta endpoint for the collections that only exist there (settings
    catalog, remediation scripts). Beta contracts can change without notice.

    Replaces (merged): seven scripts that each exported one slice of the tenant
    configuration - device configuration profiles, configuration policies, app
    protection policies, platform scripts, Win32 app detection scripts and
    application assignments - plus one that wrapped a third-party backup module.

    The first six authenticated with ADAL (Microsoft.IdentityModel.Clients.
    ActiveDirectory), retired by Microsoft; the seventh wrapped the third-party
    IntuneBackupAndRestore module, which itself depends on the retired Intune
    PowerShell SDK. All of that is replaced by Connect-MgGraph.

    A type that exports zero objects writes an empty array and says so. That means
    "none found", not "not checked" - the manifest records every type queried.
#>

[CmdletBinding()]
param(
    [string]$OutputRoot = (Join-Path $PSScriptRoot 'Exports\IntuneConfig'),

    [ValidateSet('All', 'ConfigurationProfiles', 'SettingsCatalog', 'CompliancePolicies',
                 'AppProtection', 'PlatformScripts', 'RemediationScripts',
                 'Win32AppScripts', 'Applications', 'AutopilotProfiles')]
    [string[]]$Include = @('All'),

    [switch]$IncludeAssignments,

    [switch]$NoTimestampFolder
)

$ErrorActionPreference = 'Stop'

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot '..\..\_Shared\Modules\M365.Common.psm1') -Force -ErrorAction Stop

function Write-Step { param($m) Write-Host "`n[>] $m" -ForegroundColor Cyan }
function Write-OK   { param($m) Write-Host "[OK] $m"  -ForegroundColor Green }
function Write-Warn { param($m) Write-Host "[!!] $m"  -ForegroundColor Yellow }

Connect-MgGraph -Scopes @(
    'DeviceManagementConfiguration.Read.All'
    'DeviceManagementApps.Read.All'
    'DeviceManagementManagedDevices.Read.All'
    'DeviceManagementServiceConfig.Read.All'
) -NoWelcome

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

# Windows-safe file name from an object's display name
function ConvertTo-SafeName {
    param([string]$Name, [string]$Fallback)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $Fallback }
    $bad  = [System.IO.Path]::GetInvalidFileNameChars() + [char[]]@('/', '\')
    $safe = ($Name.ToCharArray() | ForEach-Object { if ($bad -contains $_) { '_' } else { $_ } }) -join ''
    $safe = $safe.Trim(' ', '.')
    if ($safe.Length -gt 90) { $safe = $safe.Substring(0, 90) }
    if ([string]::IsNullOrWhiteSpace($safe)) { return $Fallback }
    return $safe
}

# ---------------------------------------------------------------------------
# What to export.
#   ScriptProp = property holding a base64 script body, decoded to a .ps1 alongside.
# ---------------------------------------------------------------------------
$catalog = @(
    @{ Key='ConfigurationProfiles'; Folder='ConfigurationProfiles'
       Uri='https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations'
       NameProp=@('displayName','name'); Expandable=$true }

    @{ Key='SettingsCatalog';       Folder='SettingsCatalog'
       Uri='https://graph.microsoft.com/beta/deviceManagement/configurationPolicies'
       NameProp=@('name','displayName'); Expandable=$true }

    @{ Key='CompliancePolicies';    Folder='CompliancePolicies'
       Uri='https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies'
       NameProp=@('displayName','name'); Expandable=$true }

    @{ Key='AppProtection';         Folder='AppProtection\iOS'
       Uri='https://graph.microsoft.com/beta/deviceAppManagement/iosManagedAppProtections'
       NameProp=@('displayName','name'); Expandable=$true }

    @{ Key='AppProtection';         Folder='AppProtection\Android'
       Uri='https://graph.microsoft.com/beta/deviceAppManagement/androidManagedAppProtections'
       NameProp=@('displayName','name'); Expandable=$true }

    @{ Key='PlatformScripts';       Folder='PlatformScripts'
       Uri='https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts'
       NameProp=@('displayName','name'); Expandable=$false; ScriptProp='scriptContent'; NeedsDetail=$true }

    @{ Key='RemediationScripts';    Folder='RemediationScripts'
       Uri='https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts'
       NameProp=@('displayName','name'); Expandable=$false
       ScriptProps=@('detectionScriptContent','remediationScriptContent'); NeedsDetail=$true }

    @{ Key='Applications';          Folder='Applications'
       Uri='https://graph.microsoft.com/beta/deviceAppManagement/mobileApps'
       NameProp=@('displayName','name'); Expandable=$true }

    @{ Key='Win32AppScripts';       Folder='Win32AppScripts'
       Uri='https://graph.microsoft.com/beta/deviceAppManagement/mobileApps'
       NameProp=@('displayName','name'); Expandable=$false; Win32Only=$true }

    @{ Key='AutopilotProfiles';     Folder='AutopilotProfiles'
       Uri='https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeploymentProfiles'
       NameProp=@('displayName','name'); Expandable=$false }
)

$wanted = if ($Include -contains 'All') { $catalog } else { $catalog | Where-Object { $Include -contains $_.Key } }
if (-not $wanted) { throw "Nothing selected by -Include: $($Include -join ', ')" }

$root = if ($NoTimestampFolder) { $OutputRoot }
        else { Join-Path $OutputRoot (Get-Date -Format 'yyyyMMdd-HHmmss') }
if (-not (Test-Path $root)) { New-Item -ItemType Directory -Path $root -Force | Out-Null }
Write-OK "Snapshot root: $root"

function Write-ScriptBody {
    param([string]$Base64, [string]$Path)
    if ([string]::IsNullOrWhiteSpace($Base64)) { return $false }
    try {
        $bytes = [Convert]::FromBase64String($Base64)
        [System.IO.File]::WriteAllBytes($Path, $bytes)
        return $true
    }
    catch {
        Write-Warn "Could not decode script body for $(Split-Path $Path -Leaf): $($_.Exception.Message)"
        return $false
    }
}

$manifest = New-Object System.Collections.Generic.List[object]

foreach ($c in $wanted) {
    Write-Step "$($c.Key) -> $($c.Folder)"
    $dir = Join-Path $root $c.Folder
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

    $count = 0; $scriptFiles = 0; $status = 'OK'; $detail = ''

    try {
        $uri   = if ($c.Expandable -and $IncludeAssignments) { "$($c.Uri)?`$expand=assignments" } else { $c.Uri }
        $items = Invoke-GraphPaged -Uri $uri

        foreach ($item in $items) {
            $id = [string](Get-Prop $item @('id'))

            # Win32AppScripts: only Win32 LOB apps carry detection/requirement rules
            if ($c.Win32Only) {
                $odata = [string](Get-Prop $item @('@odata.type'))
                if ($odata -notlike '*win32LobApp*') { continue }
            }

            $obj = $item

            # Some collections only return the script body on a single-object GET
            if ($c.NeedsDetail -and $id) {
                try { $obj = Invoke-MgGraphRequest -Method GET -Uri "$($c.Uri)/$id" -ErrorAction Stop }
                catch { Write-Warn "Detail fetch failed for $id : $($_.Exception.Message)" }
            }
            if ($c.Win32Only -and $id) {
                try { $obj = Invoke-MgGraphRequest -Method GET -Uri "$($c.Uri)/$id" -ErrorAction Stop }
                catch { Write-Warn "Detail fetch failed for $id : $($_.Exception.Message)" }
            }

            # assignments for the types that cannot expand them
            if ($IncludeAssignments -and -not $c.Expandable -and $id) {
                try {
                    $asg = Invoke-GraphPaged -Uri "$($c.Uri)/$id/assignments"
                    $obj = @{ object = $obj; assignments = $asg }
                }
                catch { Write-Warn "Assignment fetch failed for $id : $($_.Exception.Message)" }
            }

            $baseName = ConvertTo-SafeName ([string](Get-Prop $item $c.NameProp)) $id
            $stem     = if ($id) { "$baseName`_$($id.Substring(0, [Math]::Min(8, $id.Length)))" } else { $baseName }

            $obj | ConvertTo-Json -Depth 25 |
                Set-Content -LiteralPath (Join-Path $dir "$stem.json") -Encoding UTF8
            $count++

            # --- decode script bodies ---
            $source = if ($obj -is [System.Collections.IDictionary] -and $obj.Contains('object')) { $obj['object'] } else { $obj }

            foreach ($sp in @($c.ScriptProp)) {
                if (-not $sp) { continue }
                if (Write-ScriptBody ([string](Get-Prop $source @($sp))) (Join-Path $dir "$stem.ps1")) { $scriptFiles++ }
            }
            foreach ($sp in @($c.ScriptProps)) {
                if (-not $sp) { continue }
                $suffix = if ($sp -like 'detection*') { 'detect' } else { 'remediate' }
                if (Write-ScriptBody ([string](Get-Prop $source @($sp))) (Join-Path $dir "$stem.$suffix.ps1")) { $scriptFiles++ }
            }
            if ($c.Win32Only) {
                foreach ($r in @(Get-Prop $source @('detectionRules'))) {
                    $b64 = [string](Get-Prop $r @('scriptContent'))
                    if (Write-ScriptBody $b64 (Join-Path $dir "$stem.detectionRule.ps1")) { $scriptFiles++ }
                }
                foreach ($r in @(Get-Prop $source @('requirementRules'))) {
                    $b64 = [string](Get-Prop $r @('scriptContent'))
                    if (Write-ScriptBody $b64 (Join-Path $dir "$stem.requirementRule.ps1")) { $scriptFiles++ }
                }
            }
        }

        if ($count -eq 0) { Write-Warn "$($c.Key): 0 objects (none found)." }
        else { Write-OK "$($c.Key): $count object(s)$(if($scriptFiles){", $scriptFiles script file(s)"})" }
    }
    catch {
        # One collection failing must not lose the whole snapshot, but the manifest
        # has to record that this type is incomplete.
        $status = 'FAILED'
        $detail = $_.Exception.Message
        Write-Warn "$($c.Key): FAILED - $detail"
    }

    $manifest.Add([pscustomobject]@{
        Type = $c.Key; Folder = $c.Folder; Uri = $c.Uri
        Objects = $count; ScriptFiles = $scriptFiles
        AssignmentsIncluded = [bool]$IncludeAssignments
        Status = $status; Detail = $detail
    })
}

# ---------------------------------------------------------------------------
# Manifest: what was queried, what came back, what failed
# ---------------------------------------------------------------------------
$ctx = Get-MgContext
[pscustomobject]@{
    ExportedUtc         = (Get-Date).ToUniversalTime().ToString('s')
    TenantId            = if ($ctx) { $ctx.TenantId } else { '' }
    Account             = if ($ctx) { $ctx.Account }  else { '' }
    IncludeAssignments  = [bool]$IncludeAssignments
    Requested           = $Include
    Types               = $manifest
    Caveats             = @(
        'Secrets are not returned by Graph: certificate payloads, VPN pre-shared keys, Wi-Fi passwords and .intunewin contents are absent.'
        'Assignments reference group object IDs, which are not portable across tenants.'
        'Applications export their definition, not their installer package.'
        'A type with 0 objects means none were found, not that it was skipped.'
    )
} | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $root 'manifest.json') -Encoding UTF8

Write-Step 'Summary'
$manifest | ForEach-Object { Write-Host ("    {0,-22} {1,5} obj  {2}" -f $_.Type, $_.Objects, $_.Status) }

$failed = @($manifest | Where-Object Status -eq 'FAILED')
if ($failed.Count -gt 0) {
    Write-Warn "$($failed.Count) type(s) failed. This snapshot is INCOMPLETE - see manifest.json."
}
Write-OK "Snapshot written: $root"

$manifest
