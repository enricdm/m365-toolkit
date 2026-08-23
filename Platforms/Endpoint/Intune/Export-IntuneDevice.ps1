<#
.SYNOPSIS
    Exports managed devices from Intune with filters for platform, ownership and
    Android Enterprise enrollment type (BYOD / work profile / fully managed /
    dedicated), optionally as an incremental delta against a previous run.

.DESCRIPTION
    One export to replace eight. The scripts this merges were the same query with a
    different hard-coded filter each time - one for BYOD, one for COPE, one for
    dedicated Android, one for "registered devices", and three variations on a delta
    export. All of them are now parameters.

    Two modes:

      Full  (default) - every device matching the filters.
      Delta (-DeltaStatePath) - only devices added or changed since the last run.
                                The first run with -DeltaStatePath does a full pass
                                and writes the delta token; later runs are incremental.

    HOW THE DELTA ACTUALLY WORKS, AND ITS LIMIT
    Intune's managedDevices collection has no delta endpoint. The delta lives on the
    Entra ID devices collection (/v1.0/devices/delta), which returns Entra device
    objects, not Intune ones. This script therefore uses the Entra delta to learn
    WHICH devices changed, then fetches those from Intune.

    The consequence is worth stating plainly: a change that happens only on the Intune
    side and never touches the Entra device object - a compliance state flip, a new
    last-sync timestamp - may not surface in the delta. Delta mode is for tracking
    devices appearing and disappearing, not for tracking every attribute. If you need
    current compliance for everything, run a full export.

.PARAMETER Platform
    Windows, iOS, Android, macOS, or All (default).

.PARAMETER Ownership
    company, personal, or All (default).

.PARAMETER EnrollmentType
    Android Enterprise enrollment type, or the friendly names for them:
      BYOD             -> androidEnterpriseWorkProfile (personal device, work profile)
      COPE             -> androidEnterpriseCorporateWorkProfile
      FullyManaged     -> androidEnterpriseFullyManaged
      Dedicated        -> androidEnterpriseDedicatedDevice (kiosk)
    Only meaningful for Android. Ignored for other platforms.

.PARAMETER ComplianceState
    compliant, noncompliant, unknown, or All (default).

.PARAMETER DeltaStatePath
    JSON file holding the delta token between runs. Supplying it enables delta mode.

.PARAMETER OutputPath
    CSV path. Defaults to .\Exports\IntuneDevices-<timestamp>.csv

.EXAMPLE
    # Everything
    .\Export-IntuneDevice.ps1

.EXAMPLE
    # Android BYOD only
    .\Export-IntuneDevice.ps1 -Platform Android -EnrollmentType BYOD

.EXAMPLE
    # Corporate Windows, non-compliant only
    .\Export-IntuneDevice.ps1 -Platform Windows -Ownership company -ComplianceState noncompliant

.EXAMPLE
    # Incremental: first run seeds the token, later runs return only changes
    .\Export-IntuneDevice.ps1 -DeltaStatePath .\State\devices-delta.json

.NOTES
    Requires : Microsoft.Graph.Authentication  (Install-Module Microsoft.Graph -Scope CurrentUser)
    Scopes   : DeviceManagementManagedDevices.Read.All, Device.Read.All (delta mode only)
    Rights   : read-only. This script never writes to Intune.

    Replaces (merged): Export-DevicesIntune.ps1, Export-DevicesIntune-Delta.ps1,
    Export-DevicesIntune-Delta-incremental.ps1, Get-RegisteredIntuneDevices.ps1
    (two divergent copies), "Export BYOD.ps1", "Export COPE.ps1",
    Export-AndroidDedicated.ps1

    The originals authenticated with Connect-MSGraph from the retired Intune
    PowerShell SDK and refreshed the bearer token by hand on HTTP 401. Connect-MgGraph
    handles token lifetime, so that code is gone rather than ported.

    An empty result is an empty result, not an error. If a filter combination returns
    nothing, the CSV is written with headers and zero rows, and the console says so.
#>

[CmdletBinding()]
param(
    [ValidateSet('All', 'Windows', 'iOS', 'Android', 'macOS')]
    [string]$Platform = 'All',

    [ValidateSet('All', 'company', 'personal')]
    [string]$Ownership = 'All',

    [ValidateSet('All', 'BYOD', 'COPE', 'FullyManaged', 'Dedicated')]
    [string]$EnrollmentType = 'All',

    [ValidateSet('All', 'compliant', 'noncompliant', 'unknown')]
    [string]$ComplianceState = 'All',

    [string]$DeltaStatePath,

    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot '..\..\_Shared\Modules\M365.Common.psm1') -Force -ErrorAction Stop

function Write-Step { param($m) Write-Host "`n[>] $m" -ForegroundColor Cyan }
function Write-OK   { param($m) Write-Host "[OK] $m"  -ForegroundColor Green }
function Write-Warn { param($m) Write-Host "[!!] $m"  -ForegroundColor Yellow }

if (-not $OutputPath) {
    $ts = Get-Date -Format 'yyyyMMdd-HHmmss'
    $OutputPath = Join-Path $PSScriptRoot "Exports\IntuneDevices-$ts.csv"
}

$scopes = @('DeviceManagementManagedDevices.Read.All')
if ($DeltaStatePath) { $scopes += 'Device.Read.All' }
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

# ---------------------------------------------------------------------------
# Server-side filter. Only properties Graph actually supports in $filter go here;
# everything else is filtered client-side after the fetch.
# ---------------------------------------------------------------------------
$filters = @()
if ($Platform  -ne 'All') { $filters += "operatingSystem eq '$Platform'" }
if ($Ownership -ne 'All') { $filters += "managedDeviceOwnerType eq '$Ownership'" }
if ($ComplianceState -ne 'All') { $filters += "complianceState eq '$ComplianceState'" }

$query = if ($filters.Count) { '?$filter=' + ($filters -join ' and ') } else { '' }

# Friendly enrollment-type names -> the values Graph reports in deviceEnrollmentType
$enrollMap = @{
    'BYOD'         = @('androidEnterpriseWorkProfile', 'androidWorkProfile')
    'COPE'         = @('androidEnterpriseCorporateWorkProfile')
    'FullyManaged' = @('androidEnterpriseFullyManaged')
    'Dedicated'    = @('androidEnterpriseDedicatedDevice')
}

# ---------------------------------------------------------------------------
# Delta mode: work out which device IDs changed
# ---------------------------------------------------------------------------
$changedEntraIds = $null
$isFirstDeltaRun = $false

if ($DeltaStatePath) {
    Write-Step 'Delta mode'
    $deltaUri = 'https://graph.microsoft.com/v1.0/devices/delta?$select=id,deviceId,displayName'

    if (Test-Path -LiteralPath $DeltaStatePath) {
        $state = Get-Content -LiteralPath $DeltaStatePath -Raw | ConvertFrom-Json
        if ($state.deltaLink) {
            $deltaUri = $state.deltaLink
            Write-OK "Resuming from stored delta token (written $($state.timestamp))."
        } else {
            Write-Warn 'Delta state file has no deltaLink; treating this as a first run.'
            $isFirstDeltaRun = $true
        }
    } else {
        Write-Warn 'No delta state file yet. This run is a FULL pass; it also seeds the token for next time.'
        $isFirstDeltaRun = $true
    }

    # Walk the delta chain by hand: unlike nextLink paging, the last page carries
    # a deltaLink that has to be stored for the next run.
    $changed  = New-Object System.Collections.Generic.List[string]
    $next     = $deltaUri
    $newDelta = $null

    while ($next) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop
        foreach ($v in @(Get-Prop $resp @('value'))) {
            if ($null -eq $v) { continue }
            $devId = [string](Get-Prop $v @('deviceId'))
            if ($devId) { $changed.Add($devId) }
        }
        $newDelta = Get-Prop $resp @('@odata.deltaLink', 'odata.deltaLink')
        $next     = Get-Prop $resp @('@odata.nextLink', 'odata.nextLink')
    }

    if ($newDelta) {
        $dir = Split-Path -Parent $DeltaStatePath
        if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        [pscustomobject]@{ deltaLink = [string]$newDelta; timestamp = (Get-Date).ToString('s') } |
            ConvertTo-Json | Set-Content -LiteralPath $DeltaStatePath -Encoding UTF8
        Write-OK "Delta token stored: $DeltaStatePath"
    } else {
        Write-Warn 'Graph returned no deltaLink; the next run will not be incremental.'
    }

    if (-not $isFirstDeltaRun) {
        $changedEntraIds = [System.Collections.Generic.HashSet[string]]::new(
            [string[]]$changed, [System.StringComparer]::OrdinalIgnoreCase)
        Write-OK "Devices changed since last run: $($changedEntraIds.Count)"
        if ($changedEntraIds.Count -eq 0) {
            Write-OK 'Nothing changed. Writing an empty export.'
        }
    }
}

# ---------------------------------------------------------------------------
# Fetch from Intune
# ---------------------------------------------------------------------------
Write-Step "Loading managed devices (platform: $Platform, ownership: $Ownership)..."
$devices = Invoke-GraphPaged -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices$query"
Write-OK "Devices returned by Graph: $($devices.Count)"

$rows    = New-Object System.Collections.Generic.List[object]
$skipped = 0

foreach ($d in $devices) {
    # --- client-side filters ---
    if ($EnrollmentType -ne 'All') {
        $et = [string](Get-Prop $d @('deviceEnrollmentType'))
        if ($et -notin $enrollMap[$EnrollmentType]) { $skipped++; continue }
    }

    if ($null -ne $changedEntraIds) {
        $azId = [string](Get-Prop $d @('azureADDeviceId'))
        if (-not $azId -or -not $changedEntraIds.Contains($azId)) { $skipped++; continue }
    }

    $rows.Add([pscustomobject]@{
        DeviceName          = [string](Get-Prop $d @('deviceName'))
        DeviceId            = [string](Get-Prop $d @('id'))
        EntraDeviceId       = [string](Get-Prop $d @('azureADDeviceId'))
        SerialNumber        = [string](Get-Prop $d @('serialNumber'))
        Manufacturer        = [string](Get-Prop $d @('manufacturer'))
        Model               = [string](Get-Prop $d @('model'))
        OperatingSystem     = [string](Get-Prop $d @('operatingSystem'))
        OSVersion           = [string](Get-Prop $d @('osVersion'))
        Ownership           = [string](Get-Prop $d @('managedDeviceOwnerType'))
        EnrollmentType      = [string](Get-Prop $d @('deviceEnrollmentType'))
        EnrollmentProfile   = [string](Get-Prop $d @('enrollmentProfileName'))
        JoinType            = [string](Get-Prop $d @('joinType'))
        ManagementAgent     = [string](Get-Prop $d @('managementAgent'))
        ComplianceState     = [string](Get-Prop $d @('complianceState'))
        IsEncrypted         = [string](Get-Prop $d @('isEncrypted'))
        IsSupervised        = [string](Get-Prop $d @('isSupervised'))
        PrimaryUser         = [string](Get-Prop $d @('userPrincipalName'))
        DeviceCategory      = [string](Get-Prop $d @('deviceCategoryDisplayName'))
        EnrolledUtc         = [string](Get-Prop $d @('enrolledDateTime'))
        LastSyncUtc         = [string](Get-Prop $d @('lastSyncDateTime'))
        TotalStorageGB      = [math]::Round(([double](Get-Prop $d @('totalStorageSpaceInBytes')) / 1GB), 1)
        FreeStorageGB       = [math]::Round(([double](Get-Prop $d @('freeStorageSpaceInBytes'))  / 1GB), 1)
    })
}

if ($skipped -gt 0) { Write-OK "Filtered out client-side: $skipped" }

$dir = Split-Path -Parent $OutputPath
if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
$rows | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8

Write-Step "Exported rows: $($rows.Count)"
if ($rows.Count -eq 0) {
    Write-Warn 'Zero rows. That is a result, not a failure - check the filters if it is unexpected.'
}
Write-OK "CSV written: $OutputPath"

$rows
