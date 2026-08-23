<#
.SYNOPSIS
    Finds Intune devices that have not checked in for a given number of days and,
    with -Execute, retires or deletes them. Reports only by default.

.DESCRIPTION
    Stale device records inflate compliance reporting, hold licences and hide the
    devices that genuinely need attention. This finds them and can act on them.

    Three actions, deliberately different in blast radius:

      Report  (default) - list only, change nothing.
      Retire            - removes company data and unenrolls. The device record
                          leaves Intune; the hardware keeps working, personal data
                          on it is untouched.
      Delete            - removes the device record from Intune only. Does NOT
                          touch the device itself. If the device ever checks in
                          again it may simply re-enrol.

    Wipe is deliberately NOT offered here. Factory-resetting a fleet from a
    staleness query is not a cleanup operation, and a stale record is very often a
    device that is merely switched off - a laptop in a drawer, someone on extended
    leave, a machine awaiting reassignment. Wipe destroys user data and cannot be
    undone. If you need to wipe a specific device, do it deliberately, one at a
    time, from the portal.

.PARAMETER StaleDays
    Minimum days since lastSyncDateTime for a device to be considered stale.
    Default 90. Values below 30 are rejected: normal absence (holiday, sick leave,
    a spare machine) routinely exceeds a few weeks.

.PARAMETER Action
    Report (default), Retire, or Delete.

.PARAMETER Platform
    Restrict to one operating system. Default All.

.PARAMETER ExcludeOwnership
    Ownership types to leave alone. Defaults to 'personal' so BYOD devices are not
    retired by a fleet-wide cleanup.

.PARAMETER MaxDevices
    Safety cap on how many devices a single -Execute run may act on. Default 50.
    The run stops once the cap is reached, and says so. Raise it deliberately.

.PARAMETER Execute
    Applies the chosen Action. Without this switch the script only reports.

.PARAMETER OutputPath
    CSV path for the report / action log. Defaults to
    .\Exports\StaleDevices-<timestamp>.csv

.EXAMPLE
    # See what is stale. Changes nothing.
    .\Invoke-IntuneStaleDeviceCleanup.ps1 -StaleDays 120

.EXAMPLE
    # Delete stale corporate Windows records, capped at 25, after reviewing the report
    .\Invoke-IntuneStaleDeviceCleanup.ps1 -StaleDays 180 -Platform Windows `
        -Action Delete -MaxDevices 25 -Execute

.NOTES
    Requires : Microsoft.Graph.Authentication  (Install-Module Microsoft.Graph -Scope CurrentUser)
    Scopes   : DeviceManagementManagedDevices.ReadWrite.All
    Rights   : DESTRUCTIVE when -Execute is used. Retire and Delete cannot be undone.

    Run the report first and keep the CSV. It is the only record of what the device
    estate looked like before the cleanup.

    A stale lastSyncDateTime is evidence the device has not contacted Intune. It is
    NOT evidence the device is gone, lost or decommissioned. Confirm against your
    CMDB / asset register before acting on anything you cannot re-enrol.

    Replaces (merged): Invoke-IntuneCleanup.ps1, Invoke-IntuneCleanup-exportCSV.ps1
    (two divergent copies), Retire_DevicesIntune.ps1, Wipe_DevicesIntune.ps1
#>

[CmdletBinding()]
param(
    [ValidateRange(30, 3650)]
    [int]$StaleDays = 90,

    [ValidateSet('Report', 'Retire', 'Delete')]
    [string]$Action = 'Report',

    [ValidateSet('All', 'Windows', 'iOS', 'Android', 'macOS')]
    [string]$Platform = 'All',

    [string[]]$ExcludeOwnership = @('personal'),

    [ValidateRange(1, 100000)]
    [int]$MaxDevices = 50,

    [switch]$Execute,

    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot '..\..\_Shared\Modules\M365.Common.psm1') -Force -ErrorAction Stop

function Write-Step { param($m) Write-Host "`n[>] $m" -ForegroundColor Cyan }
function Write-OK   { param($m) Write-Host "[OK] $m"  -ForegroundColor Green }
function Write-Warn { param($m) Write-Host "[!!] $m"  -ForegroundColor Yellow }
function Write-Err  { param($m) Write-Host "[XX] $m"  -ForegroundColor Red }

if (-not $OutputPath) {
    $ts = Get-Date -Format 'yyyyMMdd-HHmmss'
    $OutputPath = Join-Path $PSScriptRoot "Exports\StaleDevices-$ts.csv"
}

Connect-MgGraph -Scopes 'DeviceManagementManagedDevices.ReadWrite.All' -NoWelcome

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
# Find the stale devices
# ---------------------------------------------------------------------------
$osFilter = switch ($Platform) {
    'Windows' { "?`$filter=operatingSystem eq 'Windows'" }
    'iOS'     { "?`$filter=operatingSystem eq 'iOS'" }
    'Android' { "?`$filter=operatingSystem eq 'Android'" }
    'macOS'   { "?`$filter=operatingSystem eq 'macOS'" }
    default   { '' }
}

Write-Step "Loading managed devices (platform: $Platform)..."
$devices = Invoke-GraphPaged -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices$osFilter"
Write-OK "Devices returned: $($devices.Count)"

$cutoff = (Get-Date).AddDays(-$StaleDays)
$stale  = New-Object System.Collections.Generic.List[object]
$noSync = 0

foreach ($d in $devices) {
    $raw  = Get-Prop $d @('lastSyncDateTime')
    $sync = $null
    if ($raw) { [void][datetime]::TryParse([string]$raw, [ref]$sync) }

    if (-not $sync) {
        # Never synced, or the property is absent. Reported, never acted on:
        # "no timestamp" is not the same as "old timestamp".
        $noSync++
        continue
    }
    if ($sync -ge $cutoff) { continue }

    $ownership = [string](Get-Prop $d @('managedDeviceOwnerType'))
    if ($ExcludeOwnership -contains $ownership) { continue }

    $stale.Add([pscustomobject]@{
        DeviceName        = [string](Get-Prop $d @('deviceName'))
        DeviceId          = [string](Get-Prop $d @('id'))
        SerialNumber      = [string](Get-Prop $d @('serialNumber'))
        OperatingSystem   = [string](Get-Prop $d @('operatingSystem'))
        OSVersion         = [string](Get-Prop $d @('osVersion'))
        Ownership         = $ownership
        PrimaryUser       = [string](Get-Prop $d @('userPrincipalName'))
        EnrolledUtc       = [string](Get-Prop $d @('enrolledDateTime'))
        LastSyncUtc       = $sync.ToString('yyyy-MM-dd HH:mm:ss')
        DaysSinceLastSync = [math]::Round(((Get-Date) - $sync).TotalDays, 1)
        ComplianceState   = [string](Get-Prop $d @('complianceState'))
    })
}

$stale = $stale | Sort-Object -Property DaysSinceLastSync -Descending

Write-Step "Stale (> $StaleDays days): $($stale.Count)"
if ($noSync -gt 0) {
    Write-Warn "$noSync device(s) have no usable lastSyncDateTime. Reported below as NO-SYNC-DATA and never acted on."
}
if ($ExcludeOwnership.Count -gt 0) {
    Write-OK "Ownership excluded from action: $($ExcludeOwnership -join ', ')"
}

$log = New-Object System.Collections.Generic.List[object]

if ($stale.Count -eq 0) {
    Write-OK 'No stale devices matched. Nothing to do.'
}
elseif ($Action -eq 'Report' -or -not $Execute) {
    if ($Action -ne 'Report' -and -not $Execute) {
        Write-Warn "DRY RUN - action '$Action' NOT applied. Re-run with -Execute to apply."
    }
    foreach ($s in $stale) {
        $row = $s.PSObject.Copy()
        $row | Add-Member -NotePropertyName Result -NotePropertyValue $(
            if ($Action -eq 'Report') { 'REPORTED' } else { "WOULD $($Action.ToUpper())" }) -Force
        $row | Add-Member -NotePropertyName Detail -NotePropertyValue '' -Force
        $log.Add($row)
    }
}
else {
    # ---- destructive path ----
    $toAct = @($stale | Select-Object -First $MaxDevices)
    if ($stale.Count -gt $MaxDevices) {
        Write-Warn "Safety cap: acting on $MaxDevices of $($stale.Count) stale devices. Raise -MaxDevices to go further."
    }

    Write-Warn "About to $($Action.ToUpper()) $($toAct.Count) device(s). This cannot be undone."
    $ok = 0; $fail = 0

    foreach ($s in $toAct) {
        try {
            if ($Action -eq 'Retire') {
                Invoke-MgGraphRequest -Method POST -ErrorAction Stop `
                    -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$($s.DeviceId)/retire"
            } else {
                Invoke-MgGraphRequest -Method DELETE -ErrorAction Stop `
                    -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$($s.DeviceId)"
            }
            $ok++
            Write-OK "$($Action): $($s.DeviceName)  (idle $($s.DaysSinceLastSync)d)"
            $row = $s.PSObject.Copy()
            $row | Add-Member -NotePropertyName Result -NotePropertyValue $Action.ToUpper() -Force
            $row | Add-Member -NotePropertyName Detail -NotePropertyValue '' -Force
            $log.Add($row)
        }
        catch {
            $fail++
            Write-Err "$($s.DeviceName): $($_.Exception.Message)"
            $row = $s.PSObject.Copy()
            $row | Add-Member -NotePropertyName Result -NotePropertyValue 'FAILED' -Force
            $row | Add-Member -NotePropertyName Detail -NotePropertyValue $_.Exception.Message -Force
            $log.Add($row)
        }
    }

    # devices past the cap are still logged, so the CSV shows the full picture
    foreach ($s in ($stale | Select-Object -Skip $MaxDevices)) {
        $row = $s.PSObject.Copy()
        $row | Add-Member -NotePropertyName Result -NotePropertyValue 'SKIPPED (MaxDevices cap)' -Force
        $row | Add-Member -NotePropertyName Detail -NotePropertyValue '' -Force
        $log.Add($row)
    }

    Write-Step "$($Action) succeeded: $ok   failed: $fail"
}

$dir = Split-Path -Parent $OutputPath
if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
$log | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
Write-OK "Log written: $OutputPath"

$log
