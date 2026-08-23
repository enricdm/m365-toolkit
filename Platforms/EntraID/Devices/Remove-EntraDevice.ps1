<#
.SYNOPSIS
    Removes device objects from Entra ID, from an explicit list or from a saved
    report. Reports by default; deletes nothing unless -Execute.

.DESCRIPTION
    Deleting the Entra device object is the last step of decommissioning, after the
    Intune record has gone. It is also the step people get wrong, because an Entra
    device object and an Intune managed device are two different things:

      Intune managed device - the MDM enrollment. Removing it unenrolls the device.
      Entra device object   - the directory identity. Removing it breaks Conditional
                              Access evaluation, device-based licensing, BitLocker
                              key escrow and Windows Hello for Business for that device.

    ORDER MATTERS: retire or delete in Intune FIRST, then remove the Entra object.
    Removing the Entra object while the device is still enrolled and in use leaves an
    orphan that will usually re-register on next sign-in - so the cleanup looks like
    it failed, when in fact it was done in the wrong order.

    BITLOCKER RECOVERY KEYS ARE ESCROWED AGAINST THE ENTRA DEVICE OBJECT. Deleting it
    can make those keys unrecoverable. If the hardware is being reused rather than
    destroyed, export the keys first. This script warns but cannot check for you.

.PARAMETER DeviceId
    One or more Entra device IDs or object IDs to remove.

.PARAMETER InputCsv
    CSV with a DeviceId column (object ID or deviceId). Extra columns are ignored.

.PARAMETER StaleDays
    Instead of an explicit list, select devices whose approximateLastSignInDateTime is
    older than this. Minimum 90: an Entra device object going quiet for a few weeks is
    normal, and this deletion is not cheap to undo.

.PARAMETER ExcludeOwnership
    Trust types to leave alone. Defaults to 'Workplace' (personal/registered devices).

.PARAMETER MaxDevices
    Safety cap for a single -Execute run. Default 25.

.PARAMETER Execute
    Performs the deletion. Without this switch the script only reports.

.PARAMETER OutputPath
    CSV path for the report / result log.

.EXAMPLE
    # What would be removed, by staleness. Deletes nothing.
    .\Remove-EntraDevice.ps1 -StaleDays 180

.EXAMPLE
    # Remove a reviewed list
    .\Remove-EntraDevice.ps1 -InputCsv .\decommissioned.csv -Execute

.NOTES
    Requires : Microsoft.Graph.Authentication  (Install-Module Microsoft.Graph -Scope CurrentUser)
    Scopes   : Device.ReadWrite.All  (Directory.AccessAsUser.All in some tenants)
    Rights   : DESTRUCTIVE when -Execute is used. Deleting a device object cannot be undone.

    Replaces (merged): Remove-EntraIDDevice.ps1, Retire_DevicesAzure.ps1
    Both used the AzureAD / MSOnline modules (Remove-AzureADDevice, Get-MsolDevice),
    retired by Microsoft on 30 March 2025 and no longer functional. Rewritten on
    Microsoft Graph.

    A device that has not signed in recently is not necessarily gone. Confirm against
    the asset register before deleting anything you cannot re-enrol.
#>

[CmdletBinding(DefaultParameterSetName = 'ByStale')]
param(
    [Parameter(Mandatory, ParameterSetName = 'ById')]
    [string[]]$DeviceId,

    [Parameter(Mandatory, ParameterSetName = 'ByCsv')]
    [string]$InputCsv,

    [Parameter(ParameterSetName = 'ByStale')]
    [ValidateRange(90, 3650)]
    [int]$StaleDays = 365,

    [Parameter(ParameterSetName = 'ByStale')]
    [string[]]$ExcludeOwnership = @('Workplace'),

    [ValidateRange(1, 100000)]
    [int]$MaxDevices = 25,

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
    $OutputPath = Join-Path $PSScriptRoot "Exports\EntraDeviceRemoval-$ts.csv"
}

Connect-MgGraph -Scopes 'Device.ReadWrite.All' -NoWelcome

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

function New-Row {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        'PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Pure function: builds and returns a report row, changes no system state.')]
    [CmdletBinding()]
    param($Device, [string]$Status, [string]$Detail = '')
    $raw  = Get-Prop $Device @('approximateLastSignInDateTime')
    $when = $null
    if ($raw) { [void][datetime]::TryParse([string]$raw, [ref]$when) }
    [pscustomobject]@{
        DisplayName     = [string](Get-Prop $Device @('displayName'))
        ObjectId        = [string](Get-Prop $Device @('id'))
        DeviceId        = [string](Get-Prop $Device @('deviceId'))
        OperatingSystem = [string](Get-Prop $Device @('operatingSystem'))
        OSVersion       = [string](Get-Prop $Device @('operatingSystemVersion'))
        TrustType       = [string](Get-Prop $Device @('trustType'))
        IsManaged       = [string](Get-Prop $Device @('isManaged'))
        IsCompliant     = [string](Get-Prop $Device @('isCompliant'))
        LastSignInUtc   = if ($when) { $when.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
        DaysSinceSignIn = if ($when) { [math]::Round(((Get-Date) - $when).TotalDays, 1) } else { $null }
        Status          = $Status
        Detail          = $Detail
    }
}

# ---------------------------------------------------------------------------
# Build the target list
# ---------------------------------------------------------------------------
$targets = New-Object System.Collections.Generic.List[object]
$log     = New-Object System.Collections.Generic.List[object]

switch ($PSCmdlet.ParameterSetName) {

    { $_ -in 'ById', 'ByCsv' } {
        $ids = if ($PSCmdlet.ParameterSetName -eq 'ById') { $DeviceId } else {
            if (-not (Test-Path -LiteralPath $InputCsv)) { throw "Input CSV not found: $InputCsv" }
            $rows = @(Import-Csv -LiteralPath $InputCsv)
            if ($rows.Count -eq 0) { throw "Input CSV is empty: $InputCsv" }
            if ('DeviceId' -notin $rows[0].PSObject.Properties.Name) {
                throw "Input CSV must have a 'DeviceId' column."
            }
            @($rows | Select-Object -ExpandProperty DeviceId | Where-Object { $_ })
        }

        Write-Step "Resolving $($ids.Count) device(s)..."
        foreach ($id in $ids) {
            $found = $null
            # accept either the object id or the deviceId GUID
            try { $found = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/devices/$id" -ErrorAction Stop }
            catch {
                $safe = ([string]$id).Replace("'", "''")
                $byDeviceId = Invoke-GraphPaged -Uri "https://graph.microsoft.com/v1.0/devices?`$filter=deviceId eq '$safe'"
                if ($byDeviceId.Count -eq 1) { $found = $byDeviceId[0] }
                elseif ($byDeviceId.Count -gt 1) {
                    Write-Warn "'$id' matches $($byDeviceId.Count) device objects - skipped, resolve by object id."
                    $log.Add((New-Row @{ id = $id } 'SKIPPED' 'ambiguous: multiple device objects share this deviceId'))
                    continue
                }
            }
            if ($null -eq $found) {
                Write-Warn "'$id' not found - skipped."
                $log.Add((New-Row @{ id = $id } 'NOT-FOUND' 'no device object with this id'))
                continue
            }
            $targets.Add($found)
        }
    }

    'ByStale' {
        Write-Step "Loading Entra device objects (stale > $StaleDays days)..."
        $all = Invoke-GraphPaged -Uri 'https://graph.microsoft.com/v1.0/devices'
        Write-OK "Device objects: $($all.Count)"

        $cutoff = (Get-Date).AddDays(-$StaleDays)
        $noDate = 0
        foreach ($d in $all) {
            $raw  = Get-Prop $d @('approximateLastSignInDateTime')
            $when = $null
            if ($raw) { [void][datetime]::TryParse([string]$raw, [ref]$when) }

            if (-not $when) {
                # No timestamp is not an old timestamp. Report, never delete.
                $noDate++
                $log.Add((New-Row $d 'NO-SIGNIN-DATA' 'no approximateLastSignInDateTime; never acted on'))
                continue
            }
            if ($when -ge $cutoff) { continue }

            $trust = [string](Get-Prop $d @('trustType'))
            if ($ExcludeOwnership -contains $trust) { continue }

            $targets.Add($d)
        }
        if ($noDate -gt 0) { Write-Warn "$noDate device object(s) have no sign-in timestamp. Reported, not acted on." }
    }
}

Write-Step "Devices selected: $($targets.Count)"
if ($targets.Count -eq 0) {
    Write-OK 'Nothing to do.'
}
else {
    Write-Warn 'BitLocker recovery keys are escrowed against the Entra device object. Deleting it can make them unrecoverable.'
    Write-Warn 'Ensure the device is already retired/deleted in Intune BEFORE removing the Entra object.'
}

# ---------------------------------------------------------------------------
# Delete (or not)
# ---------------------------------------------------------------------------
if ($targets.Count -gt 0 -and -not $Execute) {
    Write-Warn "DRY RUN - nothing deleted. Re-run with -Execute to apply."
    foreach ($t in $targets) { $log.Add((New-Row $t 'WOULD-DELETE')) }
}
elseif ($targets.Count -gt 0) {
    $toAct = @($targets | Select-Object -First $MaxDevices)
    if ($targets.Count -gt $MaxDevices) {
        Write-Warn "Safety cap: deleting $MaxDevices of $($targets.Count). Raise -MaxDevices to go further."
        foreach ($t in ($targets | Select-Object -Skip $MaxDevices)) {
            $log.Add((New-Row $t 'SKIPPED' 'MaxDevices cap'))
        }
    }

    Write-Step "Deleting $($toAct.Count) device object(s)... this cannot be undone."
    $ok = 0; $fail = 0
    foreach ($t in $toAct) {
        $oid = [string](Get-Prop $t @('id'))
        try {
            Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/devices/$oid" -ErrorAction Stop
            $ok++
            Write-OK "Deleted: $([string](Get-Prop $t @('displayName')))"
            $log.Add((New-Row $t 'DELETED'))
        }
        catch {
            $fail++
            Write-Err "$([string](Get-Prop $t @('displayName'))): $($_.Exception.Message)"
            $log.Add((New-Row $t 'FAILED' $_.Exception.Message))
        }
    }
    Write-Step "Deleted: $ok   Failed: $fail"
}

$dir = Split-Path -Parent $OutputPath
if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
$log | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
Write-OK "Log written: $OutputPath"

$log
