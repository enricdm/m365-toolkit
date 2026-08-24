<#
.SYNOPSIS
    Aligns the Intune primary user of a device with the user who actually signs in,
    from a CSV or from the live drift report. Reports by default; changes nothing
    unless -Execute is supplied.

.DESCRIPTION
    Primary user drives licence attribution, Company Portal experience and support
    routing. When someone changes desk or a device is reassigned, the primary user
    is what goes stale.

    Two ways to drive it:
      -InputCsv   an explicit list you have already reviewed. Safest.
      -FromDrift  query the tenant live and propose the last logged-on user as the
                  new primary wherever the two disagree.

    Without -Execute the script only reports what it would change. This is the
    default on purpose: setting the primary user in bulk from a drift query can
    reassign hundreds of devices at once, and on shared devices the "last user"
    is whoever happened to sign in last, which is not an owner.

.PARAMETER InputCsv
    CSV with columns DeviceId and NewPrimaryUserUpn. Extra columns are ignored.

.PARAMETER FromDrift
    Build the change set live: any device where the last logged-on user differs
    from the primary user becomes a proposed change.

.PARAMETER ExcludeAutopilotProfile
    Autopilot deployment profile names to skip entirely (classrooms, kiosks, meeting
    rooms). Strongly recommended with -FromDrift. Only applies with -FromDrift.

.PARAMETER Execute
    Applies the changes. Without this switch the script only reports.

.PARAMETER OutputPath
    CSV path for the change set / result log. Defaults to
    .\Exports\PrimaryUser-Changes-<timestamp>.csv

.EXAMPLE
    # Dry run from a reviewed CSV - reports only
    .\Set-IntuneDevicePrimaryUser.ps1 -InputCsv .\changes.csv

.EXAMPLE
    # Apply that reviewed CSV
    .\Set-IntuneDevicePrimaryUser.ps1 -InputCsv .\changes.csv -Execute

.EXAMPLE
    # Live drift, excluding shared devices, dry run first
    .\Set-IntuneDevicePrimaryUser.ps1 -FromDrift `
        -ExcludeAutopilotProfile 'AP-Classrooms','AP-MeetingRooms'

.NOTES
    Requires : Microsoft.Graph.Authentication  (Install-Module Microsoft.Graph -Scope CurrentUser)
    Scopes   : DeviceManagementManagedDevices.ReadWrite.All, User.Read.All,
               DeviceManagementServiceConfig.Read.All (with -ExcludeAutopilotProfile)
    Rights   : WRITES to Intune when -Execute is used.

    There is no bulk undo. Export the current state before applying:
        .\Get-IntuneDeviceUser.ps1 -OutputPath .\before.csv

    A device with no last logged-on user is skipped, never cleared. Absence of
    sign-in data means "not known", not "no user".

    Replaces: one earlier script that updated the primary user on Intune devices.
#>

[CmdletBinding(DefaultParameterSetName = 'FromDrift')]
param(
    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [string]$InputCsv,

    [Parameter(Mandatory, ParameterSetName = 'FromDrift')]
    [switch]$FromDrift,

    [Parameter(ParameterSetName = 'FromDrift')]
    [string[]]$ExcludeAutopilotProfile = @(),

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
    $OutputPath = Join-Path $PSScriptRoot "Exports\PrimaryUser-Changes-$ts.csv"
}

$scopes = @('DeviceManagementManagedDevices.ReadWrite.All', 'User.Read.All')
if ($ExcludeAutopilotProfile.Count -gt 0) { $scopes += 'DeviceManagementServiceConfig.Read.All' }
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

$userCache = @{}
function Resolve-UserId {
    <# UPN -> object id. Returns $null when the user cannot be resolved. #>
    param([string]$Upn)
    if ([string]::IsNullOrWhiteSpace($Upn)) { return $null }
    if ($userCache.ContainsKey($Upn)) { return $userCache[$Upn] }
    try {
        $u = Invoke-MgGraphRequest -Method GET -ErrorAction Stop `
             -Uri "https://graph.microsoft.com/v1.0/users/$([uri]::EscapeDataString($Upn))`?`$select=id,userPrincipalName"
        $userCache[$Upn] = [string](Get-Prop $u @('id'))
    }
    catch { $userCache[$Upn] = $null }
    return $userCache[$Upn]
}

# ---------------------------------------------------------------------------
# Build the change set
# ---------------------------------------------------------------------------
$changes = New-Object System.Collections.Generic.List[object]

if ($PSCmdlet.ParameterSetName -eq 'FromCsv') {
    Write-Step "Reading change set: $InputCsv"
    if (-not (Test-Path -LiteralPath $InputCsv)) { throw "Input CSV not found: $InputCsv" }
    $rows = Import-Csv -LiteralPath $InputCsv

    $missing = @('DeviceId', 'NewPrimaryUserUpn') |
               Where-Object { $_ -notin $rows[0].PSObject.Properties.Name }
    if ($missing) { throw "Input CSV is missing required column(s): $($missing -join ', ')" }

    foreach ($r in $rows) {
        if (-not $r.DeviceId -or -not $r.NewPrimaryUserUpn) { continue }
        $changes.Add([pscustomobject]@{
            DeviceId    = $r.DeviceId
            DeviceName  = if ($r.PSObject.Properties['DeviceName']) { $r.DeviceName } else { '' }
            CurrentUser = if ($r.PSObject.Properties['PrimaryUser']) { $r.PrimaryUser } else { '' }
            NewUser     = $r.NewPrimaryUserUpn
            Reason      = 'from CSV'
        })
    }
}
else {
    Write-Step 'Building change set from live drift...'

    $apBySerial = @{}
    if ($ExcludeAutopilotProfile.Count -gt 0) {
        $ap = Invoke-GraphPaged -Uri 'https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities?$expand=deploymentProfile'
        foreach ($a in $ap) {
            $s = [string](Get-Prop $a @('serialNumber'))
            if ($s) { $apBySerial[$s] = [string](Get-Prop (Get-Prop $a @('deploymentProfile')) @('displayName')) }
        }
        Write-OK "Autopilot identities loaded: $($apBySerial.Count)"
    }

    $devices = Invoke-GraphPaged -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=operatingSystem eq 'Windows'"
    Write-OK "Windows devices: $($devices.Count)"

    $skippedShared = 0
    foreach ($d in $devices) {
        $serial = [string](Get-Prop $d @('serialNumber'))
        if ($serial -and $apBySerial.ContainsKey($serial) -and
            $ExcludeAutopilotProfile -contains $apBySerial[$serial]) {
            $skippedShared++
            continue
        }

        $primaryUpn = [string](Get-Prop $d @('userPrincipalName'))

        $lastId = ''; $lastWhen = $null
        foreach ($e in @(Get-Prop $d @('usersLoggedOn'))) {
            if ($null -eq $e) { continue }
            $w = $null
            $raw = Get-Prop $e @('lastLogOnDateTime')
            if ($raw) { [void][datetime]::TryParse([string]$raw, [ref]$w) }
            if ($null -eq $lastWhen -or ($w -and $w -gt $lastWhen)) { $lastWhen = $w; $lastId = [string](Get-Prop $e @('userId')) }
        }
        if (-not $lastId) { continue }   # no sign-in data: skip, never clear

        $lastUpn = ''
        try {
            $u = Invoke-MgGraphRequest -Method GET -ErrorAction Stop `
                 -Uri "https://graph.microsoft.com/v1.0/users/$lastId`?`$select=userPrincipalName"
            $lastUpn = [string](Get-Prop $u @('userPrincipalName'))
        } catch { continue }

        if (-not $lastUpn) { continue }
        if ($primaryUpn -and ($lastUpn -ieq $primaryUpn)) { continue }

        $changes.Add([pscustomobject]@{
            DeviceId    = [string](Get-Prop $d @('id'))
            DeviceName  = [string](Get-Prop $d @('deviceName'))
            CurrentUser = $primaryUpn
            NewUser     = $lastUpn
            Reason      = if ($primaryUpn) { 'drift: last logon differs from primary' } else { 'no primary user set' }
        })
    }
    if ($skippedShared -gt 0) { Write-OK "Skipped by Autopilot profile exclusion: $skippedShared" }
}

Write-Step "Devices proposed for change: $($changes.Count)"
if ($changes.Count -eq 0) { Write-OK 'Nothing to do.'; return }

# ---------------------------------------------------------------------------
# Apply (or not)
# ---------------------------------------------------------------------------
$log = New-Object System.Collections.Generic.List[object]

if (-not $Execute) {
    Write-Warn 'DRY RUN - no changes made. Re-run with -Execute to apply.'
    foreach ($c in $changes) {
        $log.Add([pscustomobject]@{
            DeviceName = $c.DeviceName; DeviceId = $c.DeviceId
            CurrentUser = $c.CurrentUser; NewUser = $c.NewUser
            Reason = $c.Reason; Result = 'WOULD CHANGE'; Detail = ''
        })
    }
}
else {
    Write-Step "Applying $($changes.Count) change(s)..."
    $ok = 0; $fail = 0

    foreach ($c in $changes) {
        $uid = Resolve-UserId $c.NewUser
        if (-not $uid) {
            $fail++
            Write-Err "$($c.DeviceName): cannot resolve user '$($c.NewUser)' - skipped"
            $log.Add([pscustomobject]@{
                DeviceName = $c.DeviceName; DeviceId = $c.DeviceId
                CurrentUser = $c.CurrentUser; NewUser = $c.NewUser
                Reason = $c.Reason; Result = 'SKIPPED'; Detail = 'user could not be resolved'
            })
            continue
        }

        try {
            $body = @{ '@odata.id' = "https://graph.microsoft.com/beta/users/$uid" } | ConvertTo-Json -Compress
            Invoke-MgGraphRequest -Method POST -ErrorAction Stop `
                -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($c.DeviceId)/users/`$ref" `
                -Body $body -ContentType 'application/json'
            $ok++
            Write-OK "$($c.DeviceName): $($c.CurrentUser) -> $($c.NewUser)"
            $log.Add([pscustomobject]@{
                DeviceName = $c.DeviceName; DeviceId = $c.DeviceId
                CurrentUser = $c.CurrentUser; NewUser = $c.NewUser
                Reason = $c.Reason; Result = 'CHANGED'; Detail = ''
            })
        }
        catch {
            $fail++
            Write-Err "$($c.DeviceName): $($_.Exception.Message)"
            $log.Add([pscustomobject]@{
                DeviceName = $c.DeviceName; DeviceId = $c.DeviceId
                CurrentUser = $c.CurrentUser; NewUser = $c.NewUser
                Reason = $c.Reason; Result = 'FAILED'; Detail = $_.Exception.Message
            })
        }
    }
    Write-Step "Changed: $ok   Failed/skipped: $fail"
}

$dir = Split-Path -Parent $OutputPath
if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
$log | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
Write-OK "Log written: $OutputPath"

$log
