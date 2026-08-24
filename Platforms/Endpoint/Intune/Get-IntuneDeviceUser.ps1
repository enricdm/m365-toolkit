<#
.SYNOPSIS
    Reports, for every managed device in Intune, the last user who signed in, the
    assigned primary user, and whether the two agree. Optionally enriches Windows
    devices with their Autopilot deployment profile.

.DESCRIPTION
    The drift between "who the device says used it last" and "who Intune thinks owns
    it" is what makes licence reclamation, device retirement and support routing go
    wrong. This reports both and flags the mismatch.

    Shared devices (classrooms, meeting rooms, kiosks) mismatch by design: many people
    sign in, the primary user is either unset or a resource account. Use
    -SharedProfileName so those rows are labelled Shared rather than Mismatch, instead
    of drowning the real mismatches in expected noise.

.PARAMETER Platform
    Filter by operating system: Windows, iOS, Android, macOS, or All (default).

.PARAMETER IncludeAutopilot
    Also look up the Autopilot deployment profile for Windows devices. Costs one extra
    paged query. Non-Windows rows leave the Autopilot columns blank.

.PARAMETER SharedProfileName
    Autopilot deployment profile names that identify shared/kiosk devices. Devices on
    these profiles are classified 'Shared' instead of 'Mismatch'. Requires
    -IncludeAutopilot; without it there is no profile to compare against.

.PARAMETER StaleDays
    Devices whose lastSyncDateTime is older than this many days are flagged Stale.
    Default 30. Use 0 to disable the check.

.PARAMETER OutputPath
    Optional CSV path. Nothing is written unless supplied.

.EXAMPLE
    # Everything, to the pipeline
    .\Get-IntuneDeviceUser.ps1

.EXAMPLE
    # Windows only, with Autopilot, treating two profiles as shared devices
    .\Get-IntuneDeviceUser.ps1 -Platform Windows -IncludeAutopilot `
        -SharedProfileName 'AP-Classrooms','AP-MeetingRooms' `
        -OutputPath .\Exports\device-users.csv

.EXAMPLE
    # Only the rows that need attention
    .\Get-IntuneDeviceUser.ps1 -Platform Windows | Where-Object Status -eq 'Mismatch'

.NOTES
    Requires : Microsoft.Graph.Authentication  (Install-Module Microsoft.Graph -Scope CurrentUser)
    Scopes   : DeviceManagementManagedDevices.Read.All, User.Read.All,
               DeviceManagementServiceConfig.Read.All (only with -IncludeAutopilot)
    Rights   : read-only. This script never writes to Intune.

    'usersLoggedOn' is only exposed on the Graph BETA endpoint, so the device query
    targets /beta deliberately. User lookups stay on /v1.0.

    A blank LastUser is NOT proof that nobody used the device. usersLoggedOn is
    populated by the Intune agent and is routinely empty on freshly enrolled devices,
    on devices that have not checked in since the property was introduced, and on
    most non-Windows platforms. Those rows are reported as 'NoLoginData', which means
    "not known", not "unused".

    Replaces (merged): eight scripts that all answered some version of "who last
    signed in to this device". Several were copies of one another at different
    revisions, one compared the last logged-on user against the assigned primary
    user, and one produced a shortened form of the same report. The comparison and
    the output shape are options here rather than separate files.
#>

[CmdletBinding()]
param(
    [ValidateSet('All', 'Windows', 'iOS', 'Android', 'macOS')]
    [string]$Platform = 'All',

    [switch]$IncludeAutopilot,

    [string[]]$SharedProfileName = @(),

    [int]$StaleDays = 30,

    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot '..\..\_Shared\Modules\M365.Common.psm1') -Force -ErrorAction Stop

function Write-Step { param($m) Write-Host "`n[>] $m" -ForegroundColor Cyan }
function Write-OK   { param($m) Write-Host "[OK] $m"  -ForegroundColor Green }
function Write-Warn { param($m) Write-Host "[!!] $m"  -ForegroundColor Yellow }

if ($SharedProfileName.Count -gt 0 -and -not $IncludeAutopilot) {
    Write-Warn '-SharedProfileName was supplied without -IncludeAutopilot; there is no profile to compare against, so no row will be classified Shared.'
}

$scopes = @('DeviceManagementManagedDevices.Read.All', 'User.Read.All')
if ($IncludeAutopilot) { $scopes += 'DeviceManagementServiceConfig.Read.All' }
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
# User lookup cache. A 2,000-device fleet resolves to far fewer distinct users;
# without the cache this is thousands of redundant Graph calls.
# Note: a failed lookup is cached as $null on purpose, so a deleted user object
# is not retried once per device that referenced it.
# ---------------------------------------------------------------------------
$userCache = @{}
function Resolve-User {
    param([string]$UserId)
    if ([string]::IsNullOrWhiteSpace($UserId)) { return $null }
    if ($userCache.ContainsKey($UserId)) { return $userCache[$UserId] }
    try {
        $u = Invoke-MgGraphRequest -Method GET -ErrorAction Stop `
             -Uri "https://graph.microsoft.com/v1.0/users/$UserId`?`$select=id,userPrincipalName,displayName"
        $userCache[$UserId] = $u
    }
    catch {
        # Most often the user object was deleted after the device recorded the sign-in.
        $userCache[$UserId] = $null
    }
    return $userCache[$UserId]
}

# ---------------------------------------------------------------------------
# Autopilot: serial number -> deployment profile
# ---------------------------------------------------------------------------
$apBySerial = @{}
if ($IncludeAutopilot) {
    Write-Step 'Loading Autopilot identities...'
    try {
        $ap = Invoke-GraphPaged -Uri 'https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities?$expand=deploymentProfile'
        foreach ($a in $ap) {
            $serial = [string](Get-Prop $a @('serialNumber'))
            if (-not $serial) { continue }
            $apBySerial[$serial] = [string](Get-Prop (Get-Prop $a @('deploymentProfile')) @('displayName'))
        }
        Write-OK "Autopilot identities: $($apBySerial.Count)"
    }
    catch {
        Write-Warn "Autopilot lookup failed - profile columns will be blank. $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------------
# Devices
# ---------------------------------------------------------------------------
$osFilter = switch ($Platform) {
    'Windows' { "?`$filter=operatingSystem eq 'Windows'" }
    'iOS'     { "?`$filter=operatingSystem eq 'iOS'" }
    'Android' { "?`$filter=operatingSystem eq 'Android'" }
    'macOS'   { "?`$filter=operatingSystem eq 'macOS'" }
    default   { '' }
}

Write-Step "Loading managed devices (platform: $Platform)..."
$devices = Invoke-GraphPaged -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices$osFilter"
Write-OK "Devices: $($devices.Count)"

$now     = Get-Date
$results = New-Object System.Collections.Generic.List[object]
$i = 0

foreach ($d in $devices) {
    $i++
    if ($i % 250 -eq 0) { Write-Host "    ...$i/$($devices.Count)" -ForegroundColor DarkGray }

    $serial = [string](Get-Prop $d @('serialNumber'))

    # --- primary user ---
    $primaryId  = [string](Get-Prop $d @('userId'))
    $primaryUpn = [string](Get-Prop $d @('userPrincipalName'))
    if (-not $primaryUpn -and $primaryId) {
        $primaryUpn = [string](Get-Prop (Resolve-User $primaryId) @('userPrincipalName'))
    }

    # --- last logged-on user: most recent entry in usersLoggedOn ---
    $lastId   = ''
    $lastWhen = $null
    foreach ($e in @(Get-Prop $d @('usersLoggedOn'))) {
        if ($null -eq $e) { continue }
        $whenRaw = Get-Prop $e @('lastLogOnDateTime')
        $when    = $null
        if ($whenRaw) { [void][datetime]::TryParse([string]$whenRaw, [ref]$when) }
        if ($null -eq $lastWhen -or ($when -and $when -gt $lastWhen)) {
            $lastWhen = $when
            $lastId   = [string](Get-Prop $e @('userId'))
        }
    }
    $lastUpn = ''
    if ($lastId) { $lastUpn = [string](Get-Prop (Resolve-User $lastId) @('userPrincipalName')) }

    # --- last sync / staleness ---
    $syncRaw = Get-Prop $d @('lastSyncDateTime')
    $sync    = $null
    if ($syncRaw) { [void][datetime]::TryParse([string]$syncRaw, [ref]$sync) }
    $daysSinceSync = if ($sync) { [math]::Round(($now - $sync).TotalDays, 1) } else { $null }

    # --- Autopilot profile ---
    $apProfile = ''
    if ($IncludeAutopilot -and $serial -and $apBySerial.ContainsKey($serial)) {
        $apProfile = $apBySerial[$serial]
    }

    # --- classification ---
    # Order matters: a shared device is shared even if it also mismatches, and
    # "we have no data" must never be reported as "they match".
    $status =
        if ($apProfile -and $SharedProfileName -contains $apProfile) { 'Shared' }
        elseif (-not $lastUpn -and -not $primaryUpn)                 { 'NoUserData' }
        elseif (-not $lastUpn)                                       { 'NoLoginData' }
        elseif (-not $primaryUpn)                                    { 'NoPrimaryUser' }
        elseif ($lastUpn -ieq $primaryUpn)                           { 'Match' }
        else                                                         { 'Mismatch' }

    $results.Add([pscustomobject]@{
        DeviceName        = [string](Get-Prop $d @('deviceName'))
        DeviceId          = [string](Get-Prop $d @('id'))
        SerialNumber      = $serial
        OperatingSystem   = [string](Get-Prop $d @('operatingSystem'))
        OSVersion         = [string](Get-Prop $d @('osVersion'))
        Ownership         = [string](Get-Prop $d @('managedDeviceOwnerType'))
        PrimaryUser       = $primaryUpn
        LastLoggedOnUser  = $lastUpn
        LastLogonUtc      = if ($lastWhen) { $lastWhen.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
        Status            = $status
        LastSyncUtc       = if ($sync) { $sync.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
        DaysSinceLastSync = $daysSinceSync
        IsStale           = if ($StaleDays -gt 0 -and $null -ne $daysSinceSync) { $daysSinceSync -gt $StaleDays } else { $null }
        AutopilotProfile  = $apProfile
    })
}

if ($OutputPath) {
    $dir = Split-Path -Parent $OutputPath
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $results | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
    Write-OK "CSV written: $OutputPath"
}

Write-Step 'Summary'
$results | Group-Object Status | Sort-Object Count -Descending |
    ForEach-Object { Write-Host ("    {0,-14} {1,6}" -f $_.Name, $_.Count) }

$results
