<#
.SYNOPSIS
    Reports all Windows devices in Intune, translates OSVersion build numbers into the
    friendly "feature update" names (matching PingCastle-style reports), flags devices
    running an out-of-support Windows version, and enriches each device with primary-user
    country/usageLocation and last-logged-on-user info.

.NOTES
    When to use  : A Windows version reaches end of servicing and you have to say how many machines are still behind it, and in which country.
    Why it exists: Translates the build number into the feature-update name and joins it to the end-of-servicing date, which Intune does not do. The feature-update build is the THIRD segment of osVersion, not the first, and usersLoggedOn only exists on the beta managedDevice type - both are why earlier versions reported zero.
    - Requires module: Microsoft.Graph.Authentication  (Install-Module Microsoft.Graph -Scope CurrentUser)
    - Scopes needed:   DeviceManagementManagedDevices.Read.All, User.Read.All
    - EOL dates below assume Enterprise/Education editions (standard for corp-managed devices).
      If a meaningful chunk of your fleet is Home/Pro, those versions go EOL earlier — adjust
      the $eolMap dates if needed, or add an edition lookup if you track it elsewhere.
    - "Last logged on user" comes from the managedDevice 'usersLoggedOn' collection, which is
      ONLY available on the Graph BETA endpoint. The device query below targets /beta for that
      reason; per-user lookups stay on /v1.0. If your tenant/devices don't populate it, that
      column will just be blank for those rows.
#>

param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'Exports\OutOfSupport_Devices_Report.csv')
)

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All", "User.Read.All" -NoWelcome

# ---------------------------------------------------------------------------
# Build number -> friendly version name (matches the PingCastle report labels)
# ---------------------------------------------------------------------------
$osVersionMap = @{
    '10240' = 'Windows 10 1507'
    '10586' = 'Windows 10 1511'
    '14393' = 'Windows 10 1607'
    '15063' = 'Windows 10 1703'
    '16299' = 'Windows 10 1709'
    '17134' = 'Windows 10 1803'
    '17763' = 'Windows 10 1809'
    '18362' = 'Windows 10 1903'
    '18363' = 'Windows 10 1909'
    '19041' = 'Windows 10 2004'
    '19042' = 'Windows 10 20H2'
    '19043' = 'Windows 10 21H1'
    '19044' = 'Windows 10 21H2'
    '19045' = 'Windows 10 22H2'
    '22000' = 'Windows 11 21H2'
    '22621' = 'Windows 11 22H2'
    '22631' = 'Windows 11 23H2'
    '26100' = 'Windows 11 24H2'
    '26200' = 'Windows 11 25H2'
}

# ---------------------------------------------------------------------------
# End-of-servicing date per friendly version (Enterprise/Education timelines)
# ---------------------------------------------------------------------------
$eolMap = @{
    'Windows 10 1507' = '2017-05-09'
    'Windows 10 1511' = '2017-10-10'
    'Windows 10 1607' = '2019-04-09'
    'Windows 10 1703' = '2019-10-08'
    'Windows 10 1709' = '2020-10-13'
    'Windows 10 1803' = '2021-05-11'
    'Windows 10 1809' = '2021-05-11'
    'Windows 10 1903' = '2020-12-08'
    'Windows 10 1909' = '2022-05-10'
    'Windows 10 2004' = '2021-12-14'
    'Windows 10 20H2' = '2022-05-10'
    'Windows 10 21H1' = '2022-12-13'
    'Windows 10 21H2' = '2024-06-11'
    'Windows 10 22H2' = '2025-10-14'   # Windows 10 EOL overall
    'Windows 11 21H2' = '2024-10-08'
    'Windows 11 22H2' = '2025-10-14'
    'Windows 11 23H2' = '2026-11-10'
    'Windows 11 24H2' = '2027-10-12'
    'Windows 11 25H2' = '2028-10-10'
}

$today = Get-Date

# ---------------------------------------------------------------------------
# Pull all Windows managed devices (paged)
# ---------------------------------------------------------------------------
$select = @(
    'id', 'deviceName', 'serialNumber', 'userId', 'userPrincipalName',
    'operatingSystem', 'osVersion', 'complianceState', 'deviceEnrollmentType',
    'managementAgent', 'enrolledDateTime', 'lastSyncDateTime', 'azureADDeviceId',
    'usersLoggedOn', 'model', 'manufacturer'
) -join ','

# NOTE: 'usersLoggedOn' only exists on managedDevice in the BETA endpoint - it is NOT a
#       property of the v1.0 managedDevice type. Calling v1.0 with it in $select returns
#       HTTP 400 "Could not find a property named 'usersLoggedOn'". This is an API surface
#       issue, NOT a broken Graph module. Beta is used only for the device list; the per-user
#       lookups below stay on v1.0 (usageLocation/country/department are all v1.0 properties).
$uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=operatingSystem eq 'Windows'&`$select=$select&`$top=999"

$devices = @()
do {
    $response = Invoke-MgGraphRequest -Method GET -Uri $uri
    $devices += $response.value
    $uri = $response.'@odata.nextLink'
} while ($uri)

Write-Host "Retrieved $($devices.Count) Windows devices from Intune." -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# Cache user lookups (usageLocation / country) to avoid repeat calls
# ---------------------------------------------------------------------------
$userCache = @{}

function Get-CachedUser {
    param([string]$UserId)
    if (-not $UserId) { return $null }
    if ($userCache.ContainsKey($UserId)) { return $userCache[$UserId] }

    try {
        $u = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$UserId`?`$select=userPrincipalName,usageLocation,country,department"
    } catch {
        $u = $null
    }
    $userCache[$UserId] = $u
    return $u
}

# ---------------------------------------------------------------------------
# Build the report
# ---------------------------------------------------------------------------
$report = foreach ($d in $devices) {

    # osVersion is "<major>.<minor>.<build>.<revision>" (e.g. 10.0.26100.8037).
    # The feature-update build is the THIRD segment; [0] is always the literal "10".
    $osParts = @($d.osVersion -split '\.')
    $buildNumber = if ($osParts.Count -ge 3) { $osParts[2] } else { '' }
    $friendlyVer = if ($buildNumber) { $osVersionMap[$buildNumber] } else { $null }
    if (-not $friendlyVer) {
        $friendlyVer = if ($buildNumber) { "Unknown (build $buildNumber)" }
        else { "Unknown (unparseable osVersion '$($d.osVersion)')" }
    }

    $eolDateString = $eolMap[$friendlyVer]
    $isOutOfSupport = $false
    if ($eolDateString) {
        $isOutOfSupport = $today -gt [datetime]$eolDateString
    }

    # Primary user
    $primaryUser = Get-CachedUser -UserId $d.userId

    # Last logged on user (usersLoggedOn is an array of {userId, lastLogOnDateTime})
    $lastLogon = $null
    if ($d.usersLoggedOn -and $d.usersLoggedOn.Count -gt 0) {
        $lastLogon = $d.usersLoggedOn | Sort-Object lastLogOnDateTime -Descending | Select-Object -First 1
    }
    $lastLogonUser = $null
    if ($lastLogon) {
        $lastLogonUser = Get-CachedUser -UserId $lastLogon.userId
    }

    [PSCustomObject]@{
        DeviceName           = $d.deviceName
        SerialNumber         = $d.serialNumber
        IntuneDeviceId       = $d.id
        AzureADDeviceId      = $d.azureADDeviceId
        Manufacturer         = $d.manufacturer
        Model                = $d.model
        OSVersionRaw         = $d.osVersion
        TranslatedVersion    = $friendlyVer
        EndOfServicingDate   = $eolDateString
        IsOutOfSupport       = $isOutOfSupport
        ComplianceState      = $d.complianceState
        EnrollmentType       = $d.deviceEnrollmentType
        ManagementAgent      = $d.managementAgent
        EnrolledDateTime     = $d.enrolledDateTime
        LastSyncDateTime     = $d.lastSyncDateTime
        UserPrincipalName    = $d.userPrincipalName
        UserUsageLocation    = $primaryUser.usageLocation
        UserCountry          = $primaryUser.country
        UserDepartment       = $primaryUser.department
        LastLoggedOnUPN      = $lastLogonUser.userPrincipalName
        LastLoggedOnDateTime = $lastLogon.lastLogOnDateTime
        LastLoggedOnUsageLoc = $lastLogonUser.usageLocation
        LastLoggedOnCountry  = $lastLogonUser.country
    }
}

$outDir = Split-Path -Parent $OutputPath
if ($outDir -and -not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
$report | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "Full report exported to $OutputPath" -ForegroundColor Green

# ---------------------------------------------------------------------------
# Quick summary: out-of-support devices grouped by primary user's UsageLocation
# ---------------------------------------------------------------------------
$outOfSupportPath = [System.IO.Path]::ChangeExtension($OutputPath, $null).TrimEnd('.') + "_OutOfSupportOnly.csv"
$report | Where-Object { $_.IsOutOfSupport } | Export-Csv -Path $outOfSupportPath -NoTypeInformation -Encoding UTF8
Write-Host "Out-of-support-only report exported to $outOfSupportPath" -ForegroundColor Green

$report | Where-Object { $_.IsOutOfSupport } |
Group-Object UserUsageLocation |
Select-Object @{N = 'UsageLocation'; E = { $_.Name } }, Count |
Sort-Object Count -Descending |
Format-Table -AutoSize
