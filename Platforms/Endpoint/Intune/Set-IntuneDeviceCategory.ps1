<#
.SYNOPSIS
    Assigns Intune device categories from a mapping file, matching on network subnet
    or on device name. Reports by default; changes nothing unless -Execute.

.DESCRIPTION
    Device category is what lets dynamic groups, reporting and support routing tell a
    warehouse tablet from a director's laptop. Set by hand it drifts immediately; this
    keeps it derived from something factual.

    Two matching strategies:

      Subnet     - reads hardwareInformation.subnetAddress from each device and looks
                   it up in the mapping. Use when physical location determines the
                   category (site, branch, floor).
      DeviceName - matches the device name against a wildcard pattern in the mapping.
                   Use when naming convention carries the meaning.

    THE COST OF SUBNET MATCHING, STATED UP FRONT
    subnetAddress is not returned by the managedDevices list endpoint. It only appears
    when you GET a single device. So Subnet mode issues one extra Graph call per
    device: a 2,000-device fleet is 2,000 calls and will take minutes, not seconds.
    That is a property of the API, not of this script. DeviceName mode needs no extra
    calls. Progress is reported every 100 devices so a long run does not look hung.

    A device whose subnet is blank is REPORTED AND SKIPPED, never assigned a default.
    A blank subnet means the device has not reported its hardware inventory yet -
    that is missing information, not membership of some fallback category.

.PARAMETER MappingCsv
    CSV with columns Key and Category.
      Subnet mode     Key = subnet address exactly as Intune reports it (e.g. 10.20.30.0)
      DeviceName mode Key = wildcard pattern matched against the device name (e.g. LAP-BCN-*)
    In DeviceName mode rows are evaluated top to bottom and the FIRST match wins, so
    put specific patterns above general ones.

.PARAMETER MatchOn
    Subnet (default) or DeviceName.

.PARAMETER Platform
    Restrict to one operating system. Default Windows.

.PARAMETER CreateMissingCategory
    Create any category named in the mapping that does not exist in Intune. Off by
    default: a typo in the mapping would otherwise silently create a junk category.

.PARAMETER Execute
    Applies the changes. Without this switch the script only reports.

.PARAMETER OutputPath
    CSV path for the change set / result log.

.EXAMPLE
    # Dry run - what would change, and why
    .\Set-IntuneDeviceCategory.ps1 -MappingCsv .\subnet-to-site.csv

.EXAMPLE
    # Apply, creating categories that do not exist yet
    .\Set-IntuneDeviceCategory.ps1 -MappingCsv .\subnet-to-site.csv -CreateMissingCategory -Execute

.EXAMPLE
    # Match on naming convention instead - no per-device calls
    .\Set-IntuneDeviceCategory.ps1 -MappingCsv .\name-to-category.csv -MatchOn DeviceName -Execute

.NOTES
    Requires : Microsoft.Graph.Authentication  (Install-Module Microsoft.Graph -Scope CurrentUser)
    Scopes   : DeviceManagementManagedDevices.ReadWrite.All,
               DeviceManagementConfiguration.ReadWrite.All (with -CreateMissingCategory)
    Rights   : WRITES to Intune when -Execute is used.

    Example mapping (Subnet mode):
        Key,Category
        10.20.30.0,Site-North
        10.20.40.0,Site-South

    Changing a device category can move the device between dynamic groups, and
    therefore change which policies and apps target it. Run without -Execute first
    and read the CSV.

    Replaces (merged): three scripts that each assigned device categories a different
    way - one derived the category from the device's subnet using a site mapping read
    from a hard-coded local path, one matched on device name, and one set a single
    category by hand from a pasted list of device IDs. The mapping file, the match
    strategy and the category names are all parameters here.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$MappingCsv,

    [ValidateSet('Subnet', 'DeviceName')]
    [string]$MatchOn = 'Subnet',

    [ValidateSet('All', 'Windows', 'iOS', 'Android', 'macOS')]
    [string]$Platform = 'Windows',

    [switch]$CreateMissingCategory,

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
    $OutputPath = Join-Path $PSScriptRoot "Exports\DeviceCategory-$ts.csv"
}

$scopes = @('DeviceManagementManagedDevices.ReadWrite.All')
if ($CreateMissingCategory) { $scopes += 'DeviceManagementConfiguration.ReadWrite.All' }
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
# Mapping
# ---------------------------------------------------------------------------
if (-not (Test-Path -LiteralPath $MappingCsv)) { throw "Mapping CSV not found: $MappingCsv" }
$mapRows = @(Import-Csv -LiteralPath $MappingCsv)
if ($mapRows.Count -eq 0) { throw "Mapping CSV is empty: $MappingCsv" }

foreach ($col in @('Key', 'Category')) {
    if ($col -notin $mapRows[0].PSObject.Properties.Name) {
        throw "Mapping CSV must have columns 'Key' and 'Category'. Missing: $col"
    }
}
$mapRows = @($mapRows | Where-Object { $_.Key -and $_.Category })
Write-OK "Mapping rows: $($mapRows.Count)  (match on: $MatchOn)"

# exact lookup for subnet; ordered list for wildcard name matching
$subnetMap = @{}
if ($MatchOn -eq 'Subnet') {
    foreach ($r in $mapRows) {
        $k = $r.Key.Trim()
        if ($subnetMap.ContainsKey($k)) {
            Write-Warn "Duplicate subnet '$k' in mapping; keeping the first ('$($subnetMap[$k])')."
            continue
        }
        $subnetMap[$k] = $r.Category.Trim()
    }
}

# ---------------------------------------------------------------------------
# Existing categories
# ---------------------------------------------------------------------------
Write-Step 'Loading device categories...'
$catByName = @{}
foreach ($c in (Invoke-GraphPaged -Uri 'https://graph.microsoft.com/beta/deviceManagement/deviceCategories')) {
    $n = [string](Get-Prop $c @('displayName'))
    if ($n) { $catByName[$n] = [string](Get-Prop $c @('id')) }
}
Write-OK "Categories in tenant: $($catByName.Count)"

$wanted = @($mapRows | Select-Object -ExpandProperty Category -Unique | ForEach-Object { $_.Trim() })
$missing = @($wanted | Where-Object { -not $catByName.ContainsKey($_) })

if ($missing.Count -gt 0) {
    if (-not $CreateMissingCategory) {
        Write-Warn "$($missing.Count) category name(s) in the mapping do not exist in Intune: $($missing -join ', ')"
        Write-Warn 'Devices matching those rows will be reported as NO-CATEGORY and skipped. Use -CreateMissingCategory to create them.'
    }
    elseif (-not $Execute) {
        Write-Warn "-CreateMissingCategory given without -Execute: would create $($missing.Count) categor(y/ies): $($missing -join ', ')"
    }
    else {
        Write-Step "Creating $($missing.Count) missing categor(y/ies)..."
        foreach ($m in $missing) {
            $body = @{ displayName = $m; description = 'Created by Set-IntuneDeviceCategory.ps1' } | ConvertTo-Json -Compress
            $new  = Invoke-MgGraphRequest -Method POST -ErrorAction Stop `
                    -Uri 'https://graph.microsoft.com/beta/deviceManagement/deviceCategories' `
                    -Body $body -ContentType 'application/json'
            $catByName[$m] = [string](Get-Prop $new @('id'))
            Write-OK "Created category '$m'"
        }
    }
}

# ---------------------------------------------------------------------------
# Devices
# ---------------------------------------------------------------------------
$osFilter = if ($Platform -ne 'All') { "?`$filter=operatingSystem eq '$Platform'" } else { '' }
Write-Step "Loading managed devices (platform: $Platform)..."
$devices = Invoke-GraphPaged -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices$osFilter"
Write-OK "Devices: $($devices.Count)"

if ($MatchOn -eq 'Subnet') {
    Write-Warn "Subnet mode issues one extra Graph call per device ($($devices.Count) calls). This will take a while."
}

$plan = New-Object System.Collections.Generic.List[object]
$i = 0

foreach ($d in $devices) {
    $i++
    if ($i % 100 -eq 0) { Write-Host "    ...$i/$($devices.Count)" -ForegroundColor DarkGray }

    $id      = [string](Get-Prop $d @('id'))
    $name    = [string](Get-Prop $d @('deviceName'))
    $current = [string](Get-Prop $d @('deviceCategoryDisplayName'))

    $matchValue = ''
    $target     = ''
    $note       = ''

    if ($MatchOn -eq 'Subnet') {
        # subnetAddress is only present on a single-device GET
        try {
            $one = Invoke-MgGraphRequest -Method GET -ErrorAction Stop `
                   -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$id`?`$select=id,hardwareInformation"
            $matchValue = [string](Get-Prop (Get-Prop $one @('hardwareInformation')) @('subnetAddress'))
        }
        catch {
            $note = "lookup failed: $($_.Exception.Message)"
        }

        if (-not $note) {
            if (-not $matchValue) {
                $note = 'no subnet reported (hardware inventory not yet received)'
            }
            elseif ($subnetMap.ContainsKey($matchValue)) {
                $target = $subnetMap[$matchValue]
            }
            else {
                $note = 'subnet not in mapping'
            }
        }
    }
    else {
        $matchValue = $name
        foreach ($r in $mapRows) {           # first match wins, by document order
            if ($name -like $r.Key.Trim()) { $target = $r.Category.Trim(); break }
        }
        if (-not $target) { $note = 'device name matched no pattern' }
    }

    $status =
        if     ($note)                          { 'SKIPPED' }
        elseif (-not $catByName.ContainsKey($target)) { 'NO-CATEGORY' }
        elseif ($current -eq $target)           { 'ALREADY-CORRECT' }
        else                                    { 'TO-CHANGE' }

    if ($status -eq 'NO-CATEGORY') { $note = "category '$target' does not exist in Intune" }

    $plan.Add([pscustomobject]@{
        DeviceName      = $name
        DeviceId        = $id
        MatchedOn       = $matchValue
        CurrentCategory = $current
        TargetCategory  = $target
        Status          = $status
        Detail          = $note
    })
}

Write-Step 'Plan'
$plan | Group-Object Status | Sort-Object Count -Descending |
    ForEach-Object { Write-Host ("    {0,-16} {1,6}" -f $_.Name, $_.Count) }

# ---------------------------------------------------------------------------
# Apply
# ---------------------------------------------------------------------------
$toChange = @($plan | Where-Object Status -eq 'TO-CHANGE')

if ($toChange.Count -eq 0) {
    Write-OK 'No category changes needed.'
}
elseif (-not $Execute) {
    Write-Warn "DRY RUN - $($toChange.Count) device(s) would change category. Re-run with -Execute to apply."
}
else {
    Write-Step "Applying $($toChange.Count) change(s)..."
    $ok = 0; $fail = 0
    foreach ($p in $toChange) {
        try {
            $body = @{ '@odata.id' = "https://graph.microsoft.com/beta/deviceManagement/deviceCategories/$($catByName[$p.TargetCategory])" } |
                    ConvertTo-Json -Compress
            Invoke-MgGraphRequest -Method PUT -ErrorAction Stop `
                -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($p.DeviceId)/deviceCategory/`$ref" `
                -Body $body -ContentType 'application/json'
            $ok++
            $p.Status = 'CHANGED'
            Write-OK "$($p.DeviceName): '$($p.CurrentCategory)' -> '$($p.TargetCategory)'"
        }
        catch {
            $fail++
            $p.Status = 'FAILED'
            $p.Detail = $_.Exception.Message
            Write-Err "$($p.DeviceName): $($_.Exception.Message)"
        }
    }
    Write-Step "Changed: $ok   Failed: $fail"
}

$dir = Split-Path -Parent $OutputPath
if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
$plan | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
Write-OK "Log written: $OutputPath"

$plan
