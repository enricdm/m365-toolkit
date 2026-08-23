<#
.SYNOPSIS
  Fetches usageLocation from Microsoft Graph for a list of UPNs and writes a CSV
  that plugs directly into classify_exceptions_by_country.py --usage-location.

.DESCRIPTION
  READ-ONLY. One Graph lookup per UPN; rows that do not resolve are kept with
  Found = "No - <reason>" so nothing silently disappears from the list.

.PARAMETER TenantId
  Directory (tenant) ID to connect to.

.PARAMETER InputCsv
  CSV with a "UserPrincipalName" column - export it from the
  "All Exceptions MFA (by country)" sheet (or just the "Needs Review" sheet if
  you only want to resolve the flagged ones). Default: users.csv next to the script.

.PARAMETER OutputFolder
  Where the timestamped CSV is written. Default: <script folder>\Exports.

.EXAMPLE
  .\Get-ExemptionUsageLocation.ps1 -TenantId '<tenant-id>' -InputCsv .\needs-review.csv

.NOTES
  Install-Module Microsoft.Graph -Scope CurrentUser
  Graph scope: User.Read.All
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$TenantId,
    [string]$InputCsv     = (Join-Path $PSScriptRoot 'users.csv'),
    [string]$OutputFolder = (Join-Path $PSScriptRoot 'Exports')
)

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputCsv = Join-Path $OutputFolder "UsageLocation_$Timestamp.csv"

function Write-Step($msg) { Write-Host "[..] $msg" -ForegroundColor Cyan }
function Write-OK($msg)   { Write-Host "[OK] $msg" -ForegroundColor Green }
function Write-Warn($msg) { Write-Host "[!!] $msg" -ForegroundColor Yellow }

if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder | Out-Null }

Write-Step "Connecting to Microsoft Graph ($TenantId)"
Connect-MgGraph -TenantId $TenantId -Scopes "User.Read.All" -NoWelcome
Write-OK "Connected"

Write-Step "Loading UPN list from $InputCsv"
if (-not (Test-Path $InputCsv)) { Write-Warn "Input file not found: $InputCsv"; exit 1 }
$users = Import-Csv $InputCsv
Write-OK "Loaded $($users.Count) UPNs"

$results = [System.Collections.Generic.List[object]]::new()
$i = 0
foreach ($u in $users) {
    $i++
    $upn = $u.UserPrincipalName
    if (-not $upn) { continue }
    Write-Progress -Activity "Querying Graph" -Status $upn -PercentComplete (($i / $users.Count) * 100)
    try {
        $mgUser = Get-MgUser -UserId $upn -Property Id,UserPrincipalName,UsageLocation,AccountEnabled,UserType -ErrorAction Stop
        $results.Add([PSCustomObject]@{
            UserPrincipalName = $mgUser.UserPrincipalName
            usageLocation     = $mgUser.UsageLocation
            AccountEnabled    = $mgUser.AccountEnabled
            UserType          = $mgUser.UserType
            Found             = "Yes"
        })
    }
    catch {
        $results.Add([PSCustomObject]@{
            UserPrincipalName = $upn
            usageLocation     = $null
            AccountEnabled    = $null
            UserType          = $null
            Found             = "No - $($_.Exception.Message.Split("`n")[0])"
        })
        Write-Warn "Not resolved: $upn"
    }
}
Write-Progress -Activity "Querying Graph" -Completed

$results | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
Write-OK "Wrote $($results.Count) rows to $OutputCsv"

$missing = ($results | Where-Object { $_.usageLocation -eq $null -or $_.usageLocation -eq "" }).Count
Write-Step "Users with no usageLocation set in Entra: $missing"

Disconnect-MgGraph | Out-Null
