<#
.SYNOPSIS
    Password-age report for the enabled users of one country. READ-ONLY.

.DESCRIPTION
    Lists every enabled user with the given usageLocation, their last password
    change, the resulting age in days and whether that exceeds the policy maximum.
    Includes last successful sign-in and the on-premises OU for context.

.PARAMETER UsageLocation
    ISO 3166-1 alpha-2 code to filter on, e.g. 'ES'.

.PARAMETER MaxAgeDays
    Password-policy maximum. Users above it are reported as 'Over policy'.

.PARAMETER ExportDir
    Output folder. Default: <script folder>\Exports.

.EXAMPLE
    .\Get-PasswordAge.ps1 -UsageLocation 'ES' -MaxAgeDays 90

.NOTES
    Scopes: User.Read.All, AuditLog.Read.All (signInActivity needs Entra ID P1+).
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^[A-Za-z]{2}$')]
    [string]$UsageLocation,

    [int]$MaxAgeDays = 90,

    [string]$ExportDir = (Join-Path $PSScriptRoot 'Exports')
)

Connect-MgGraph -Scopes "User.Read.All", "AuditLog.Read.All"

if (-not (Test-Path $ExportDir)) { New-Item -ItemType Directory -Path $ExportDir -Force | Out-Null }

$UsageLocation = $UsageLocation.ToUpper()
$maxAgeDays = $MaxAgeDays
$now = (Get-Date).ToUniversalTime()

$props = "id,displayName,userPrincipalName,accountEnabled,lastPasswordChangeDateTime," +
"usageLocation,onPremisesSyncEnabled,onPremisesDistinguishedName,signInActivity"

Get-MgUser -All -ConsistencyLevel eventual -CountVariable c `
  -Filter "usageLocation eq '$UsageLocation' and accountEnabled eq true" `
  -Property $props |
Select-Object displayName, userPrincipalName, accountEnabled,
@{n = "LastPwdChangeUTC"; e = { $_.lastPasswordChangeDateTime } },
@{n = "PwdAgeDays"; e = { if ($_.lastPasswordChangeDateTime) { [int]($now - $_.lastPasswordChangeDateTime.ToUniversalTime()).TotalDays } } },
@{n = "Status"; e = {
    if (-not $_.lastPasswordChangeDateTime) { "Unknown" }
    elseif (($now - $_.lastPasswordChangeDateTime.ToUniversalTime()).TotalDays -gt $maxAgeDays) { "Over policy" }
    else { "OK" }
  }
},
@{n = "LastSuccessfulSignIn"; e = { $_.signInActivity.lastSuccessfulSignInDateTime } },
onPremisesSyncEnabled,
@{n = "OU"; e = { $_.onPremisesDistinguishedName } } |
Sort-Object Status, PwdAgeDays -Descending |
Export-Csv (Join-Path $ExportDir "${UsageLocation}_Entra_PasswordAge.csv") -NoTypeInformation -Encoding UTF8