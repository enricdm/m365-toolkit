<#
.SYNOPSIS
  Export Entra ID sign-in logs for a VPN application (FortiGate SSL-VPN /
  FortiClient by default) plus Identity Protection risk data, for a credential
  -exposure review. Read-only — no tenant changes are made.

.DESCRIPTION
  Pulls from Microsoft Graph and writes flattened CSVs to <script folder>\Exports:
    1. Service principals matching -AppFilter (confirms the VPN app[s])
    2. Sign-in logs for those app(s) over the last -DaysBack days
    3. Identity Protection risk detections (leaked creds, password spray, etc.)
    4. Currently at-risk users
  Plus a short console summary to flag obvious items before deeper analysis.

  RETENTION: Graph keeps interactive sign-in logs ~30 days (Entra ID P1/P2), so
  -DaysBack above 30 returns nothing older than retention. For a longer window,
  query SigninLogs in Log Analytics / Sentinel instead, if the sign-ins are
  archived there.

.PARAMETER TenantId
  Directory (tenant) ID to connect to.

.PARAMETER AppFilter
  displayName search term used to discover the VPN service principal(s).

.PARAMETER AppId
  Explicit application id(s). Supplying these skips discovery entirely.

.EXAMPLE
  .\Export-VpnSignInLog.ps1 -TenantId '<tenant-id>' -Interactive

.EXAMPLE
  .\Export-VpnSignInLog.ps1 -TenantId '<tenant-id>' -AppId '<client-id>' -DaysBack 14 `
      -ClientId '<client-id>' -CertThumbprint '<cert-thumbprint>'

.NOTES
  Required Graph permissions (delegated via -Interactive, or app-only):
    AuditLog.Read.All, Directory.Read.All,
    IdentityRiskyUser.Read.All, IdentityRiskEvent.Read.All
  Module: Microsoft.Graph.Authentication
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$TenantId,
    [string]$AppFilter   = 'Forti',
    [string[]]$AppId,                                 # explicit app id(s) -> skip discovery
    [int]$DaysBack       = 30,
    [string]$OutputPath  = (Join-Path $PSScriptRoot 'Exports'),

    # App-only cert auth (default convention). Or use -Interactive for delegated.
    [string]$ClientId,
    [string]$CertThumbprint,
    [switch]$Interactive
)

$ErrorActionPreference = 'Stop'

# --- helpers -------------------------------------------------------------
function Write-Step($m){ Write-Host "`n==> $m" -ForegroundColor Cyan }
function Write-OK($m){   Write-Host "  [OK] $m" -ForegroundColor Green }
function Write-Warn($m){ Write-Host "  [!]  $m" -ForegroundColor Yellow }
function Die($m){        Write-Host "  [X]  $m" -ForegroundColor Red; exit 1 }

function Invoke-GraphAll($uri, $headers){
    $items = New-Object System.Collections.Generic.List[object]
    while($uri){
        if($headers){ $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $headers }
        else        { $resp = Invoke-MgGraphRequest -Method GET -Uri $uri }
        if($resp.value){ $items.AddRange($resp.value) }
        $uri = $resp.'@odata.nextLink'
    }
    return $items
}

# --- preflight -----------------------------------------------------------
if(-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)){
    Die "Microsoft.Graph.Authentication not found. Install: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
}

# --- connect -------------------------------------------------------------
Write-Step "Connecting to Microsoft Graph"
$scopes = @('AuditLog.Read.All','Directory.Read.All','IdentityRiskyUser.Read.All','IdentityRiskEvent.Read.All')
if($ClientId -and $CertThumbprint -and -not $Interactive){
    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertThumbprint -NoWelcome
} else {
    Connect-MgGraph -TenantId $TenantId -Scopes $scopes -NoWelcome
}
$ctx = Get-MgContext
if(-not $ctx){ Die "Graph connection failed." }
Write-OK "Connected as $($ctx.Account ?? $ctx.AppName) ($($ctx.AuthType))"

# --- resolve VPN app(s) --------------------------------------------------
Write-Step "Resolving VPN application(s)"
$appIds = @()
if($AppId){
    $appIds = $AppId
    Write-OK "Using supplied AppId(s): $($appIds -join ', ')"
} else {
    $searchVal = $AppFilter.Replace('"','')
    $uri = "https://graph.microsoft.com/v1.0/servicePrincipals?`$search=`"displayName:$searchVal`"&`$select=appId,displayName,servicePrincipalType&`$top=999"
    $sp = Invoke-GraphAll $uri @{ ConsistencyLevel = 'eventual' }
    if(-not $sp -or $sp.Count -eq 0){
        Die "No service principals matched '$AppFilter'. Re-run with -AppFilter <name> or -AppId <guid>."
    }
    $sp | ForEach-Object { Write-Host ("    - {0}  ({1})" -f $_.displayName, $_.appId) }
    $appIds = @($sp.appId)
    Write-OK "$($appIds.Count) candidate app(s) resolved"
}

# --- export sign-ins -----------------------------------------------------
$since = (Get-Date).ToUniversalTime().AddDays(-$DaysBack).ToString('yyyy-MM-ddTHH:mm:ssZ')
Write-Step "Exporting sign-ins since $since UTC (last $DaysBack days)"

$signIns = New-Object System.Collections.Generic.List[object]
foreach($id in $appIds){
    $f   = "createdDateTime ge $since and appId eq '$id'"
    $u   = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$filter=$([uri]::EscapeDataString($f))&`$top=999"
    $b   = Invoke-GraphAll $u
    Write-OK "$($b.Count) sign-ins for appId $id"
    if($b.Count){ $signIns.AddRange($b) }
}

$rows = foreach($s in $signIns){
    [pscustomobject]@{
        createdDateTime           = $s.createdDateTime
        userPrincipalName         = $s.userPrincipalName
        userDisplayName           = $s.userDisplayName
        userId                    = $s.userId
        appDisplayName            = $s.appDisplayName
        appId                     = $s.appId
        clientAppUsed             = $s.clientAppUsed
        ipAddress                 = $s.ipAddress
        city                      = $s.location.city
        state                     = $s.location.state
        countryOrRegion           = $s.location.countryOrRegion
        autonomousSystemNumber    = $s.autonomousSystemNumber
        statusErrorCode           = $s.status.errorCode
        statusFailureReason       = $s.status.failureReason
        conditionalAccessStatus   = $s.conditionalAccessStatus
        authenticationRequirement = $s.authenticationRequirement
        mfaAuthMethod             = $s.mfaDetail.authMethod
        isInteractive             = $s.isInteractive
        riskLevelDuringSignIn     = $s.riskLevelDuringSignIn
        riskState                 = $s.riskState
        riskDetail                = $s.riskDetail
        deviceOS                  = $s.deviceDetail.operatingSystem
        deviceBrowser             = $s.deviceDetail.browser
        correlationId             = $s.correlationId
    }
}

if(-not (Test-Path $OutputPath)){ New-Item -ItemType Directory -Path $OutputPath | Out-Null }
$stamp     = Get-Date -Format 'yyyyMMdd-HHmmss'
$signInCsv = Join-Path $OutputPath "VPN-SignIns_$stamp.csv"
$rows | Sort-Object createdDateTime | Export-Csv -Path $signInCsv -NoTypeInformation -Encoding UTF8
Write-OK "Sign-ins exported: $signInCsv ($($rows.Count) rows)"

# --- export risk detections ---------------------------------------------
Write-Step "Exporting Identity Protection risk detections"
try {
    $rd = Invoke-GraphAll "https://graph.microsoft.com/v1.0/identityProtection/riskDetections?`$filter=$([uri]::EscapeDataString("detectedDateTime ge $since"))"
} catch {
    Write-Warn "Filtered query failed ($($_.Exception.Message)); pulling unfiltered and filtering client-side."
    $rd = Invoke-GraphAll "https://graph.microsoft.com/v1.0/identityProtection/riskDetections" |
          Where-Object { $_.detectedDateTime -ge $since }
}
$rdRows = foreach($r in $rd){
    [pscustomobject]@{
        detectedDateTime    = $r.detectedDateTime
        userPrincipalName   = $r.userPrincipalName
        riskType            = $r.riskType
        riskEventType       = $r.riskEventType
        riskLevel           = $r.riskLevel
        riskState           = $r.riskState
        ipAddress           = $r.ipAddress
        city                = $r.location.city
        countryOrRegion     = $r.location.countryOrRegion
        detectionTimingType = $r.detectionTimingType
        source              = $r.source
        correlationId       = $r.correlationId
    }
}
$rdCsv = Join-Path $OutputPath "RiskDetections_$stamp.csv"
$rdRows | Sort-Object detectedDateTime | Export-Csv -Path $rdCsv -NoTypeInformation -Encoding UTF8
Write-OK "Risk detections exported: $rdCsv ($($rdRows.Count) rows)"

# --- export risky users --------------------------------------------------
Write-Step "Exporting currently at-risk users"
$ru = Invoke-GraphAll "https://graph.microsoft.com/v1.0/identityProtection/riskyUsers?`$filter=$([uri]::EscapeDataString("riskState eq 'atRisk'"))"
$ruRows = foreach($u in $ru){
    [pscustomobject]@{
        userPrincipalName     = $u.userPrincipalName
        userDisplayName       = $u.userDisplayName
        riskLevel             = $u.riskLevel
        riskState             = $u.riskState
        riskDetail            = $u.riskDetail
        riskLastUpdatedDateTime = $u.riskLastUpdatedDateTime
    }
}
$ruCsv = Join-Path $OutputPath "RiskyUsers_$stamp.csv"
$ruRows | Export-Csv -Path $ruCsv -NoTypeInformation -Encoding UTF8
Write-OK "Risky users exported: $ruCsv ($($ruRows.Count) rows)"

# --- console summary -----------------------------------------------------
Write-Step "Quick summary (full analysis on the CSVs)"
$ok   = ($rows | Where-Object { $_.statusErrorCode -eq 0 }).Count
$fail = $rows.Count - $ok
$risky= ($rows | Where-Object { $_.riskLevelDuringSignIn -and $_.riskLevelDuringSignIn -ne 'none' }).Count
Write-Host ("    Sign-ins:        {0}  (success {1} / fail {2})" -f $rows.Count, $ok, $fail)
Write-Host ("    Distinct users:  {0}" -f ($rows.userPrincipalName | Sort-Object -Unique).Count)
Write-Host ("    Countries:       {0}" -f (($rows.countryOrRegion | Where-Object {$_} | Sort-Object -Unique) -join ', '))
Write-Host ("    Risky sign-ins:  {0}" -f $risky)
Write-Host ("    Risk detections: {0}" -f $rdRows.Count)
Write-Host ("    At-risk users:   {0}" -f $ruRows.Count)
Write-Host ""
Write-OK "Done. Attach the three CSVs in $OutputPath for analysis."
Disconnect-MgGraph | Out-Null
