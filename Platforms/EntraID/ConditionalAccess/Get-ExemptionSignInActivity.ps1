<#
.SYNOPSIS
    Enriches the MFA-exempt accounts with 30-day sign-in telemetry to judge whether
    each exemption is still needed - or could be scoped to specific apps/locations.

.DESCRIPTION
    For each exempted user (interactive AND non-interactive sign-ins) over the window:
      1. Activity volume + last sign-in in the past N days
      2. Distinct applications signed into
      3. Distinct source IPs + countries
      4. Admin (ADM) accounts are skipped by default - most will be removed from the
         exemption group anyway (-IncludeAdmin to override)
      5. A derived 'Signal' that flags: no activity (drop exemption), single/few apps
         (scope exemption to those apps), non-interactive only (move to cert/managed
         identity), already satisfying MFA, legacy auth clients, and foreign IPs.

    Source set: pass -InputCsv (the exclusions/triage CSV, needs UserPrincipalName and
    optionally ObjectId) or let it expand -ExceptionGroup transitively.

    READ-ONLY. Sign-in logs require Entra ID P1+ and are retained ~30 days, so -Days
    above 30 will return nothing older than retention.

.PARAMETER TenantId       Directory (tenant) ID to connect to.
.PARAMETER ClientId       App-only auth: application (client) ID. Pair with -CertThumbprint.
.PARAMETER CertThumbprint App-only auth: certificate thumbprint. Pair with -ClientId.
.PARAMETER InputCsv       CSV of exempted users (UserPrincipalName[,ObjectId]).
.PARAMETER ExceptionGroup Group expanded transitively if no InputCsv. Rename to match your tenant.
.PARAMETER Days           Look-back window. Default 30.
.PARAMETER IncludeAdmin   Include ADM accounts (skipped by default).
.PARAMETER AdminPattern   Regex marking admin UPNs. Default matches -ADM-, adm., ADM_.
.PARAMETER AllRows        Keep every row of -InputCsv instead of narrowing to true exemptions.
.PARAMETER Detailed       Also write a raw per-sign-in CSV for deep dives.

.EXAMPLE
    .\Get-ExemptionSignInActivity.ps1 -TenantId '<tenant-id>' `
        -InputCsv .\Exports\CA-MFA-Exclusions_20260101-120000.csv
    Chained run: consumes the CSV produced by Get-MfaExclusion.ps1.

.EXAMPLE
    .\Get-ExemptionSignInActivity.ps1 -TenantId '<tenant-id>' `
        -ExceptionGroup 'CA-Exception-MFA' -Days 30 -Detailed
    Standalone run: expands the exception group itself and writes per-sign-in detail.

.NOTES
    When to use  : You already have the list of exempt accounts and now have to decide whose exemption is dropped without breaking their Monday.
    Why it exists: Turns 'these people are exempt' into 'drop this one, scope that one to a single app, this one is already completing MFA anyway'. It also separates exclusion from a block policy - which is not an exemption - from a real MFA exemption, which is the error that inflates the total.
    Graph permissions: AuditLog.Read.All, Directory.Read.All (+ Group.Read.All if expanding a group)
    Modules: Microsoft.Graph.Authentication, .Reports (Get-MgAuditLogSignIn), .Groups
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertThumbprint,
    [string]$InputCsv,
    [string]$ExceptionGroup = 'CA-Exception-MFA',
    [int]$Days = 30,
    [switch]$IncludeAdmin,
    [string]$AdminPattern = '(?i)(-ADM-|^adm[._]|ADM_)',
    [switch]$AllRows,
    [switch]$Detailed,
    [string]$ExportDir = (Join-Path $PSScriptRoot 'Exports')
)

function Write-Step($m) { Write-Host "`n=> $m" -ForegroundColor Cyan }
function Write-OK  ($m) { Write-Host "   [OK]  $m" -ForegroundColor Green }
function Write-Warn($m) { Write-Host "   [!]   $m" -ForegroundColor Yellow }
function Die       ($m) { Write-Host "   [X]   $m" -ForegroundColor Red; exit 1 }

$LegacyClients = @('Other clients', 'IMAP', 'IMAP4', 'POP', 'POP3', 'SMTP', 'Authenticated SMTP', 'Exchange ActiveSync', 'MAPI Over HTTP', 'Offline Address Book', 'Exchange Web Services', 'AutoDiscover')

# ---- connect --------------------------------------------------------------
Write-Step "Connecting to Microsoft Graph (tenant $TenantId)"
try {
    if ($ClientId -and $CertThumbprint) {
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertThumbprint -NoWelcome -ErrorAction Stop
        Write-OK "App-only certificate auth"
    } else {
        Connect-MgGraph -TenantId $TenantId -NoWelcome -ErrorAction Stop -Scopes 'AuditLog.Read.All', 'Directory.Read.All', 'Group.Read.All'
        Write-Warn "Interactive delegated auth"
    }
} catch { Die "Graph connect failed: $($_.Exception.Message)" }

# ---- build target set -----------------------------------------------------
Write-Step "Building target user set"
$targets = @()   # objects with UPN + Id
if ($InputCsv -and (Test-Path $InputCsv)) {
    $rows = @(Import-Csv $InputCsv)
    # If handed the full exclusions dump, narrow to TRUE MFA exemptions (exception
    # group + direct user exclusions). Groups that *enforce* MFA on their members are
    # typically excluded from BLOCK policies too, which does not make them exempt —
    # counting those rows inflates the exemption total. Use -AllRows to keep them.
    if (-not $AllRows -and ($rows | Get-Member -Name ExcludedVia)) {
        $before = $rows.Count
        $rows = $rows | Where-Object { $_.ExcludedVia -match $ExceptionGroup -or $_.ExcludedVia -match 'Direct user exclusion' }
        if ($before -ne $rows.Count) { Write-Warn ("Narrowed {0} exclusion rows to {1} true MFA exemptions (use -AllRows to keep all)" -f $before, $rows.Count) }
    }
    foreach ($r in $rows) {
        $upn = $r.UserPrincipalName
        if ($upn) { $targets += [pscustomobject]@{ UPN = $upn.Trim(); Id = $r.ObjectId } }
    }
    Write-OK ("{0} users from {1}" -f $targets.Count, (Split-Path $InputCsv -Leaf))
} else {
    $g = Get-MgGroup -Filter "displayName eq '$ExceptionGroup'" -Property Id, DisplayName -ErrorAction SilentlyContinue
    if (-not $g) { Die "Group '$ExceptionGroup' not found and no -InputCsv supplied." }
    Get-MgGroupTransitiveMember -GroupId $g.Id -All |
    Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.user' } |
    ForEach-Object { $targets += [pscustomobject]@{ UPN = $_.AdditionalProperties['userPrincipalName']; Id = $_.Id } }
    Write-OK ("{0} transitive members of {1}" -f $targets.Count, $ExceptionGroup)
}

if (-not $IncludeAdmin) {
    $before = $targets.Count
    $targets = $targets | Where-Object { $_.UPN -notmatch $AdminPattern }
    Write-OK ("Skipped {0} ADM accounts (use -IncludeAdmin to keep) -> {1} to query" -f ($before - $targets.Count), $targets.Count)
}

# ---- query sign-ins -------------------------------------------------------
$start = (Get-Date).ToUniversalTime().AddDays(-$Days).ToString('yyyy-MM-ddTHH:mm:ssZ')
Write-Step "Querying sign-ins since $start (interactive + non-interactive)"
$agg = New-Object System.Collections.Generic.List[object]
$raw = New-Object System.Collections.Generic.List[object]
$n = 0
foreach ($t in $targets) {
    $n++; Write-Progress -Activity "Sign-in pull" -Status "$n / $($targets.Count)  $($t.UPN)" -PercentComplete ($n * 100 / [Math]::Max(1, $targets.Count))
    $key = if ($t.Id) { "userId eq '$($t.Id)'" } else { "userPrincipalName eq '$($t.UPN)'" }
    $base = "$key and createdDateTime ge $start"
    # Both queries must be tracked, because the Signal derived below turns "no events"
    # into "drop the exemption". A failed query looks exactly like a quiet account, and
    # the non-interactive leg is the one that matters for service accounts - it is the
    # difference between an unused exemption and the only reason a nightly job still works.
    $events = @()
    $queryFailed = $false
    try { $events += Get-MgAuditLogSignIn -Filter $base -All -ErrorAction Stop }
    catch { $queryFailed = $true; Write-Warn "interactive query failed for $($t.UPN): $($_.Exception.Message)" }
    try { $events += Get-MgAuditLogSignIn -Filter "$base and signInEventTypes/any(x: x eq 'nonInteractiveUser')" -All -ErrorAction Stop }
    catch { $queryFailed = $true; Write-Warn "non-interactive query failed for $($t.UPN): $($_.Exception.Message)" }

    if ($Detailed) {
        foreach ($e in $events) {
            $raw.Add([pscustomobject]@{
                    UserPrincipalName = $t.UPN; CreatedDateTime = $e.CreatedDateTime; App = $e.AppDisplayName; Resource = $e.ResourceDisplayName
                    IP = $e.IPAddress; Country = $e.Location.CountryOrRegion; City = $e.Location.City; ClientApp = $e.ClientAppUsed
                    Interactive = $e.IsInteractive; AuthReq = $e.AuthenticationRequirement; CAStatus = $e.ConditionalAccessStatus
                    ErrorCode = $e.Status.ErrorCode 
                }) 
        } 
    }

    $apps = @($events | ForEach-Object { $_.AppDisplayName } | Where-Object { $_ } | Sort-Object -Unique)
    $ips = @($events | ForEach-Object { $_.IPAddress } | Where-Object { $_ } | Sort-Object -Unique)
    $countries = @($events | ForEach-Object { $_.Location.CountryOrRegion } | Where-Object { $_ } | Sort-Object -Unique)
    $clients = @($events | ForEach-Object { $_.ClientAppUsed } | Where-Object { $_ } | Sort-Object -Unique)
    $interactiveCount = @($events | Where-Object { $_.IsInteractive }).Count
    $mfaSatisfied = @($events | Where-Object { $_.AuthenticationRequirement -eq 'multiFactorAuthentication' }).Count
    $legacy = @($clients | Where-Object { $_ -in $LegacyClients })
    $last = ($events | Sort-Object CreatedDateTime -Descending | Select-Object -First 1).CreatedDateTime

    # ---- derive signal ----
    # Order matters. "Non-interactive only" is tested BEFORE the app-count branches:
    # a service account typically talks to one or two apps, so under the old order it was
    # labelled "single app - scope the exemption" and the fact that no human ever signs in
    # - the thing that actually tells you to move it to a certificate or a managed identity
    # - was never reached.
    $signal = switch ($true) {
        ($queryFailed) { "UNKNOWN - a sign-in query failed for this account; do not act on this row"; break }
        ($events.Count -eq 0) { "No sign-ins in $Days d - exemption likely unneeded"; break }
        ($interactiveCount -eq 0) { "Non-interactive only - service auth; move to cert/managed identity"; break }
        ($apps.Count -eq 1) { "Single app ($($apps[0])) - scope exemption to this app"; break }
        ($apps.Count -le 3) { "Few apps ($($apps.Count)) - candidate to scope to named apps"; break }
        default { "Broad usage ($($apps.Count) apps) - review; exemption may be justified" }
    }
    $flags = @()
    if ($mfaSatisfied -gt 0) { $flags += "Already completing MFA on $mfaSatisfied sign-in(s) - exemption moot" }
    if ($legacy.Count) { $flags += "Legacy auth client: $($legacy -join ',')" }
    if ($ips.Count -eq 1) { $flags += "Single source IP - scope to trusted location" }
    if ($countries.Count -gt 1) { $flags += "Multiple countries: $($countries -join ',')" }

    $agg.Add([pscustomobject]@{
            UserPrincipalName = $t.UPN
            SignIns           = $events.Count
            Interactive       = $interactiveCount
            LastSignIn        = $last
            AppCount          = $apps.Count
            Apps              = (($apps | Select-Object -First 10) -join '; ')
            IPCount           = $ips.Count
            IPs               = (($ips | Select-Object -First 10) -join '; ')
            Countries         = ($countries -join '; ')
            ClientApps        = ($clients -join '; ')
            MfaSatisfiedCount = $mfaSatisfied
            Signal            = $signal
            Flags             = ($flags -join ' | ')
            ObjectId          = $t.Id
        })
}
Write-Progress -Activity "Sign-in pull" -Completed

# ---- sort: best removal/scoping candidates first --------------------------
# UNKNOWN sorts to the top, not the bottom: those rows are the ones a reviewer must not
# act on, so they belong where they will be seen rather than at the end of the table.
$order = @{ 'UNKNOWN' = -1; 'No' = 0; 'Non' = 1; 'Single' = 2; 'Few' = 3; 'Broad' = 4 }
$agg = $agg | Sort-Object `
@{E = { $k = ($_.Signal -split ' ')[0]; if ($order.ContainsKey($k)) { $order[$k] }else { 5 } } }, `
@{E = { $_.SignIns } }, UserPrincipalName

# ---- export ---------------------------------------------------------------
if (-not (Test-Path $ExportDir)) { New-Item -ItemType Directory -Path $ExportDir | Out-Null }
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$out = Join-Path $ExportDir "Exemption-SignInActivity_$stamp.csv"
$agg | Export-Csv -Path $out -NoTypeInformation -Encoding UTF8
Write-Step "Done"
Write-OK "Exported $($agg.Count) users -> $out"
if ($Detailed) {
    $det = Join-Path $ExportDir "Exemption-SignIns-Detail_$stamp.csv"
    $raw | Export-Csv -Path $det -NoTypeInformation -Encoding UTF8
    Write-OK "Per-sign-in detail -> $det"
}
$unknown = @($agg | Where-Object { $_.Signal -like 'UNKNOWN*' })
Write-Warn ("{0} have NO activity in {1}d (drop the exemption)" -f @($agg | Where-Object { $_.Signal -notlike 'UNKNOWN*' -and $_.SignIns -eq 0 }).Count, $Days)
Write-Warn ("{0} use a single app (scope to that app)" -f @($agg | Where-Object { $_.AppCount -eq 1 }).Count)
if ($unknown.Count -gt 0) {
    Write-Warn ("{0} row(s) could not be queried and are NOT counted above. Zero sign-ins there means 'not measured', not 'unused' - re-run before acting on them." -f $unknown.Count)
}
$agg | Select-Object UserPrincipalName, SignIns, AppCount, IPCount, Signal | Format-Table -AutoSize