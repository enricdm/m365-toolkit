# ═══════════════════════════════════════════════════════════════════════════
# Export SAML Certificate Notification Emails and Expiry Dates
# ═══════════════════════════════════════════════════════════════════════════
# Fix notes:
#   - Use -OutputType PSObject so Graph returns real PSObjects, not hashtables
#   - Graph returns "0001-01-01T00:00:00Z" for non-SAML apps, not $null
#     This sentinel value must be filtered out explicitly
# ═══════════════════════════════════════════════════════════════════════════

param(
    [string]$OutputPath = ".\SAML_Notification_Emails_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Graph sentinel value for "no SAML cert configured"
$EMPTY_DATE = "0001-01-01T00:00:00Z"

Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host " SAML Certificate Notification Email Export" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

# ─── Connect ──────────────────────────────────────────────────────────────────
Write-Host "[1/4] Connecting to Microsoft Graph..." -ForegroundColor Yellow
try {
    $context = Get-MgContext
    if ($null -eq $context) { throw "Not connected" }
    Write-Host "      Already connected as: $($context.Account)" -ForegroundColor Gray
} catch {
    Connect-MgGraph -Scopes "Application.Read.All", "Directory.Read.All" -NoWelcome
}

# ─── Fetch all Service Principals ─────────────────────────────────────────────
Write-Host "[2/4] Fetching all Service Principals..." -ForegroundColor Yellow

$allSPs = [System.Collections.Generic.List[object]]::new()
$uri = "https://graph.microsoft.com/v1.0/servicePrincipals" +
       "?`$select=id,appId,displayName,preferredTokenSigningKeyEndDateTime,notificationEmailAddresses,signInAudience" +
       "&`$top=100"

do {
    # -OutputType PSObject returns proper PSObjects instead of hashtables
    $response = Invoke-MgGraphRequest -Uri $uri -Method GET -OutputType PSObject
    foreach ($sp in $response.value) { $allSPs.Add($sp) }
    $uri = $response.'@odata.nextLink'
    Write-Host "      Fetched $($allSPs.Count) SPs so far..." -ForegroundColor Gray
} while ($null -ne $uri)

Write-Host "      Total Service Principals: $($allSPs.Count)" -ForegroundColor Green

# ─── Filter to SAML apps ──────────────────────────────────────────────────────
Write-Host "[3/4] Filtering SAML-enabled apps..." -ForegroundColor Yellow

$samlApps = $allSPs | Where-Object {
    $v = $_.preferredTokenSigningKeyEndDateTime
    ($null -ne $v) -and ($v -ne "") -and ($v -ne $EMPTY_DATE)
}

Write-Host "      SAML apps found: $($samlApps.Count)" -ForegroundColor $(if ($samlApps.Count -gt 0) { "Green" } else { "Red" })

if ($samlApps.Count -eq 0) {
    Write-Host ""
    Write-Host "  No SAML apps found. Run this to inspect raw values:" -ForegroundColor Yellow
    Write-Host "  `$allSPs[0..4] | Select displayName, preferredTokenSigningKeyEndDateTime" -ForegroundColor Cyan
    exit
}

# ─── Build export data ────────────────────────────────────────────────────────
Write-Host "[4/4] Building export dataset..." -ForegroundColor Yellow

$exportData = [System.Collections.Generic.List[PSCustomObject]]::new()
$i = 0

foreach ($sp in $samlApps) {
    $i++
    Write-Progress -Activity "Processing SAML apps" `
                   -Status "$i of $($samlApps.Count): $($sp.displayName)" `
                   -PercentComplete (($i / $samlApps.Count) * 100)

    # Notification emails
    $emails = ""
    if ($null -ne $sp.notificationEmailAddresses -and $sp.notificationEmailAddresses.Count -gt 0) {
        $emails = ($sp.notificationEmailAddresses) -join "; "
    }

    # Certificate expiry and status
    $certExpiry = ""
    $certStatus = ""
    $daysLeft   = $null

    $rawDate = $sp.preferredTokenSigningKeyEndDateTime
    if ($null -ne $rawDate -and $rawDate -ne "" -and $rawDate -ne $EMPTY_DATE) {
        try {
            $expiryDate = [DateTimeOffset]::Parse($rawDate).UtcDateTime
            $certExpiry = $expiryDate.ToString("yyyy-MM-dd")
            $daysLeft   = [int]($expiryDate - (Get-Date).ToUniversalTime()).TotalDays

            $certStatus = switch ($true) {
                ($daysLeft -lt 0)  { "Expired" }
                ($daysLeft -le 30) { "Expires soon (<30d)" }
                ($daysLeft -le 90) { "Expires soon (<90d)" }
                default            { "Current" }
            }
        } catch {
            $certExpiry = $rawDate
            $certStatus = "Parse error"
        }
    }

    $exportData.Add([PSCustomObject]@{
        AppId              = $sp.appId
        DisplayName        = $sp.displayName
        NotificationEmails = $emails
        CertificateExpiry  = $certExpiry
        CertificateStatus  = $certStatus
        DaysUntilExpiry    = if ($null -ne $daysLeft) { $daysLeft } else { "" }
        SignInAudience     = $sp.signInAudience
        ServicePrincipalId = $sp.id
    })
}

Write-Progress -Activity "Processing SAML apps" -Completed

# ─── Export ───────────────────────────────────────────────────────────────────
$exportData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -Force

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host " Export Complete!" -ForegroundColor Green
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host ""
Write-Host "  Output : $OutputPath" -ForegroundColor Cyan
Write-Host "  Records: $($exportData.Count)" -ForegroundColor Cyan
Write-Host ""

# ─── Summary ──────────────────────────────────────────────────────────────────
$withEmail    = ($exportData | Where-Object { $_.NotificationEmails -ne "" }).Count
$withoutEmail = $exportData.Count - $withEmail
$expired      = ($exportData | Where-Object { $_.CertificateStatus -eq "Expired" }).Count
$soonLt30     = ($exportData | Where-Object { $_.CertificateStatus -eq "Expires soon (<30d)" }).Count
$soonLt90     = ($exportData | Where-Object { $_.CertificateStatus -eq "Expires soon (<90d)" }).Count
$current      = ($exportData | Where-Object { $_.CertificateStatus -eq "Current" }).Count

Write-Host "  Notification email configured : $withEmail"    -ForegroundColor $(if ($withEmail -gt 0)    {"Green"}  else {"Gray"})
Write-Host "  No notification email         : $withoutEmail" -ForegroundColor $(if ($withoutEmail -gt 0) {"Red"}    else {"Gray"})
Write-Host ""
Write-Host "  Expired                       : $expired"      -ForegroundColor $(if ($expired -gt 0)      {"Red"}    else {"Gray"})
Write-Host "  Expires < 30 days             : $soonLt30"     -ForegroundColor $(if ($soonLt30 -gt 0)     {"Red"}    else {"Gray"})
Write-Host "  Expires < 90 days             : $soonLt90"     -ForegroundColor $(if ($soonLt90 -gt 0)     {"Yellow"} else {"Gray"})
Write-Host "  Current                       : $current"      -ForegroundColor $(if ($current -gt 0)      {"Green"}  else {"Gray"})
Write-Host ""

if ($expired -gt 0) {
    Write-Host "⚠  EXPIRED certificates:" -ForegroundColor Red
    $exportData |
        Where-Object  { $_.CertificateStatus -eq "Expired" } |
        Sort-Object     DaysUntilExpiry |
        Select-Object   DisplayName, CertificateExpiry, NotificationEmails |
        Format-Table   -AutoSize
}

if ($withoutEmail -gt 0) {
    Write-Host "⚠  No notification email (expiry warnings won't reach anyone):" -ForegroundColor Yellow
    $exportData |
        Where-Object  { $_.NotificationEmails -eq "" } |
        Sort-Object     CertificateExpiry |
        Select-Object   DisplayName, CertificateStatus, CertificateExpiry |
        Format-Table   -AutoSize
}