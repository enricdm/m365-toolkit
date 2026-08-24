# ============================================================
# Test-PhishingSimulationRule.ps1
#
# Read-only. Proves whether the transport rule that forwards user-reported
# phishing to the simulation vendor is actually firing, by tracing both legs:
#   inbound  -> the trap mailbox users report to
#   outbound -> the vendor's reporting address the rule adds as a recipient
# A rule that "exists" is not a rule that works; only the outbound leg proves it.
#
# Uses Get-MessageTraceV2 (replaces the deprecated Get-MessageTrace).
# Message trace only retains ~10 days, so -DaysBack above that returns nothing.
#
# Produces a CSV of the traced messages plus a short text evidence report,
# both under .\Evidence\ - the kind of artefact you attach to a change record.
# ============================================================

<#
.NOTES
    When to use  : The phishing-simulation vendor says user reports are not reaching them, or you have just changed the transport rule and need evidence it works.
    Why it exists: A rule that exists is not a rule that works. Only the outbound leg to the vendor proves it fired, so both legs are traced and the result is written as an evidence report of the kind you attach to a change record.
#>

param(
    [string]$AdminUPN          = "",                      # empty = prompt / reuse session
    [Parameter(Mandatory)]
    [string]$TrapMailbox,                                 # e.g. phishing.report@contoso.com
    [Parameter(Mandatory)]
    [string]$SimulationReportAddress,                     # vendor address the rule forwards to
    [string]$RuleName          = "Report phishing to simulation vendor",
    [string]$SimulationHeader  = "x-threatsim-header",    # header/marker the rule matches on
    [int]$DaysBack             = 10
)

# ── Output folder (works whether script is run as file OR pasted) ──
$baseFolder = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$OutputFolder = Join-Path $baseFolder "Evidence"
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder | Out-Null }

$timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$csvPath    = Join-Path $OutputFolder "MessageTrace_$timestamp.csv"
$reportPath = Join-Path $OutputFolder "EvidenceReport_$timestamp.txt"

# ── Connect ──────────────────────────────────────────────────
Write-Host "`n[1/3] Connecting to Exchange Online..." -ForegroundColor Cyan
$connectArgs = @{ ShowBanner = $false }
if ($AdminUPN) { $connectArgs['UserPrincipalName'] = $AdminUPN }
Connect-ExchangeOnline @connectArgs

# ── Run V2 trace (search BOTH directions) ────────────────────
Write-Host "[2/3] Running Get-MessageTraceV2..." -ForegroundColor Cyan

$startDate = (Get-Date).AddDays(-$DaysBack)
$endDate   = Get-Date

# Inbound to TRAP mailbox
$inbound = Get-MessageTraceV2 `
    -RecipientAddress $TrapMailbox `
    -StartDate $startDate `
    -EndDate $endDate

# Outbound forwards to the simulation vendor
$outbound = Get-MessageTraceV2 `
    -RecipientAddress $SimulationReportAddress `
    -StartDate $startDate `
    -EndDate $endDate

$all = @($inbound) + @($outbound) | Sort-Object Received -Descending |
       Select-Object Received, SenderAddress, RecipientAddress, Subject, Status, MessageTraceId

Write-Host "      Inbound to trap    : $($inbound.Count)" -ForegroundColor Green
Write-Host "      Outbound to vendor : $($outbound.Count)" -ForegroundColor Green

$all | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

# ── Build short text report ──────────────────────────────────
Write-Host "[3/3] Generating report..." -ForegroundColor Cyan

$ruleFired = $outbound.Count -gt 0

$report = @"
============================================================
  PHISHING SIMULATION TRANSPORT RULE – EVIDENCE REPORT
  Generated : $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")
============================================================

RULE
  Name      : $RuleName
  Condition : To '$TrapMailbox' AND body/subject contains '$SimulationHeader'
  Action    : Add recipient '$SimulationReportAddress'

SEARCH
  Range     : $($startDate.ToString("dd/MM/yyyy HH:mm")) – $($endDate.ToString("dd/MM/yyyy HH:mm"))
  Inbound to trap       : $($inbound.Count)
  Forwarded to vendor   : $($outbound.Count)

RESULT
  Rule working : $(if ($ruleFired) { "✓ YES — $($outbound.Count) message(s) forwarded to $SimulationReportAddress" } else { "✗ NOT CONFIRMED — no forwards detected. Send a test email containing '$SimulationHeader' to $TrapMailbox and re-run." })

------------------------------------------------------------
MESSAGES
------------------------------------------------------------
"@

foreach ($m in $all) {
    $report += "`n  $($m.Received) | $($m.SenderAddress) → $($m.RecipientAddress)"
    $report += "`n    Subject: $($m.Subject)"
    $report += "`n    Status : $($m.Status)`n"
}

$report | Out-File -FilePath $reportPath -Encoding UTF8

# ── Console output ───────────────────────────────────────────
Write-Host "`n============================================================" -ForegroundColor Yellow
Write-Host "  RESULT: $(if ($ruleFired) { 'RULE WORKING ✓' } else { 'NOT CONFIRMED ✗' })" -ForegroundColor $(if ($ruleFired) { "Green" } else { "Red" })
Write-Host "============================================================" -ForegroundColor Yellow
$all | Format-Table Received, SenderAddress, RecipientAddress, Subject, Status -AutoSize
Write-Host "`n  Report : $reportPath"
Write-Host "  CSV    : $csvPath`n"

Disconnect-ExchangeOnline -Confirm:$false