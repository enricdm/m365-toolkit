# ==============================================================================
# Get-MailboxReceiveVolume.ps1
#
# Counts, per mailbox, how much mail was received over a period, and splits the
# tenant into "under the threshold" and "at or over it" — the input any per-mailbox
# mail-security or archiving licensing decision needs.
#
# Uses the Microsoft Graph "Email Activity User Detail" report — the SAME data
# source as the M365 admin center's Email Activity report, so the numbers can be
# reconciled against the portal instead of argued about.
#
# Why this is better than looping Get-MessageTraceV2 per mailbox:
#   - ONE API call returns aggregated Receive Count per mailbox for D7/D30/D90/D180
#   - No pagination, no chunking, no rate-limit dance
#   - Aggregation is done server-side by Microsoft (matches admin center exactly)
#   - Runs in seconds instead of minutes
# The message-trace approach needs one trace per mailbox and still only reaches
# back ~10 days; the Graph report reaches 180 and costs a single request.
#
# Read-only: it changes nothing in the tenant.
#
# Requirements:
#   Install-Module Microsoft.Graph.Reports -Scope CurrentUser
#   Permission: Reports.Read.All  (admin-consent required)
#
# Note: Graph usage data is always ~2 days behind real-time, and the admin
# center report has the same lag. For a licensing decision that is fine.
#
# Caveat: if the tenant has "Concealed names" enabled, UPNs come back hashed.
# The counts stay valid but you cannot map them to users — the script detects
# this and tells you where to turn it off.
# ==============================================================================

<#
.NOTES
    When to use  : You have to decide how many per-mailbox licences to buy and the number will be checked by someone who wants to reconcile it against the portal.
    Why it exists: Uses the Graph Email Activity User Detail report rather than one message trace per mailbox: one call, server-side aggregation by Microsoft, so the figures match the admin centre exactly and reach 180 days instead of ten. It detects concealed report names and recommends re-running at D90 and D180 before committing to a licence count.
#>

[CmdletBinding()]
param (
    [ValidateSet('D7','D30','D90','D180')]
    [string]$Period         = 'D30',

    [int]   $MailThreshold  = 150,

    [string]$ExportCsvPath  = ".\MailboxReceiveVolume_${Period}_$(Get-Date -Format 'yyyy-MM-dd').csv"
)

Write-Host "`n=== Mailbox Receive Volume (Graph Reports API) ===" -ForegroundColor Cyan
Write-Host "Period    : $Period"
Write-Host "Threshold : < $MailThreshold (strictly less than)`n"

# ------------------------------------------------------------------------------
# 1. Connect to Microsoft Graph
# ------------------------------------------------------------------------------
try {
    $ctx = Get-MgContext -ErrorAction Stop
    if (-not $ctx) { throw "no context" }
    if ($ctx.Scopes -notcontains 'Reports.Read.All') {
        Write-Warning "Current Graph session lacks Reports.Read.All — reconnecting."
        Disconnect-MgGraph | Out-Null
        Connect-MgGraph -Scopes 'Reports.Read.All' -NoWelcome
    }
} catch {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
    Connect-MgGraph -Scopes 'Reports.Read.All' -NoWelcome
}

# ------------------------------------------------------------------------------
# 2. CHECK: does the tenant have "Concealed Names" turned on?
#    If yes, UPN/DisplayName come back as "5C09F00..." hashes and the report
#    is useless for per-user identification. Warn so the user can flip the
#    setting in M365 admin center (Settings -> Org settings -> Reports).
# ------------------------------------------------------------------------------
# (We can't query the setting directly without extra perms; we detect it from
# the first data row after download.)

# ------------------------------------------------------------------------------
# 3. Pull the report — single Graph call, returns CSV
# ------------------------------------------------------------------------------
$tempCsv = Join-Path $env:TEMP "EmailActivityUserDetail_$([guid]::NewGuid()).csv"
$uri     = "https://graph.microsoft.com/v1.0/reports/getEmailActivityUserDetail(period='$Period')"

Write-Host "Downloading Email Activity report (period=$Period)..." -ForegroundColor Yellow
try {
    Invoke-MgGraphRequest -Method GET -Uri $uri -OutputFilePath $tempCsv -ErrorAction Stop
} catch {
    Write-Error "Graph request failed: $($_.Exception.Message)"
    exit 1
}

$fileSize = (Get-Item $tempCsv).Length
Write-Host "  Downloaded $([math]::Round($fileSize/1KB,1)) KB to $tempCsv"

# ------------------------------------------------------------------------------
# 4. Parse CSV
#    Graph prepends a UTF-8 BOM + sometimes garbage characters in front of
#    "Report Refresh Date". Strip them so Import-Csv works cleanly.
# ------------------------------------------------------------------------------
$raw    = Get-Content -Path $tempCsv -Raw -Encoding UTF8
$raw    = $raw -replace '^[^\w]*Report Refresh Date', 'Report Refresh Date'
$report = $raw | ConvertFrom-Csv

Write-Host "  Parsed $($report.Count) rows from report.`n"

# Concealed-names detection
$firstUpn = ($report | Select-Object -First 1).'User Principal Name'
if ($firstUpn -match '^[0-9A-F]{30,}$') {
    Write-Warning "Tenant has 'Concealed names' enabled — UPNs are hashed."
    Write-Warning "Counts are still valid, but you can't map them to real users."
    Write-Warning "Fix: M365 admin center -> Settings -> Org settings -> Services -> Reports"
    Write-Warning "     -> uncheck 'Display concealed user, group, and site names'"
}

# ------------------------------------------------------------------------------
# 5. Build results — keep ALL rows from the report
# ------------------------------------------------------------------------------
$results = foreach ($row in $report) {
    $rcv = 0
    [int]::TryParse($row.'Receive Count', [ref]$rcv) | Out-Null

    [PSCustomObject]@{
        DisplayName       = $row.'Display Name'
        UserPrincipalName = $row.'User Principal Name'
        ReceiveCount      = $rcv
        SendCount         = $row.'Send Count'
        ReadCount         = $row.'Read Count'
        LastActivityDate  = $row.'Last Activity Date'
        AssignedProducts  = $row.'Assigned Products'
        LiteEligible      = $rcv -lt $MailThreshold
    }
}

# ------------------------------------------------------------------------------
# 6. Summary
# ------------------------------------------------------------------------------
$eligible   = @($results | Where-Object {  $_.LiteEligible })
$ineligible = @($results | Where-Object { -not $_.LiteEligible })

Write-Host "--- Summary ($Period) ---" -ForegroundColor Cyan
Write-Host ("{0,-15} {1,12} {2,12} {3,12}" -f "Report", "Rcvd <$MailThreshold", "Rcvd >=$MailThreshold", "Total")
Write-Host ("{0,-15} {1,12} {2,12} {3,12}" -f $Period, $eligible.Count, $ineligible.Count, $results.Count)

Write-Host "`nRe-run with -Period D90 / D180 to see how stable the split is before" -ForegroundColor DarkGray
Write-Host "committing to a licence count; a single 30-day window can mislead." -ForegroundColor DarkGray

Write-Host "`nTop 10 highest-volume mailboxes:" -ForegroundColor Yellow
$results | Sort-Object ReceiveCount -Descending | Select-Object -First 10 |
    Format-Table DisplayName, UserPrincipalName, ReceiveCount, LiteEligible -AutoSize

# ------------------------------------------------------------------------------
# 7. Export & cleanup
# ------------------------------------------------------------------------------
$results | Sort-Object ReceiveCount -Descending |
    Export-Csv -Path $ExportCsvPath -NoTypeInformation -Encoding unicode

Write-Host "Full results exported to: $ExportCsvPath" -ForegroundColor Cyan
Remove-Item $tempCsv -Force -ErrorAction SilentlyContinue
Write-Host ""
