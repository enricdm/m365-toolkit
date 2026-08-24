<#
.SYNOPSIS
    Analyzes Proofpoint users with missing Location attribute against Entra ID.

.DESCRIPTION
    Reads a Proofpoint user export CSV and identifies users where the "Location" field
    (mapped from Entra ID UsageLocation) is not populated. Queries Microsoft Graph to
    categorize each case and determine the required action.

    EXO dependency removed — mailbox type is resolved via Graph mailboxSettings.userPurpose,
    which avoids WAM/MSAL broker issues in modern terminal hosts (Windows Terminal, VS Code).

    Output categories:
      A  – UsageLocation IS set in Entra (Proofpoint sync lag)           → No action needed
      B  – Non-user mailbox (shared/room/equipment)                      → Expected, not required
      C  – Licensed user mailbox, UsageLocation missing                  → MUST FIX
      D  – Unlicensed Member, UsageLocation missing                      → Should fix
      E  – Disabled account, UsageLocation missing                       → Low priority
      F  – Guest / B2B user                                              → Not required
      G  – No SSO ID in Proofpoint export                                → Not a synced Entra user
      X  – Object not found in Entra (stale Proofpoint record)           → Investigate

.PARAMETER CsvPath
    Path to the Proofpoint user export CSV file.

.PARAMETER OutputFolder
    Folder where the analysis report CSV will be saved. Defaults to the script folder.

.PARAMETER BatchSize
    Users per processing batch (throttle courtesy pause). Default: 20.

.PARAMETER DomainMapPath
    Domain -> country data file used by the optional remediation snippet at the
    bottom of this script. Defaults to the sample map in Platforms\_Shared\Data,
    which ships with placeholder contoso.* domains — replace them with your own.

.EXAMPLE
    .\Get-MissingLocationReport.ps1 -CsvPath ".\proofpoint_export.csv"

.NOTES
    When to use  : Proofpoint (or anything else fed from Entra) has users with no Location and you need to know which are a real problem and which are noise.
    Why it exists: Eight actionable categories with a recommended action each, not a list of UPNs: sync lag, shared mailbox, guest, disabled, and 'licensed with no usageLocation, fix now' are different problems. Mailbox type comes from Graph mailboxSettings.userPurpose so Exchange Online is not needed, and Unknown stays distinct from 'not a user mailbox'.
    READ-ONLY. The bulk remediation block at the end of the file is commented out
    and has to be run deliberately.

    Prerequisites:
        Install-Module Microsoft.Graph.Users -Scope CurrentUser

    Graph scopes requested:
        User.Read.All, Directory.Read.All, MailboxSettings.Read
    (the commented remediation block additionally needs User.ReadWrite.All)
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CsvPath,

    [string]$OutputFolder = $PSScriptRoot,

    [ValidateRange(1, 20)]
    [int]$BatchSize = 20,

    [string]$DomainMapPath = (Join-Path $PSScriptRoot '..\..\_Shared\Data\domain-country-map.psd1')
)

#region ── Bootstrap ────────────────────────────────────────────────────────────

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptStart = Get-Date
$timestamp = $scriptStart.ToString('yyyyMMdd_HHmmss')
$reportPath = Join-Path $OutputFolder "ProofpointMissingLocation_$timestamp.csv"
$logPath = Join-Path $OutputFolder "ProofpointMissingLocation_$timestamp.log"

function Write-Log {
    param ([string]$Message, [string]$Level = 'INFO')
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message
    Write-Host $line -ForegroundColor $(switch ($Level) {
            'WARN' { 'Yellow' }
            'ERROR' { 'Red' }
            'OK' { 'Green' }
            default { 'Cyan' }
        })
    Add-Content -Path $logPath -Value $line
}

#endregion

#region ── Connect (Graph only — no EXO) ────────────────────────────────────────

Write-Log "Connecting to Microsoft Graph..."
try {
    Connect-MgGraph `
        -Scopes 'User.Read.All', 'Directory.Read.All', 'MailboxSettings.Read' `
        -NoWelcome
    Write-Log "Graph connected." 'OK'
} catch {
    Write-Log "Graph connection failed: $_" 'ERROR'
    throw
}

#endregion

#region ── Helper: Get-MailboxPurpose ───────────────────────────────────────────
#
#  Calls /users/{id}/mailboxSettings and returns the userPurpose string.
#  Graph values: user | shared | room | equipment | linked | unknownFutureValue
#  Returns 'NoMailbox' when the user has no Exchange mailbox at all.
#  Returns 'Unknown'   on any other error (e.g. throttling, transient).
#
function Get-MailboxPurpose {
    param ([string]$ObjectId)
    try {
        $resp = Invoke-MgGraphRequest `
            -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/users/$ObjectId/mailboxSettings" `
            -ErrorAction Stop
        # userPurpose may be absent on very old/hybrid objects — default to 'user'
        return ($resp.userPurpose ?? 'user')
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match 'MailboxNotEnabledForRESTAPI|ResourceNotFound|does not have a mailbox') {
            return 'NoMailbox'
        }
        Write-Log "  mailboxSettings call failed for $ObjectId — $msg" 'WARN'
        return 'Unknown'
    }
}

#endregion

#region ── Load & split CSV ──────────────────────────────────────────────────────

Write-Log "Loading CSV: $CsvPath"
$allRows = Import-Csv -Path $CsvPath
Write-Log "Total rows in export: $($allRows.Count)"

$missingLocation = $allRows | Where-Object {
    [string]::IsNullOrWhiteSpace($_.Location)
}
Write-Log "Rows with missing Location: $($missingLocation.Count)"

$noSsoId = @($missingLocation | Where-Object { [string]::IsNullOrWhiteSpace($_.'SSO ID') })
$hasSsoId = @($missingLocation | Where-Object { -not [string]::IsNullOrWhiteSpace($_.'SSO ID') })

Write-Log "  Without SSO ID (Group G — shared mailboxes / bots / local-tenant): $($noSsoId.Count)"
Write-Log "  With SSO ID to investigate: $($hasSsoId.Count)"

#endregion

#region ── Pre-build Group G (no SSO ID) ────────────────────────────────────────

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($row in $noSsoId) {
    $results.Add([PSCustomObject]@{
            ObjectId            = ''
            UPN                 = $row.Email
            DisplayName         = "$($row.'First Name') $($row.'Last Name')".Trim()
            ProofpointCountry   = $row.Country
            ProofpointCompany   = $row.Company
            ProofpointAdGroup   = $row.adGroup
            EntraUsageLocation  = ''
            EntraUserType       = ''
            EntraAccountEnabled = ''
            EntraLicensed       = ''
            MailboxPurpose      = ''
            Category            = 'G'
            CategoryDescription = 'No SSO ID — not a synced Entra user (shared mailbox, bot, or local-tenant object)'
            RecommendedAction   = 'Verify in EXO. If shared mailbox or DL, no action. If real user, check source tenant.'
        })
}
Write-Log "Group G pre-populated ($($noSsoId.Count) rows)."

#endregion

#region ── Query Graph for SSO ID group ─────────────────────────────────────────

$total = $hasSsoId.Count
$processed = 0

Write-Log "Starting Graph lookup for $total users (batch size: $BatchSize)..."

for ($i = 0; $i -lt $total; $i += $BatchSize) {

    $batch = $hasSsoId[$i .. [Math]::Min($i + $BatchSize - 1, $total - 1)]

    foreach ($row in $batch) {

        $objectId = $row.'SSO ID'.Trim()
        $upn = $row.Email.Trim()
        $processed++

        Write-Progress -Activity "Querying Microsoft Graph" `
            -Status   "$processed / $total : $upn" `
            -PercentComplete (($processed / $total) * 100)

        # ── Entra user lookup ─────────────────────────────────────────────────
        $graphUser = $null
        try {
            $graphUser = Get-MgUser -UserId $objectId `
                -Property Id, UserPrincipalName, DisplayName, UsageLocation,
            UserType, AccountEnabled, AssignedLicenses `
                -ErrorAction Stop
        } catch {
            $results.Add([PSCustomObject]@{
                    ObjectId            = $objectId
                    UPN                 = $upn
                    DisplayName         = "$($row.'First Name') $($row.'Last Name')".Trim()
                    ProofpointCountry   = $row.Country
                    ProofpointCompany   = $row.Company
                    ProofpointAdGroup   = $row.adGroup
                    EntraUsageLocation  = ''
                    EntraUserType       = ''
                    EntraAccountEnabled = ''
                    EntraLicensed       = ''
                    MailboxPurpose      = ''
                    Category            = 'X'
                    CategoryDescription = 'Object not found in Entra (stale Proofpoint record)'
                    RecommendedAction   = 'Remove from Proofpoint or verify ObjectId. Object may have been deleted.'
                })
            Write-Log "  [X] $upn — not found in Entra." 'WARN'
            continue
        }

        $usageLocation = $graphUser.UsageLocation
        $userType = $graphUser.UserType        # Member | Guest
        $accountEnabled = $graphUser.AccountEnabled
        $isLicensed = ($graphUser.AssignedLicenses.Count -gt 0)

        # ── Mailbox purpose via Graph ─────────────────────────────────────────
        $mailboxPurpose = Get-MailboxPurpose -ObjectId $objectId
        # shared | room | equipment | user | linked | NoMailbox | Unknown

        $isNonUserMailbox = $mailboxPurpose -in @('shared', 'room', 'equipment', 'linked')

        # ── Categorise ────────────────────────────────────────────────────────
        $category = $description = $action = ''

        if (-not [string]::IsNullOrWhiteSpace($usageLocation)) {
            $category = 'A'
            $description = "UsageLocation set in Entra ($usageLocation) — Proofpoint sync lag"
            $action = 'No action required. Will resolve on next Proofpoint directory sync.'
        } elseif ($userType -eq 'Guest') {
            $category = 'F'
            $description = 'Guest / B2B user — UsageLocation not required'
            $action = 'No action required.'
        } elseif ($isNonUserMailbox) {
            $category = 'B'
            $description = "Non-user mailbox ($mailboxPurpose) — UsageLocation not required"
            $action = 'No action required. Confirm no active license is assigned.'
        } elseif (-not $accountEnabled) {
            $category = 'E'
            $description = 'Disabled Entra account — UsageLocation missing'
            $action = 'Set UsageLocation before re-enabling. Review for deletion if stale.'
        } elseif ($isLicensed) {
            $category = 'C'
            $description = "Licensed user, UsageLocation missing (licenses: $($graphUser.AssignedLicenses.Count))"
            $action = 'MUST FIX: Set UsageLocation immediately. MS requires this for all licensed users.'
        } else {
            $category = 'D'
            $description = "Unlicensed Member — UsageLocation missing (mailbox: $mailboxPurpose)"
            $action = 'Should fix: stamp UsageLocation before assigning any license.'
        }

        $results.Add([PSCustomObject]@{
                ObjectId            = $objectId
                UPN                 = $graphUser.UserPrincipalName
                DisplayName         = $graphUser.DisplayName
                ProofpointCountry   = $row.Country
                ProofpointCompany   = $row.Company
                ProofpointAdGroup   = $row.adGroup
                EntraUsageLocation  = $usageLocation
                EntraUserType       = $userType
                EntraAccountEnabled = $accountEnabled
                EntraLicensed       = $isLicensed
                MailboxPurpose      = $mailboxPurpose
                Category            = $category
                CategoryDescription = $description
                RecommendedAction   = $action
            })

        if ($category -eq 'C') {
            Write-Log "  [C] $upn — LICENSED, UsageLocation missing." 'WARN'
        }
    }

    if ($i + $BatchSize -lt $total) { Start-Sleep -Milliseconds 200 }
}

Write-Progress -Activity "Querying Microsoft Graph" -Completed

#endregion

#region ── Export ────────────────────────────────────────────────────────────────

$results | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8
Write-Log "Report saved: $reportPath" 'OK'

#endregion

#region ── Summary ──────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  ANALYSIS SUMMARY" -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Cyan

$categoryLabels = @{
    A = 'UsageLocation set in Entra (Proofpoint sync lag)'
    B = 'Non-user mailbox (shared/room/equipment)'
    C = 'Licensed user — UsageLocation MISSING  ◄ FIX'
    D = 'Unlicensed Member — UsageLocation missing'
    E = 'Disabled account — UsageLocation missing'
    F = 'Guest / B2B user'
    G = 'No SSO ID (not a synced Entra user)'
    X = 'Not found in Entra (stale Proofpoint record)'
}

$results | Group-Object Category | Sort-Object Name | ForEach-Object {
    $label = $categoryLabels[$_.Name]
    $color = if ($_.Name -eq 'C') { 'Red' }
    elseif ($_.Name -in 'A', 'B', 'F', 'G') { 'Green' }
    else { 'Yellow' }
    Write-Host ("  [{0}] {1,-52}  {2,5}" -f $_.Name, $label, $_.Count) -ForegroundColor $color
}

$catC = @($results | Where-Object { $_.Category -eq 'C' })
if ($catC.Count -gt 0) {
    Write-Host ""
    Write-Host "  Category C breakdown by email domain:" -ForegroundColor Red
    $catC | Group-Object { ($_.UPN -split '@')[1] } | Sort-Object Count -Descending |
    ForEach-Object {
        Write-Host ("    {0,-42}  {1}" -f $_.Name, $_.Count) -ForegroundColor Red
    }
}

Write-Host ""
Write-Host ("  Total analysed : {0}" -f $results.Count)
Write-Host ("  Elapsed        : {0}" -f ((Get-Date) - $scriptStart).ToString('hh\:mm\:ss'))
Write-Host ("  Report         : {0}" -f $reportPath)
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Cyan

#endregion

#region ── Optional: bulk remediation for Category C ────────────────────────────
<#
    WRITES TO THE DIRECTORY. Run separately, and only after reviewing the CSV report.
    Requires the User.ReadWrite.All scope, which the read-only run above does not request.

    The domain → ISO UsageLocation mapping lives in the shared data file pointed at by
    -DomainMapPath (Platforms\_Shared\Data\domain-country-map.psd1). Central domains are
    absent from it on purpose: those accounts are multi-country and must be resolved by
    hand, so they are skipped with a warning instead of being stamped with a wrong country.

    $domainMap  = (Import-PowerShellDataFile -Path $DomainMapPath).DomainToCountry
    $catCUsers  = Import-Csv -Path $reportPath | Where-Object { $_.Category -eq 'C' }

    foreach ($user in $catCUsers) {
        $domain = ($user.UPN -split '@')[1]
        $loc    = $domainMap[$domain]
        if ($loc) {
            Update-MgUser -UserId $user.ObjectId -UsageLocation $loc
            Write-Host "  SET $loc → $($user.UPN)"
        } else {
            Write-Warning "No mapping for domain '$domain' — skipping $($user.UPN)"
        }
    }
#>
#endregion