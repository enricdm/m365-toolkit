#
#    New-SharedMailbox.ps1 — Shared mailbox creator with member access provisioning
#    ---------------------------------------------------------------------------------
#    Creates a shared mailbox and grants FullAccess + SendAs to every member
#    listed in a CSV.
#    External members (non-tenant addresses) are skipped for mailbox permissions
#    with a clear warning — they cannot be granted Exchange mailbox access directly.
#
#    CSV format: one column, header 'Email' (or 'Member')
#        Email
#        user@contoso.com
#        ...
#
#    Dry run by default. Nothing is written until you pass -Execute.
#
#    Prerequisites:
#        Connect-ExchangeOnline
#

<#
.NOTES
    When to use  : Creating a shared mailbox that more than three or four people need access to.
    Why it exists: Creates the mailbox and grants FullAccess and SendAs to every member of a CSV in one pass, skipping external addresses with a clear warning because they cannot be granted mailbox access. Dry run by default.
#>

param(
    [Parameter(Mandatory)][string]$MailboxName,        # Name / sAMAccountName-style identifier
    [string]$DisplayName,                              # defaults to -MailboxName
    [Parameter(Mandatory)][string]$PrimarySmtp,        # e.g. team.inbox@contoso.com

    # AutoMapping: $true  = mailbox appears automatically in Outlook
    #              $false = user must add it manually
    [bool]$AutoMapping = $true,

    [string]$CsvPath = ".\members.csv",                # column: 'Email' or 'Member'

    [switch]$Execute                                   # omit = preview only (no changes)
)

#region ── CONFIG ─────────────────────────────────────────────────────────────
# PowerShell variable names are case-insensitive, so the parameters above are
# used directly below under their lower-case spellings.
if (-not $DisplayName) { $DisplayName = $MailboxName }
$dryRun = -not $Execute   # preview unless -Execute was passed
#endregion

#region ── LOAD MEMBERS CSV ───────────────────────────────────────────────────
if (-not (Test-Path $csvPath)) {
    Write-Host "ERROR: CSV not found at '$csvPath'" -ForegroundColor Red
    return 1
}

$csv = Import-Csv -Path $csvPath
$colName = if ($csv[0].PSObject.Properties.Name -contains 'Email') { 'Email' } else { 'Member' }
$members = $csv.$colName |
Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
Sort-Object -Unique

Write-Host "Loaded $($members.Count) unique addresses (column: '$colName')"
#endregion

#region ── HELPERS ────────────────────────────────────────────────────────────
# Cache accepted domains once
$acceptedDomains = (Get-AcceptedDomain).DomainName

function IsInternalAddress ([string]$Email) {
    $domain = ($Email -split '@')[-1]
    $domain -in $script:acceptedDomains
}
#endregion

#region ── CREATE SHARED MAILBOX ──────────────────────────────────────────────
Write-Host "`nCreating shared mailbox '$displayName' <$primarySmtp>..."

if ($dryRun) {
    Write-Host "[DRY-RUN] Would create shared mailbox: $displayName <$primarySmtp>" -ForegroundColor Cyan
} else {
    # Check if it already exists
    $existing = Get-Mailbox -Identity $primarySmtp -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "Mailbox already exists — skipping creation." -ForegroundColor DarkYellow
    } else {
        New-Mailbox `
            -Shared `
            -Name               $mailboxName `
            -DisplayName        $displayName `
            -PrimarySmtpAddress $primarySmtp `
            -ErrorAction Stop

        Write-Host "Shared mailbox created." -ForegroundColor Green
        Start-Sleep -Seconds 10   # allow replication before permissions
    }
}
#endregion

#region ── GRANT PERMISSIONS ──────────────────────────────────────────────────
Write-Host "`nGranting permissions to members..."

$skipped = @()
$success = @()
$failures = @()

foreach ($email in $members) {

    # ── External addresses: cannot hold Exchange mailbox permissions ──────────
    if (-not (IsInternalAddress $email)) {
        Write-Host "  SKIPPED (external): $email" -ForegroundColor DarkYellow
        $skipped += $email
        continue
    }

    if ($dryRun) {
        Write-Host "  [DRY-RUN] Would grant FullAccess + SendAs: $email" -ForegroundColor Cyan
        continue
    }

    # ── FullAccess ────────────────────────────────────────────────────────────
    try {
        Add-MailboxPermission `
            -Identity    $primarySmtp `
            -User        $email `
            -AccessRights FullAccess `
            -AutoMapping $autoMapping `
            -ErrorAction Stop | Out-Null
        Write-Host "  FullAccess granted : $email" -ForegroundColor Green
    } catch {
        Write-Host "  FAILED FullAccess  : $email  ->  $($_.Exception.Message)" -ForegroundColor Red
        $failures += "FullAccess::$email"
    }

    # ── SendAs ────────────────────────────────────────────────────────────────
    try {
        Add-RecipientPermission `
            -Identity    $primarySmtp `
            -Trustee     $email `
            -AccessRights SendAs `
            -Confirm:$false `
            -ErrorAction Stop | Out-Null
        Write-Host "  SendAs granted     : $email" -ForegroundColor Green
        $success += $email
    } catch {
        Write-Host "  FAILED SendAs      : $email  ->  $($_.Exception.Message)" -ForegroundColor Red
        $failures += "SendAs::$email"
    }
}
#endregion

#region ── SUMMARY ────────────────────────────────────────────────────────────
Write-Host "`n── Summary ───────────────────────────────────────────────" -ForegroundColor Cyan

if ($skipped.Count -gt 0) {
    Write-Host "Skipped (external — cannot grant mailbox permissions):" -ForegroundColor DarkYellow
    $skipped | ForEach-Object { Write-Host "  • $_" }
    Write-Host "  → These users must access the mailbox via a different method" -ForegroundColor DarkYellow
    Write-Host "    (e.g. shared credentials, forwarding rule, or guest account)" -ForegroundColor DarkYellow
}

if ($failures.Count -gt 0) {
    Write-Host "Failures:" -ForegroundColor Red
    $failures | ForEach-Object { Write-Host "  • $_" }
}

Write-Host "Provisioned successfully: $($success.Count) user(s)" -ForegroundColor Green
#endregion

#region ── VERIFY ─────────────────────────────────────────────────────────────
if (-not $dryRun) {
    Write-Host "`n── Verification ──────────────────────────────────────────" -ForegroundColor Cyan

    Get-Mailbox $primarySmtp | Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails

    Write-Host "`nFullAccess permissions:"
    Get-MailboxPermission -Identity $primarySmtp | Where-Object { $_.AccessRights -like "*FullAccess*" -and $_.User -notlike "NT AUTHORITY\*" } | Select-Object User, AccessRights, Deny | Sort-Object User

    Write-Host "`nSendAs permissions:"
    Get-RecipientPermission -Identity $primarySmtp | Where-Object { $_.Trustee -notlike "NT AUTHORITY\*" } | Select-Object Trustee, AccessRights | Sort-Object Trustee
}
#endregion
