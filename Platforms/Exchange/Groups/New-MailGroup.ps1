<#
    New-MailGroup.ps1 — Universal group creator (DL / Mail-Enabled SG / M365)
    -------------------------------------------------------------------------
    One script for the three group types, with the type chosen interactively at
    runtime:
      - External-member support: auto-creates MailContacts for non-tenant addresses
      - Sender-restriction enforcement (AcceptMessagesOnlyFromSendersOrMembers)
      - Flexible CSV: accepts 'Email' or 'Member' header

    CSV format: one column, header 'Email' (or 'Member')
        Email
        user@contoso.com
        ...

    Dry run by default. Nothing is written until you pass -Execute.

    Prerequisites:
        Connect-ExchangeOnline
        (M365 Group only: Connect-MgGraph -Scopes "Group.ReadWrite.All")

#>

param(
    [Parameter(Mandatory)][string]$GroupName,        # e.g. CC.CITY.Example
    [string]$DisplayName,                            # defaults to -GroupName
    [Parameter(Mandatory)][string]$PrimarySmtp,      # e.g. CC.CITY.Example@contoso.com
    [Parameter(Mandatory)][string]$Owner,            # ManagedBy / group owner

    # Only these addresses may send to the group (leave empty to allow anyone)
    [string[]]$AllowedSenders = @(),

    [string]$CsvPath = ".\members.csv",              # column: 'Email' or 'Member'

    [switch]$Execute                                 # omit = preview only (no changes)
)

#region ── CONFIG ─────────────────────────────────────────────────────────────
# PowerShell variable names are case-insensitive, so the parameters above are
# used directly below under their lower-case spellings.
if (-not $DisplayName) { $DisplayName = $GroupName }
$dryRun = -not $Execute   # preview unless -Execute was passed
#endregion

#region ── GROUP TYPE SELECTION ───────────────────────────────────────────────
Write-Host ""
Write-Host "Select group type:" -ForegroundColor Cyan
Write-Host "  [1] Distribution List        (email routing only; supports external members)"
Write-Host "  [2] Mail-Enabled Sec. Group  (email + Entra access control)"
Write-Host "  [3] Microsoft 365 Group      (Teams / SharePoint / Outlook; internal members)"
Write-Host ""

do { $choice = Read-Host "Enter 1, 2 or 3" } while ($choice -notin @('1', '2', '3'))

$groupType = switch ($choice) {
    '1' { 'Distribution' }
    '2' { 'Security' }
    '3' { 'M365' }
}
Write-Host "→ Type selected: $groupType" -ForegroundColor Yellow
#endregion

#region ── LOAD MEMBERS CSV ───────────────────────────────────────────────────
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

function Ensure-MailContact ([string]$Email) {
    # 1. Existing MailContact?
    $existing = Get-MailContact -Filter "ExternalEmailAddress -eq 'smtp:$Email'" `
        -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  Contact exists : $Email" -ForegroundColor DarkGray
        return $existing.Identity
    }

    # 2. Already in tenant as a Guest / MailUser? (avoids proxy conflict)
    $existing = Get-Recipient -Filter "EmailAddresses -eq 'smtp:$Email'" `
        -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  Guest/MailUser exists: $Email ($($existing.DisplayName))" -ForegroundColor DarkYellow
        return $existing.Identity
    }

    # 3. Nothing found → create new contact
    if ($script:dryRun) {
        Write-Host "  [DRY-RUN] Would create MailContact: $Email" -ForegroundColor Cyan
        return $Email
    }
    $alias = ($Email -replace '[^a-zA-Z0-9]', '-') -replace '-+', '-'
    $contact = New-MailContact -Name $Email -ExternalEmailAddress $Email `
        -Alias $alias -ErrorAction Stop
    Write-Host "  Created contact: $Email" -ForegroundColor Green
    return $contact.Identity
}
#endregion

#region ── CREATE GROUP ───────────────────────────────────────────────────────
Write-Host "`nCreating group '$groupName'..."
if ($dryRun) {
    Write-Host "[DRY-RUN] Would create $groupType group: $displayName <$primarySmtp>" -ForegroundColor Cyan
} else {
    switch ($groupType) {
        'Distribution' {
            New-DistributionGroup `
                -Name               $groupName `
                -DisplayName        $displayName `
                -PrimarySmtpAddress $primarySmtp `
                -Type               Distribution `
                -ManagedBy          $owner
        }
        'Security' {
            New-DistributionGroup `
                -Name               $groupName `
                -DisplayName        $displayName `
                -PrimarySmtpAddress $primarySmtp `
                -Type               Security `
                -ManagedBy          $owner
        }
        'M365' {
            $alias = $groupName -replace '[^a-zA-Z0-9\-]', ''
            New-UnifiedGroup `
                -DisplayName        $displayName `
                -Alias              $alias `
                -PrimarySmtpAddress $primarySmtp `
                -Owner              $owner `
                -AccessType         Private
        }
    }
    Write-Host "Group created." -ForegroundColor Green
    Start-Sleep -Seconds 10
}
#endregion

#region ── ADD MEMBERS ────────────────────────────────────────────────────────
Write-Host "`nAdding members..."

foreach ($email in $members) {
    if ($dryRun) {
        Write-Host "  [DRY-RUN] Would add: $email" -ForegroundColor Cyan
        continue
    }

    # External addresses need a MailContact object first (DL and SG only)
    $identity = $email
    if ($groupType -ne 'M365' -and -not (IsInternalAddress $email)) {
        try { $identity = Ensure-MailContact -Email $email }
        catch {
            Write-Host "FAILED (contact): $email  ->  $($_.Exception.Message)" -ForegroundColor Red
            continue
        }
    }

    try {
        switch ($groupType) {
            { $_ -in 'Distribution', 'Security' } {
                Add-DistributionGroupMember `
                    -Identity  $groupName `
                    -Member    $identity `
                    -BypassSecurityGroupManagerCheck `
                    -ErrorAction Stop
            }
            'M365' {
                Add-UnifiedGroupLinks `
                    -Identity  $groupName `
                    -LinkType  Members `
                    -Links     $identity `
                    -ErrorAction Stop
            }
        }
        Write-Host "Added:  $email" -ForegroundColor Green
    } catch {
        Write-Host "FAILED: $email  ->  $($_.Exception.Message)" -ForegroundColor Red
    }
}
#endregion

#region ── SENDER RESTRICTIONS ───────────────────────────────────────────────
if ($allowedSenders.Count -gt 0) {
    Write-Host "`nApplying sender restrictions: $($allowedSenders -join ', ')"
    if ($dryRun) {
        Write-Host "[DRY-RUN] Would restrict senders to the above." -ForegroundColor Cyan
    } else {
        switch ($groupType) {
            { $_ -in 'Distribution', 'Security' } {
                Set-DistributionGroup `
                    -Identity                              $groupName `
                    -AcceptMessagesOnlyFromSendersOrMembers $allowedSenders `
                    -RequireSenderAuthenticationEnabled    $false
                Write-Host "Sender restriction applied." -ForegroundColor Green
            }
            'M365' {
                # M365 Groups use Set-UnifiedGroup -AcceptMessagesOnlyFrom
                Set-UnifiedGroup `
                    -Identity            $groupName `
                    -AcceptMessagesOnlyFrom $allowedSenders
                Write-Host "Sender restriction applied." -ForegroundColor Green
            }
        }
    }
}
#endregion

#region ── VERIFY ─────────────────────────────────────────────────────────────
if (-not $dryRun) {
    Write-Host "`n── Verification ──────────────────────────────────────────" -ForegroundColor Cyan
    switch ($groupType) {
        { $_ -in 'Distribution', 'Security' } {
            Get-DistributionGroup $groupName |
            Select-Object Name, PrimarySmtpAddress, GroupType, ManagedBy,
            AcceptMessagesOnlyFromSendersOrMembers

            Write-Host "`nMembers ($( (Get-DistributionGroupMember $groupName).Count )):"
            Get-DistributionGroupMember $groupName |
            Select-Object DisplayName, PrimarySmtpAddress |
            Sort-Object DisplayName
        }
        'M365' {
            Get-UnifiedGroup $groupName |
            Select-Object DisplayName, PrimarySmtpAddress, AccessType,
            AcceptMessagesOnlyFrom

            Write-Host "`nMembers:"
            Get-UnifiedGroupLinks -Identity $groupName -LinkType Members |
            Select-Object DisplayName, PrimarySmtpAddress |
            Sort-Object DisplayName
        }
    }
}
#endregion
