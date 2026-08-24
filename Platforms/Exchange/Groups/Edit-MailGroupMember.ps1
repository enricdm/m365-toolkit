<#
.SYNOPSIS
    Adds and removes members on EXISTING mail groups, creating MailContacts for
    external addresses that do not exist yet.

.DESCRIPTION
    The groups already exist, so this script only manages membership
    (no group creation, no sender-restriction logic).

    External-member handling:
      - Resolves each ADD target via Get-Recipient first
      - If it already exists (MailContact / GuestMailUser / Mailbox) -> add as-is
      - If it does NOT exist and is external -> creates a MailContact, then adds
      - REMOVE of a non-member is treated as "already absent" (not an error),
        so re-running is safe and pre-done removals are fine.

    Group type is auto-detected per group:
      - DL / Mail-Enabled SG -> Add/Remove-DistributionGroupMember
      - M365 Group           -> Add/Remove-UnifiedGroupLinks

    Prereq:  Connect-ExchangeOnline

    Dry run by default. Nothing is written until you pass -Execute.
    Edit the $Jobs list below: one entry per group, with its Add / Remove sets.

.NOTES
    When to use  : A batch of adds and removals across several mail groups arrives and some of the addresses are external.
    Why it exists: Resolves every add through Get-Recipient first and creates the MailContact if the address is external and absent, detects the group type to pick the right cmdlet, and treats removing a non-member as 'already absent' rather than an error, so the run is repeatable.
#>

param(
    [switch]$Execute,                                  # omit = preview only; pass -Execute to apply
    [string[]]$InternalDomains = @('contoso.com'),     # never auto-create a MailContact for these
    [string]$LogDir = (Join-Path $PSScriptRoot 'logs')
)

#region ── CONFIG ─────────────────────────────────────────────────────────────
$DryRun = -not $Execute

# One entry per group. 'Add' addresses that do not exist in the tenant and are
# NOT in -InternalDomains are created as MailContacts first, then added.
$Jobs = @(
    @{
        Group     = 'example.group@contoso.com'
        Reference = 'change-request reference (free text, logged only)'
        Add       = @(
            'external.partner@fabrikam.com'
            'new.member@contoso.com'
        )
        Remove    = @(
            'former.member@contoso.com'
        )
    }
)
#endregion

#region ── HELPERS ────────────────────────────────────────────────────────────
function Write-Step($m) { Write-Host "`n=== $m ===" -ForegroundColor Cyan }
function Write-OK  ($m) { Write-Host "  [OK]   $m" -ForegroundColor Green }
function Write-Warn($m) { Write-Host "  [WARN] $m" -ForegroundColor Yellow }
function Write-Die ($m) { Write-Host "  [FAIL] $m" -ForegroundColor Red; throw $m }

function Get-DisplayNameFromAddress {
    param([string]$Address)
    $local = ($Address -split '@')[0]
    $parts = $local -split '[._-]' | Where-Object { $_ }
    $nice = ($parts | ForEach-Object { $_.Substring(0, 1).ToUpper() + $_.Substring(1) }) -join ' '
    if (-not $nice) { $nice = $Address }
    $dom = ($Address -split '@')[1]
    "$nice ($dom)"
}

function Test-Internal {
    param([string]$Address)
    $dom = ($Address -split '@')[1]
    $InternalDomains -contains $dom
}

# Returns the recipient object if any object in the tenant already represents
# this address (MailContact, Guest MailUser, Mailbox, etc.), else $null.
function Resolve-Recipient {
    param([string]$Address)
    try { return Get-Recipient -Identity $Address -ErrorAction Stop }
    catch { return $null }
}
#endregion

#region ── PRE-FLIGHT ─────────────────────────────────────────────────────────
Write-Step "Pre-flight"
try {
    $null = Get-ConnectionInformation -ErrorAction Stop
    Write-OK "Connected to Exchange Online"
} catch {
    Write-Die "Not connected. Run Connect-ExchangeOnline first."
}

if ($DryRun) { Write-Warn "DRY RUN — no changes will be made. Use -Execute to apply." }

if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir | Out-Null }
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$logFile = Join-Path $LogDir "Edit-MailGroupMember-$stamp.csv"
$log = New-Object System.Collections.Generic.List[object]

function Log {
    param($Group, $Action, $Member, $Result, $Detail)
    $log.Add([pscustomobject]@{
            Timestamp = (Get-Date -Format s)
            Group     = $Group
            Action    = $Action
            Member    = $Member
            Result    = $Result
            Detail    = $Detail
        })
}
#endregion

#region ── MAIN ───────────────────────────────────────────────────────────────
$added = 0; $removed = 0; $created = 0; $skipped = 0; $failed = 0

foreach ($job in $Jobs) {
    $gAddr = $job.Group
    Write-Step "Group: $gAddr   ($($job.Reference))"

    # --- resolve the group + type ---
    $grp = Resolve-Recipient $gAddr
    if (-not $grp) { Write-Warn "Group not found — skipping."; Log $gAddr 'GROUP' '-' 'ERROR' 'Group not found'; $failed++; continue }

    $type = $grp.RecipientTypeDetails
    $isM365 = ($type -eq 'GroupMailbox')
    Write-OK "Resolved as $type"

    # --- current members (for dry-run preview + membership checks) ---
    $memberAddrs = @{}
    try {
        $members = if ($isM365) { Get-UnifiedGroupLinks -Identity $gAddr -LinkType Members -ResultSize Unlimited }
        else { Get-DistributionGroupMember -Identity $gAddr -ResultSize Unlimited }
        foreach ($m in $members) {
            foreach ($a in @($m.PrimarySmtpAddress) + @($m.EmailAddresses)) {
                if ($a) { $memberAddrs[($a -replace '^smtp:', '').ToLower()] = $true }
            }
        }
        Write-OK "Current member count: $($members.Count)"
    } catch { Write-Warn "Could not enumerate members: $($_.Exception.Message)" }

    $isMember = { param($addr) $memberAddrs.ContainsKey($addr.ToLower()) }

    # ----------------------------- ADDS -----------------------------
    foreach ($addr in $job.Add) {
        if (& $isMember $addr) { Write-OK "ADD  $addr — already a member, skipping"; Log $gAddr 'ADD' $addr 'SKIP' 'Already member'; $skipped++; continue }

        $rcpt = Resolve-Recipient $addr

        # create contact if external + missing
        if (-not $rcpt) {
            if (Test-Internal $addr) {
                Write-Warn "ADD  $addr — internal address not found in tenant (typo?). Skipping."
                Log $gAddr 'ADD' $addr 'ERROR' 'Internal recipient not found'; $failed++; continue
            }
            $dn = Get-DisplayNameFromAddress $addr
            if ($DryRun) {
                Write-Warn "ADD  $addr — WOULD create MailContact '$dn' then add"
                Log $gAddr 'ADD' $addr 'DRYRUN' "Would create contact '$dn' + add"; continue
            }
            try {
                $rcpt = New-MailContact -Name $addr -DisplayName $dn -ExternalEmailAddress $addr -ErrorAction Stop
                Write-OK "Created MailContact '$dn'"; Log $gAddr 'CONTACT' $addr 'CREATED' $dn; $created++
            } catch {
                Write-Warn "ADD  $addr — contact creation failed: $($_.Exception.Message)"
                Log $gAddr 'CONTACT' $addr 'ERROR' $_.Exception.Message; $failed++; continue
            }
        }

        if ($DryRun) { Write-Warn "ADD  $addr — WOULD add (exists as $($rcpt.RecipientTypeDetails))"; Log $gAddr 'ADD' $addr 'DRYRUN' "Would add ($($rcpt.RecipientTypeDetails))"; continue }

        try {
            if ($isM365) { Add-UnifiedGroupLinks -Identity $gAddr -LinkType Members -Links $addr -ErrorAction Stop }
            else { Add-DistributionGroupMember -Identity $gAddr -Member $addr -BypassSecurityGroupManagerCheck -ErrorAction Stop }
            Write-OK "ADD  $addr"; Log $gAddr 'ADD' $addr 'ADDED' $rcpt.RecipientTypeDetails; $added++
        } catch {
            if ($_.Exception.Message -match 'already a member') { Write-OK "ADD  $addr — already a member"; Log $gAddr 'ADD' $addr 'SKIP' 'Already member'; $skipped++ }
            else { Write-Warn "ADD  $addr — failed: $($_.Exception.Message)"; Log $gAddr 'ADD' $addr 'ERROR' $_.Exception.Message; $failed++ }
        }
    }

    # ---------------------------- REMOVES ----------------------------
    foreach ($addr in $job.Remove) {
        if (-not (& $isMember $addr)) { Write-OK "RMV  $addr — already absent, skipping"; Log $gAddr 'REMOVE' $addr 'SKIP' 'Already absent'; $skipped++; continue }

        if ($DryRun) { Write-Warn "RMV  $addr — WOULD remove"; Log $gAddr 'REMOVE' $addr 'DRYRUN' 'Would remove'; continue }

        try {
            if ($isM365) { Remove-UnifiedGroupLinks -Identity $gAddr -LinkType Members -Links $addr -Confirm:$false -ErrorAction Stop }
            else { Remove-DistributionGroupMember -Identity $gAddr -Member $addr -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop }
            Write-OK "RMV  $addr"; Log $gAddr 'REMOVE' $addr 'REMOVED' '-'; $removed++
        } catch {
            if ($_.Exception.Message -match "isn't a member|not a member") { Write-OK "RMV  $addr — already absent"; Log $gAddr 'REMOVE' $addr 'SKIP' 'Already absent'; $skipped++ }
            else { Write-Warn "RMV  $addr — failed: $($_.Exception.Message)"; Log $gAddr 'REMOVE' $addr 'ERROR' $_.Exception.Message; $failed++ }
        }
    }
}
#endregion

#region ── SUMMARY ────────────────────────────────────────────────────────────
$log | Export-Csv -Path $logFile -NoTypeInformation -Encoding UTF8

Write-Step "Summary$(if($DryRun){' (DRY RUN — nothing changed)'})"
Write-Host ("  Added:            {0}" -f $added)
Write-Host ("  Removed:          {0}" -f $removed)
Write-Host ("  Contacts created: {0}" -f $created)
Write-Host ("  Skipped:          {0}" -f $skipped)
Write-Host ("  Failed:           {0}" -f $failed) -ForegroundColor $(if ($failed) { 'Red' }else { 'Gray' })
Write-Host ("  Log:              {0}" -f $logFile)
#endregion
