<#
.SYNOPSIS
    Configures server-side forwarding on one or more Exchange Online mailboxes.

.DESCRIPTION
    Auto-forward all incoming mail for a set of source mailboxes to a single
    target address. Dry-run by default: reports current vs. intended state and
    changes nothing until you pass -Execute.

    ############################################################################
    #  READ THIS BEFORE USING IT                                               #
    #                                                                          #
    #  Server-side forwarding to an external address is the exact mechanism    #
    #  attackers configure after a Business Email Compromise, and it is the    #
    #  first thing an incident responder looks for. Setting it deliberately    #
    #  means you are creating the artefact that a BEC investigation hunts.     #
    #                                                                          #
    #  Therefore:                                                              #
    #   - Have a written, approved request before you run this with -Execute.  #
    #   - Prefer an internal target. Many tenants block external auto-forward  #
    #     outright via an outbound spam policy, and for good reason.           #
    #   - Keep -KeepCopyInMailbox $true so mail is still discoverable in the   #
    #     source mailbox; forward-only leaves no local record.                 #
    #   - Keep the exported CSV. It is your evidence that the forward was      #
    #     configured by IT on request, not by an intruder.                     #
    #   - Remove the forward when the reason for it ends.                      #
    #                                                                          #
    #  Guard rail: if the target domain is NOT an accepted domain of the       #
    #  tenant, the script refuses to apply anything unless you also pass       #
    #  -AllowExternalTarget. That switch is the "yes, I mean it" confirmation. #
    ############################################################################

    Forwarding is applied with -ForwardingSMTPAddress (works for internal or
    external targets, no recipient object required). If you prefer the AD-level
    forward that shows in EAC, swap to -ForwardingAddress below.

.NOTES
    When to use  : A request arrives to forward the mailbox of someone on long-term leave, and you want the paper trail that says IT configured it on request.
    Why it exists: Server-side forwarding to an external address is the exact mechanism an attacker configures after a business email compromise, and the first thing an incident responder looks for. If the target domain is not an accepted domain the script refuses to apply anything without -AllowExternalTarget, and the exported CSV is the evidence.
    Connect first:  Connect-ExchangeOnline
    The account needs Exchange Recipient Administrator or equivalent.

.EXAMPLE
    # Preview only - internal target
    .\Set-MailboxForwarding.ps1 -SourceMailboxes 'user.one@contoso.com','user.two@contoso.com' `
        -ForwardTo 'shared.inbox@contoso.com'

.EXAMPLE
    # Apply to an external target - requires the explicit confirmation switch
    .\Set-MailboxForwarding.ps1 -SourceMailboxes 'user.one@contoso.com' `
        -ForwardTo 'someone@fabrikam.com' -AllowExternalTarget -Execute
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string[]]$SourceMailboxes,
    [Parameter(Mandatory)][string]$ForwardTo,

    # $true  = deliver to source AND forward (recommended - keeps mail discoverable)
    # $false = forward only, no local copy retained
    [bool]$KeepCopyInMailbox = $true,

    # Required to apply forwarding to a domain that is not an accepted domain
    # of this tenant. See the BEC warning above.
    [switch]$AllowExternalTarget,

    [string]$ExportDir = (Join-Path $PSScriptRoot 'Exports'),
    [switch]$Execute
)

# ---- Helpers ----
function Write-Step($m) { Write-Host "[*] $m" -ForegroundColor Cyan }
function Write-OK  ($m) { Write-Host "[OK] $m" -ForegroundColor Green }
function Write-Warn($m) { Write-Host "[!] $m" -ForegroundColor Yellow }
function Die       ($m) { Write-Host "[X] $m" -ForegroundColor Red; exit 1 }

# ---- Preconditions ----
if (-not (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
    Die "ExchangeOnline session not found. Run Connect-ExchangeOnline first."
}
if (-not (Test-Path $ExportDir)) { New-Item -ItemType Directory -Path $ExportDir | Out-Null }

$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outFile = Join-Path $ExportDir "MailboxForwarding_$stamp.csv"
$mode = if ($Execute) { 'EXECUTE' } else { 'DRY-RUN' }

Write-Step "Mode: $mode  |  Forward target: $ForwardTo  |  Keep local copy: $KeepCopyInMailbox"

# ---- Is the target inside the tenant? ----
# Fail safe: if the accepted-domain list cannot be read, treat the target as
# external rather than assuming it is internal.
$targetDomain = ($ForwardTo -split '@')[-1]
$acceptedDomains = @()
try { $acceptedDomains = @((Get-AcceptedDomain -ErrorAction Stop).DomainName) }
catch { Write-Warn "Could not enumerate accepted domains: $($_.Exception.Message)" }

$isExternalTarget = -not ($acceptedDomains -contains $targetDomain)

if ($isExternalTarget) {
    Write-Host ""
    Write-Warn "EXTERNAL FORWARDING TARGET: '$ForwardTo' ($targetDomain) is not an accepted domain of this tenant."
    Write-Warn "External auto-forward is the classic Business Email Compromise artefact and is"
    Write-Warn "what incident response hunts for. Only proceed against an approved written request."
    if ($KeepCopyInMailbox -ne $true) {
        Write-Warn "-KeepCopyInMailbox is `$false: mail will leave the tenant with NO local copy retained."
    }
    if ($Execute -and -not $AllowExternalTarget) {
        Die "Refusing to forward outside the tenant without -AllowExternalTarget. Re-run with that switch if this is intended and approved."
    }
    if (-not $Execute) {
        Write-Warn "[dry-run] -Execute would additionally require -AllowExternalTarget for this target."
    } else {
        Write-Warn "-AllowExternalTarget was supplied - proceeding with an external forward."
    }
    Write-Host ""
} else {
    Write-OK "Target '$ForwardTo' is inside an accepted domain of this tenant."
}

# ---- Validate target resolves (warn only; ForwardingSMTPAddress doesn't require it) ----
$targetObj = Get-Recipient -Identity $ForwardTo -ErrorAction SilentlyContinue
if ($targetObj) { Write-OK  "Target resolves: $($targetObj.RecipientType) - $($targetObj.PrimarySmtpAddress)" }
else { Write-Warn "Target '$ForwardTo' did not resolve to a recipient object. Forwarding will still be set via SMTP address." }

# ---- Process mailboxes ----
$results = foreach ($id in $SourceMailboxes) {
    Write-Step "Processing $id"
    $mbx = Get-Mailbox -Identity $id -ErrorAction SilentlyContinue
    if (-not $mbx) {
        Write-Warn "  Mailbox not found - skipping."
        [pscustomobject]@{
            Mailbox = $id; Found = $false; PreviousForward = ''; PreviousDeliverAndForward = '';
            NewForward = $ForwardTo; DeliverAndForward = $KeepCopyInMailbox; ExternalTarget = $isExternalTarget;
            Action = 'SKIPPED (not found)'; Applied = $false
        }
        continue
    }

    $prevFwd = $mbx.ForwardingSMTPAddress
    $prevDeliver = $mbx.DeliverToMailboxAndForward
    Write-Host  "  Current ForwardingSMTPAddress : $prevFwd"
    Write-Host  "  Current DeliverToMailboxAndForward : $prevDeliver"

    $applied = $false
    if ($Execute) {
        try {
            Set-Mailbox -Identity $id -ForwardingSMTPAddress $ForwardTo -DeliverToMailboxAndForward $KeepCopyInMailbox -ErrorAction Stop
            $verify = Get-Mailbox -Identity $id
            if ($verify.ForwardingSMTPAddress -match [regex]::Escape($ForwardTo)) {
                Write-OK "  Forwarding set -> $ForwardTo"
                $applied = $true
            } else {
                Write-Warn "  Set-Mailbox ran but verification did not match expected target."
            }
        } catch {
            Write-Warn "  Failed to set forwarding: $($_.Exception.Message)"
        }
    } else {
        Write-Host "  [dry-run] Would set ForwardingSMTPAddress=$ForwardTo, DeliverToMailboxAndForward=$KeepCopyInMailbox" -ForegroundColor DarkGray
    }

    [pscustomobject]@{
        Mailbox                   = $id
        Found                     = $true
        PreviousForward           = $prevFwd
        PreviousDeliverAndForward = $prevDeliver
        NewForward                = $ForwardTo
        DeliverAndForward         = $KeepCopyInMailbox
        ExternalTarget            = $isExternalTarget
        Action                    = if ($Execute) { 'APPLIED' } else { 'DRY-RUN' }
        Applied                   = $applied
    }
}

# ---- Export + summary ----
$results | Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8
Write-OK "Results exported: $outFile"

$applied = ($results | Where-Object Applied).Count
$total = ($results | Where-Object Found).Count
Write-Step "Summary: $applied/$total mailbox(es) forwarded ($mode)."
if ($applied -gt 0 -and $isExternalTarget) {
    Write-Warn "External forwarding is now active. Keep $outFile as the record of who approved it, and remove the forward when it is no longer needed."
}
if (-not $Execute) { Write-Warn "No changes made. Re-run with -Execute to apply." }

# ---- To undo, per mailbox: ----
#   Set-Mailbox -Identity '<mailbox>' -ForwardingSMTPAddress $null -DeliverToMailboxAndForward $false
