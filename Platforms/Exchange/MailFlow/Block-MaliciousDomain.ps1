#requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Blocks a malicious domain in the Exchange Online Tenant Allow/Block List (TABL)
    as both a Sender domain and a URL entry.

.DESCRIPTION
    Adds block entries to the Tenant Allow/Block List for a phishing/malware domain.
    Checks for existing entries before writing (no duplicates), verifies after, and
    logs results to Exports/. Defaults to a DRY RUN; pass -Execute to apply changes.

    Blocking as BOTH Sender and Url matters: a Sender block stops mail claiming to
    come from the domain, but does nothing about a link to it inside a message that
    arrives from somewhere else. Most phishing needs both entries.

.PARAMETER Domain
    Domain to block, without protocol. Mandatory - there is deliberately no default,
    because a default here would mean shipping someone's indicator of compromise
    inside the script.

.PARAMETER ListType
    Which TABL list(s) to write. Default: both Sender and Url.

.PARAMETER Notes
    Free-text note stored on each TABL entry. Use it for the change reference.

.PARAMETER AdminUpn
    UPN to connect with. Optional - omit it to let Connect-ExchangeOnline prompt,
    or to reuse an existing session.

.PARAMETER Execute
    Apply changes. Without this switch the script runs in dry-run mode.

.EXAMPLE
    .\Block-MaliciousDomain.ps1 -Domain 'malicious.example'
    Dry run (shows what would be blocked).

.EXAMPLE
    .\Block-MaliciousDomain.ps1 -Domain 'malicious.example' -ListType Url `
        -Notes 'Phishing campaign, ref 12345' -Execute
    Creates the Url block entry only.

.NOTES
    When to use  : An indicator of compromise arrives at three in the morning and the domain has to be blocked with a record of what was done, without duplicating an entry someone else already made.
    Why it exists: Blocks as both Sender and Url: a sender block does nothing about a link to the domain inside a message arriving from somewhere else, and most phishing needs both entries. It also inspects Get-ConnectionInformation properly, which does not throw when there is no session - it returns nothing.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^[A-Za-z0-9]([A-Za-z0-9\-\.]*[A-Za-z0-9])?\.[A-Za-z]{2,}$')]
    [string]   $Domain,
    [ValidateSet('Sender', 'Url')]
    [string[]] $ListType = @('Sender', 'Url'),
    [string]   $Notes = 'Phishing/malware domain - blocked by IT',
    [string]   $AdminUpn = '',
    [switch]   $Execute
)

# --- Helpers -------------------------------------------------------------
function Write-Step { param([string]$m) Write-Host "[*] $m" -ForegroundColor Cyan }
function Write-OK { param([string]$m) Write-Host "[+] $m" -ForegroundColor Green }
function Write-Warn { param([string]$m) Write-Host "[!] $m" -ForegroundColor Yellow }
function Die { param([string]$m) Write-Host "[x] $m" -ForegroundColor Red; exit 1 }

# --- Setup ---------------------------------------------------------------
$ExportDir = Join-Path $PSScriptRoot 'Exports'
if (-not (Test-Path $ExportDir)) { New-Item -ItemType Directory -Path $ExportDir | Out-Null }
$Stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$LogPath = Join-Path $ExportDir ("TABL-Block_{0}_{1}.csv" -f ($Domain -replace '\.', '_'), $Stamp)

if (-not $Execute) { Write-Warn 'DRY RUN - no changes will be made. Re-run with -Execute to apply.' }

# --- Connect -------------------------------------------------------------
# Get-ConnectionInformation does NOT throw when there is no session - it returns
# nothing. Testing it with try/catch therefore always lands in the "reuse" branch
# and the script announces a session it does not have, only to fail later on the
# first real cmdlet. Inspect the returned object instead.
Write-Step 'Checking for an existing Exchange Online session'
$existingSession = @(Get-ConnectionInformation -ErrorAction SilentlyContinue |
    Where-Object { $_.State -eq 'Connected' -and $_.TokenStatus -ne 'Expired' })

if ($existingSession.Count -gt 0) {
    Write-OK "Reusing existing Exchange Online session ($($existingSession[0].UserPrincipalName))."
} else {
    Write-Step "Connecting to Exchange Online$(if ($AdminUpn) { " as $AdminUpn" })"
    try {
        $connectArgs = @{ ShowBanner = $false; ErrorAction = 'Stop' }
        if ($AdminUpn) { $connectArgs['UserPrincipalName'] = $AdminUpn }
        Connect-ExchangeOnline @connectArgs
        Write-OK 'Connected to Exchange Online.'
    } catch {
        Die "Failed to connect to Exchange Online: $($_.Exception.Message)"
    }
}

# --- Block ---------------------------------------------------------------
$Results = @()

foreach ($lt in $ListType) {
    Write-Step "Processing $lt block for '$Domain'"

    # Check for an existing block so we don't duplicate
    # Both a SilentlyContinue and an empty catch used to sit on this one call, so a failed
    # check looked identical to "no entry exists" and the script went on to add a duplicate.
    # Adding one is harmless; being unable to tell is not, so it says which happened.
    $existing = $null
    try   { $existing = Get-TenantAllowBlockListItems -ListType $lt -Block -Entry $Domain -ErrorAction Stop }
    catch {
        if ($_.Exception.Message -match 'not\s+found|no\s+entries|NoMatch') { $existing = $null }
        else { Write-Warn "$lt duplicate check failed ($($_.Exception.Message)). Continuing; a duplicate entry is possible." }
    }

    if ($existing) {
        Write-Warn "$lt entry already exists (Id: $($existing.Identity)). Skipping."
        $Results += [pscustomobject]@{
            Domain = $Domain; ListType = $lt; Action = 'AlreadyBlocked'
            Identity = $existing.Identity; Notes = $existing.Notes
        }
        continue
    }

    if ($Execute) {
        try {
            $new = New-TenantAllowBlockListItems -ListType $lt -Block -Entries $Domain `
                -NoExpiration -Notes $Notes -ErrorAction Stop
            Write-OK "$lt block created (Id: $($new.Identity))."
            $Results += [pscustomobject]@{
                Domain = $Domain; ListType = $lt; Action = 'Blocked'
                Identity = $new.Identity; Notes = $Notes
            }
        } catch {
            Write-Warn "Failed to block $lt '$Domain': $($_.Exception.Message)"
            $Results += [pscustomobject]@{
                Domain = $Domain; ListType = $lt; Action = "Error: $($_.Exception.Message)"
                Identity = ''; Notes = $Notes
            }
        }
    } else {
        Write-Warn "[DRY RUN] Would create $lt block for '$Domain' (NoExpiration)."
        $Results += [pscustomobject]@{
            Domain = $Domain; ListType = $lt; Action = 'WouldBlock (dry run)'
            Identity = ''; Notes = $Notes
        }
    }
}

# --- Verify --------------------------------------------------------------
if ($Execute) {
    Write-Step 'Verifying block entries'
    foreach ($lt in $ListType) {
        # Three outcomes, not two. A verification that could not run is not the same as a
        # block that is absent, and reporting the first as the second either sends someone
        # to re-apply a block that is already there, or - worse, if read the other way -
        # gets treated as confirmation that nothing was applied.
        $v = $null; $verifyError = $null
        try   { $v = Get-TenantAllowBlockListItems -ListType $lt -Block -Entry $Domain -ErrorAction Stop }
        catch { $verifyError = $_.Exception.Message }

        if     ($verifyError) { Write-Warn "$lt verification could not run for '$Domain': $verifyError. The block may or may not be in place - check it in the portal." }
        elseif ($v)           { Write-OK   "$lt block present for '$Domain' (Id: $($v.Identity))." }
        else                  { Write-Warn "$lt block NOT present for '$Domain' - the write did not take effect." }
    }
}

# --- Export --------------------------------------------------------------
$Results | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8
Write-OK "Log written to $LogPath"
Write-Step 'Done.'
