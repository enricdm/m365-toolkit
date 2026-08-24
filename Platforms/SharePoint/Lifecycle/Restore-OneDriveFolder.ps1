<#
.SYNOPSIS
    Restore a deleted folder subtree from a OneDrive/SPO first-stage recycle bin.
    Dry-run by default; pass -Execute to actually restore. Throttle-aware, idempotent.

.DESCRIPTION
    Recovery tool for a mass-delete: enumerates the target site's recycle bin, filters
    to one folder path, reports a summary (item count, size, date range, who deleted
    what), and only with -Execute restores each item, retrying with exponential back-off
    on throttling. Writes a timestamped CSV of the targeted set and a separate CSV of
    any failures.

    Restoring is idempotent, so a failed run can simply be re-run with -Execute.

.PARAMETER SiteUrl
    Target site or OneDrive URL. Mandatory: this script acts on someone else's
    content, so there is deliberately no default target.

.PARAMETER PathFilter
    Wildcard matched against each recycle bin item's DirName, e.g. '*Documents/Reports*'.

.PARAMETER Execute
    Perform the restore. Without it the script only reports.

.PARAMETER Interactive
    Sign in interactively instead of using app-only certificate auth.

.EXAMPLE
    # Report only (no changes) - confirm count, size, date range
    .\Restore-OneDriveFolder.ps1 -SiteUrl 'https://contoso-my.sharepoint.com/personal/user_contoso_com' `
        -PathFilter '*Documents/Reports*' -ClientId '<client-id>' -Tenant 'contoso.onmicrosoft.com'

.EXAMPLE
    # Perform the restore
    .\Restore-OneDriveFolder.ps1 -SiteUrl 'https://contoso-my.sharepoint.com/personal/user_contoso_com' `
        -PathFilter '*Documents/Reports*' -ClientId '<client-id>' -Tenant 'contoso.onmicrosoft.com' -Execute

.NOTES
    When to use  : Someone deletes an entire folder from their OneDrive or a library and tells you three days later.
    Why it exists: Reports item count, size, date range and who deleted what before touching anything, then restores with exponential back-off on throttling. Restoring is idempotent, so a failed run can simply be re-run. -SiteUrl is mandatory on purpose because this acts on someone else's content.
    Prereq: site-collection admin on the target OneDrive (Set-SPOUser ... -IsSiteCollectionAdmin $true).
    Auth: app-only cert, or -Interactive.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $SiteUrl,
    [Parameter(Mandatory)][string] $PathFilter,
    [switch] $Execute,
    [switch] $Interactive,
    [string] $ClientId,
    [string] $Tenant,
    [string] $CertSubject    = 'CN=PnP-SPO-Snapshot',
    [bool]   $FirstStageOnly = $true,   # ignore anything already in 2nd stage
    [int]    $RowLimit       = 250000,
    [int]    $MaxRetry       = 5,
    [int]    $BaseDelaySec   = 2,       # exponential back-off seed
    [string] $ExportDir      = (Join-Path $PSScriptRoot 'Exports'),
    [string] $Label          = 'RecycleBinRestore'   # used in output file names
)

$UseInteractive = [bool]$Interactive
if (-not $UseInteractive -and (-not $ClientId -or -not $Tenant)) {
    throw "App-only auth needs -ClientId and -Tenant. Pass them, or use -Interactive."
}

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------
function Write-Step($m) { Write-Host "==> $m" -ForegroundColor Cyan }
function Write-OK  ($m) { Write-Host "  OK  $m" -ForegroundColor Green }
function Write-Warn($m) { Write-Host "  !!  $m" -ForegroundColor Yellow }

function Connect-Target {
    if ($UseInteractive) {
        Connect-PnPOnline -Url $SiteUrl -Interactive
    } else {
        $cert = Get-ChildItem Cert:\CurrentUser\My, Cert:\LocalMachine\My -ErrorAction SilentlyContinue |
        Where-Object { $_.Subject -like "$CertSubject*" -and $_.HasPrivateKey -and $_.NotAfter -gt (Get-Date) } |
        Sort-Object NotAfter -Descending | Select-Object -First 1
        if (-not $cert) { throw "No valid '$CertSubject' cert with private key. Pass -Interactive or fix the cert context." }
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $Tenant -Thumbprint $cert.Thumbprint
    }
}

# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
if (-not (Test-Path $ExportDir)) { New-Item -ItemType Directory -Path $ExportDir | Out-Null }

Write-Step "Connecting to $SiteUrl"
Connect-Target
Write-OK "Connected"

Write-Step "Enumerating recycle bin (RowLimit $RowLimit)"
$all = Get-PnPRecycleBinItem -RowLimit $RowLimit
$set = $all | Where-Object { $_.DirName -like $PathFilter }
if ($FirstStageOnly) { $set = $set | Where-Object { $_.ItemState -eq 'FirstStageRecycleBin' } }

$files = @($set | Where-Object ItemType -eq 'File')
$folders = @($set | Where-Object ItemType -eq 'Folder')
$bytes = ($files | Measure-Object Size -Sum).Sum
$dates = $set | Where-Object DeletedDate | Sort-Object DeletedDate
Write-OK ("Matched {0} items ({1} files, {2} folders), {3:N2} GB" -f $set.Count, $files.Count, $folders.Count, ($bytes / 1GB))
if ($dates) {
    Write-OK ("Deleted between {0} and {1}" -f $dates[0].DeletedDate, $dates[-1].DeletedDate)
}
$byWho = $set | Group-Object DeletedByName | Sort-Object Count -Descending
$byWho | ForEach-Object { Write-Host ("       {0,6}  {1}" -f $_.Count, $_.Name) }

# Snapshot the targeted set for the audit trail
$snapCsv = Join-Path $ExportDir "Restore-Target-${Label}_$stamp.csv"
$set | Select-Object Id, LeafName, ItemType, Size, DeletedDate, DeletedByName, DirName |
Export-Csv $snapCsv -NoTypeInformation -Encoding UTF8
Write-OK "Target snapshot -> $snapCsv"

if (-not $Execute) {
    Write-Warn "DRY-RUN. No items restored. Re-run with -Execute to restore the $($set.Count) items above."
    return
}

# Restore folders first (so child paths exist), then files. Retry with back-off on throttle.
Write-Step "Restoring $($set.Count) items"
$ordered = @($folders) + @($files)
$failures = New-Object System.Collections.Generic.List[object]
$done = 0
foreach ($it in $ordered) {
    $attempt = 0
    while ($true) {
        $attempt++
        try {
            Restore-PnPRecycleBinItem -Identity $it.Id -Force -ErrorAction Stop
            $done++
            break
        } catch {
            if ($attempt -ge $MaxRetry) {
                $failures.Add([pscustomobject]@{ Id = $it.Id; LeafName = $it.LeafName; DirName = $it.DirName; Error = $_.Exception.Message })
                Write-Warn ("Gave up on {0} after {1} tries: {2}" -f $it.LeafName, $attempt, $_.Exception.Message)
                break
            }
            $delay = $BaseDelaySec * [math]::Pow(2, $attempt - 1)
            Start-Sleep -Seconds $delay
        }
    }
    if ($done % 250 -eq 0 -and $done -gt 0) { Write-Host "       restored $done / $($ordered.Count)" -ForegroundColor DarkGray }
}

Write-OK ("Restored {0} / {1} items" -f $done, $ordered.Count)
if ($failures.Count -gt 0) {
    $failCsv = Join-Path $ExportDir "Restore-Failures-${Label}_$stamp.csv"
    $failures | Export-Csv $failCsv -NoTypeInformation -Encoding UTF8
    Write-Warn "$($failures.Count) failures -> $failCsv  (re-run -Execute to retry; restore is idempotent)"
} else {
    Write-OK "No failures. Verify the live folder count matches before closing the request."
}