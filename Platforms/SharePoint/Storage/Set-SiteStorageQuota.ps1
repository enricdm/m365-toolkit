<#
.SYNOPSIS
    INCREASE - sets the storage quota on one or more site collections.

.DESCRIPTION
    Writes. Records the current quota of every target site to a CSV before changing
    anything, so the change can be rolled back. Supports -WhatIf / -Confirm; an
    unreadable site is recorded as 'unknown', never as 0.

    DO NOT RUN until:
      1. Get-SiteStorage.ps1 has been run and the -SiteUrl values are confirmed
         against its output.
      2. Whatever approval your storage policy requires for the total increase
         has been obtained. Raising quota consumes tenant-pooled storage.

.PARAMETER AdminUrl
    SharePoint Online admin center URL, e.g. https://contoso-admin.sharepoint.com.

.PARAMETER SiteUrl
    One or more site collection URLs to update.

.PARAMETER QuotaMB
    New storage quota in MB. Default 102400 (100 GB).

.PARAMETER WarningQuotaMB
    Warning threshold in MB. Default 97280 (~95 GB).

.PARAMETER ExportDir
    Folder for the pre-change rollback CSV.

.EXAMPLE
    # Dry run first - shows what would change, writes the rollback CSV, changes nothing
    .\Set-SiteStorageQuota.ps1 -AdminUrl 'https://contoso-admin.sharepoint.com' `
        -SiteUrl 'https://contoso.sharepoint.com/sites/Example' -WhatIf

.EXAMPLE
    .\Set-SiteStorageQuota.ps1 -AdminUrl 'https://contoso-admin.sharepoint.com' `
        -SiteUrl 'https://contoso.sharepoint.com/sites/Example','https://contoso.sharepoint.com/sites/Example2' `
        -QuotaMB 102400

.NOTES
    Module : Microsoft.Online.SharePoint.PowerShell
    Rights : SharePoint Administrator
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)][string]   $AdminUrl,
    [Parameter(Mandatory)][string[]] $SiteUrl,
    [int]    $QuotaMB        = 102400,  # 100 GB
    [int]    $WarningQuotaMB = 97280,   # ~95 GB
    [string] $ExportDir      = (Join-Path $PSScriptRoot 'Exports')
)

Connect-SPOService -Url $AdminUrl

$siteUrls       = $SiteUrl
$maxQuotaMB     = $QuotaMB
$warningQuotaMB = $WarningQuotaMB

# --- Record the CURRENT quota before changing anything (rollback evidence) ---
if (-not (Test-Path $ExportDir)) { New-Item -ItemType Directory -Path $ExportDir -Force | Out-Null }
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backupCsv = Join-Path $ExportDir "SiteStorageQuota-Before-$stamp.csv"

$previous = foreach ($url in $siteUrls) {
    try {
        $s = Get-SPOSite -Identity $url -ErrorAction Stop
        [pscustomobject]@{
            Url                        = $url
            StorageQuotaMB             = $s.StorageQuota
            StorageQuotaWarningLevelMB = $s.StorageQuotaWarningLevel
            Retrieved                  = 'OK'
        }
    } catch {
        # Never record a made-up value: an unreadable site is 'unknown', not 0.
        Write-Warning "Could not read current quota for $url : $($_.Exception.Message)"
        [pscustomobject]@{
            Url                        = $url
            StorageQuotaMB             = 'unknown'
            StorageQuotaWarningLevelMB = 'unknown'
            Retrieved                  = "FAILED: $($_.Exception.Message)"
        }
    }
}
$previous | Export-Csv -Path $backupCsv -NoTypeInformation -Encoding UTF8
Write-Host "Pre-change quotas saved to: $backupCsv" -ForegroundColor Cyan

foreach ($url in $siteUrls) {
    if (-not $PSCmdlet.ShouldProcess($url, "Set StorageQuota=$maxQuotaMB MB (warning $warningQuotaMB MB)")) { continue }
    try {
        Set-SPOSite -Identity $url -StorageQuota $maxQuotaMB -StorageQuotaWarningLevel $warningQuotaMB
        Write-Host "Updated: $url -> Quota $maxQuotaMB MB (warning $warningQuotaMB MB)" -ForegroundColor Green
    } catch {
        Write-Error "Failed to update $url : $_"
    }
}

# --- Verify ---
Write-Host "`n=== Post-change quota ===" -ForegroundColor Cyan
foreach ($url in $siteUrls) {
    Get-SPOSite -Identity $url |
        Select-Object Url,
            @{N='QuotaGB';   E={[math]::Round($_.StorageQuota/1024,2)}},
            @{N='WarningGB'; E={[math]::Round($_.StorageQuotaWarningLevel/1024,2)}}
}
