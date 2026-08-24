<#
.SYNOPSIS
    DETECTION - reports current storage usage and quota for the site collections
    matching a title/URL filter.

.DESCRIPTION
    Read-only. Run this FIRST and confirm the exact site URLs in the output before
    running Set-SiteStorageQuota.ps1 against them.

    Exports a "before" snapshot CSV so a quota change has a rollback record.

.PARAMETER AdminUrl
    SharePoint Online admin center URL, e.g. https://contoso-admin.sharepoint.com.

.PARAMETER TitleFilter
    Regex matched against both site Title and Url. Defaults to '.' (every site).

.PARAMETER OutputPath
    CSV written with the pre-change snapshot.

.EXAMPLE
    # Every site in the tenant
    .\Get-SiteStorage.ps1 -AdminUrl 'https://contoso-admin.sharepoint.com'

.EXAMPLE
    # Only the sites whose title or URL matches a project name
    .\Get-SiteStorage.ps1 -AdminUrl 'https://contoso-admin.sharepoint.com' -TitleFilter 'FinancialPlanning'

.NOTES
    When to use  : Step 1 of the storage chain - before Set-SiteStorageQuota.ps1, to confirm the exact site URLs and capture the "before" snapshot.
    Why it exists: Reading quota and writing quota deserve different levels of care, so they are separate scripts and the read one has no way to change anything. Its export doubles as the rollback reference for the write half. One caveat it states: PercentUsed is computed against whatever cap is set on the site, and in a pooled-storage tenant that cap is often a large shared ceiling rather than a real per-site allocation - so it spots sites with an explicit quota, and is not a ranking of who is actually full.
    Module : Microsoft.Online.SharePoint.PowerShell
    Rights : SharePoint Administrator
    Makes no changes.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $AdminUrl,
    [string] $TitleFilter = '.',
    [string] $OutputPath  = ".\SiteStorage-Before_$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
)

# --- Connect to SharePoint Online Admin Center ---
Connect-SPOService -Url $AdminUrl

Write-Host "`n=== Searching for site collections matching '$TitleFilter' ===" -ForegroundColor Cyan

$allSites = Get-SPOSite -Limit All -IncludePersonalSite $false

# NOTE: not named $matches - that is an automatic variable populated by -match.
$matchedSites = $allSites | Where-Object {
    $_.Title -match $TitleFilter -or $_.Url -match $TitleFilter
}

if (-not $matchedSites) {
    Write-Warning "No sites found matching '$TitleFilter'. Check naming/URL convention and re-run with an adjusted -TitleFilter."
} else {
    $matchedSites |
        Select-Object Title, Url,
            @{N='UsedGB';    E={[math]::Round($_.StorageUsageCurrent/1024,2)}},
            @{N='QuotaGB';   E={[math]::Round($_.StorageQuota/1024,2)}},
            @{N='WarningGB'; E={[math]::Round($_.StorageQuotaWarningLevel/1024,2)}},
            @{N='UsedPct';   E={ if ($_.StorageQuota -gt 0) { [math]::Round(($_.StorageUsageCurrent/$_.StorageQuota)*100,1) } else { "N/A" } }} |
        Format-Table -AutoSize

    # Export snapshot as the before/after record
    $matchedSites |
        Select-Object Title, Url, StorageUsageCurrent, StorageQuota, StorageQuotaWarningLevel |
        Export-Csv -Path $OutputPath -NoTypeInformation

    Write-Host "`nSnapshot exported to $OutputPath" -ForegroundColor Green
}
