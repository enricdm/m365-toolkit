<#
.SYNOPSIS
    Finds 'Everyone except external users' (EEEU) direct permission grants across
    the site collections matching a URL pattern.

.DESCRIPTION
    Read-only. Scans at site and list/library level; per-item scanning is opt-in via
    -ScanItems because it is very slow on large sites.

    EEEU is the grant that silently makes content readable by every licensed user in
    the tenant, so finding it where it was not intended is the point of this script.

    DETECTION IS CLAIM-BASED, NOT NAME-BASED.
    EEEU's display name is localised: an English tenant shows "Everyone except external
    users", a German one "Jeder, außer externen Benutzern", and so on. Matching only on
    display name therefore misses real grants in any tenant whose locale is not covered
    by the list - a silent false negative, which in a permissions scanner is worse than
    no scan at all, because it reports "clean". The reliable signal is the well-known
    login claim 'c:0-.f|rolemanager|spo-grid-all-users', which is identical in every
    locale. The display-name list is kept only as a secondary signal.

.PARAMETER ClientId
    Entra app (client) ID used for the interactive PnP sign-in.

.PARAMETER AdminCenterUrl
    SharePoint Online admin center URL, e.g. https://contoso-admin.sharepoint.com.

.PARAMETER UrlPattern
    Wildcard pattern selecting which site collections to scan, matched against site URL.
    Default '*' = every site in the tenant.

.PARAMETER ScanItems
    Also scan per-item permissions. Slow on large sites.

.PARAMETER OutputCsv
    CSV of findings.

.EXAMPLE
    # Whole tenant
    .\Get-EveryoneExceptExternalGrant.ps1 -ClientId '<client-id>' `
        -AdminCenterUrl 'https://contoso-admin.sharepoint.com'

.EXAMPLE
    # Only one group of sites, including item-level grants
    .\Get-EveryoneExceptExternalGrant.ps1 -ClientId '<client-id>' `
        -AdminCenterUrl 'https://contoso-admin.sharepoint.com' `
        -UrlPattern '*/sites/Example*' -ScanItems

.NOTES
    Requires:
      - PnP.PowerShell module installed
      - Entra app ClientId (registered by your SharePoint admin)
      - SharePoint Admin OR Site Collection Admin on every site to be scanned
        (a tenant DAG/permissions report does not need per-site SCA - only this
         direct-scan approach does)
    Makes no changes.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $ClientId,
    [Parameter(Mandatory)][string] $AdminCenterUrl,
    [string] $UrlPattern = '*',
    [switch] $ScanItems,
    [string] $OutputCsv  = (Join-Path $PSScriptRoot "Exports\EEEU_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv")
)

# EEEU display names, English + German. Locale-dependent, so this list alone
# produces silent false negatives — the login claim below is the reliable signal.
$TargetPrincipals = @(
    "Everyone except external users",
    "Jeder, außer externen Benutzern",
    "Jeder außer externe Benutzer"   # alt German variant
)

# Locale-independent EEEU identification: the well-known SharePoint login claim.
$TargetLoginClaims = @(
    'c:0-.f|rolemanager|spo-grid-all-users'
)

# True when a role assignment's member is EEEU, by claim OR by display name.
function Test-IsEeeuMember {
    param($Member)
    if ($TargetPrincipals -contains $Member.Title) { return $true }
    foreach ($claim in $TargetLoginClaims) {
        if ($Member.LoginName -like "*$claim*") { return $true }
    }
    return $false
}

# =======================================================================
# INIT
# =======================================================================
$ErrorActionPreference = 'Stop'

function Write-Step($m) { Write-Host "`n=== $m ===" -ForegroundColor Cyan }
function Write-Info($m) { Write-Host "  $m" -ForegroundColor Gray }
function Write-OK  ($m) { Write-Host "  $m" -ForegroundColor Green }
function Write-Warn($m) { Write-Host "  ! $m" -ForegroundColor Yellow }

New-Item -ItemType Directory -Force -Path (Split-Path $OutputCsv) | Out-Null

# =======================================================================
# STEP 1: Enumerate target sites (via SPO admin connection)
# =======================================================================
Write-Step "Enumerating sites"
Connect-PnPOnline -Url $AdminCenterUrl -Interactive -ClientId $ClientId

$allSites = Get-PnPTenantSite -Detailed
$sitesToScan = $allSites | Where-Object { $_.Url -like $UrlPattern }
Write-Info "Found $($sitesToScan.Count) sites matching '$UrlPattern'"
$sitesToScan | ForEach-Object { Write-Host "    $($_.Url)" -ForegroundColor DarkGray }

Disconnect-PnPOnline

# =======================================================================
# STEP 2: Scan each site
# =======================================================================
$report = [System.Collections.Generic.List[object]]::new()
$siteCount = 0
foreach ($site in $sitesToScan) {
    $siteCount++
    Write-Step "[$siteCount/$($sitesToScan.Count)] Scanning $($site.Url)"

    try {
        Connect-PnPOnline -Url $site.Url -Interactive -ClientId $ClientId -ErrorAction Stop

        $web = Get-PnPWeb -Includes RoleAssignments, Title

        # ---------- SITE LEVEL ----------
        foreach ($ra in $web.RoleAssignments) {
            Get-PnPProperty -ClientObject $ra -Property Member, RoleDefinitionBindings | Out-Null
            if (Test-IsEeeuMember $ra.Member) {
                $roles = ($ra.RoleDefinitionBindings | Select-Object -ExpandProperty Name) -join ', '
                $report.Add([pscustomobject]@{
                    SiteUrl           = $site.Url
                    SiteTitle         = $web.Title
                    ObjectType        = 'Site'
                    ObjectName        = $web.Title
                    ObjectUrl         = $site.Url
                    UniquePermissions = $true
                    Principal         = $ra.Member.Title
                    PrincipalLogin    = $ra.Member.LoginName
                    PermissionLevels  = $roles
                })
                Write-OK "  Site-level: $($ra.Member.Title) -> $roles"
            }
        }

        # ---------- LIST / LIBRARY LEVEL ----------
        $lists = Get-PnPList
        foreach ($list in $lists) {
            Get-PnPProperty -ClientObject $list -Property HasUniqueRoleAssignments | Out-Null
            if (-not $list.HasUniqueRoleAssignments) { continue }

            Get-PnPProperty -ClientObject $list -Property RoleAssignments | Out-Null
            foreach ($ra in $list.RoleAssignments) {
                Get-PnPProperty -ClientObject $ra -Property Member, RoleDefinitionBindings | Out-Null
                if (Test-IsEeeuMember $ra.Member) {
                    $roles = ($ra.RoleDefinitionBindings | Select-Object -ExpandProperty Name) -join ', '
                    $report.Add([pscustomobject]@{
                        SiteUrl           = $site.Url
                        SiteTitle         = $web.Title
                        ObjectType        = 'List/Library'
                        ObjectName        = $list.Title
                        ObjectUrl         = $list.DefaultViewUrl
                        UniquePermissions = $true
                        Principal         = $ra.Member.Title
                        PrincipalLogin    = $ra.Member.LoginName
                        PermissionLevels  = $roles
                    })
                    Write-OK "  List '$($list.Title)': $($ra.Member.Title) -> $roles"
                }
            }
        }

        # ---------- ITEM LEVEL (optional, slow) ----------
        if ($ScanItems) {
            foreach ($list in $lists) {
                try {
                    $items = Get-PnPListItem -List $list.Title -PageSize 500 -ErrorAction Stop
                    foreach ($item in $items) {
                        Get-PnPProperty -ClientObject $item -Property HasUniqueRoleAssignments | Out-Null
                        if (-not $item.HasUniqueRoleAssignments) { continue }
                        Get-PnPProperty -ClientObject $item -Property RoleAssignments | Out-Null
                        foreach ($ra in $item.RoleAssignments) {
                            Get-PnPProperty -ClientObject $ra -Property Member, RoleDefinitionBindings | Out-Null
                            if (Test-IsEeeuMember $ra.Member) {
                                $roles = ($ra.RoleDefinitionBindings | Select-Object -ExpandProperty Name) -join ', '
                                $report.Add([pscustomobject]@{
                                    SiteUrl           = $site.Url
                                    SiteTitle         = $web.Title
                                    ObjectType        = 'Folder/Item'
                                    ObjectName        = $item.FieldValues.FileLeafRef
                                    ObjectUrl         = $item.FieldValues.FileRef
                                    UniquePermissions = $true
                                    Principal         = $ra.Member.Title
                                    PrincipalLogin    = $ra.Member.LoginName
                                    PermissionLevels  = $roles
                                })
                                Write-OK "  Item '$($item.FieldValues.FileLeafRef)': $($ra.Member.Title) -> $roles"
                            }
                        }
                    }
                } catch { Write-Warn "  Skipped list '$($list.Title)': $($_.Exception.Message)" }
            }
        }

        Disconnect-PnPOnline
    } catch {
        Write-Warn "Site failed: $($_.Exception.Message)"
    }
}

# =======================================================================
# STEP 3: Export
# =======================================================================
Write-Step "Exporting"
if ($report.Count -gt 0) {
    $report | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
    Write-OK "Saved $($report.Count) findings to: $OutputCsv"
    Write-Host ""
    Write-Host "Summary by site:" -ForegroundColor Cyan
    $report | Group-Object SiteUrl | Sort-Object Count -Descending | ForEach-Object {
        Write-Host ("  {0,4}  {1}" -f $_.Count, $_.Name) -ForegroundColor DarkGray
    }
} else {
    Write-OK "No EEEU grants found on the scanned sites."
}
