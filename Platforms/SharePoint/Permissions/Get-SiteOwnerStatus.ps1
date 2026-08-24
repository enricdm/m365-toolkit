<#
.SYNOPSIS
    SharePoint Online site owner snapshot: owners + owner account-enabled status.

.DESCRIPTION
    Read-only. App-only (certificate) auth via PnP.PowerShell. One row per site.

    Two modes:
      - Targeted (recommended): pass -InputCsv (e.g. a quota-based ">85% full" export)
        or -SiteUrl. Owners are resolved for exactly those sites. Fast, and avoids the
        storage-% problem below.
      - Full tenant: no list passed -> every site is enumerated.

    Owner resolution by site type:
      - Group-connected : M365 group owners via Graph (accountEnabled inline).
      - Non-group       : the site's Owners group members + Site Collection Admins,
                          read per-site via PnP (what the admin UI 'Site owners' /
                          'Site admins' tabs show). These are SharePoint groups, NOT
                          M365 groups, so Graph /groups cannot see them.

    accountEnabled is tri-state: True / False / Unknown. A failed/missing lookup is
    NEVER reported as False.

    NOTE ON STORAGE %: PercentUsed is computed from StorageMaximumLevel. In a pooled-
    storage tenant that is a large shared cap, so the value is NOT a real "% full" and
    must not be used to filter. Use a quota-based export for the threshold and let this
    script add the owners. Join the two on SiteUrl.

    Makes no changes to any site or account.

.PARAMETER AdminUrl
    SharePoint Online admin center URL, e.g. https://contoso-admin.sharepoint.com.

.PARAMETER ClientId
    App (client) ID of the Entra app registration used for app-only auth.

.PARAMETER Tenant
    Tenant name, e.g. contoso.onmicrosoft.com.

.PARAMETER Thumbprint
    Auth certificate thumbprint. If omitted, resolved from -CertSubject.

.PARAMETER CertSubject
    Subject of the auth certificate to look up when -Thumbprint is not supplied.

.PARAMETER InputCsv
    CSV containing a site URL column (Url / SiteUrl / 'Site URL'). Owners resolved for
    those sites only.

.PARAMETER SiteUrl
    One or more explicit site URLs to resolve (alternative to -InputCsv).

.EXAMPLE
    # Targeted: resolve owners for the sites listed in a CSV
    .\Get-SiteOwnerStatus.ps1 -AdminUrl 'https://contoso-admin.sharepoint.com' `
        -ClientId '<client-id>' -Tenant 'contoso.onmicrosoft.com' -InputCsv .\over-85-percent.csv

.EXAMPLE
    # Full tenant sweep
    .\Get-SiteOwnerStatus.ps1 -AdminUrl 'https://contoso-admin.sharepoint.com' `
        -ClientId '<client-id>' -Tenant 'contoso.onmicrosoft.com' -OutputPath .\Exports\owners.csv

.NOTES
    When to use  : You need to notify site owners about something and first need to know how many of those owners no longer work here.
    Why it exists: Owner resolution differs by site type: group-connected sites expose owners through Graph, non-group sites keep them in the SharePoint Owners group and site collection admins, which Graph cannot see. accountEnabled is tri-state and a failed lookup is never reported as False. It also warns that PercentUsed is not a real fill percentage in a pooled-storage tenant.
    Module : PnP.PowerShell
    Auth   : App-only certificate (ClientId + Tenant + cert)
    Rights : Application permissions, admin-consented -
             Microsoft Graph : Group.Read.All, User.Read.All
             SharePoint      : Sites.FullControl.All
    Tested with PowerShell 7.6.x.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $AdminUrl,
    [Parameter(Mandatory)][string] $ClientId,
    [Parameter(Mandatory)][string] $Tenant,
    [string]   $Thumbprint,
    [string]   $CertSubject     = 'CN=PnP-SPO-Snapshot',
    [string]   $InputCsv,
    [string[]] $SiteUrl,
    [string]   $OutputPath      = ".\SPO-SiteOwnerStatus_$(Get-Date -Format 'yyyyMMdd').csv",
    [double]   $MinPercentUsed  = 0,
    [int]      $MaxOwnersListed = 15,
    [bool]     $DeepScanNonGroup = $true
)

$ErrorActionPreference = 'Stop'

$SystemTemplates = @('SPSMSITEHOST#0', 'REDIRECTSITE#0', 'PWA#0', 'PWA#1', 'APPCATALOG#0', 'SRCHCEN#0', 'EHS#1')

# --- helpers ---------------------------------------------------------------

function Join-Capped {
    param([string[]] $Items, [int] $Max)
    if (-not $Items -or $Items.Count -eq 0) { return '' }
    if ($Items.Count -le $Max) { return ($Items -join '; ') }
    $shown = $Items[0..($Max - 1)] -join '; '
    return "$shown; ... (+$($Items.Count - $Max) more)"
}

function Get-UpnFromPrincipal {
    param($Principal)
    if ($Principal.Email) { return $Principal.Email }
    $ln = [string]$Principal.LoginName
    if ($ln -match 'membership\|(.+)$') { return $Matches[1] }
    if ($ln -match '\|([^|]+@[^|]+)$')   { return $Matches[1] }
    return $null
}

$enabledCache = @{}
function Resolve-Enabled {
    param([string] $Upn, $Conn)
    if (-not $Upn) { return $null }
    if ($enabledCache.ContainsKey($Upn)) { return $enabledCache[$Upn] }
    $res = $null
    try {
        $u = Invoke-PnPGraphMethod -Connection $Conn -Method Get `
            -Url ('v1.0/users/{0}?$select=accountEnabled' -f [uri]::EscapeDataString($Upn))
        $res = $u.accountEnabled
    } catch { $res = $null }   # unknown, NOT disabled
    $enabledCache[$Upn] = $res
    return $res
}

# --- resolve cert -----------------------------------------------------------

if (-not $Thumbprint) {
    $cert = Get-ChildItem Cert:\CurrentUser\My |
        Where-Object { $_.Subject -eq $CertSubject } |
        Sort-Object NotAfter -Descending | Select-Object -First 1
    if (-not $cert) { throw "No cert with subject '$CertSubject' in CurrentUser\My. Pass -Thumbprint." }
    $Thumbprint = $cert.Thumbprint
}

# --- connect (admin) --------------------------------------------------------

Write-Host 'Connecting to SharePoint admin (app-only cert)...' -ForegroundColor Cyan
$adminConn = Connect-PnPOnline -Url $AdminUrl -ClientId $ClientId -Tenant $Tenant -Thumbprint $Thumbprint -ReturnConnection

$warnings = New-Object System.Collections.Generic.List[string]

# --- build target site list -------------------------------------------------

if ($InputCsv) {
    if (-not (Test-Path $InputCsv)) { throw "InputCsv not found: $InputCsv" }
    $csv = Import-Csv $InputCsv
    $urlCol = $csv[0].psobject.Properties.Name |
        Where-Object { $_ -match '^(site\s*url|siteurl|url)$' } | Select-Object -First 1
    if (-not $urlCol) { throw "No URL column in $InputCsv (looked for Url / SiteUrl / 'Site URL')." }
    $SiteUrl = $csv | ForEach-Object { $_.$urlCol } | Where-Object { $_ } | Select-Object -Unique
    Write-Host "Loaded $($SiteUrl.Count) site URLs from $InputCsv" -ForegroundColor Cyan
}

if ($SiteUrl) {
    Write-Host "Resolving $($SiteUrl.Count) targeted sites..." -ForegroundColor Cyan
    $sites = foreach ($u in $SiteUrl) {
        try { Get-PnPTenantSite -Identity $u -Connection $adminConn }
        catch { $warnings.Add("Site not found / inaccessible: $u") }
    }
} else {
    Write-Host 'Enumerating all sites...' -ForegroundColor Cyan
    $sites = Get-PnPTenantSite -IncludeOneDriveSites:$false -Connection $adminConn
}

$total   = @($sites).Count
$results = New-Object System.Collections.Generic.List[object]
$i = 0

foreach ($site in $sites) {
    if (-not $site) { continue }
    $i++
    Write-Progress -Activity 'Resolving site owners' -Status "$i / $total : $($site.Url)" `
        -PercentComplete ([int](($i / [math]::Max($total, 1)) * 100))

    $usedGb  = [math]::Round($site.StorageUsageCurrent / 1024, 2)
    $quotaGb = if ($site.StorageMaximumLevel) { [math]::Round($site.StorageMaximumLevel / 1024, 2) } else { $null }
    $pct     = if ($site.StorageMaximumLevel) {
        [math]::Round(($site.StorageUsageCurrent / $site.StorageMaximumLevel) * 100, 2)
    } else { $null }

    if ($pct -ne $null -and $pct -lt $MinPercentUsed) { continue }

    $isGroup  = [string]::IsNullOrWhiteSpace($site.GroupId.Guid) -eq $false `
        -and $site.GroupId.Guid -ne '00000000-0000-0000-0000-000000000000'
    $isSystem = $SystemTemplates -contains $site.Template

    $resolved = @()

    # 1) Group-connected -> M365 group owners (enabled inline)
    if ($isGroup) {
        $url = 'v1.0/groups/{0}/owners?$select=id,displayName,userPrincipalName,accountEnabled' -f $site.GroupId.Guid
        try {
            $owners = (Invoke-PnPGraphMethod -Connection $adminConn -Url $url -Method Get -All).value
            foreach ($o in $owners) {
                $upn = if ($o.userPrincipalName) { $o.userPrincipalName } else { $o.displayName }
                $resolved += [pscustomobject]@{ Upn = $upn; Name = $o.displayName; Enabled = $o.accountEnabled; Source = 'GroupOwner' }
                if ($o.userPrincipalName -and -not $enabledCache.ContainsKey($o.userPrincipalName)) {
                    $enabledCache[$o.userPrincipalName] = $o.accountEnabled
                }
            }
        } catch {
            $warnings.Add("Group owners failed for $($site.Url): $($_.Exception.Message)")
        }
    }

    # 2) Non-group -> Owners group members + Site Collection Admins (per-site)
    if (-not $isGroup -and $DeepScanNonGroup -and -not $isSystem) {
        $siteConn = $null
        try {
            $siteConn = Connect-PnPOnline -Url $site.Url -ClientId $ClientId -Tenant $Tenant -Thumbprint $Thumbprint -ReturnConnection
            $principals = @()

            $web = Get-PnPWeb -Includes AssociatedOwnerGroup -Connection $siteConn
            if ($web.AssociatedOwnerGroup -and $web.AssociatedOwnerGroup.Id -gt 0) {
                $principals += Get-PnPGroupMember -Group $web.AssociatedOwnerGroup -Connection $siteConn
            }
            $principals += Get-PnPSiteCollectionAdmin -Connection $siteConn

            $seen = @{}
            foreach ($p in $principals) {
                $upn = Get-UpnFromPrincipal -Principal $p
                $key = if ($upn) { $upn } else { [string]$p.LoginName }
                if (-not $key -or $seen.ContainsKey($key)) { continue }
                $seen[$key] = $true

                $isUser  = ($p.PrincipalType -eq 'User') -or ($upn -and $p.PrincipalType -ne 'SecurityGroup')
                $enabled = if ($isUser -and $upn) { Resolve-Enabled -Upn $upn -Conn $adminConn } else { $null }
                $resolved += [pscustomobject]@{
                    Upn     = if ($upn) { $upn } else { $p.Title }
                    Name    = $p.Title
                    Enabled = $enabled
                    Source  = "$($p.PrincipalType)"
                }
            }
        } catch {
            $warnings.Add("Owner-group read failed for $($site.Url): $($_.Exception.Message)")
        }
        # NOTE: per-site connections are left for GC. Disconnect-PnPOnline has no
        # -Connection param in this PnP build, and a bare Disconnect would drop the
        # admin (current) connection. Fine for targeted lists; avoid full-tenant +
        # DeepScan in one run if memory is tight.
    }

    # 3) Last-resort fallback: tenant Owner field
    if ($resolved.Count -eq 0 -and $site.Owner) {
        $raw = ($site.Owner -split '\|')[-1]
        $enabled = if ($raw -match '@') { Resolve-Enabled -Upn $raw -Conn $adminConn } else { $null }
        $resolved += [pscustomobject]@{ Upn = $raw; Name = $raw; Enabled = $enabled; Source = 'OwnerField' }
    }

    # --- summarise (tri-state enabled) ---
    $enabledTrue = @($resolved | Where-Object { $_.Enabled -eq $true })
    $knownStatus = @($resolved | Where-Object { $_.Enabled -ne $null })
    $primary = if ($enabledTrue.Count) { $enabledTrue[0] }
    elseif ($resolved.Count) { $resolved[0] }
    else { $null }

    $hasEnabled = if ($resolved.Count -eq 0) { '' }
    elseif ($enabledTrue.Count -gt 0) { $true }
    elseif ($knownStatus.Count -gt 0) { $false }
    else { 'Unknown' }

    $primaryActive = if (-not $primary) { '' }
    elseif ($primary.Enabled -eq $true) { 'True' }
    elseif ($primary.Enabled -eq $false) { 'False' }
    else { 'Unknown' }

    $results.Add([pscustomobject]@{
            SiteTitle            = $site.Title
            SiteUrl              = $site.Url
            Template             = $site.Template
            IsGroupConnected     = $isGroup
            IsSystemSite         = $isSystem
            StorageUsedGB        = $usedGb
            StorageQuotaGB       = $quotaGb
            PercentUsed          = $pct
            OwnerCount           = $resolved.Count
            EnabledOwnerCount    = $enabledTrue.Count
            HasEnabledOwner      = $hasEnabled
            PrimaryContact       = if ($primary) { $primary.Upn } else { '' }
            PrimaryContactName   = if ($primary) { $primary.Name } else { '' }
            PrimaryContactActive = $primaryActive
            AllOwners            = Join-Capped -Items ($resolved | ForEach-Object { $_.Upn }) -Max $MaxOwnersListed
            FollowUpStatus       = ''
            FollowUpNotes        = ''
            SnapshotDate         = (Get-Date -Format 'yyyy-MM-dd')
        })
}

Write-Progress -Activity 'Resolving site owners' -Completed

$results | Sort-Object PercentUsed -Descending |
    Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

# --- summary ----------------------------------------------------------------

$noOwner   = @($results | Where-Object { $_.OwnerCount -eq 0 -and -not $_.IsSystemSite })
$noEnabled = @($results | Where-Object { $_.HasEnabledOwner -eq $false })
$unknown   = @($results | Where-Object { $_.HasEnabledOwner -eq 'Unknown' })
Write-Host ''
Write-Host "Sites written         : $($results.Count)" -ForegroundColor Green
Write-Host "No owner (non-system) : $($noOwner.Count)  <-- genuinely ownerless" -ForegroundColor Yellow
Write-Host "All owners disabled   : $($noEnabled.Count)  <-- orphaned, escalate" -ForegroundColor Yellow
Write-Host "Enabled status UNKNOWN: $($unknown.Count)  <-- check Graph permissions if high" -ForegroundColor DarkYellow
if ($warnings.Count) {
    Write-Host "Warnings              : $($warnings.Count) (run with -Verbose for detail)" -ForegroundColor Red
    $warnings | Select-Object -First 8 | ForEach-Object { Write-Verbose $_ }
}
Write-Host "Output                : $OutputPath" -ForegroundColor Green

try { Disconnect-PnPOnline } catch {}