<#
.SYNOPSIS
    Connect to SharePoint Online via PnP using app-only certificate auth.

.DESCRIPTION
    Opens an app-only (non-interactive) PnP connection to a site, an admin center,
    or a OneDrive. -Url is mandatory on purpose: there is no safe default target for
    a connection cmdlet, and a wrong implicit target is worse than no default.

    If -Thumbprint is omitted the newest non-expired certificate whose subject matches
    -CertSubject is taken from CurrentUser\My.

.PARAMETER Url
    Target URL. A site collection, the tenant admin center, or a OneDrive.

.PARAMETER ClientId
    App (client) ID of the Entra app registration used for app-only auth.

.PARAMETER Tenant
    Tenant name, e.g. contoso.onmicrosoft.com.

.PARAMETER Thumbprint
    Certificate thumbprint. If omitted, resolved from -CertSubject.

.PARAMETER CertSubject
    Subject of the auth certificate to look up when -Thumbprint is not supplied.

.EXAMPLE
    .\Connect-PnPSite.ps1 -Url 'https://contoso.sharepoint.com/sites/Example' `
        -ClientId '<client-id>' -Tenant 'contoso.onmicrosoft.com'

.EXAMPLE
    .\Connect-PnPSite.ps1 -Url 'https://contoso-admin.sharepoint.com' `
        -ClientId '<client-id>' -Tenant 'contoso.onmicrosoft.com' -Thumbprint '<cert-thumbprint>'

.NOTES
    Module : PnP.PowerShell
    Auth   : app-only certificate. No interactive sign-in.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $Url,
    [Parameter(Mandatory)][string] $ClientId,
    [Parameter(Mandatory)][string] $Tenant,
    [string] $Thumbprint,
    [string] $CertSubject = 'CN=PnP-SPO-Snapshot'
)

if (-not $Thumbprint) {
    $cert = Get-ChildItem Cert:\CurrentUser\My |
    Where-Object { $_.Subject -eq $CertSubject } |
    Sort-Object NotAfter -Descending | Select-Object -First 1
    if (-not $cert) { throw "Cert '$CertSubject' not found in CurrentUser\My. Pass -Thumbprint." }
    $Thumbprint = $cert.Thumbprint
}

Connect-PnPOnline -Url $Url -ClientId $ClientId -Tenant $Tenant -Thumbprint $Thumbprint
Write-Host "Connected (app-only) to $Url" -ForegroundColor Green
