<#
.SYNOPSIS
    Clears the local MSAL/WAM token caches, then reconnects to the SharePoint Online
    admin center with modern auth.

.DESCRIPTION
    Fixes the "stuck identity" case: Connect-SPOService keeps reusing a cached token
    for the wrong account or a stale tenant, and no amount of Disconnect-SPOService
    clears it because the token lives in the local identity caches, not in the module.

    Deletes the cache directories, then connects. Close the current PowerShell window
    and open a fresh one before running this: a session that has already loaded the
    identity assemblies can rewrite the caches on exit.

.PARAMETER AdminUrl
    SharePoint Online admin center URL, e.g. https://contoso-admin.sharepoint.com.

.PARAMETER SkipConnect
    Only purge the caches; do not connect afterwards.

.EXAMPLE
    .\Clear-SpoTokenCache.ps1 -AdminUrl 'https://contoso-admin.sharepoint.com'

.EXAMPLE
    .\Clear-SpoTokenCache.ps1 -AdminUrl 'https://contoso-admin.sharepoint.com' -SkipConnect

.NOTES
    Module : Microsoft.Online.SharePoint.PowerShell
    Affects the CURRENT USER's token caches only. Other Microsoft 365 tools on the
    machine will prompt for sign-in again after this runs.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $AdminUrl,
    [switch] $SkipConnect
)

# Clear MSAL/WAM token caches
Remove-Item "$env:LOCALAPPDATA\.IdentityService\*" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "$env:LOCALAPPDATA\Microsoft\TokenBroker\Cache\*" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "$env:LOCALAPPDATA\Microsoft\IdentityCache\*" -Recurse -Force -ErrorAction SilentlyContinue

Write-Host 'Token caches cleared.' -ForegroundColor Green

if ($SkipConnect) { return }

# Connect WITHOUT importing SPO first: let the cmdlet auto-load with the right context
Connect-SPOService -Url $AdminUrl -ModernAuth $true
Write-Host "Connected to $AdminUrl" -ForegroundColor Green
