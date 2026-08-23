<#
.SYNOPSIS
    Intune remediation DETECTION: verifies that the managed-bookmarks policy for
    Microsoft Edge or Google Chrome is present and matches the expected set.

.DESCRIPTION
    Detection half of a device-side remediation pair. Pair with
    Repair-ManagedBookmark.ps1.

    Exit 0 = compliant, exit 1 = remediate.

    Checks the browser POLICY value, not the user's own bookmark file:

      Edge   HKLM:\SOFTWARE\Policies\Microsoft\Edge   ManagedFavorites
      Chrome HKLM:\SOFTWARE\Policies\Google\Chrome    ManagedBookmarks

    Comparison is on the SET of URLs, order-insensitive and case-insensitive.
    Bookmark titles are ignored: renaming a folder label is not a compliance
    problem, a missing destination is.

.PARAMETER Browser
    Edge (default) or Chrome.

.PARAMETER ExpectedUrl
    URLs that must be present in the managed set. Edit the default before
    deploying - Intune remediation scripts run with no arguments.

.PARAMETER Exact
    Require the managed set to contain EXACTLY these URLs. By default extra
    bookmarks beyond -ExpectedUrl are tolerated.

.EXAMPLE
    # Deploy as-is after editing the default ExpectedUrl list
    .\Detect-ManagedBookmark.ps1

.EXAMPLE
    # Chrome, exact match required
    .\Detect-ManagedBookmark.ps1 -Browser Chrome -Exact `
        -ExpectedUrl 'https://intranet.contoso.com','https://mail.contoso.com'

.NOTES
    Context  : run as SYSTEM (Intune remediation default), 64-bit
    Requires : Windows PowerShell 5.1
    Rights   : read-only. This script never writes to the registry.

    Only the machine-wide policy is inspected. A user who has bookmarked the same
    sites manually will still be reported NON-COMPLIANT, because the policy is what
    survives a profile reset - which is the point on shared devices.

    Replaces (merged): four detection scripts that each hard-coded one
    organisation's bookmark set - two divergent copies of the same Chrome check,
    a per-tenant variant, and a separate "bookmarks fix" detection.
#>

[CmdletBinding()]
param(
    [ValidateSet('Edge', 'Chrome')]
    [string]$Browser = 'Edge',

    [string[]]$ExpectedUrl = @(
        'https://intranet.contoso.com'
        'https://mail.contoso.com'
        'https://portal.office.com/onedrive'
    ),

    [switch]$Exact
)

$policy = switch ($Browser) {
    'Edge'   { @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Edge'; Name = 'ManagedFavorites' } }
    'Chrome' { @{ Path = 'HKLM:\SOFTWARE\Policies\Google\Chrome';  Name = 'ManagedBookmarks' } }
}

if (-not (Test-Path $policy.Path)) {
    Write-Output "NON-COMPLIANT: policy key not present ($($policy.Path))."
    exit 1
}

$raw = $null
try {
    $raw = (Get-ItemProperty -Path $policy.Path -Name $policy.Name -ErrorAction Stop).$($policy.Name)
}
catch {
    Write-Output "NON-COMPLIANT: policy value '$($policy.Name)' not set."
    exit 1
}

if ([string]::IsNullOrWhiteSpace($raw)) {
    Write-Output "NON-COMPLIANT: policy value '$($policy.Name)' is empty."
    exit 1
}

# The value is a JSON array. A malformed value must remediate, not crash.
try   { $parsed = $raw | ConvertFrom-Json -ErrorAction Stop }
catch {
    Write-Output "NON-COMPLIANT: policy value is not valid JSON - $($_.Exception.Message)"
    exit 1
}

# Managed bookmarks nest: entries carry either 'url' or 'children'. Walk the tree.
function Get-UrlFromNode {
    param($Node)
    $acc = New-Object System.Collections.Generic.List[string]
    foreach ($n in @($Node)) {
        if ($null -eq $n) { continue }
        if ($n.PSObject.Properties['url'] -and $n.url) { $acc.Add([string]$n.url) }
        if ($n.PSObject.Properties['children'] -and $n.children) {
            foreach ($c in (Get-UrlFromNode $n.children)) { $acc.Add($c) }
        }
    }
    return $acc
}

# Normalise for comparison: trim, lowercase, drop one trailing slash.
function ConvertTo-ComparableUrl {
    param([string]$Url)
    if ([string]::IsNullOrWhiteSpace($Url)) { return '' }
    return ($Url.Trim().TrimEnd('/').ToLowerInvariant())
}

$presentRaw = @(Get-UrlFromNode $parsed)
$present    = @($presentRaw | ForEach-Object { ConvertTo-ComparableUrl $_ } | Where-Object { $_ })
$expected   = @($ExpectedUrl | ForEach-Object { ConvertTo-ComparableUrl $_ } | Where-Object { $_ })

$missing = @($expected | Where-Object { $_ -notin $present })
if ($missing.Count -gt 0) {
    Write-Output "NON-COMPLIANT: $($missing.Count) expected bookmark(s) missing: $($missing -join ', ')"
    exit 1
}

if ($Exact) {
    $extra = @($present | Where-Object { $_ -notin $expected })
    if ($extra.Count -gt 0) {
        Write-Output "NON-COMPLIANT (-Exact): $($extra.Count) unexpected bookmark(s): $($extra -join ', ')"
        exit 1
    }
}

Write-Output "COMPLIANT: $Browser managed bookmarks contain all $($expected.Count) expected URL(s)."
exit 0
