<#
.SYNOPSIS
    Clears the signed-in user's MSAL, WAM and IdentityService token caches, and reports
    per cache what it actually managed to remove. Optionally reconnects to SharePoint
    Online afterwards.

.DESCRIPTION
    Fixes the "stuck identity" case: a module keeps reusing a cached token for the wrong
    account or a stale tenant, and no amount of Disconnect-* clears it, because the token
    does not live in the module. It lives in the local identity caches.

    The symptom shows up most often with Connect-SPOService, which is why this used to sit
    under SharePoint and be called Clear-SpoTokenCache. That name described where you
    noticed the problem rather than what the fix touches, so both moved.

    SCOPE. These caches are under %LOCALAPPDATA%, so this affects the profile you are
    signed in as and no other user on the machine. Within that profile it is broad:
    clearing them signs you out of every Microsoft 365 tool that uses them, which
    includes Teams, Outlook, OneDrive and the Graph SDK, not just the module that was
    misbehaving. That is the intended effect, not a side effect, but it is worth knowing
    before you run it on a machine in the middle of somebody's working day.

    WHY IT REPORTS INSTEAD OF ANNOUNCING. The previous version printed "Token caches
    cleared" whether or not anything had been cleared, which is the worst thing this
    particular tool can do: you run it precisely when you are already twenty minutes into
    an authentication problem you do not understand, and a false confirmation sends you
    off to debug the wrong thing. A cache file held open by a running process is the
    normal case here, not an edge case. So each cache is reported separately, and the
    summary says removed, already empty, or failed, with the reason.

    OPEN A FRESH CONSOLE FIRST. A session that has already loaded the identity assemblies
    can write the caches back out on exit, which looks exactly like the clear not working.

.PARAMETER AdminUrl
    SharePoint Online admin center URL. Optional. Supply it to reconnect after clearing;
    omit it and the script only clears, which is the common case now that this is not a
    SharePoint-specific tool.

.PARAMETER IncludeProcessCheck
    Look for running processes known to hold these caches open and name them before
    clearing. On by default; use -IncludeProcessCheck:$false to skip it.

.EXAMPLE
    # Just clear the caches
    .\Clear-M365TokenCache.ps1

.EXAMPLE
    # See what would be removed, change nothing
    .\Clear-M365TokenCache.ps1 -WhatIf

.EXAMPLE
    # Clear, then reconnect to SharePoint admin
    .\Clear-M365TokenCache.ps1 -AdminUrl 'https://contoso-admin.sharepoint.com'

.NOTES
    When to use  : You have spent twenty minutes fighting an access-denied that should not happen, signed in as the wrong account, and Disconnect has not helped.
    Why it exists: The token is not in the module, it is in the local identity caches, and nothing in the module reaches them. It reports per cache what it removed rather than announcing success, because a file held open by a running Teams or Outlook is the normal case and a false confirmation sends you to debug the wrong thing.
    Requires : PowerShell 5.1 or 7.x. Microsoft.Online.SharePoint.PowerShell only if you pass -AdminUrl.
    Rights   : none. It only touches the current profile's own cache folders.
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [string]$AdminUrl,
    [switch]$IncludeProcessCheck = $true
)

$ErrorActionPreference = 'Stop'

function Write-Step { param([string]$m) Write-Host "`n[*] $m" -ForegroundColor Cyan }
function Write-OK   { param([string]$m) Write-Host "[+] $m" -ForegroundColor Green }
function Write-Warn { param([string]$m) Write-Host "[!] $m" -ForegroundColor Yellow }

# The caches are addressed relative to LOCALAPPDATA. If that is not set we stop, rather
# than building a path that starts at the root of the current drive and deleting there.
$localAppData = $env:LOCALAPPDATA
if ([string]::IsNullOrWhiteSpace($localAppData) -or -not (Test-Path -LiteralPath $localAppData)) {
    throw "LOCALAPPDATA is not set or does not exist. Refusing to guess at cache paths."
}

$caches = @(
    [pscustomobject]@{ Name = 'MSAL / IdentityService'; Path = Join-Path $localAppData '.IdentityService' }
    [pscustomobject]@{ Name = 'WAM TokenBroker';        Path = Join-Path $localAppData 'Microsoft\TokenBroker\Cache' }
    [pscustomobject]@{ Name = 'IdentityCache';          Path = Join-Path $localAppData 'Microsoft\IdentityCache' }
)

Write-Step "Token cache clear - profile $env:USERNAME"

# ---- who is holding these open ----
if ($IncludeProcessCheck) {
    $holders = @('Teams','ms-teams','OUTLOOK','OneDrive','Microsoft.SharePoint','pwsh','powershell','Code')
    $running = @(Get-Process -Name $holders -ErrorAction Ignore | Select-Object -ExpandProperty Name -Unique | Sort-Object)
    # 'Ignore' rather than SilentlyContinue: a process simply not running is the expected
    # answer here, not an error worth recording.
    if ($running.Count) {
        Write-Warn "Running and likely to hold these caches open: $($running -join ', ')"
        Write-Host '    Files they have open cannot be removed, and they may write the caches' -ForegroundColor DarkGray
        Write-Host '    back out. Close them for a clean result, or expect partial removal.' -ForegroundColor DarkGray
    }
    else { Write-OK 'No known cache-holding process is running.' }
}

# ---- clear, one cache at a time, reporting each ----
Write-Step 'Clearing'
$results = New-Object System.Collections.Generic.List[object]

foreach ($cache in $caches) {
    $row = [ordered]@{ Cache = $cache.Name; Path = $cache.Path; Items = 0; Status = ''; Detail = '' }

    if (-not (Test-Path -LiteralPath $cache.Path)) {
        $row.Status = 'ABSENT'
        $row.Detail = 'directory does not exist'
        $results.Add([pscustomobject]$row)
        continue
    }

    $items = @(Get-ChildItem -LiteralPath $cache.Path -Force -ErrorAction Ignore)
    $row.Items = $items.Count

    if ($items.Count -eq 0) {
        $row.Status = 'EMPTY'
        $results.Add([pscustomobject]$row)
        continue
    }

    if (-not $PSCmdlet.ShouldProcess($cache.Path, "Remove $($items.Count) cache item(s)")) {
        $row.Status = 'WHATIF'
        $row.Detail = "$($items.Count) item(s) would be removed"
        $results.Add([pscustomobject]$row)
        continue
    }

    # Delete the CONTENTS, never the directory itself: the parent folders are recreated
    # by the identity stack and removing them outright causes its own class of problem.
    $failed = New-Object System.Collections.Generic.List[string]
    foreach ($item in $items) {
        try { Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop }
        catch { $failed.Add("$($item.Name): $($_.Exception.Message)") }
    }

    if ($failed.Count -eq 0) {
        $row.Status = 'CLEARED'
    }
    else {
        $row.Status = 'PARTIAL'
        $row.Detail = "$($failed.Count) of $($items.Count) could not be removed - $($failed[0])"
    }
    $results.Add([pscustomobject]$row)
}

$results | Select-Object Cache, Items, Status, Detail | Format-Table -AutoSize -Wrap

# ---- an honest summary ----
$cleared = @($results | Where-Object Status -eq 'CLEARED').Count
$partial = @($results | Where-Object Status -eq 'PARTIAL').Count
$whatIf  = @($results | Where-Object Status -eq 'WHATIF').Count

if ($whatIf) {
    Write-Warn "Dry run. Nothing was removed. Re-run without -WhatIf to clear."
}
elseif ($partial) {
    Write-Warn "$partial cache(s) only partially cleared. Close the processes named above and run this again."
    Write-Host '    Until they are clear, the stale identity may still come back.' -ForegroundColor DarkGray
}
elseif ($cleared) {
    Write-OK "$cleared cache(s) cleared. Every Microsoft 365 tool on this profile will ask you to sign in again."
}
else {
    Write-OK 'Nothing to clear: the caches were already empty or absent.'
    Write-Host '    If the wrong account is still being used, the token is not in these caches.' -ForegroundColor DarkGray
    Write-Host '    Check Windows Credential Manager, and any -UseDeviceAuthentication or' -ForegroundColor DarkGray
    Write-Host '    service-principal credentials the script itself supplies.' -ForegroundColor DarkGray
}

# ---- optional reconnect ----
if (-not $AdminUrl) { return }
if ($whatIf) { Write-Warn 'Skipping the reconnect on a dry run.'; return }

Write-Step "Connecting to $AdminUrl"
# Do not Import-Module first: let the cmdlet auto-load so it builds a fresh auth context.
Connect-SPOService -Url $AdminUrl -ModernAuth $true
Write-OK "Connected to $AdminUrl"
