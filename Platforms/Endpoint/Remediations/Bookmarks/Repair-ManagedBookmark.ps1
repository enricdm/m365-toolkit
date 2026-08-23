<#
.SYNOPSIS
    Intune remediation REPAIR: writes the managed-bookmarks policy for Microsoft
    Edge or Google Chrome.

.DESCRIPTION
    Repair half of a device-side remediation pair. Pair with
    Detect-ManagedBookmark.ps1.

    Writes the browser POLICY value:

      Edge   HKLM:\SOFTWARE\Policies\Microsoft\Edge   ManagedFavorites
      Chrome HKLM:\SOFTWARE\Policies\Google\Chrome    ManagedBookmarks

    Managed bookmarks appear in a dedicated, read-only folder on the bookmarks bar.
    They survive profile resets and cannot be deleted by the user - which is what
    you want on shared classroom, kiosk and meeting-room devices.

    WHY POLICY AND NOT THE BOOKMARKS FILE
    Several of the scripts this replaces edited
    %LOCALAPPDATA%\...\User Data\Default\Bookmarks directly, assigning
    $bar.children = @() before inserting the new set. That is destructive - it
    discards every bookmark the user had - and unreliable for three reasons:
      - Chrome/Edge hold the file open and rewrite it on exit, silently reverting
        the change if the browser was running.
      - The file carries a checksum; an externally edited file can be discarded.
      - It only ever affects the Default profile of the user who happened to run it.
    The policy has none of those problems. File seeding is still available via
    -SeedUserFile for the narrow case where users must be able to EDIT the seeded
    bookmarks, but it is off by default and warns about what it overwrites.

.PARAMETER Browser
    Edge (default) or Chrome.

.PARAMETER FolderName
    Top-level folder label shown on the bookmarks bar.

.PARAMETER Bookmark
    Ordered hashtables with Name and Url. Edit the default before deploying -
    Intune remediation scripts run with no arguments.

.PARAMETER SeedUserFile
    ALSO seed the current user's bookmark file so the entries are user-editable.
    Off by default. Appends only; never clears existing bookmarks. Skipped, with a
    message, if the browser is running.

.EXAMPLE
    # Deploy as-is after editing the default Bookmark list
    .\Repair-ManagedBookmark.ps1

.EXAMPLE
    # Chrome, custom set
    .\Repair-ManagedBookmark.ps1 -Browser Chrome -FolderName 'Contoso' -Bookmark @(
        @{ Name = 'Intranet'; Url = 'https://intranet.contoso.com' },
        @{ Name = 'Webmail';  Url = 'https://mail.contoso.com' }
    )

.NOTES
    Context  : run as SYSTEM (Intune remediation default), 64-bit
    Requires : Windows PowerShell 5.1
    Rights   : WRITES to HKLM. -SeedUserFile additionally writes to the user profile.

    The policy applies at next browser launch. Already-open windows keep the old
    set until restarted - a device that still shows the old bookmarks right after
    remediation has not necessarily failed.

    Replaces (merged): Set-ChromeBookmarks.ps1, Set-EdgeBookmarks.ps1,
    Remediate-ChromeBookmarks.ps1, ChromeBookmarksFix_Remediation.ps1
#>

[CmdletBinding()]
param(
    [ValidateSet('Edge', 'Chrome')]
    [string]$Browser = 'Edge',

    [string]$FolderName = 'Contoso',

    [hashtable[]]$Bookmark = @(
        @{ Name = 'Intranet';  Url = 'https://intranet.contoso.com' },
        @{ Name = 'Webmail';   Url = 'https://mail.contoso.com' },
        @{ Name = 'OneDrive';  Url = 'https://portal.office.com/onedrive' }
    ),

    [switch]$SeedUserFile
)

$cfg = switch ($Browser) {
    'Edge' {
        @{ Path='HKLM:\SOFTWARE\Policies\Microsoft\Edge'; Name='ManagedFavorites'
           Process='msedge'; File="$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Bookmarks" }
    }
    'Chrome' {
        @{ Path='HKLM:\SOFTWARE\Policies\Google\Chrome';  Name='ManagedBookmarks'
           Process='chrome'; File="$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Bookmarks" }
    }
}

try {
    # ---- validate input before touching anything ----
    $entries = New-Object System.Collections.Generic.List[object]
    foreach ($b in $Bookmark) {
        if (-not $b.Name -or -not $b.Url) {
            throw "Every -Bookmark entry needs both Name and Url. Offending entry: $($b | ConvertTo-Json -Compress)"
        }
        $entries.Add([ordered]@{ name = [string]$b.Name; url = [string]$b.Url })
    }
    if ($entries.Count -eq 0) { throw 'No bookmarks supplied; refusing to write an empty policy.' }

    # ---- build the policy payload ----
    # First element carries the folder label; the rest are the bookmarks.
    $payload = @([ordered]@{ toplevel_name = $FolderName }) + $entries.ToArray()
    $json    = $payload | ConvertTo-Json -Depth 10 -Compress

    if (-not (Test-Path $cfg.Path)) {
        New-Item -Path $cfg.Path -Force -ErrorAction Stop | Out-Null
        Write-Output "Created policy key $($cfg.Path)"
    }

    New-ItemProperty -Path $cfg.Path -Name $cfg.Name -Value $json `
                     -PropertyType String -Force -ErrorAction Stop | Out-Null

    Write-Output "Set $Browser policy '$($cfg.Name)' with $($entries.Count) bookmark(s) under folder '$FolderName'."

    # ---- optional: seed the user's own bookmark file ----
    if ($SeedUserFile) {
        $running = @(Get-Process -Name $cfg.Process -ErrorAction Ignore)
        if ($running.Count -gt 0) {
            Write-Output "SKIPPED -SeedUserFile: $($cfg.Process) is running and would overwrite the file on exit."
        }
        elseif (-not (Test-Path -LiteralPath $cfg.File)) {
            Write-Output "SKIPPED -SeedUserFile: no bookmark file yet at $($cfg.File) (profile never opened)."
        }
        else {
            # Back up before touching a file we did not create.
            $backup = "$($cfg.File).bak-$(Get-Date -Format 'yyyyMMddHHmmss')"
            Copy-Item -LiteralPath $cfg.File -Destination $backup -ErrorAction Stop
            Write-Output "Backed up existing bookmarks to $backup"

            $doc = Get-Content -LiteralPath $cfg.File -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            $bar = $doc.roots.bookmark_bar
            if ($null -eq $bar) { throw 'Bookmark file has no roots.bookmark_bar node.' }

            $existing = @($bar.children | ForEach-Object { $_.url })
            $added    = 0
            foreach ($e in $entries) {
                # append only - never clear what the user already had
                if ($existing -contains $e.url) { continue }
                $bar.children += [pscustomobject]@{ type = 'url'; name = $e.name; url = $e.url }
                $added++
            }

            if ($added -gt 0) {
                # Chrome/Edge validate a checksum over the file. Removing the stale
                # value makes the browser recompute it instead of rejecting the file.
                if ($doc.PSObject.Properties['checksum']) { $doc.PSObject.Properties.Remove('checksum') }
                $doc | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $cfg.File -Encoding UTF8 -ErrorAction Stop
                Write-Output "Seeded $added new bookmark(s) into the user's bookmark bar."
            } else {
                Write-Output 'User bookmark file already contained every URL; left unchanged.'
            }
        }
    }

    Write-Output 'Policy applies at next browser launch.'
    exit 0
}
catch {
    Write-Output "ERROR: $($_.Exception.Message)"
    exit 1
}
