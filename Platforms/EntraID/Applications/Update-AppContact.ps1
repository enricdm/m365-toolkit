<#
.SYNOPSIS
    Enriches the App Registration tracking Excel with a "Responsible contact" per app.

.DESCRIPTION
    Resolves a contact email per app using hierarchical signal priority:

      Tier 1 (95) — SAML notificationEmails  (admin accounts resolved via cache)
      Tier 2 (85) — Owner, corporate mailbox (real human, real domain)
      Tier 3 (75) — Owner, admin account     (resolved via Graph user lookup)
      Tier 4 (55) — Country-team alias       (from naming convention scope)
      Tier 5 (35) — Homepage TLD             (country-team via TLD inference)
      Tier 6 (0)  — None — manual review

    Reinforcement bonus: +5 confidence when a lower tier independently agrees
    with the chosen tier's country.

    Writes three columns: "Responsible contact", "Contact source", "Contact confidence".
    Existing data in those columns is preserved unless -Force is specified.

.PARAMETER ExcelPath
    App-registration tracking workbook (.xlsx) with an 'App Tracking' worksheet.

.PARAMETER SamlCsvPath
    Path to the SAML notification-email CSV produced by Export-SamlNotificationEmail.ps1.

.PARAMETER SpCsvPath
    Optional service-principal export CSV (needs 'appId' and 'homepage' columns).
    Without it Tier 5 (homepage TLD inference) is simply skipped.

.PARAMETER DomainMapPath
    Domain/country data file. Defaults to the sample map in Platforms\_Shared\Data —
    replace it with your own domains before relying on Tiers 3-5.

.PARAMETER AdminAccountDomain
    Domain that admin (non-mailbox) accounts live on. Contacts on this domain are
    resolved back to the human's real mailbox through Graph.

.PARAMETER CorporateDomainPattern
    Regex identifying a real corporate mailbox address.

.PARAMETER OutputPath
    Where to write the enriched Excel. Defaults to <ExcelPath>_enriched.xlsx.

.PARAMETER WhatIf
    Run full resolution and print summary, but do not write the Excel.

.PARAMETER Force
    Overwrite existing values in "Responsible contact" column.

.EXAMPLE
    .\Update-AppContact.ps1 -ExcelPath .\AppTracking.xlsx -SamlCsvPath .\SAML_Notification_Emails.csv -WhatIf
    Dry run. Review the resolution summary before committing.

.EXAMPLE
    .\Update-AppContact.ps1 -ExcelPath .\AppTracking.xlsx -SamlCsvPath .\SAML_Notification_Emails.csv `
        -SpCsvPath .\servicePrincipals.csv -DomainMapPath .\my-domains.psd1
    Real run with all signals. Writes <input>_enriched.xlsx alongside the source.

.NOTES
    When to use  : The app inventory has 200 rows with no owner and someone has to be called before anything is touched.
    Why it exists: Resolves a responsible human through five ranked signals with a confidence score, resolves admin accounts back to the person's real mailbox through Graph, and outputs the ones it could not resolve as a separate manual-review list. That is a decision, not a data dump.
    Requires: ImportExcel, Microsoft.Graph.Authentication, Microsoft.Graph.Users
    Permissions: Directory.Read.All (Global Reader role is sufficient)
    Connect to the tenant that holds the app registrations. Admin accounts that
    live in a separate tenant must be present as guests for Tier 1/3 to resolve.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    # App-registration tracking workbook. No copy exists in this repo, so it must be supplied.
    [Parameter(Mandatory)]
    [string]$ExcelPath,
    # Output of Export-SamlNotificationEmail.ps1 (same folder as this script). No live copy
    # in the repo either — point this at the CSV that script produced.
    [Parameter(Mandatory)]
    [string]$SamlCsvPath,
    [string]$SpCsvPath,
    [string]$DomainMapPath = (Join-Path $PSScriptRoot '..\..\_Shared\Data\domain-country-map.psd1'),
    [string]$AdminAccountDomain = 'contoso.onmicrosoft.com',
    [string]$CorporateDomainPattern = '@contoso[.-]',
    [string]$OutputPath,
    [switch]$Force
)

# ─── Setup ──────────────────────────────────────────────────────────────────

$ErrorActionPreference = 'Stop'

if (-not $OutputPath) {
    $OutputPath = $ExcelPath -replace '\.xlsx$','_enriched.xlsx'
}

# Domain/country data lives outside the script so it can be swapped per tenant.
#   CountryToDomains  country code → preferred email domains, tried in order
#   TldToCountry      homepage TLD → country code (Tier 5 signal)
#   TeamDomainToCode  team-alias domain → country code (Tier 4 derivation)
if (-not (Test-Path $DomainMapPath)) {
    throw "Domain map not found: $DomainMapPath. Pass -DomainMapPath pointing at your own copy."
}
$DomainMap = Import-PowerShellDataFile -Path $DomainMapPath
$CountryDomainPreference = $DomainMap.CountryToDomains
$TldToCountry            = $DomainMap.TldToCountry
$TeamDomainToCode        = $DomainMap.TeamDomainToCode
Write-Host "Domain map: $DomainMapPath ($($CountryDomainPreference.Count) country codes)" -ForegroundColor DarkGray

# Fallback domains when a country code is not in the map at all.
$DefaultDomains = @($DomainMap.CentralDomains | Select-Object -First 1)

# Anything matching these is filtered out as a contact candidate
$ExcludePatterns = @(
    'noreply', 'donotreply', 'no-reply', 'do-not-reply',
    'powerautomate', 'pva-', 'flow-',
    '#ext#@'   # external guest accounts
)

# ─── Module / connection ────────────────────────────────────────────────────

foreach ($mod in 'ImportExcel','Microsoft.Graph.Authentication','Microsoft.Graph.Users') {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Host "Installing module $mod..." -ForegroundColor Yellow
        Install-Module $mod -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $mod -ErrorAction Stop
}

if (-not (Get-MgContext)) {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes 'Directory.Read.All' -NoWelcome
}
$ctx = Get-MgContext
Write-Host "Connected as $($ctx.Account) | tenant $($ctx.TenantId)" -ForegroundColor Green

# ─── Load data ──────────────────────────────────────────────────────────────

Write-Host "`nLoading data..." -ForegroundColor Cyan
$apps  = Import-Excel -Path $ExcelPath -WorksheetName 'App Tracking'
$saml  = Import-Csv  -Path $SamlCsvPath
$sps   = if ($SpCsvPath) { Import-Csv -Path $SpCsvPath } else { @() }
Write-Host "  Apps:                $($apps.Count)"
Write-Host "  SAML notif rows:     $($saml.Count)"
if ($SpCsvPath) { Write-Host "  Service principals:  $($sps.Count)" }
else            { Write-Host "  Service principals:  (no -SpCsvPath — Tier 5 disabled)" -ForegroundColor DarkGray }

# Index SAML CSV by AppId
$samlByAppId = @{}
foreach ($row in $saml) {
    if ($row.AppId) { $samlByAppId[$row.AppId] = $row }
}

# Index SP export by AppId
$spByAppId = @{}
foreach ($row in $sps) {
    if ($row.appId) { $spByAppId[$row.appId] = $row }
}

# ─── Helpers ────────────────────────────────────────────────────────────────

function Test-IsExcluded {
    param([string]$Email)
    if ([string]::IsNullOrWhiteSpace($Email)) { return $true }
    $e = $Email.ToLower()
    foreach ($p in $ExcludePatterns) {
        if ($e -like "*$p*") { return $true }
    }
    return $false
}

$AdminDomainRegex = '@' + [regex]::Escape($AdminAccountDomain.ToLower()) + '$'

function Test-IsAdminAccount {
    param([string]$Email)
    if (-not $Email) { return $false }
    return ($Email.ToLower() -match $AdminDomainRegex)
}

function Test-IsCorporateMailbox {
    param([string]$Email)
    if (-not $Email) { return $false }
    $e = $Email.ToLower()
    return ($e -match $CorporateDomainPattern) -and -not (Test-IsAdminAccount $e) -and -not (Test-IsExcluded $e)
}

function Get-EmailsFromOwnerString {
    param([string]$OwnerString)
    if ([string]::IsNullOrWhiteSpace($OwnerString)) { return @() }
    # Owner cells look like: "Display Name <email@x>; Other Name <email2@y>"
    [regex]::Matches($OwnerString, '<([^>]+)>') |
        ForEach-Object { $_.Groups[1].Value.Trim().ToLower() }
}

function Get-ScopeFromName {
    param([string]$SuggestedName)
    if (-not $SuggestedName) { return $null }
    if ($SuggestedName -match '^([A-Z]{2})-') { return $matches[1] }
    return $null
}

# ─── Admin-account resolution cache (Graph lookup) ──────────────────────────

$adminCache = @{}  # admin-account email (lower) -> corporate mailbox or $null

function Resolve-AdminAccount {
    param([string]$AdminEmail)

    if (-not $AdminEmail) { return $null }
    $key = $AdminEmail.ToLower()
    if ($adminCache.ContainsKey($key)) { return $adminCache[$key] }

    # Look up the admin account in Graph
    try {
        $u = Get-MgUser -UserId $key -Property 'givenName,surname,displayName,mail,userPrincipalName,otherMails' -ErrorAction Stop
    } catch {
        Write-Verbose "  [ADM] $key NOT FOUND in Graph — $($_.Exception.Message.Split([Environment]::NewLine)[0])"
        $adminCache[$key] = $null
        return $null
    }

    $given   = if ($u.GivenName) { $u.GivenName.Trim() } else { '' }
    $surname = if ($u.Surname)   { $u.Surname.Trim() }   else { '' }

    # If account has otherMails populated with a corporate mailbox, prefer that
    foreach ($m in @($u.OtherMails)) {
        if (Test-IsCorporateMailbox $m) {
            Write-Verbose "  [ADM] $key -> $($m.ToLower())  (from otherMails)"
            $adminCache[$key] = $m.ToLower()
            return $adminCache[$key]
        }
    }

    # Mail attribute might already be the corporate mailbox
    if ($u.Mail -and (Test-IsCorporateMailbox $u.Mail)) {
        Write-Verbose "  [ADM] $key -> $($u.Mail.ToLower())  (from mail attr)"
        $adminCache[$key] = $u.Mail.ToLower()
        return $adminCache[$key]
    }

    # Otherwise construct candidates from givenName.surname
    if (-not $given -or -not $surname) {
        Write-Verbose "  [ADM] $key — no givenName/surname populated, can't construct candidate"
        $adminCache[$key] = $null
        return $null
    }

    # Country code from the admin-account naming prefix (e.g. 'es-adm-doe@...')
    $cc = $null
    if ($key -match '^([a-z]{2,3})-adm-') { $cc = $matches[1].ToUpper() }
    $domains = $CountryDomainPreference[$cc]
    if (-not $domains) { $domains = $DefaultDomains }

    $local = "{0}.{1}" -f $given.ToLower(), $surname.ToLower()
    # Strip diacritics: NFD decomposition removes accents, then keep only ASCII letters/dots/hyphens
    $local = ($local.Normalize([Text.NormalizationForm]::FormD).ToCharArray() |
              Where-Object { [Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne [Globalization.UnicodeCategory]::NonSpacingMark }) -join ''
    $local = $local -replace '[^a-z0-9.-]',''

    foreach ($d in $domains) {
        $candidate = "$local@$d"
        try {
            $check = Get-MgUser -UserId $candidate -Property 'userPrincipalName' -ErrorAction Stop
            if ($check) {
                Write-Verbose "  [ADM] $key -> $candidate  (verified)"
                $adminCache[$key] = $candidate
                return $candidate
            }
        } catch {
            # Not found — try next domain
        }
    }

    Write-Verbose "  [ADM] $key — could not resolve (tried: $($domains -join ', '))"
    $adminCache[$key] = $null
    return $null
}

# ─── Country team map (auto-derived from SAML CSV) ──────────────────────────

# Multi-occurrence team aliases in SAML notif emails = country/team mailboxes
$countryTeamMap = @{}
$emailFreq = @{}
foreach ($row in $saml) {
    if (-not $row.NotificationEmails) { continue }
    foreach ($e in ($row.NotificationEmails -split ';')) {
        $e = $e.Trim().ToLower()
        if (-not $e -or (Test-IsAdminAccount $e) -or (Test-IsExcluded $e)) { continue }
        if (-not (Test-IsCorporateMailbox $e)) { continue }
        # Only team/alias mailboxes (no firstname.surname pattern)
        if ($e -match '^[a-z]+\.[a-z]+@') { continue }  # firstname.surname pattern
        if (-not $emailFreq.ContainsKey($e)) { $emailFreq[$e] = 0 }
        $emailFreq[$e]++
    }
}

# Country team email = team alias appearing 2+ times, mapped by domain
foreach ($e in $emailFreq.Keys) {
    if ($emailFreq[$e] -lt 2) { continue }
    $domain = ($e -split '@')[1]
    # TeamDomainToCode maps a team-alias domain to the country team that owns it;
    # the central domain maps to the global team rather than to a country.
    $cc = $null
    foreach ($d in $TeamDomainToCode.Keys) {
        if ($domain -eq $d.ToLower()) { $cc = $TeamDomainToCode[$d]; break }
    }
    if (-not $cc) { continue }
    if (-not $countryTeamMap.ContainsKey($cc)) { $countryTeamMap[$cc] = @() }
    $countryTeamMap[$cc] += $e
}

Write-Host "`nCountry team aliases auto-derived:" -ForegroundColor Cyan
foreach ($cc in ($countryTeamMap.Keys | Sort-Object)) {
    Write-Host ("  {0}: {1}" -f $cc, ($countryTeamMap[$cc] -join '; '))
}

# ─── Main resolution loop ───────────────────────────────────────────────────

Write-Host "`nResolving contacts for $($apps.Count) apps..." -ForegroundColor Cyan
Write-Host "(use -Verbose to see admin-account resolution decisions)" -ForegroundColor DarkGray

$tierCounts = @{ 1=0; 2=0; 3=0; 4=0; 5=0; 6=0 }
$results = @()

$progress = 0
foreach ($app in $apps) {
    $progress++
    if ($progress % 50 -eq 0) {
        Write-Progress -Activity "Resolving" -Status "$progress / $($apps.Count)" -PercentComplete (($progress/$apps.Count)*100)
    }

    # Skip if already filled (unless -Force)
    if ($app.'Responsible contact' -and -not $Force) {
        $results += [PSCustomObject]@{
            AppId      = $app.'App ID'
            Tier       = 0
            Source     = 'preserved'
            Contact    = $app.'Responsible contact'
            Confidence = $null
        }
        continue
    }

    $appId  = $app.'App ID'
    $scope  = Get-ScopeFromName $app.'Suggested name'
    $contact = $null
    $source  = $null
    $confidence = 0
    $tier = 6

    # ── Tier 1: SAML notificationEmails ─────────────────────────────────────
    if ($appId -and $samlByAppId.ContainsKey($appId)) {
        $samlRow = $samlByAppId[$appId]
        if ($samlRow.NotificationEmails) {
            $candidates = @()
            foreach ($e in ($samlRow.NotificationEmails -split ';')) {
                $e = $e.Trim().ToLower()
                if (-not $e -or (Test-IsExcluded $e)) { continue }
                if (Test-IsAdminAccount $e) {
                    $resolved = Resolve-AdminAccount $e
                    if ($resolved) { $candidates += $resolved }
                } elseif (Test-IsCorporateMailbox $e) {
                    $candidates += $e
                }
            }
            $candidates = $candidates | Select-Object -Unique
            if ($candidates.Count -gt 0) {
                $contact = $candidates[0]
                $source = 'SAML notif'
                $tier = 1
                $confidence = 95
            }
        }
    }

    # ── Tier 2: Owner — direct corporate mailbox ────────────────────────────
    if (-not $contact -and $app.'Current owner(s)') {
        $emails = Get-EmailsFromOwnerString $app.'Current owner(s)'
        $nominals = @($emails | Where-Object { Test-IsCorporateMailbox $_ })
        if ($nominals.Count -gt 0) {
            $contact = $nominals[0]
            $source = 'Owner (nominal)'
            $tier = 2
            $confidence = 85
        }
    }

    # ── Tier 3: Owner — admin account resolved through Graph ────────────────
    if (-not $contact -and $app.'Current owner(s)') {
        $emails = Get-EmailsFromOwnerString $app.'Current owner(s)'
        $admins = @($emails | Where-Object { Test-IsAdminAccount $_ })
        foreach ($r in $admins) {
            $resolved = Resolve-AdminAccount $r
            if ($resolved) {
                $contact = $resolved
                $source = "Owner (admin→$r)"
                $tier = 3
                $confidence = 75
                break
            }
        }
    }

    # ── Tier 4: Country-team alias from naming scope ────────────────────────
    if (-not $contact -and $scope -and $countryTeamMap.ContainsKey($scope)) {
        $contact = $countryTeamMap[$scope][0]  # first alias for that country
        $source = "Country-team ($scope)"
        $tier = 4
        $confidence = 55
    }

    # ── Tier 5: Homepage TLD inference ──────────────────────────────────────
    if (-not $contact -and $appId -and $spByAppId.ContainsKey($appId)) {
        $hp = $spByAppId[$appId].homepage
        if ($hp) {
            foreach ($tld in $TldToCountry.Keys) {
                if ($hp.ToLower() -match "\.$([regex]::Escape($tld))(/|:|$)") {
                    $cc = $TldToCountry[$tld]
                    if ($countryTeamMap.ContainsKey($cc)) {
                        $contact = $countryTeamMap[$cc][0]
                        $source = "Homepage TLD (.$tld → $cc)"
                        $tier = 5
                        $confidence = 35
                        break
                    }
                }
            }
        }
    }

    # ── Reinforcement bonus ─────────────────────────────────────────────────
    # If a different signal independently agrees with the chosen country, +5
    if ($contact -and $tier -le 3 -and $scope) {
        $contactDomain = ($contact -split '@')[1]
        $expectedDomains = $CountryDomainPreference[$scope]
        if ($expectedDomains -and $expectedDomains -contains $contactDomain) {
            $confidence = [Math]::Min(100, $confidence + 5)
        }
    }

    $tierCounts[$tier]++

    if (-not $contact) {
        $source = 'manual review needed'
    }

    $results += [PSCustomObject]@{
        AppId      = $appId
        Tier       = $tier
        Source     = $source
        Contact    = $contact
        Confidence = $confidence
    }

    # Patch the row in-place
    $app.'Responsible contact' = $contact
    if ($app.PSObject.Properties['Contact source']) {
        $app.'Contact source' = $source
    } else {
        $app | Add-Member -NotePropertyName 'Contact source' -NotePropertyValue $source -Force
    }
    if ($app.PSObject.Properties['Contact confidence']) {
        $app.'Contact confidence' = $confidence
    } else {
        $app | Add-Member -NotePropertyName 'Contact confidence' -NotePropertyValue $confidence -Force
    }
}
Write-Progress -Activity "Resolving" -Completed

# ─── Summary ────────────────────────────────────────────────────────────────

Write-Host "`n──────── Resolution summary ────────" -ForegroundColor Green
Write-Host ("Tier 1 (SAML notif, conf 95):       {0,3}" -f $tierCounts[1])
Write-Host ("Tier 2 (Owner nominal, conf 85):    {0,3}" -f $tierCounts[2])
Write-Host ("Tier 3 (Owner admin resolved, conf 75): {0,3}" -f $tierCounts[3])
Write-Host ("Tier 4 (Country team, conf 55):     {0,3}" -f $tierCounts[4])
Write-Host ("Tier 5 (Homepage TLD, conf 35):     {0,3}" -f $tierCounts[5])
Write-Host ("Tier 6 (manual review):             {0,3}" -f $tierCounts[6]) -ForegroundColor Yellow
Write-Host ("Total:                              {0,3}" -f $apps.Count)

$resolved = $apps.Count - $tierCounts[6]
$pct = [Math]::Round(100*$resolved/$apps.Count,1)
Write-Host "`n=> Resolved: $resolved / $($apps.Count) ($pct%)" -ForegroundColor Green

# Admin-account cache report
$admHits = ($adminCache.Values | Where-Object { $_ }).Count
$admMiss = ($adminCache.Values | Where-Object { -not $_ }).Count
Write-Host "`nAdmin-account lookups:  $admHits resolved, $admMiss unresolvable (out of $($adminCache.Count) unique accounts seen)"

if ($admMiss -gt 0) {
    Write-Host "`nUnresolved admin accounts (need manual handling):" -ForegroundColor Yellow
    $adminCache.GetEnumerator() | Where-Object { -not $_.Value } | ForEach-Object {
        Write-Host "  $($_.Key)"
    }
}

# Manual review apps
$manualApps = $results | Where-Object { $_.Tier -eq 6 }
if ($manualApps.Count -gt 0) {
    $manualReportPath = $OutputPath -replace '\.xlsx$','_manual_review.csv'
    $manualExport = foreach ($r in $manualApps) {
        $a = $apps | Where-Object { $_.'App ID' -eq $r.AppId } | Select-Object -First 1
        [PSCustomObject]@{
            AppId         = $r.AppId
            CurrentName   = $a.'Current name'
            SuggestedName = $a.'Suggested name'
            Status        = $a.Status
            Owners        = $a.'Current owner(s)'
        }
    }
    $manualExport | Export-Csv -Path $manualReportPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nManual review list written to: $manualReportPath" -ForegroundColor Yellow
}

# ─── Write Excel ────────────────────────────────────────────────────────────

if ($PSCmdlet.ShouldProcess($OutputPath, 'Write enriched Excel')) {
    Write-Host "`nWriting enriched Excel..." -ForegroundColor Cyan
    # Copy original to preserve Instructions sheet and formatting
    Copy-Item -Path $ExcelPath -Destination $OutputPath -Force
    # Overwrite the App Tracking sheet with enriched data
    $apps | Export-Excel -Path $OutputPath -WorksheetName 'App Tracking' -ClearSheet -AutoSize -FreezeTopRow -BoldTopRow
    Write-Host "  Saved to: $OutputPath" -ForegroundColor Green
} else {
    Write-Host "`n[WhatIf] Excel NOT written. Re-run without -WhatIf to commit." -ForegroundColor Magenta
}

Write-Host "`nDone." -ForegroundColor Green
