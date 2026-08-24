<#
.SYNOPSIS
    Sends "SharePoint Site Storage over 85%" notifications to the owners of sites
    whose Resolution Status in the tracker is still "No Reply".

.DESCRIPTION
    Reads the weekly SPO storage tracker (.xlsx), filters rows by Resolution Status,
    resolves recipients (Primary Contact + All Owners), and sends a templated HTML
    notification from the team shared mailbox.

    AUTH: app-only with a certificate, done WITHOUT the Microsoft.Graph PowerShell SDK.
    The script signs a JWT client assertion with the cert and calls the Graph REST
    endpoint directly (Invoke-RestMethod). This avoids the Microsoft.Identity.Client
    (MSAL) assembly conflicts that occur when PnP.PowerShell / Az are loaded in the
    same session as Connect-MgGraph. Only the ImportExcel module is required.

    The workbook is read with -NoHeader and the header row is auto-detected, so the
    decorative two-line column titles never trip a parser and the script keeps working
    whether or not the banner rows survive an Excel re-save.

    Supports -WhatIf: it still acquires a token (so a dry run confirms cert auth and
    permissions are healthy) but sends nothing. Re-run without -WhatIf to send.

    *** REQUIRED APP PERMISSION ***
    The app registration needs Microsoft Graph >> Mail.Send (Application) with admin
    consent. Scope it to the sender mailbox with an Application Access Policy:
        New-ApplicationAccessPolicy -AppId <clientId> `
            -PolicyScopeGroupId <mail-enabled security group containing the shared mailbox> `
            -AccessRight RestrictAccess -Description "Restrict workplace-reports app to shared mailbox"

.PARAMETER ExcelPath           Path to the tracker workbook.
.PARAMETER WorksheetName       Worksheet holding the data. Default: "Storage Tracker".
.PARAMETER HeaderRow           1-based header row, or 0 = auto-detect (default).
.PARAMETER StatusFilter        Substring matched against Resolution Status. Default: "No Reply".
.PARAMETER SenderMailbox       Mailbox the notification is sent from (mandatory).
.PARAMETER CcMailbox           CC on every notification (team copy). Default: the sender mailbox.
.PARAMETER TenantId            Entra tenant ID.
.PARAMETER ClientId            App (client) ID.
.PARAMETER CertificateThumbprint  Thumbprint of the auth certificate (uppercase hex, no spaces).
.PARAMETER CertStoreLocation   'CurrentUser' (default) or 'LocalMachine'. Both are searched if not found.
.PARAMETER MaxSends            Safety cap on notifications per run. Default: 0 (no cap).
.PARAMETER DelayMs             Pause between sends to avoid throttling. Default: 500.
.PARAMETER SupportTicketUrl    URL behind the "raise a ticket" link in the email body.
.PARAMETER TeamName            Team name used in the email signature.
.PARAMETER LogPath             CSV run log. Default: timestamped file next to the workbook.

.EXAMPLE
    # Dry run: acquires a token (proving cert auth and Mail.Send work) but sends nothing
    .\Send-StorageNotification.ps1 -ExcelPath '<path-to>\StorageTracker.xlsx' `
        -SenderMailbox 'admin@contoso.com' `
        -TenantId '<tenant-id>' -ClientId '<client-id>' `
        -CertificateThumbprint '<cert-thumbprint>' -WhatIf

.EXAMPLE
    # Send for real, capped at 10 notifications on the first live run
    .\Send-StorageNotification.ps1 -ExcelPath '<path-to>\StorageTracker.xlsx' `
        -SenderMailbox 'admin@contoso.com' `
        -TenantId '<tenant-id>' -ClientId '<client-id>' `
        -CertificateThumbprint '<cert-thumbprint>' -MaxSends 10

.NOTES
    When to use  : The weekly round of notifications to owners of sites over 85%, without opening Outlook.
    Why it exists: Signs the client assertion JWT with the certificate and calls the Graph REST endpoint directly, avoiding the MSAL assembly conflicts that happen when PnP.PowerShell or Az are loaded in the same session. -WhatIf still acquires the token, so a dry run proves certificate auth and Mail.Send are healthy while sending nothing, and -MaxSends caps the run.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory)] [string] $ExcelPath,
    [string] $WorksheetName = 'Storage Tracker',
    [int]    $HeaderRow = 0,
    [string] $StatusFilter = 'No Reply',
    [Parameter(Mandatory)] [string] $SenderMailbox,
    [string] $CcMailbox,
    [Parameter(Mandatory)] [string] $TenantId,
    [Parameter(Mandatory)] [string] $ClientId,
    [Parameter(Mandatory)] [string] $CertificateThumbprint,
    [ValidateSet('CurrentUser', 'LocalMachine')] [string] $CertStoreLocation = 'CurrentUser',
    [int]    $MaxSends = 0,
    [int]    $DelayMs = 500,
    [string] $SupportTicketUrl = 'https://contoso.sharepoint.com/sites/Example',
    [string] $TeamName         = 'IT Support',
    [string] $LogPath
)

# --------------------------- Setup ---------------------------
$ErrorActionPreference = 'Stop'
if (-not $CcMailbox) { $CcMailbox = $SenderMailbox }
if (-not $LogPath) {
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $LogPath = Join-Path (Split-Path -Parent (Resolve-Path $ExcelPath)) "SPO_Notify_Log_$stamp.csv"
}
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    throw "Required module 'ImportExcel' is not installed.  Install-Module ImportExcel -Scope CurrentUser"
}

# ---------------- Cert lookup + JWT client-assertion auth ----------------
function Get-AuthCertificate {
    param([string]$Thumbprint, [string]$PreferredLocation)
    $locations = @($PreferredLocation, 'CurrentUser', 'LocalMachine') | Select-Object -Unique
    foreach ($loc in $locations) {
        $path = "Cert:\$loc\My\$Thumbprint"
        if (Test-Path $path) { return Get-Item $path }
    }
    throw "Certificate $Thumbprint not found in CurrentUser\My or LocalMachine\My."
}

function ConvertTo-Base64Url {
    param([byte[]]$Bytes)
    [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

function Get-GraphToken {
    # Builds a signed JWT assertion and exchanges it for a Graph access token. No MSAL.
    param([string]$TenantId, [string]$ClientId, $Cert)

    $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Cert)
    if (-not $rsa) { throw "Could not access the certificate's private key (thumbprint $($Cert.Thumbprint))." }

    # x5t = base64url of the cert's SHA-1 hash (thumbprint bytes)
    $thumbBytes = for ($i = 0; $i -lt $Cert.Thumbprint.Length; $i += 2) {
        [Convert]::ToByte($Cert.Thumbprint.Substring($i, 2), 16)
    }
    $x5t = ConvertTo-Base64Url ([byte[]]$thumbBytes)

    $aud = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $now = [System.DateTimeOffset]::UtcNow.ToUnixTimeSeconds()

    $headerJson = @{ alg = 'RS256'; typ = 'JWT'; x5t = $x5t } | ConvertTo-Json -Compress
    $payloadJson = @{
        aud = $aud; iss = $ClientId; sub = $ClientId
        jti = [guid]::NewGuid().ToString(); nbf = $now; exp = $now + 600
    } | ConvertTo-Json -Compress

    $h = ConvertTo-Base64Url ([System.Text.Encoding]::UTF8.GetBytes($headerJson))
    $p = ConvertTo-Base64Url ([System.Text.Encoding]::UTF8.GetBytes($payloadJson))
    $unsigned = "$h.$p"
    $sig = $rsa.SignData(
        [System.Text.Encoding]::ASCII.GetBytes($unsigned),
        [System.Security.Cryptography.HashAlgorithmName]::SHA256,
        [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
    $jwt = "$unsigned." + (ConvertTo-Base64Url $sig)

    $body = @{
        client_id             = $ClientId
        scope                 = 'https://graph.microsoft.com/.default'
        client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
        client_assertion      = $jwt
        grant_type            = 'client_credentials'
    }
    $resp = Invoke-RestMethod -Method POST -Uri $aud -Body $body -ContentType 'application/x-www-form-urlencoded'
    return $resp.access_token
}

function Send-GraphMail {
    param([string]$Sender, [string[]]$To, [string]$Cc, [string]$Subject, [string]$BodyHtml, [string]$Token)
    $payload = @{
        message         = @{
            subject      = $Subject
            body         = @{ contentType = 'HTML'; content = $BodyHtml }
            toRecipients = @($To | ForEach-Object { @{ emailAddress = @{ address = $_ } } })
            ccRecipients = @(@{ emailAddress = @{ address = $Cc } })
        }
        saveToSentItems = $true
    }
    $json = $payload | ConvertTo-Json -Depth 8
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    Invoke-RestMethod -Method POST `
        -Uri "https://graph.microsoft.com/v1.0/users/$Sender/sendMail" `
        -Headers @{ Authorization = "Bearer $Token" } `
        -ContentType 'application/json; charset=utf-8' -Body $bytes | Out-Null
}

# ---------------- Header detection + column resolver ----------------
function Get-NormText { param([object]$v) (([string]$v) -replace '\s', '').ToLower() }

function Find-HeaderIndex {
    param([object[]]$Rows, [int]$Max = 15)
    $limit = [Math]::Min($Max, $Rows.Count)
    for ($i = 0; $i -lt $limit; $i++) {
        $joined = ($Rows[$i].PSObject.Properties.Value | ForEach-Object { Get-NormText $_ }) -join '|'
        if ($joined -match 'sitetitle' -and $joined -match 'resolutionstatus') { return $i }
    }
    return -1
}
function Build-HeaderIndex {
    param([object]$HeaderObj)
    $idx = @{}
    foreach ($pr in $HeaderObj.PSObject.Properties) {
        $n = Get-NormText $pr.Value
        if ($n -and -not $idx.ContainsKey($n)) { $idx[$n] = $pr.Name }
    }
    return $idx
}
function Get-Col {
    param([hashtable]$Index, [string]$Token)
    $t = ($Token -replace '\s', '').ToLower()
    foreach ($k in $Index.Keys) { if ($k -like "*$t*") { return $Index[$k] } }
    return $null
}
function Get-Val {
    param([object]$Row, [hashtable]$Index, [string]$Token)
    $col = Get-Col -Index $Index -Token $Token
    if ($col) { return [string]$Row.$col } else { return $null }
}

# ---------------- Recipient parsing ----------------
$emailRegex = '^[^@\s]+@[^@\s]+\.[^@\s]+$'
function Get-Recipients {
    param([string]$Primary, [string]$AllOwners)
    $list = @()
    if ($Primary) { $list += $Primary }
    if ($AllOwners) { $list += ($AllOwners -split '[;,]') }
    $list | ForEach-Object { $_.Trim() } |
    Where-Object { $_ -match $emailRegex } |
    Sort-Object -Unique
}

# ---------------- Email body ----------------
function New-Body {
    param([string]$SiteName, [string]$SiteUrl, [string]$PercentText)
    @"
<div style="font-family: Calibri,Arial,sans-serif; font-size: 11pt; color: #000;">
<p>Dear Site Owner/s,</p>

<p>Following up on our earlier notice: the SharePoint site below is <strong>still over 85% of its storage capacity</strong> and hasn't yet been actioned.</p>

<p><strong>Site</strong> – $SiteName<br>
<strong>URL</strong> – <a href="$SiteUrl">$SiteUrl</a><br>
<strong>Current usage</strong> – $PercentText</p>

<p>Once a site hits its limit, uploads and edits stop working and OneDrive/Teams sync begins to fail, so it's worth resolving before that point. There are two ways to clear it:</p>
<ul>
<li><strong>Free up space</strong> – review the site and remove outdated or unnecessary files.</li>
<li><strong>Request more storage</strong> – raise a <a href="$SupportTicketUrl">support ticket</a> if the content genuinely needs to stay.</li>
</ul>

<p>If you'd like help working out what's taking up the space, just reply and the <strong>$TeamName</strong> team will assist.</p>

<p>Kind regards,<br><strong>$TeamName</strong></p>
</div>
"@
}
function Format-Percent {
    param([string]$Raw)
    if (-not $Raw) { return 'over 85%' }
    $n = 0.0
    if ([double]::TryParse((($Raw -replace '[^\d\.,]', '') -replace ',', '.'), [ref]$n)) {
        if ($n -le 1) { $n = $n * 100 }
        return ('{0:N2}' -f $n)
    }
    return $Raw
}

# --------------------------- Run ---------------------------
$subject = '[Action Required] - SharePoint Site Storage over 85%'

Write-Host "Reading $ExcelPath  (sheet '$WorksheetName')" -ForegroundColor Cyan
$all = @(Import-Excel -Path $ExcelPath -WorksheetName $WorksheetName -NoHeader)
if (-not $all) { throw "No rows read from worksheet '$WorksheetName'." }

if ($HeaderRow -gt 0) {
    $hIdx = $HeaderRow - 1
} else {
    $hIdx = Find-HeaderIndex -Rows $all
    if ($hIdx -lt 0) { throw "Could not auto-detect header row. Pass -HeaderRow explicitly." }
    Write-Host ("Header row auto-detected at row {0}" -f ($hIdx + 1)) -ForegroundColor DarkGray
}

$index = Build-HeaderIndex -HeaderObj $all[$hIdx]
foreach ($need in 'Site Title', 'Site URL', 'Resolution Status', 'Primary Contact Email', 'All Owners') {
    if (-not (Get-Col -Index $index -Token $need)) { Write-Warning "Column matching '$need' not found." }
}

$dataRows = if ($all.Count -gt ($hIdx + 1)) { $all[($hIdx + 1)..($all.Count - 1)] } else { @() }

$targets = @($dataRows | Where-Object {
        $status = Get-Val -Row $_ -Index $index -Token 'Resolution Status'
        $url = Get-Val -Row $_ -Index $index -Token 'Site URL'
        $status -and $url -and ($status -like "*$StatusFilter*")
    })

Write-Host ("Sites matching status '*{0}*': {1}" -f $StatusFilter, $targets.Count) -ForegroundColor Yellow
if (-not $targets) { Write-Host 'Nothing to do.' -ForegroundColor Green; return }

# Acquire token up front (validates cert + Mail.Send even on a -WhatIf dry run)
Write-Host 'Acquiring Graph token (cert client-assertion, no SDK)...' -ForegroundColor Cyan
$cert = Get-AuthCertificate -Thumbprint $CertificateThumbprint -PreferredLocation $CertStoreLocation
$token = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -Cert $cert
Write-Host 'Token acquired.' -ForegroundColor Green

$log = New-Object System.Collections.Generic.List[object]
$sent = 0
$capped = $false

foreach ($row in $targets) {
    if ($MaxSends -gt 0 -and $sent -ge $MaxSends) { $capped = $true; break }

    $siteName = Get-Val -Row $row -Index $index -Token 'Site Title'
    $siteUrl = Get-Val -Row $row -Index $index -Token 'Site URL'
    $primary = Get-Val -Row $row -Index $index -Token 'Primary Contact Email'
    $owners = Get-Val -Row $row -Index $index -Token 'All Owners'
    $pctText = Format-Percent (Get-Val -Row $row -Index $index -Token '% Used')

    $recips = Get-Recipients -Primary $primary -AllOwners $owners

    $entry = [pscustomobject]@{
        Timestamp = (Get-Date).ToString('s'); Site = $siteName; Url = $siteUrl
        Recipients = ($recips -join '; '); PercentUsed = $pctText; Action = ''; Result = ''
    }

    if (-not $recips) {
        $entry.Action = 'SKIPPED'; $entry.Result = 'No valid owner email address'
        Write-Warning "SKIP  '$siteName' - no valid recipient."; $log.Add($entry); continue
    }

    $target = "$siteName  ->  $($recips -join ', ')"
    if ($PSCmdlet.ShouldProcess($target, 'Send storage notification')) {
        try {
            Send-GraphMail -Sender $SenderMailbox -To $recips -Cc $CcMailbox `
                -Subject $subject -BodyHtml (New-Body -SiteName $siteName -SiteUrl $siteUrl -PercentText $pctText) `
                -Token $token
            $entry.Action = 'SENT'; $entry.Result = 'OK'; $sent++
            Write-Host "SENT  '$siteName' -> $($recips -join ', ')" -ForegroundColor Green
            Start-Sleep -Milliseconds $DelayMs
        } catch {
            $entry.Action = 'ERROR'; $entry.Result = $_.Exception.Message
            Write-Warning "FAIL  '$siteName' - $($_.Exception.Message)"
        }
    } else {
        $entry.Action = 'WHATIF'; $entry.Result = 'Not sent (-WhatIf)'
    }
    $log.Add($entry)
}

$log | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8
$by = $log | Group-Object Action | ForEach-Object { "$($_.Name)=$($_.Count)" }
Write-Host ''
Write-Host '-------- Summary --------' -ForegroundColor Cyan
Write-Host ("Targets : {0}" -f $targets.Count)
Write-Host ("Outcome : {0}" -f ($by -join '  '))
if ($capped) { Write-Host ("NOTE    : stopped at -MaxSends {0}" -f $MaxSends) -ForegroundColor Yellow }
Write-Host ("Log     : {0}" -f $LogPath)
if ($WhatIfPreference) { Write-Host 'WHATIF mode - nothing was sent. Re-run without -WhatIf to send.' -ForegroundColor Yellow }