#Requires -Version 7.0
<#
.SYNOPSIS
    List resource mailboxes (rooms / equipment) for a given ISO country, using a
    country-origin marker in extensionAttribute12 plus name/address fallbacks so
    rooms are caught wherever the country token sits in the name - start, middle,
    end, or not at all.

.DESCRIPTION
    Read-only. In a tenant that grew by acquisition, resource mailboxes are named
    every possible way and the origin marker is only populated on some of them.
    Filtering on any single signal misses rooms; this script ORs several signals
    together and tells you which one fired, so you can see how sparse the marker
    actually is.

    Connects to EXO, pulls RoomMailbox/EquipmentMailbox, and keeps the ones that
    match the country by any of these signals:
      marker    CustomAttribute12 -eq '<MarkerPrefix>-<CC>'  (extensionAttribute12; most reliable)
      usageloc  UsageLocation      -eq '<CC>'
      name      display name contains <CC> as a standalone token (case-sensitive)
      country   display name contains the full country name
      addr      any proxy address matches <AddressPrefix><CC> / <CC>. / @<domain>.<cc> / country name
      extra     any -ExtraPattern you pass (e.g. city names)
    The MatchedBy column shows which signal(s) fired.

    Optionally checks whether specific addresses are already in use (-CheckAddress),
    across ALL recipient types (not just resources), via Get-Recipient. That check
    is what you run before proposing new room addresses.

    -MarkerPrefix / -AddressPrefix / -Domain are the organisation-specific tokens.
    Set them to your own conventions; the defaults are placeholders.

.EXAMPLE
    .\Get-CountryResource.ps1 -Country HU

.EXAMPLE
    # Country rooms only + check four proposed addresses in one pass
    .\Get-CountryResource.ps1 -Country HU -ResourceType Room -CheckAddress `
        'RESHU.MR.RoomOne@contoso.com','RESHU.MR.RoomTwo@contoso.com'

.EXAMPLE
    # Strict: only objects carrying the origin marker
    .\Get-CountryResource.ps1 -Country DE -MarkerOnly

.EXAMPLE
    # Add city names as extra signals for a country whose marker is sparse
    .\Get-CountryResource.ps1 -Country PT -ExtraPattern 'Lisboa','Porto'

.EXAMPLE
    # Unattended, app-only certificate auth
    .\Get-CountryResource.ps1 -Country DE -AppId '<client-id>' `
        -CertThumb '<cert-thumbprint>' -Organization 'contoso.onmicrosoft.com'

.NOTES
    When to use  : Before proposing addresses for new rooms in a country, or when a subsidiary asks for its room inventory.
    Why it exists: In a tenant grown by acquisition, resource mailboxes are named every possible way and the origin marker is only populated on some of them, so filtering on one signal misses rooms. Six signals are ORed together and the MatchedBy column shows which fired, which incidentally measures how sparse the marker really is.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^[A-Za-z]{2}$')]
    [string]$Country,

    [ValidateSet('Room','Equipment','Both')]
    [string]$ResourceType = 'Both',

    [switch]$MarkerOnly,                 # restrict to the origin marker only (no heuristics)
    [string[]]$ExtraPattern,             # extra regex alternatives (e.g. city names)
    [string[]]$CheckAddress,             # also report whether these addresses are in use
    [switch]$IncludeAddresses,           # include full proxy address list in the output

    # ---- organisation-specific tokens (change these to your own conventions) ----
    [string]$MarkerPrefix  = 'ORG',      # CustomAttribute12 value is <MarkerPrefix>-<CC>
    [string]$AddressPrefix = 'RES',      # address convention, e.g. RESDE.MR.<room>@...
    [string]$Domain        = 'contoso.com',

    # ---- EXO app-only cert auth (leave empty for interactive/delegated) ----
    [string]$AppId        = $env:EXO_APPID,
    [string]$CertThumb    = $env:EXO_CERTTHUMB,
    [string]$Organization = $env:EXO_ORG,

    [string]$ExportDir = (Join-Path $PSScriptRoot 'Exports')
)

# ---------- helpers ----------
function Write-Step($m){ Write-Host "`n[STEP] $m" -ForegroundColor Cyan }
function Write-OK  ($m){ Write-Host "[ OK ] $m"   -ForegroundColor Green }
function Write-Warn($m){ Write-Host "[WARN] $m"   -ForegroundColor Yellow }
function Die       ($m){ Write-Host "[FAIL] $m"   -ForegroundColor Red; try{Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue}catch{}; exit 1 }

$CC = $Country.ToUpper()

# ISO2 -> full-name regex. Only a partial list; unknown codes still work via
# marker / usageLocation / address / name-token.
$CountryName = @{
    DE='Germany'; AT='Austria'; PH='Philippines'; ES='Spain'; DK='Denmark'; HU='Hungary'
    CZ='Czech(ia|\ Republic)?'; NG='Nigeria'; AE='United\ Arab\ Emirates|UAE'; GH='Ghana'
    SK='Slovakia'; EE='Estonia'; IT='Italy'; FI='Finland'; BR='Brazil'; MX='Mexico'
    CO='Colombia'; CL='Chile'; PT='Portugal'; BE='Belgium'; PE='Peru'; SE='Sweden'
    UK='United\ Kingdom|Great\ Britain'; GB='United\ Kingdom|Great\ Britain'
}
$cname = $CountryName[$CC]
if(-not $cname -and -not $MarkerOnly){ Write-Warn "No full country-name mapping for '$CC' - relying on marker / usageLocation / address / name-token only." }

$marker  = "$MarkerPrefix-$CC"
$nameTok = "(^|[^A-Za-z])$CC([^A-Za-z]|`$)"          # -cmatch: CC as a standalone token, anywhere
$lcc     = $CC.ToLower()
$domRoot = [regex]::Escape(($Domain -split '\.')[0])
$addrRe  = (@("$AddressPrefix$CC", "(^|[^A-Za-z])$CC\.", "@$domRoot\.$lcc\b") + @(if($cname){$cname})) -join '|'

Write-Host "===== Get-CountryResource  [$CC / $ResourceType] =====" -ForegroundColor Magenta
Write-Host "  Marker: $marker   MarkerOnly: $([bool]$MarkerOnly)"

if(-not (Test-Path $ExportDir)){ New-Item -ItemType Directory -Path $ExportDir | Out-Null }
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'

# ---------- connect ----------
# App-only when all three of -AppId / -CertThumb / -Organization are supplied
# (they default to EXO_APPID / EXO_CERTTHUMB / EXO_ORG); interactive otherwise.
# Partial app-only config is rejected rather than silently falling back.
$appOnly = -not ([string]::IsNullOrWhiteSpace($AppId) -or
                 [string]::IsNullOrWhiteSpace($CertThumb) -or
                 [string]::IsNullOrWhiteSpace($Organization))
if(-not $appOnly){
    $partial = @($AppId, $CertThumb, $Organization) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    if($partial.Count -gt 0){ Die "Incomplete app-only configuration: -AppId, -CertThumb and -Organization must all be set (or all be empty for interactive auth)." }
}

Write-Step "Connecting to Exchange Online ($(if($appOnly){'app-only'}else{'interactive'}))"
try {
    if($appOnly){ Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $CertThumb -Organization $Organization -ShowBanner:$false -ErrorAction Stop | Out-Null }
    else        { Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null }
}
catch { Die "EXO connect failed: $($_.Exception.Message)" }
Write-OK "Connected$(if($appOnly){" app-only as $AppId"}else{' interactively (delegated)'})"

# ---------- pull resources ----------
$rtd = switch($ResourceType){ 'Room'{,'RoomMailbox'} 'Equipment'{,'EquipmentMailbox'} default{'RoomMailbox','EquipmentMailbox'} }
Write-Step "Retrieving $($rtd -join ' + ')"
$all = Get-EXOMailbox -RecipientTypeDetails $rtd -ResultSize Unlimited `
        -Properties DisplayName,PrimarySmtpAddress,Alias,EmailAddresses,CustomAttribute12,UsageLocation
Write-OK "$($all.Count) resource mailboxes in tenant"

# ---------- match ----------
Write-Step "Matching country = $CC"
$rows = foreach($m in $all){
    $hits = [System.Collections.Generic.List[string]]::new()
    if($m.CustomAttribute12 -eq $marker){ $hits.Add('marker') }
    if(-not $MarkerOnly){
        if($m.UsageLocation -eq $CC)                                   { $hits.Add('usageloc') }
        if($m.DisplayName   -cmatch $nameTok)                          { $hits.Add('name') }
        if($cname -and ($m.DisplayName -match $cname))                 { $hits.Add('country') }
        if(($m.EmailAddresses -match $addrRe) -or ($m.Alias -match $addrRe)) { $hits.Add('addr') }
        foreach($p in $ExtraPattern){
            if(($m.DisplayName -match $p) -or ($m.EmailAddresses -match $p)){ $hits.Add("extra:$p") }
        }
    }
    if($hits.Count -eq 0){ continue }
    [pscustomobject]@{
        DisplayName       = $m.DisplayName
        PrimarySmtp       = [string]$m.PrimarySmtpAddress
        Alias             = $m.Alias
        Type              = [string]$m.RecipientTypeDetails
        Marker            = $m.CustomAttribute12
        UsageLocation     = $m.UsageLocation
        MatchedBy         = ($hits -join ',')
        EmailAddresses    = $(if($IncludeAddresses){ (($m.EmailAddresses | Where-Object {$_ -clike 'smtp:*' -or $_ -clike 'SMTP:*'}) -join '; ') } else { $null })
    }
}
$rows = $rows | Sort-Object DisplayName

Write-Step "Results: $($rows.Count) resource(s) for $CC"
$rows | Select-Object DisplayName,PrimarySmtp,Type,Marker,UsageLocation,MatchedBy | Format-Table -AutoSize

$csv = Join-Path $ExportDir "Resources_${CC}_$stamp.csv"
$rows | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
Write-OK "List written to $csv"

# ---------- optional: address in-use check (all recipient types) ----------
if($CheckAddress){
    Write-Step "Checking $($CheckAddress.Count) address(es) across ALL recipient types"
    $chk = foreach($a in $CheckAddress){
        $r = Get-Recipient -Identity $a -ErrorAction SilentlyContinue
        [pscustomobject]@{
            Address = $a
            InUse   = [bool]$r
            On      = $(if($r){$r.DisplayName}else{''})
            Type    = $(if($r){[string]$r.RecipientTypeDetails}else{''})
        }
    }
    $chk | Format-Table -AutoSize
    $chkCsv = Join-Path $ExportDir "AddressCheck_${CC}_$stamp.csv"
    $chk | Export-Csv -Path $chkCsv -NoTypeInformation -Encoding UTF8
    Write-OK "Address check written to $chkCsv"
}

Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue