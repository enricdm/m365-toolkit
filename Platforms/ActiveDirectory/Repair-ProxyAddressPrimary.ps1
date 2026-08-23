<#
.SYNOPSIS
    Finds and repairs AD objects with no valid primary SMTP address.

.DESCRIPTION
    Scans mail-enabled users and flags:
      - NoPrimary        : no value carrying an uppercase 'SMTP:' prefix
      - BadPrefix        : unrecognised prefix (STMP, SMPT, SMTO, misspelled smtp...)
      - MultiplePrimary  : more than one 'SMTP:'
      - PrimaryIsMOERA   : the primary address is *.onmicrosoft.com
      - MailMismatch     : the 'mail' attribute differs from the primary address

    With -Execute it repairs ONLY the safe case:
      an object with 0 primaries and exactly 1 value carrying a misspelled prefix
      -> the corrupt value is removed and re-added as 'SMTP:'.
    Everything else is reported but left untouched.

    The scope of the automatic repair is deliberately narrow. An object with two
    primaries, or with several malformed values, needs a human to decide which
    address should win; guessing would silently change someone's reply-to address.

.PARAMETER Server
    On-premises domain controller to query. Mandatory.

.PARAMETER AcceptedRoot
    Domain the primary address is expected to belong to. The 'mail' attribute is
    only updated when the repaired address falls inside this domain.

.PARAMETER SearchBase
    Distinguished name to search under. "" = domain root.

.PARAMETER Execute
    Applies the repairs. Without this switch the script only reports.

.PARAMETER BadPrefixes
    Misspellings treated as repairable. Extend if your directory has others.

.PARAMETER ValidPrefixes
    Prefixes considered legitimate. Anything outside this list counts as BadPrefix.

.EXAMPLE
    # Dry run: report only, exports the CSV
    .\Repair-ProxyAddressPrimary.ps1 -Server 'DC01.corp.local' -AcceptedRoot 'contoso.com'

.EXAMPLE
    # Apply the safe repairs, scoped to a single OU
    .\Repair-ProxyAddressPrimary.ps1 -Server 'DC01.corp.local' -AcceptedRoot 'contoso.com' `
        -SearchBase 'OU=Users,DC=corp,DC=local' -Execute

.NOTES
    Requires : ActiveDirectory module (RSAT)
    Rights   : write access to proxyAddresses and mail on the target user objects
    After    : run a delta sync on AAD Connect and verify in Exchange Online
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Server,
    [Parameter(Mandatory)][string]$AcceptedRoot,
    [string]$SearchBase = "",
    [switch]$Execute,
    [string]$ExportDir     = (Join-Path $PSScriptRoot 'Exports'),
    [string[]]$BadPrefixes = @('STMP','SMPT','SMTO','SMT','STMTP','SPTM'),
    [string[]]$ValidPrefixes = @('smtp','x500','sip','eum','eas','x400','notes','ccmail','msmail','mailto')
)

$ErrorActionPreference = 'Stop'

function Write-Step { param($m) Write-Host "`n[>] $m" -ForegroundColor Cyan }
function Write-OK   { param($m) Write-Host "[OK] $m"  -ForegroundColor Green }
function Write-Warn { param($m) Write-Host "[!!] $m"  -ForegroundColor Yellow }
function Write-Die  { param($m) Write-Host "[XX] $m"  -ForegroundColor Red; exit 1 }

Write-Step "Repair-ProxyAddressPrimary  -  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
if (-not $Execute) { Write-Warn "DRY-RUN MODE. Use -Execute to apply changes." }

try { Import-Module ActiveDirectory -ErrorAction Stop }
catch { Write-Die "Could not load the ActiveDirectory module (RSAT)." }

# --------------------------- COLLECT ----------------------------
Write-Step "Querying objects on $Server ..."
$params = @{
    Filter     = 'proxyAddresses -like "*"'
    Properties = 'proxyAddresses','mail','userPrincipalName','distinguishedName','enabled','msExchRecipientTypeDetails'
    Server     = $Server
}
if ($SearchBase) { $params['SearchBase'] = $SearchBase }

$users = Get-ADUser @params
Write-OK "$($users.Count) objects with proxyAddresses retrieved."

# --------------------------- ANALYSE ----------------------------
Write-Step "Analysing prefixes..."
$findings = New-Object System.Collections.Generic.List[object]

foreach ($u in $users) {

    $addrs   = @($u.proxyAddresses)
    # Case-sensitive comparison: only 'SMTP:' counts as the primary address
    $primary = @($addrs | Where-Object { $_ -cmatch '^SMTP:' })
    $bad     = @($addrs | Where-Object {
                    $p = ($_ -split ':',2)[0]
                    $p -and ($ValidPrefixes -notcontains $p.ToLower())
               })

    $issues = New-Object System.Collections.Generic.List[string]

    if ($primary.Count -eq 0) { $issues.Add('NoPrimary') }
    if ($primary.Count -gt 1) { $issues.Add('MultiplePrimary') }
    if ($bad.Count -gt 0)     { $issues.Add('BadPrefix') }

    if ($primary.Count -eq 1) {
        $pAddr = ($primary[0] -split ':',2)[1]
        if ($pAddr -like '*.onmicrosoft.com') { $issues.Add('PrimaryIsMOERA') }
        if ($u.mail -and $u.mail -ne $pAddr)  { $issues.Add('MailMismatch') }
    }

    if ($issues.Count -eq 0) { continue }

    # Repair candidate: 0 primaries + exactly 1 value with a corrupt prefix
    $fixable  = $false
    $fixFrom  = $null
    $fixTo    = $null

    if ($primary.Count -eq 0 -and $bad.Count -eq 1) {
        $parts  = $bad[0] -split ':',2
        $prefix = $parts[0]
        $addr   = $parts[1]
        if (($BadPrefixes -contains $prefix.ToUpper()) -and $addr -match '^[^@]+@[^@]+\.[^@]+$') {
            # the same address must not already exist as a secondary alias
            $dupe = $addrs | Where-Object { $_ -imatch "^smtp:$([regex]::Escape($addr))$" }
            if (-not $dupe) {
                $fixable = $true
                $fixFrom = $bad[0]
                $fixTo   = "SMTP:$addr"
            }
        }
    }

    $findings.Add([pscustomobject]@{
        SamAccountName    = $u.SamAccountName
        DisplayName       = $u.Name
        UPN               = $u.userPrincipalName
        Enabled           = $u.Enabled
        Mail              = $u.mail
        Issues            = ($issues -join '; ')
        PrimaryCount      = $primary.Count
        CurrentPrimary    = ($primary -join ' | ')
        BadValues         = ($bad -join ' | ')
        AllProxyAddresses = ($addrs -join ' | ')
        Fixable           = $fixable
        FixFrom           = $fixFrom
        FixTo             = $fixTo
        DN                = $u.distinguishedName
    })
}

Write-OK "$($findings.Count) objects with issues."
$fixables = @($findings | Where-Object { $_.Fixable })
Write-OK "$($fixables.Count) automatically repairable."

if ($findings.Count -eq 0) { Write-Step "Nothing to do."; return }

$findings | Group-Object Issues | Sort-Object Count -Descending |
    Select-Object @{n='Issue';e={$_.Name}}, Count | Format-Table -AutoSize

# --------------------------- EXPORT -----------------------------
if (-not (Test-Path $ExportDir)) { New-Item -ItemType Directory -Path $ExportDir | Out-Null }
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$csv   = Join-Path $ExportDir "ProxyAddressAudit-$stamp.csv"
$findings | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8 -Delimiter ';'
Write-OK "Report exported: $csv"

# --------------------------- APPLY ------------------------------
if (-not $Execute) {
    Write-Step "Dry run finished. Review the CSV and re-run with -Execute."
    return
}
if ($fixables.Count -eq 0) { Write-Step "No automatic candidates."; return }

Write-Step "Applying $($fixables.Count) repairs..."
$results = New-Object System.Collections.Generic.List[object]

foreach ($f in $fixables) {
    try {
        Set-ADUser -Identity $f.DN -Server $Server `
                   -Remove @{proxyAddresses = $f.FixFrom} `
                   -Add    @{proxyAddresses = $f.FixTo}

        $newMail = ($f.FixTo -split ':',2)[1]
        if ($newMail -like "*@$AcceptedRoot" -and $f.Mail -ne $newMail) {
            Set-ADUser -Identity $f.DN -Server $Server -Replace @{mail = $newMail}
        }

        Write-OK "$($f.SamAccountName): $($f.FixFrom) -> $($f.FixTo)"
        $results.Add([pscustomobject]@{ Sam=$f.SamAccountName; Status='OK'; From=$f.FixFrom; To=$f.FixTo; Error='' })
    }
    catch {
        Write-Warn "$($f.SamAccountName): $($_.Exception.Message)"
        $results.Add([pscustomobject]@{ Sam=$f.SamAccountName; Status='FAILED'; From=$f.FixFrom; To=$f.FixTo; Error=$_.Exception.Message })
    }
}

$log = Join-Path $ExportDir "ProxyAddressFix-$stamp.csv"
$results | Export-Csv -Path $log -NoTypeInformation -Encoding UTF8 -Delimiter ';'
Write-OK "Change log: $log"
Write-Step "Run a delta sync on the domain's AAD Connect and verify in Exchange Online."
