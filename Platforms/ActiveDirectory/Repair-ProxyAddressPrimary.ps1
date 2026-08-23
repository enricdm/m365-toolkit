<#
.SYNOPSIS
    Detecta y corrige objetos de AD sin direccion SMTP primaria valida.

.DESCRIPTION
    Escanea usuarios habilitados para correo y marca:
      - NoPrimary        : ningun valor con prefijo 'SMTP:' en mayusculas
      - BadPrefix        : prefijo no reconocido (STMP, SMPT, SMTO, smtp mal escrito...)
      - MultiplePrimary  : mas de un 'SMTP:'
      - PrimaryIsMOERA   : la primaria es *.onmicrosoft.com
      - MailMismatch     : atributo 'mail' distinto de la primaria

    Con -Execute repara UNICAMENTE el caso seguro:
      objeto con 0 primarias + exactamente 1 valor con prefijo mal escrito
      -> se elimina el valor corrupto y se anade como 'SMTP:'.
    El resto se reporta pero no se toca.

.PARAMETER Server
    Controlador de dominio on-prem contra el que se consulta. Obligatorio.

.PARAMETER AcceptedRoot
    Dominio que debe llevar la direccion primaria. Solo se actualiza el atributo
    'mail' cuando la direccion reparada pertenece a este dominio.

.PARAMETER SearchBase
    DN donde buscar. "" = raiz del dominio.

.PARAMETER Execute
    Aplica las correcciones. Sin este switch solo se informa.

.EXAMPLE
    # Dry-run: solo informa y exporta el CSV
    .\Repair-ProxyAddressPrimary.ps1 -Server 'DC01.corp.local' -AcceptedRoot 'contoso.com'

.EXAMPLE
    # Aplica las correcciones seguras, acotado a una OU
    .\Repair-ProxyAddressPrimary.ps1 -Server 'DC01.corp.local' -AcceptedRoot 'contoso.com' `
        -SearchBase 'OU=Users,DC=corp,DC=local' -Execute
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

Write-Step "Repair-ProxyAddressPrimary  -  $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')"
if (-not $Execute) { Write-Warn "MODO DRY-RUN. Usa -Execute para aplicar cambios." }

try { Import-Module ActiveDirectory -ErrorAction Stop }
catch { Write-Die "No se pudo cargar el modulo ActiveDirectory (RSAT)." }

# --------------------------- RECOGIDA ---------------------------
Write-Step "Consultando objetos en $Server ..."
$params = @{
    Filter     = 'proxyAddresses -like "*"'
    Properties = 'proxyAddresses','mail','userPrincipalName','distinguishedName','enabled','msExchRecipientTypeDetails'
    Server     = $Server
}
if ($SearchBase) { $params['SearchBase'] = $SearchBase }

$users = Get-ADUser @params
Write-OK "$($users.Count) objetos con proxyAddresses recuperados."

# --------------------------- ANALISIS ---------------------------
Write-Step "Analizando prefijos..."
$findings = New-Object System.Collections.Generic.List[object]

foreach ($u in $users) {

    $addrs   = @($u.proxyAddresses)
    # Comparacion sensible a mayusculas: solo 'SMTP:' cuenta como primaria
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

    # Candidato de reparacion: 0 primarias + 1 solo valor con prefijo corrupto
    $fixable  = $false
    $fixFrom  = $null
    $fixTo    = $null

    if ($primary.Count -eq 0 -and $bad.Count -eq 1) {
        $parts  = $bad[0] -split ':',2
        $prefix = $parts[0]
        $addr   = $parts[1]
        if (($BadPrefixes -contains $prefix.ToUpper()) -and $addr -match '^[^@]+@[^@]+\.[^@]+$') {
            # no debe existir ya el mismo correo como alias secundario
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

Write-OK "$($findings.Count) objetos con incidencias."
$fixables = @($findings | Where-Object { $_.Fixable })
Write-OK "$($fixables.Count) reparables automaticamente."

if ($findings.Count -eq 0) { Write-Step "Nada que hacer."; return }

$findings | Group-Object Issues | Sort-Object Count -Descending |
    Select-Object @{n='Incidencia';e={$_.Name}}, Count | Format-Table -AutoSize

# --------------------------- EXPORT -----------------------------
if (-not (Test-Path $ExportDir)) { New-Item -ItemType Directory -Path $ExportDir | Out-Null }
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$csv   = Join-Path $ExportDir "ProxyAddressAudit-$stamp.csv"
$findings | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8 -Delimiter ';'
Write-OK "Informe exportado: $csv"

# --------------------------- APLICAR ----------------------------
if (-not $Execute) {
    Write-Step "Dry-run finalizado. Revisa el CSV y relanza con -Execute."
    return
}
if ($fixables.Count -eq 0) { Write-Step "Sin candidatos automaticos."; return }

Write-Step "Aplicando $($fixables.Count) correcciones..."
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
Write-OK "Log de cambios: $log"
Write-Step "Lanza un delta sync en el AAD Connect del dominio y verifica en EXO."
