<#
.SYNOPSIS
    Reports which PowerShell modules this toolkit needs, which of them you already have,
    and installs the missing ones. Reports by default; installs only with -Execute.

.DESCRIPTION
    Two ways of working out what is needed, and it does both, because they disagree and
    the disagreement is the useful part:

      1. A curated map of module -> platform, kept alongside this script.
      2. A scan of the actual .ps1 files for the cmdlet prefixes they call
         (Get-Mg*, Get-EXO*, Get-PnP*, Get-AD*, Import-Excel, and so on).

    The scan catches what the map forgets. A script that quietly started calling
    Get-PnPTenantSite is a new dependency whether anybody wrote it down or not, and this
    is how you find out before a user does. Where the two sources disagree, it says so
    rather than silently preferring one.

    Scope it with -Platform so you install what you need rather than everything. Somebody
    who only wants the Exchange scripts should not be made to pull the whole Graph SDK.

    NOTHING IS EVER REMOVED. It installs and it reports. If you have an old version of a
    module it will tell you and offer to update it with -Execute, but it will not
    uninstall anything, because the side effects of removing a module somebody else's
    script depends on are not this script's to accept.

.PARAMETER Platform
    Which platform folders to resolve. Default All. Repeatable.

.PARAMETER Execute
    Install what is missing. Without it, nothing is installed and nothing changes.

.PARAMETER Scope
    CurrentUser (default) or AllUsers. AllUsers needs an elevated session.

.PARAMETER IncludeOptional
    Also include modules only one or two scripts need, which most people can skip.

.PARAMETER SkipScan
    Use only the curated map. Faster, and it stops the scan reporting drift you already
    know about.

.EXAMPLE
    # What would I need for everything, and what have I got?
    .\Install-ToolkitPrerequisite.ps1

.EXAMPLE
    # Just the Exchange and Entra ID scripts, installed for me
    .\Install-ToolkitPrerequisite.ps1 -Platform Exchange,EntraID -Execute

.EXAMPLE
    # Machine-wide, everything, including the rarely-needed ones
    .\Install-ToolkitPrerequisite.ps1 -Execute -Scope AllUsers -IncludeOptional

.NOTES
    When to use  : First thing, on a new machine, before you try to run anything else in this repository.
    Why it exists: Resolves dependencies per platform instead of making you install the whole Graph SDK to run one Exchange script, and cross-checks the curated list against the cmdlets the scripts actually call so the list cannot quietly drift out of date. It reports before it installs and never removes anything.
    Requires : PowerShell 5.1 or 7.x. PowerShellGet, which ships with both.
    Rights   : none for -Scope CurrentUser. Elevation for -Scope AllUsers.
#>

[CmdletBinding()]
param(
    [ValidateSet('All','EntraID','Exchange','Endpoint','SharePoint','ActiveDirectory','Purview','Automation')]
    [string[]]$Platform = @('All'),

    [switch]$Execute,

    [ValidateSet('CurrentUser','AllUsers')]
    [string]$Scope = 'CurrentUser',

    [switch]$IncludeOptional,
    [switch]$SkipScan
)

$ErrorActionPreference = 'Stop'

function Write-Step { param([string]$m) Write-Host "`n[*] $m" -ForegroundColor Cyan }
function Write-OK   { param([string]$m) Write-Host "[+] $m" -ForegroundColor Green }
function Write-Warn { param([string]$m) Write-Host "[!] $m" -ForegroundColor Yellow }
function Write-Die  { param([string]$m) Write-Host "[X] $m" -ForegroundColor Red; exit 1 }

# ==================== THE CURATED MAP ====================
# MinVersion is only set where a specific version is genuinely required.
# WindowsPowerShellOnly means it will not load in PowerShell 7 at all.

$Catalogue = @(
    [pscustomobject]@{ Name='Microsoft.Graph.Authentication';  Platforms=@('EntraID','Endpoint','Exchange','Automation'); MinVersion='2.0.0'; Optional=$false; WinPSOnly=$false; Note='Connect-MgGraph and Invoke-MgGraphRequest. The one module almost everything here needs.' }
    [pscustomobject]@{ Name='Microsoft.Graph.Users';           Platforms=@('EntraID');                                    MinVersion='2.0.0'; Optional=$false; WinPSOnly=$false; Note='' }
    [pscustomobject]@{ Name='Microsoft.Graph.Groups';          Platforms=@('EntraID');                                    MinVersion='2.0.0'; Optional=$false; WinPSOnly=$false; Note='' }
    [pscustomobject]@{ Name='Microsoft.Graph.Applications';    Platforms=@('EntraID','Automation');                       MinVersion='2.0.0'; Optional=$false; WinPSOnly=$false; Note='App registrations, and the managed-identity permission grant in Automation/.' }
    [pscustomobject]@{ Name='Microsoft.Graph.Identity.SignIns';Platforms=@('EntraID');                                    MinVersion='2.0.0'; Optional=$false; WinPSOnly=$false; Note='Conditional Access and sign-in data.' }
    [pscustomobject]@{ Name='Microsoft.Graph.Reports';         Platforms=@('EntraID','Exchange');                         MinVersion='2.0.0'; Optional=$false; WinPSOnly=$false; Note='Usage reports. Get-MailboxReceiveVolume needs this one.' }
    [pscustomobject]@{ Name='ExchangeOnlineManagement';        Platforms=@('Exchange','EntraID');                         MinVersion='3.0.0'; Optional=$false; WinPSOnly=$false; Note='EntraID needs it only for the shared-mailbox conversion tier of licence reclamation.' }
    [pscustomobject]@{ Name='PnP.PowerShell';                  Platforms=@('SharePoint');                                 MinVersion='';      Optional=$false; WinPSOnly=$false; Note='Version matters: 2.x is PowerShell 7 only, 1.x is Windows PowerShell only.' }
    [pscustomobject]@{ Name='ActiveDirectory';                 Platforms=@('ActiveDirectory');                            MinVersion='';      Optional=$false; WinPSOnly=$false; Note='Ships with RSAT, not from the gallery. See below if it is missing.' }
    [pscustomobject]@{ Name='PurviewInformationProtection';    Platforms=@('Purview');                                    MinVersion='';      Optional=$false; WinPSOnly=$true;  Note='Windows PowerShell 5.1 only. Will not load in PowerShell 7.' }
    [pscustomobject]@{ Name='AIPService';                      Platforms=@('Purview');                                    MinVersion='';      Optional=$false; WinPSOnly=$true;  Note='Windows PowerShell 5.1 only.' }
    [pscustomobject]@{ Name='ImportExcel';                     Platforms=@('Exchange','SharePoint');                      MinVersion='';      Optional=$true;  WinPSOnly=$false; Note='Only for scripts that read or write .xlsx.' }
    [pscustomobject]@{ Name='Microsoft.Online.SharePoint.PowerShell'; Platforms=@('SharePoint');                          MinVersion='';      Optional=$true;  WinPSOnly=$false; Note='The SPO module. Only needed if you re-add storage tooling.' }
)

# Cmdlet prefix -> module, for the scan. Deliberately conservative: a false positive here
# tells somebody to install a module they do not need, which is worse than a gap.
$PrefixMap = @{
    'Get-Mg'          = 'Microsoft.Graph.Authentication'
    'Invoke-MgGraph'  = 'Microsoft.Graph.Authentication'
    'Connect-MgGraph' = 'Microsoft.Graph.Authentication'
    'Get-EXO'         = 'ExchangeOnlineManagement'
    'Connect-Exchange'= 'ExchangeOnlineManagement'
    'Get-Mailbox'     = 'ExchangeOnlineManagement'
    'Set-Mailbox'     = 'ExchangeOnlineManagement'
    'Get-DistributionGroup' = 'ExchangeOnlineManagement'
    'New-DistributionGroup' = 'ExchangeOnlineManagement'
    'Get-PnP'         = 'PnP.PowerShell'
    'Connect-PnP'     = 'PnP.PowerShell'
    'Get-AD'          = 'ActiveDirectory'
    'Set-AD'          = 'ActiveDirectory'
    'Get-SPO'         = 'Microsoft.Online.SharePoint.PowerShell'
    'Connect-SPO'     = 'Microsoft.Online.SharePoint.PowerShell'
    'Set-SPO'         = 'Microsoft.Online.SharePoint.PowerShell'
    'Import-Excel'    = 'ImportExcel'
    'Export-Excel'    = 'ImportExcel'
    'Connect-AipService' = 'AIPService'
    'Get-AipService'  = 'AIPService'
}

# ==================== RESOLVE ====================

$RepoRoot     = Split-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -Parent
$PlatformRoot = Join-Path $RepoRoot 'Platforms'
$wanted       = if ($Platform -contains 'All') { @('EntraID','Exchange','Endpoint','SharePoint','ActiveDirectory','Purview','Automation') } else { $Platform }

Write-Step "Toolkit prerequisites - $($wanted -join ', ')"
Write-Host "    PowerShell  : $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))"
Write-Host "    Scope       : $Scope"
if (-not $Execute) { Write-Warn 'Report only. Nothing will be installed. Add -Execute to install.' }

$needed = @($Catalogue | Where-Object {
    ($_.Platforms | Where-Object { $wanted -contains $_ }) -and ($IncludeOptional -or -not $_.Optional)
})

# ==================== SCAN ====================

if (-not $SkipScan) {
    Write-Step 'Scanning the scripts for what they actually call'
    if (-not (Test-Path $PlatformRoot)) {
        Write-Warn "Platforms folder not found at $PlatformRoot - skipping the scan."
    }
    else {
        $files = @(Get-ChildItem $PlatformRoot -Recurse -Filter *.ps1 -File |
                   Where-Object { $wanted -contains ($_.FullName -replace [regex]::Escape($PlatformRoot + [IO.Path]::DirectorySeparatorChar), '' -replace '[\\/].*$', '') })
        Write-OK "$($files.Count) script(s) in scope"

        $seen = @{}
        foreach ($file in $files) {
            $text = Get-Content $file.FullName -Raw
            foreach ($prefix in $PrefixMap.Keys) {
                if ($text -match [regex]::Escape($prefix)) {
                    $mod = $PrefixMap[$prefix]
                    if (-not $seen.ContainsKey($mod)) { $seen[$mod] = New-Object System.Collections.Generic.List[string] }
                    if (-not $seen[$mod].Contains($file.Name)) { $seen[$mod].Add($file.Name) }
                }
            }
        }

        foreach ($mod in ($seen.Keys | Sort-Object)) {
            if ($needed.Name -notcontains $mod) {
                $sample = ($seen[$mod] | Select-Object -First 3) -join ', '
                Write-Warn "DRIFT: '$mod' is called by $($seen[$mod].Count) script(s) but is not in the curated list for this scope."
                Write-Host  "         e.g. $sample" -ForegroundColor DarkGray
                $entry = $Catalogue | Where-Object Name -eq $mod | Select-Object -First 1
                if (-not $entry) {
                    $entry = [pscustomobject]@{ Name=$mod; Platforms=@('(detected)'); MinVersion=''; Optional=$false; WinPSOnly=$false; Note='Found by scan, not in the catalogue.' }
                }
                $needed += $entry
            }
        }
        if (-not ($seen.Keys | Where-Object { $needed.Name -notcontains $_ })) { Write-OK 'No drift: the scan agrees with the curated list.' }
    }
}

$needed = @($needed | Sort-Object Name -Unique)

# ==================== CHECK ====================

Write-Step 'Checking what is installed'
$report = New-Object System.Collections.Generic.List[object]

foreach ($m in $needed) {
    $have    = @(Get-Module -ListAvailable -Name $m.Name | Sort-Object Version -Descending)
    $newest  = if ($have.Count) { $have[0].Version } else { $null }
    $status  = 'MISSING'

    if ($newest) {
        $status = 'OK'
        if ($m.MinVersion -and [version]$newest -lt [version]$m.MinVersion) { $status = 'OLD' }
    }

    $blocked = $null
    if ($m.WinPSOnly -and $PSVersionTable.PSEdition -eq 'Core') {
        $blocked = 'Windows PowerShell 5.1 only - it will install but not load in this session'
    }
    if ($m.Name -eq 'ActiveDirectory' -and $status -eq 'MISSING') {
        $blocked = 'Comes from RSAT, not the gallery - see the note at the end'
    }

    $report.Add([pscustomobject]@{
        Module    = $m.Name
        Required  = if ($m.MinVersion) { ">= $($m.MinVersion)" } else { 'any' }
        Installed = if ($newest) { $newest.ToString() } else { '-' }
        Status    = $status
        Caveat    = $blocked
        Note      = $m.Note
        Entry     = $m
    })
}

$report | Select-Object Module, Required, Installed, Status | Format-Table -AutoSize

foreach ($r in ($report | Where-Object { $_.Caveat })) { Write-Warn "$($r.Module): $($r.Caveat)" }
foreach ($r in ($report | Where-Object { $_.Note })) { Write-Host "    $($r.Module): $($r.Note)" -ForegroundColor DarkGray }

# ==================== INSTALL ====================

$todo = @($report | Where-Object { $_.Status -in @('MISSING','OLD') -and $_.Module -ne 'ActiveDirectory' })

if ($todo.Count -eq 0) {
    Write-Step 'Nothing to install.'
}
elseif (-not $Execute) {
    Write-Step "$($todo.Count) module(s) would be installed or updated"
    foreach ($r in $todo) { Write-Host "    $($r.Module)  [$($r.Status)]" }
    Write-Warn 'Re-run with -Execute to install them.'
}
else {
    if ($Scope -eq 'AllUsers') {
        $elevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
                    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if (-not $elevated) { Write-Die '-Scope AllUsers needs an elevated session.' }
    }

    Write-Step "Installing $($todo.Count) module(s)"
    $failed = New-Object System.Collections.Generic.List[string]

    foreach ($r in $todo) {
        Write-Host ("    {0,-46} " -f $r.Module) -NoNewline
        try {
            $args = @{ Name = $r.Module; Scope = $Scope; Force = $true; AllowClobber = $true; ErrorAction = 'Stop' }
            if ($r.Entry.MinVersion) { $args['MinimumVersion'] = $r.Entry.MinVersion }
            Install-Module @args
            $now = (Get-Module -ListAvailable -Name $r.Module | Sort-Object Version -Descending | Select-Object -First 1).Version
            Write-Host "OK  $now" -ForegroundColor Green
        }
        catch {
            Write-Host "FAILED" -ForegroundColor Red
            Write-Host "        $($_.Exception.Message)" -ForegroundColor DarkGray
            $failed.Add($r.Module)
        }
    }

    if ($failed.Count) {
        Write-Warn "$($failed.Count) module(s) failed: $($failed -join ', ')"
        Write-Host '    The usual causes: no network to the PowerShell Gallery, TLS 1.2 not enabled on' -ForegroundColor DarkGray
        Write-Host '    Windows PowerShell 5.1, or an untrusted repository. For the second one:' -ForegroundColor DarkGray
        Write-Host '      [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12' -ForegroundColor DarkGray
    }
    else { Write-OK 'All requested modules installed.' }
}

# ==================== NOTES THAT SAVE TIME ====================

Write-Step 'Worth knowing'

if ($report.Module -contains 'ActiveDirectory' -and ($report | Where-Object Module -eq 'ActiveDirectory').Status -eq 'MISSING') {
    Write-Host '    ActiveDirectory does not come from the PowerShell Gallery. On Windows 10/11:' -ForegroundColor DarkGray
    Write-Host '      Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0' -ForegroundColor DarkGray
    Write-Host '    On a server: Install-WindowsFeature RSAT-AD-PowerShell' -ForegroundColor DarkGray
}

Write-Host '    The Microsoft.Graph sub-modules are installed individually here on purpose. The' -ForegroundColor DarkGray
Write-Host '    Microsoft.Graph meta-module pulls in dozens of them, takes a long time, and is the' -ForegroundColor DarkGray
Write-Host '    usual cause of an import that fails with an assembly conflict. Install only what you' -ForegroundColor DarkGray
Write-Host '    need. If you already have the meta-module and things break, that is where to look.' -ForegroundColor DarkGray
Write-Host ''
Write-Host '    Mixing PnP.PowerShell with the Graph SDK or Az in one session can conflict over MSAL' -ForegroundColor DarkGray
Write-Host '    assemblies. If a script authenticates fine on its own and fails after you loaded' -ForegroundColor DarkGray
Write-Host '    another module, open a fresh session before you start debugging the script.' -ForegroundColor DarkGray

Write-Host "`nFinished $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))`n"
