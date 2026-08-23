<#
.SYNOPSIS
    Endpoint detection: is Visio / Project installed, and when was it last used.

.DESCRIPTION
    Runs ON a device, in the LOGGED-ON USER's context, and reports per-app:
      - Installed (yes/no), edition, exe version
      - Last used date  = max of two endpoint signals:
            * Office File MRU  (last file opened in the app, with timestamp)
            * UserAssist       (last VISIO.EXE / WINPROJ.EXE launch + run count)

    There is NO Microsoft cloud API for desktop Visio/Project usage. These two
    local signals are the supported, app-specific way to get a real last-used
    date. The script is read-only (HKCU/HKLM reads only) and emits ONE line of
    JSON to stdout - shaped for Intune Remediations, which captures the last
    stdout line. Optionally also appends a CSV row to a share for logon-script
    collection.

.PARAMETER OutputShare
    Optional UNC path. If set, appends one CSV row (<Computer>_<User>.csv style
    file) for collection via GPO logon script. Omit for Intune Remediations
    (which captures stdout instead).

.NOTES
    DEPLOYMENT
      Intune Remediations: upload as the *detection* script, set
        "Run this script using the logged-on credentials = Yes"
        "Enforce script signature check = No", "Run in 64-bit PowerShell = Yes".
      Collect results from the Remediations report (Pre-remediation detection
      output column), or via Graph deviceHealthScripts run states.
      GPO logon script: deploy with -OutputShare \\server\share\VPUsage

    CONTEXT REQUIREMENT
      Must run as the signed-in user (reads that user's HKCU). If it runs as
      SYSTEM it will say so in the output and skip the per-user signals.

    SIGNALS ARE A FLOOR
      MRU only populates when a file is opened/saved; a user who only ever opens
      blank docs is caught by UserAssist instead. "No signal" is not proof of
      non-use - pair with device coverage in the report.

    READ-ONLY. No writes to the device.
#>

[CmdletBinding()]
param(
    [string]$OutputShare
)

$ErrorActionPreference = 'SilentlyContinue'

$Apps = @(
    [pscustomobject]@{ Name='Visio';   Exe='VISIO.EXE';   OfficeKey='Visio';      Edition='Visio'   }
    [pscustomobject]@{ Name='Project'; Exe='WINPROJ.EXE'; OfficeKey='MS Project'; Edition='Project' }
)

# --- helpers ---------------------------------------------------------------
function ConvertFrom-Rot13([string]$s){
    -join ($s.ToCharArray() | ForEach-Object {
        $c = $_
        if ($c -cmatch '[A-Z]') { [char]((([int]$c - 65 + 13) % 26) + 65) }
        elseif ($c -cmatch '[a-z]') { [char]((([int]$c - 97 + 13) % 26) + 97) }
        else { $c }
    })
}

function Get-AppPathExe([string]$exe){
    foreach($root in 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths',
                     'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths'){
        $p = (Get-ItemProperty "$root\$exe" -ErrorAction SilentlyContinue).'(default)'
        if($p -and (Test-Path $p)){ return $p }
    }
    return $null
}

function Get-ClickToRunProducts {
    (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue).ProductReleaseIds
}

# Last file-open time from Office File MRU (any identity, Office 15.0/16.0)
function Get-OfficeMruLastUsed([string]$officeKey){
    $max = $null
    foreach($ver in '16.0','15.0'){
        $base = "HKCU:\Software\Microsoft\Office\$ver\$officeKey"
        $mruRoots = @("$base\File MRU")
        $userMru = "$base\User MRU"
        if(Test-Path $userMru){
            Get-ChildItem $userMru -ErrorAction SilentlyContinue | ForEach-Object { $mruRoots += "$($_.PSPath)\File MRU" }
        }
        foreach($mru in $mruRoots){
            if(-not (Test-Path $mru)){ continue }
            $k = Get-Item $mru -ErrorAction SilentlyContinue
            foreach($name in $k.GetValueNames()){
                $val = [string]$k.GetValue($name)
                $m = [regex]::Match($val,'\[T([0-9A-Fa-f]{16})\]')
                if($m.Success){
                    try {
                        $ft = [Convert]::ToInt64($m.Groups[1].Value,16)
                        $dt = [DateTime]::FromFileTime($ft)
                        if((-not $max) -or ($dt -gt $max)){ $max = $dt }
                    } catch {}
                }
            }
        }
    }
    return $max
}

# Last launch + run count from UserAssist (executable execution GUID)
function Get-UserAssist([string]$exe){
    $guid = '{CEBFF5CD-ACE2-4F4F-9178-9926F41749EA}'
    $key  = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\UserAssist\$guid\Count"
    if(-not (Test-Path $key)){ return $null }
    $k = Get-Item $key -ErrorAction SilentlyContinue
    $best = $null
    foreach($name in $k.GetValueNames()){
        $decoded = ConvertFrom-Rot13 $name
        if($decoded -match [regex]::Escape($exe) + '\s*$'){
            $bytes = [byte[]]$k.GetValue($name)
            if($bytes.Length -ge 68){
                $count = [BitConverter]::ToInt32($bytes,4)
                $ft    = [BitConverter]::ToInt64($bytes,60)
                $dt    = if($ft -gt 0){ [DateTime]::FromFileTime($ft) } else { $null }
                if($dt -and ((-not $best) -or ($dt -gt $best.LastRun))){
                    $best = [pscustomobject]@{ LastRun=$dt; RunCount=$count }
                }
            }
        }
    }
    return $best
}

# --- context check ---------------------------------------------------------
$whoUpn = (whoami /upn) 2>$null
$whoNt  = (whoami) 2>$null
$isSystem = $whoNt -match 'nt authority\\system'

$c2r = Get-ClickToRunProducts
$result = [ordered]@{
    Computer    = $env:COMPUTERNAME
    User        = if($whoUpn){ $whoUpn } else { $env:USERNAME }
    Context     = if($isSystem){ 'SYSTEM (per-user signals skipped)' } else { 'User' }
    CollectedAt = (Get-Date).ToString('s')
}

foreach($app in $Apps){
    $exePath = Get-AppPathExe $app.Exe
    $installed = [bool]$exePath -or ($c2r -match $app.Edition)
    $ver = if($exePath){ (Get-Item $exePath).VersionInfo.ProductVersion } else { $null }

    $mru = $null; $ua = $null
    if(-not $isSystem){
        $mru = Get-OfficeMruLastUsed $app.OfficeKey
        $ua  = Get-UserAssist $app.Exe
    }
    $candidates = @($mru, $ua.LastRun) | Where-Object { $_ }
    $lastUsed = if($candidates){ ($candidates | Sort-Object -Descending)[0] } else { $null }

    $result["$($app.Name)Installed"]  = $installed
    $result["$($app.Name)Version"]    = $ver
    $result["$($app.Name)LastUsed"]   = if($lastUsed){ $lastUsed.ToString('yyyy-MM-dd') } else { $null }
    $result["$($app.Name)RunCount"]   = if($ua){ $ua.RunCount } else { $null }
    $src = @(); if($mru){ $src += 'MRU' }; if($ua){ $src += 'UserAssist' }
    $result["$($app.Name)Source"]     = ($src -join '+')
}

# --- optional share write for logon-script collection ----------------------
if($OutputShare -and (Test-Path $OutputShare)){
    try {
        $row  = [pscustomobject]$result
        $file = Join-Path $OutputShare ("{0}_{1}.csv" -f $env:COMPUTERNAME, ($result.User -replace '[\\@:]','_'))
        $row | Export-Csv $file -NoTypeInformation -Encoding UTF8 -Force
    } catch {}
}

# --- single JSON line for Intune Remediations stdout capture ---------------
($result | ConvertTo-Json -Compress)
exit 0
