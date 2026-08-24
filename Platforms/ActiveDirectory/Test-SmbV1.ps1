<#
.SYNOPSIS
  Unauthenticated network probe for SMBv1 on domain controllers or file servers.

  ####################################################################
  #  AUTHORISATION REQUIRED BEFORE USE                               #
  #                                                                  #
  #  This is a NETWORK SCANNER. It opens raw TCP connections to      #
  #  hosts you name and sends a hand-built SMB protocol packet.      #
  #  Running it against infrastructure you do not own or do not      #
  #  have written permission to test may be treated as hostile       #
  #  activity, may breach computer-misuse law in your jurisdiction,  #
  #  and will very likely trip IDS/EDR alerts and an incident.       #
  #                                                                  #
  #  Only run it against hosts you are explicitly authorised to      #
  #  test, and tell your security team before you do.                #
  ####################################################################

.DESCRIPTION
  Replicates how vulnerability scanners (e.g. Tenable Identity Exposure S-SMB-v1) detect
  SMBv1: opens TCP 445 and sends an SMBv1 SMB_COM_NEGOTIATE offering ONLY SMB1 dialects,
  then inspects the response.
    - Server selects an SMB1 dialect   -> SMBv1 ENABLED (finding confirmed for that host)
    - Server resets / rejects dialects -> SMBv1 disabled

  SMB dialect negotiation happens BEFORE authentication, so this needs NO credentials and
  NO admin rights on the target - only reachability to TCP 445. Read-only: no authentication
  is attempted, nothing is exploited, nothing on the target is changed.

  Four distinct outcomes, deliberately kept apart:
    ENABLED  - the server picked an SMB1 dialect. Confirmed finding.
    DISABLED - the server refused, answered with an SMB2-only stack, or reset the
               connection after negotiation (a post-negotiate RST is the normal
               behaviour of a host with the SMB1 server driver stopped, so it is
               treated as a valid "disabled" signal rather than an error).
    NO_TCP   - port 445 unreachable or filtered. NOT a verdict: a firewalled host
               proves nothing about whether SMB1 is enabled on it.
    UNCLEAR  - answered, but not in a shape this probe understands. Inspect by hand.

  The NO_TCP / DISABLED distinction is the point of the script. Reporting a filtered
  host as "disabled" would turn a blocked scan into a false all-clear.

.PARAMETER ComputerName
  One or more hostnames or IPs to probe. Mandatory: there is no default target list,
  by design.

.PARAMETER Port
  TCP port to probe. Default 445.

.PARAMETER TimeoutMs
  Connect/read timeout in milliseconds. Default 5000.

.PARAMETER ExportDir
  Folder for the results CSV.

.EXAMPLE
  .\Test-SmbV1.ps1 -ComputerName DC01

.EXAMPLE
  .\Test-SmbV1.ps1 -ComputerName DC01,DC02,FS01 -TimeoutMs 10000

.NOTES
    When to use  : A vulnerability scan flags SMBv1 on domain controllers or file servers and you have to prove host by host whether it is genuinely enabled.
    Why it exists: Four outcomes instead of two. A post-negotiate RST counts as DISABLED (it is the normal behaviour of a host with the SMB1 server driver stopped) and a filtered port counts as nothing at all - reporting a firewalled host as 'disabled' would turn a blocked scan into a false all-clear.
  Run as a FILE. If 445 is filtered you get NO_TCP, which is not a verdict.
  Alternative if nmap is available: nmap -p445 --script smb-protocols <host>
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string[]] $ComputerName,
    [int]    $Port      = 445,
    [int]    $TimeoutMs = 5000,
    [string] $ExportDir = (Join-Path $PSScriptRoot 'Exports')
)

function Write-Step($m){ Write-Host "[*] $m" -ForegroundColor Cyan }
function Write-OK  ($m){ Write-Host "[+] $m" -ForegroundColor Green }
function Write-Warn($m){ Write-Host "[!] $m" -ForegroundColor Yellow }

function Test-SmbV1Offered {
    param([string]$ComputerName,[int]$Port=445,[int]$TimeoutMs=5000)

    # SMB1-only dialect list (deliberately no "SMB 2.???", so an SMB2-only server can't upgrade)
    $dialects = @('PC NETWORK PROGRAM 1.0','LANMAN1.0','Windows for Workgroups 3.1a','LM1.2X002','LANMAN2.1','NT LM 0.12')
    $enc = [System.Text.Encoding]::ASCII
    $dlg = New-Object System.Collections.Generic.List[byte]
    foreach($d in $dialects){ $dlg.Add(0x02); $dlg.AddRange($enc.GetBytes($d)); $dlg.Add(0x00) }

    $smb = New-Object System.Collections.Generic.List[byte]
    $smb.AddRange([byte[]](0xFF,0x53,0x4D,0x42))        # Protocol  \xffSMB
    $smb.Add(0x72)                                       # SMB_COM_NEGOTIATE
    $smb.AddRange([byte[]](0,0,0,0))                     # Status
    $smb.Add(0x18)                                       # Flags
    $smb.AddRange([byte[]](0x01,0x28))                   # Flags2
    $smb.AddRange([byte[]](0,0))                         # PIDHigh
    $smb.AddRange([byte[]](0,0,0,0,0,0,0,0))             # Signature
    $smb.AddRange([byte[]](0,0))                         # Reserved
    $smb.AddRange([byte[]](0,0))                         # TID
    $smb.AddRange([byte[]](0x2F,0x4B))                   # PIDLow
    $smb.AddRange([byte[]](0,0))                         # UID
    $smb.AddRange([byte[]](0,0))                         # MID
    $smb.Add(0x00)                                       # WordCount
    $smb.AddRange([BitConverter]::GetBytes([UInt16]$dlg.Count)) # ByteCount (LE)
    $smb.AddRange($dlg)

    $len    = $smb.Count
    $nbss   = [byte[]](0x00, (($len -shr 16) -band 0xFF), (($len -shr 8) -band 0xFF), ($len -band 0xFF))
    $packet = $nbss + $smb.ToArray()

    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $iar = $client.BeginConnect($ComputerName,$Port,$null,$null)
        if(-not $iar.AsyncWaitHandle.WaitOne($TimeoutMs)){ return [pscustomobject]@{ Result='NO_TCP'; Detail="445 unreachable/timeout" } }
        $client.EndConnect($iar)
        $ns = $client.GetStream(); $ns.ReadTimeout = $TimeoutMs
        $ns.Write($packet,0,$packet.Length); $ns.Flush()

        $buf  = New-Object byte[] 1024
        $read = $ns.Read($buf,0,$buf.Length)
        if($read -lt 39){ return [pscustomobject]@{ Result='DISABLED'; Detail="short/no response ($read bytes)" } }

        $isSmb = ($buf[4] -eq 0xFF -and $buf[5] -eq 0x53 -and $buf[6] -eq 0x4D -and $buf[7] -eq 0x42)
        if(-not $isSmb){ return [pscustomobject]@{ Result='DISABLED'; Detail='non-SMB1 response (SMB2-only stack)' } }
        if($buf[8] -ne 0x72){ return [pscustomobject]@{ Result='UNCLEAR'; Detail=("unexpected cmd 0x{0:X2}" -f $buf[8]) } }

        $statusZero   = (($buf[9] -bor $buf[10] -bor $buf[11] -bor $buf[12]) -eq 0)
        $wordCount    = $buf[36]
        $dialectIndex = [BitConverter]::ToUInt16($buf,37)

        if($statusZero -and $wordCount -ge 1 -and $dialectIndex -ne 0xFFFF){
            return [pscustomobject]@{ Result='ENABLED'; Detail="SMB1 negotiate OK, dialect idx=$dialectIndex" }
        }
        return [pscustomobject]@{ Result='DISABLED'; Detail="no dialect selected (idx=$dialectIndex, status0=$statusZero)" }
    } catch {
        # Post-negotiate RST is the normal signal for a DC with the SMB1 server driver stopped
        return [pscustomobject]@{ Result='DISABLED'; Detail=("reset/error: " + $_.Exception.Message.Split([char]10)[0]) }
    } finally { $client.Close() }
}

# ---- Main ----
if(-not (Test-Path $ExportDir)){ New-Item -ItemType Directory -Path $ExportDir | Out-Null }
$results = New-Object System.Collections.Generic.List[object]

foreach($dc in $ComputerName){
    Write-Step "Probing $dc : $Port ..."
    $r = Test-SmbV1Offered -ComputerName $dc -Port $Port -TimeoutMs $TimeoutMs
    $verdict = switch($r.Result){
        'ENABLED'  { 'SMBv1 ENABLED (finding CONFIRMED)' }
        'DISABLED' { 'SMBv1 disabled (finding NOT reproduced)' }
        'NO_TCP'   { 'PORT 445 UNREACHABLE - inconclusive' }
        default    { 'UNCLEAR - inspect manually' }
    }
    switch($r.Result){
        'ENABLED'  { Write-Warn "$dc -> $verdict" }
        'DISABLED' { Write-OK   "$dc -> $verdict" }
        default    { Write-Warn "$dc -> $verdict" }
    }
    $results.Add([pscustomobject]@{ ComputerName=$dc; Port=$Port; Result=$r.Result; Verdict=$verdict; Detail=$r.Detail })
}

$results | Format-Table ComputerName,Port,Result,Verdict,Detail -AutoSize

$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$csv   = Join-Path $ExportDir "SMBv1-NetworkProbe_$stamp.csv"
$results | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
Write-OK "Exported: $csv"
