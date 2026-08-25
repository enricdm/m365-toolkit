<#
.SYNOPSIS
    Proves whether a list of host:port destinations is actually reachable from the
    machine you run it on, in four layers, and tells you which layer failed.
    Read-only.

.DESCRIPTION
    "The firewall rule is open" and "the application can talk to it" are different
    claims, and an open socket is not evidence of the second. So each destination is
    tested in four separate layers, and each one proves something the others do not:

      1. DNS resolution           -> internal DNS can resolve the name at all
      2. TCP connect (timed)      -> the firewall rule is genuinely open
      3. TLS handshake (443 only) -> nothing is breaking and re-signing the session
      4. HTTP request (80 only)   -> the service answers, not just the socket

    Layer 3 is the one that earns its keep. A TLS-inspecting proxy will happily let the
    connection through and hand you a certificate it signed itself, which passes any
    test that stops at "did it connect". Software that pins certificates or validates
    the issuer will then fail in production with an error that looks nothing like a
    firewall problem. This reports the issuer and flags it when it looks like
    interception.

    WHERE YOU RUN IT MATTERS MORE THAN WHAT IT REPORTS. Run it locally on the machine
    that will really make the connections. Running it from an admin jump box proves the
    jump box's network path is open and tells you nothing about the one you asked for,
    while producing output that looks exactly like a successful test. The source
    hostname and local network context are recorded in every row so the evidence says
    where it came from.

.PARAMETER Destination
    Destinations as 'host:port', or bare 'host' for port 443. Repeatable.

.PARAMETER InputCsv
    CSV of destinations. Required column: Host. Optional: Port (default 443), Purpose,
    Wildcard (the original wildcard from the firewall request, if this name is an
    expansion of one).

.PARAMETER TimeoutMs
    TCP connect timeout per destination. Default 5000.

.PARAMETER SkipTls / SkipHttp
    Skip layer 3 or layer 4.

.PARAMETER UseProxy
    Follow the system proxy for the HTTP tests. Off by default, because the question is
    usually whether the direct path works.

.EXAMPLE
    # A couple of destinations
    .\Test-EndpointConnectivity.ps1 -Destination 'login.microsoftonline.com','graph.microsoft.com:443'

.EXAMPLE
    # A vendor's published endpoint list, with a longer timeout
    .\Test-EndpointConnectivity.ps1 -InputCsv .\vendor-endpoints.csv -TimeoutMs 8000

.EXAMPLE
    # Port 80 CRL/OCSP checks, no TLS layer
    .\Test-EndpointConnectivity.ps1 -Destination 'ocsp.digicert.com:80','crl3.digicert.com:80' -SkipTls

.NOTES
    When to use  : You asked for a firewall rule for a new deployment and now have to prove it is in place and that a proxy or TLS inspection is not quietly breaking it.
    Why it exists: Four layers per destination, each proving something different, and it names the layer that failed instead of reporting one pass/fail. It flags TLS interception by issuer, which is the failure that survives every simpler test and then breaks the application in production. The blocked destinations are printed at the end in a form you can paste straight into the firewall request.
    Requires : Windows PowerShell 5.1 or PowerShell 7. No modules.
    Rights   : none beyond running locally. Nothing is modified.
#>

[CmdletBinding(DefaultParameterSetName = 'Inline')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Inline', Position = 0)]
    [string[]]$Destination,

    [Parameter(Mandatory, ParameterSetName = 'Csv')]
    [string]$InputCsv,

    [int]$TimeoutMs     = 5000,
    [switch]$SkipTls,
    [switch]$SkipHttp,
    [switch]$UseProxy,
    [string]$ExportPath = (Join-Path $PSScriptRoot 'Exports')
)

# ==================== CONFIG ====================

# Certificate issuer fragments that mean the session is being intercepted and re-signed
# rather than terminated by the real public CA. Add your own perimeter product, and add
# your organisation's internal CA name: an internal issuer on a public destination is
# the clearest possible sign of break-and-inspect.
$InterceptPatterns = @(
    'Fortinet', 'FG-', 'Zscaler', 'Palo Alto', 'Netskope', 'Blue Coat',
    'Symantec Web', 'Forcepoint', 'Sophos', 'McAfee Web', 'Internal'
)

# ==================== HELPERS ====================

function Write-Step { param([string]$m) Write-Host "`n[*] $m" -ForegroundColor Cyan }
function Write-OK   { param([string]$m) Write-Host "[+] $m" -ForegroundColor Green }
function Write-Warn { param([string]$m) Write-Host "[!] $m" -ForegroundColor Yellow }
function Write-Die  { param([string]$m) Write-Host "[X] $m" -ForegroundColor Red; exit 1 }

function Test-TcpPort {
    param([string]$Target,[int]$Port,[int]$Timeout)
    $out = [pscustomobject]@{ Success=$false; LatencyMs=$null; Error=$null }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $iar = $client.BeginConnect($Target,$Port,$null,$null)
        if ($iar.AsyncWaitHandle.WaitOne($Timeout,$false)) { $client.EndConnect($iar); $out.Success = $true }
        else { $out.Error = "Timeout after $Timeout ms" }
    } catch { $out.Error = $_.Exception.Message.Trim() }
    finally { $sw.Stop(); $client.Close() }
    $out.LatencyMs = [math]::Round($sw.Elapsed.TotalMilliseconds,0)
    return $out
}

function Test-TlsHandshake {
    param([string]$Target,[int]$Port,[int]$Timeout)
    $out = [pscustomobject]@{ Success=$false; Protocol=$null; Subject=$null; Issuer=$null; Expiry=$null; Intercepted=$false; Error=$null }
    $client = New-Object System.Net.Sockets.TcpClient
    $ssl = $null
    try {
        $iar = $client.BeginConnect($Target,$Port,$null,$null)
        if (-not $iar.AsyncWaitHandle.WaitOne($Timeout,$false)) { $out.Error = 'TCP timeout'; return $out }
        $client.EndConnect($iar)
        $cb = [System.Net.Security.RemoteCertificateValidationCallback]{ param($s,$c,$ch,$e) return $true }
        $ssl = New-Object System.Net.Security.SslStream($client.GetStream(),$false,$cb)

        $protos = [System.Security.Authentication.SslProtocols]::Tls12
        try { $protos = $protos -bor [System.Security.Authentication.SslProtocols]::Tls13 } catch { }
        try   { $ssl.AuthenticateAsClient($Target,$null,$protos,$false) }
        catch { $ssl.AuthenticateAsClient($Target,$null,[System.Security.Authentication.SslProtocols]::Tls12,$false) }

        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($ssl.RemoteCertificate)
        $out.Success  = $true
        $out.Protocol = $ssl.SslProtocol.ToString()
        $out.Subject  = $cert.Subject
        $out.Issuer   = $cert.Issuer
        $out.Expiry   = $cert.NotAfter.ToString('dd/MM/yyyy')
        foreach ($p in $InterceptPatterns) { if ($cert.Issuer -match [regex]::Escape($p)) { $out.Intercepted = $true; break } }
    } catch { $out.Error = $_.Exception.Message.Trim() }
    finally { if ($ssl) { $ssl.Dispose() }; $client.Close() }
    return $out
}

function Test-HttpEndpoint {
    param([string]$Target,[int]$Timeout,[bool]$Proxy)
    $out = [pscustomobject]@{ Success=$false; StatusCode=$null; Error=$null }
    try {
        $req = [System.Net.HttpWebRequest]::Create("http://$Target/")
        $req.Method            = 'GET'
        $req.Timeout           = $Timeout
        $req.AllowAutoRedirect = $false
        $req.UserAgent         = 'EndpointConnectivityTest'
        if (-not $Proxy) { $req.Proxy = $null }
        $resp = $req.GetResponse()
        $out.Success    = $true
        $out.StatusCode = [int]$resp.StatusCode
        $resp.Close()
    } catch [System.Net.WebException] {
        if ($_.Exception.Response) {
            # A 4xx/5xx still proves the request reached the far end at layer 7
            $out.Success    = $true
            $out.StatusCode = [int]$_.Exception.Response.StatusCode
            $_.Exception.Response.Close()
        } else { $out.Error = $_.Exception.Message.Trim() }
    } catch { $out.Error = $_.Exception.Message.Trim() }
    return $out
}

# ==================== MAIN ====================

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$startedAt = Get-Date
$me = $env:COMPUTERNAME

Write-Step "Endpoint connectivity test"
Write-Host "    Source host : $me"
Write-Host "    Started     : $($startedAt.ToString('yyyy-MM-dd HH:mm:ss'))"

# --- build the destination queue ---
$queue = New-Object System.Collections.Generic.List[object]

if ($PSCmdlet.ParameterSetName -eq 'Csv') {
    if (-not (Test-Path $InputCsv)) { Write-Die "Input CSV not found: $InputCsv" }
    $rows = @(Import-Csv -Path $InputCsv)
    if (-not $rows) { Write-Die "'$InputCsv' has no rows." }
    if (-not ($rows[0].PSObject.Properties.Name -contains 'Host')) {
        Write-Die "'$InputCsv' needs a 'Host' column. Optional: Port, Purpose, Wildcard."
    }
    foreach ($r in $rows) {
        if (-not $r.Host) { continue }
        $queue.Add([pscustomobject]@{
            Host     = $r.Host.Trim()
            Port     = if ($r.PSObject.Properties.Name -contains 'Port' -and $r.Port) { [int]$r.Port } else { 443 }
            Purpose  = if ($r.PSObject.Properties.Name -contains 'Purpose')  { $r.Purpose }  else { '' }
            Wildcard = if ($r.PSObject.Properties.Name -contains 'Wildcard') { $r.Wildcard } else { '' }
        })
    }
}
else {
    foreach ($d in $Destination) {
        $spec = $d.Trim()
        if (-not $spec) { continue }
        # host:port, or bare host meaning 443. IPv6 in brackets keeps its colons.
        if ($spec -match '^\[(?<h>.+)\](:(?<p>\d+))?$' -or $spec -match '^(?<h>[^:]+)(:(?<p>\d+))?$') {
            $queue.Add([pscustomobject]@{
                Host     = $Matches['h']
                Port     = if ($Matches['p']) { [int]$Matches['p'] } else { 443 }
                Purpose  = ''
                Wildcard = ''
            })
        }
        else { Write-Die "Cannot parse destination '$spec'. Use 'host' or 'host:port'." }
    }
}

if ($queue.Count -eq 0) { Write-Die 'No destinations to test.' }
Write-OK "$($queue.Count) destination(s) queued"

# --- local network context, recorded so the evidence says where it came from ---
Write-Step 'Local network context'
try {
    $nics = @(Get-NetIPConfiguration -ErrorAction Stop | Where-Object { $_.IPv4DefaultGateway })
    if ($nics.Count -eq 0) { Write-Warn 'No interface with a default gateway.' }
    foreach ($n in $nics) {
        Write-Host "    $($n.InterfaceAlias): $($n.IPv4Address.IPAddress) | GW $($n.IPv4DefaultGateway.NextHop) | DNS $($n.DNSServer.ServerAddresses -join ', ')"
    }
}
catch { Write-Warn "Could not read the local network configuration: $($_.Exception.Message)" }

try {
    $winhttp = (netsh winhttp show proxy) -join ' '
    Write-Host "    WinHTTP proxy: $($winhttp -replace '\s+',' ')"
}
catch { Write-Warn 'Could not read the WinHTTP proxy setting.' }

# --- run the tests ---
Write-Step 'Running tests'
$results = New-Object System.Collections.Generic.List[object]

foreach ($t in $queue) {
    $label = "{0}:{1}" -f $t.Host, $t.Port
    Write-Host ("    {0,-58} " -f $label) -NoNewline

    $row = [ordered]@{
        Timestamp   = (Get-Date).ToString('dd/MM/yyyy HH:mm:ss')
        SourceHost  = $me
        Destination = $t.Host
        Port        = $t.Port
        Wildcard    = $t.Wildcard
        Purpose     = $t.Purpose
        DnsResult   = 'FAIL'
        ResolvedIPs = ''
        TcpResult   = 'FAIL'
        LatencyMs   = ''
        L7Result    = 'SKIP'
        L7Detail    = ''
        TlsProtocol = ''
        CertIssuer  = ''
        CertExpiry  = ''
        Intercepted = ''
        Verdict     = 'FAIL'
        Error       = ''
    }

    # 1. DNS
    try {
        $dns = Resolve-DnsName -Name $t.Host -Type A -ErrorAction Stop | Where-Object { $_.IPAddress }
        if ($dns) { $row.DnsResult = 'OK'; $row.ResolvedIPs = ($dns.IPAddress | Select-Object -Unique) -join ';' }
        else      { $row.Error = 'DNS returned no A record' }
    } catch { $row.Error = "DNS: $($_.Exception.Message.Trim())" }

    if ($row.DnsResult -ne 'OK') {
        $row.Verdict = 'FAIL (DNS)'
        Write-Host 'DNS FAIL' -ForegroundColor Red
        $results.Add([pscustomobject]$row); continue
    }

    # 2. TCP
    $tcp = Test-TcpPort -Target $t.Host -Port $t.Port -Timeout $TimeoutMs
    $row.LatencyMs = $tcp.LatencyMs
    if ($tcp.Success) { $row.TcpResult = 'OK' }
    else {
        $row.TcpResult = 'FAIL'; $row.Verdict = 'FAIL (TCP)'
        $row.Error = "TCP: $($tcp.Error)"
        Write-Host "TCP BLOCKED  ($($tcp.Error))" -ForegroundColor Red
        $results.Add([pscustomobject]$row); continue
    }

    # 3. Layer 7
    if ($t.Port -eq 443 -and -not $SkipTls) {
        $tls = Test-TlsHandshake -Target $t.Host -Port $t.Port -Timeout $TimeoutMs
        if ($tls.Success) {
            $row.L7Result    = 'OK'
            $row.TlsProtocol = $tls.Protocol
            $row.CertIssuer  = $tls.Issuer
            $row.CertExpiry  = $tls.Expiry
            $row.Intercepted = $tls.Intercepted
            $row.L7Detail    = $tls.Protocol
            if ($tls.Intercepted) {
                $row.Verdict = 'PASS (TLS INTERCEPTED)'
                Write-Host "TLS INTERCEPTED  <- $($tls.Issuer)" -ForegroundColor Yellow
            } else {
                $row.Verdict = 'PASS'
                Write-Host "PASS  $($tcp.LatencyMs) ms  $($tls.Protocol)" -ForegroundColor Green
            }
        } else {
            $row.L7Result = 'FAIL'; $row.Verdict = 'FAIL (TLS)'; $row.Error = "TLS: $($tls.Error)"
            Write-Host "TCP OK / TLS FAIL  ($($tls.Error))" -ForegroundColor Yellow
        }
    }
    elseif ($t.Port -eq 80 -and -not $SkipHttp) {
        $http = Test-HttpEndpoint -Target $t.Host -Timeout $TimeoutMs -Proxy:$UseProxy.IsPresent
        if ($http.Success) {
            $row.L7Result = 'OK'; $row.L7Detail = "HTTP $($http.StatusCode)"; $row.Verdict = 'PASS'
            Write-Host "PASS  $($tcp.LatencyMs) ms  HTTP $($http.StatusCode)" -ForegroundColor Green
        } else {
            $row.L7Result = 'FAIL'; $row.Verdict = 'FAIL (HTTP)'; $row.Error = "HTTP: $($http.Error)"
            Write-Host "TCP OK / HTTP FAIL  ($($http.Error))" -ForegroundColor Yellow
        }
    }
    else {
        $row.Verdict = 'PASS (TCP only)'
        Write-Host "PASS  $($tcp.LatencyMs) ms" -ForegroundColor Green
    }

    $results.Add([pscustomobject]$row)
}

# --- export ---
Write-Step 'Export'
if (-not (Test-Path $ExportPath)) { New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null }
$file = Join-Path $ExportPath ("EndpointConnectivity_{0}_{1}.csv" -f $me, (Get-Date -Format 'yyyyMMdd-HHmmss'))
$results | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8
Write-OK "Saved: $file"

# --- summary ---
Write-Step 'Summary'
$pass    = @($results | Where-Object { $_.Verdict -like 'PASS*' }).Count
$fail    = @($results | Where-Object { $_.Verdict -like 'FAIL*' }).Count
$interc  = @($results | Where-Object { $_.Intercepted -eq $true }).Count

Write-Host "    Passed : $pass / $($results.Count)"
if ($fail)   { Write-Warn "Failed : $fail" }
if ($interc) { Write-Warn "TLS interception detected on $interc destination(s) - certificate-pinning clients commonly break under SSL inspection. Raise an inspection bypass with the network team." }

if ($fail) {
    Write-Host "`n    Blocked destinations (for the firewall ticket):" -ForegroundColor Yellow
    $results | Where-Object { $_.Verdict -like 'FAIL*' } |
        Format-Table Destination, Port, ResolvedIPs, Verdict, Error -AutoSize
}

Write-Host "`nFinished $((Get-Date).ToString('dd/MM/yyyy HH:mm:ss')) - elapsed $([math]::Round(((Get-Date) - $startedAt).TotalSeconds,1))s`n"
