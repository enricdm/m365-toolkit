<#
.SYNOPSIS
  Shared Microsoft Graph helpers for the M365 tooling in this tree.

.DESCRIPTION
  Single home for Invoke-GraphPaged. Three scripts used to carry private copies of
  this function; only one of them was ever fixed for the phantom-record bug, and the
  other two kept injecting a fake record for every empty collection. That divergence
  is exactly what this module exists to prevent - fix it once, here.

  Requires an active Microsoft Graph session (Connect-MgGraph) in the caller; the
  module resolves Invoke-MgGraphRequest from the caller's session.

.NOTES
    When to use  : Imported by any new script that talks to Graph. It is the first line you write, not the last.
    Why it exists: Invoke-GraphPaged lived in three private copies, only one was ever fixed for the phantom-record bug, and the other two kept injecting a fake record for every empty collection. An empty collection returns an EMPTY list here, never a one-element list holding the raw response, and 429s are backed off.
#>

# Private helper. Dictionary-safe property read: Invoke-MgGraphRequest returns
# hashtables, not PSObjects, so .PSObject.Properties finds Keys/Count instead of
# the data. NOTE: this returns $null for an empty value - never use it to decide
# whether a key EXISTS (see the phantom-record comment below).
function Get-GraphProp {
    param($Obj, [string[]]$Names)
    if ($null -eq $Obj) { return $null }
    foreach ($n in $Names) {
        $v = $null
        if ($Obj -is [System.Collections.IDictionary]) { if ($Obj.Contains($n)) { $v = $Obj[$n] } }
        else { $p = $Obj.PSObject.Properties[$n]; if ($p) { $v = $p.Value } }
        if ($null -ne $v -and "$v" -ne '') { return $v }
    }
    return $null
}

function Invoke-GraphPaged {
<#
.SYNOPSIS
  Paged Microsoft Graph GET with 429 backoff.

.PARAMETER Uri
  Absolute Graph URI to start from. nextLink pages are followed automatically.

.PARAMETER MaxItems
  Return at most this many items. 0 (default) = no cap. The last page fetched
  may contain more; the result is trimmed to exactly this count.

.OUTPUTS
  A List[object] of items. An empty collection returns an EMPTY list - never a
  one-element list containing the raw response.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [int]$MaxItems = 0
    )

    $out  = New-Object System.Collections.Generic.List[object]
    $next = $Uri
    while ($next) {
        $resp = $null
        for ($try = 1; $try -le 5; $try++) {
            try { $resp = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop; break }
            catch {
                # Prefer the real status code when the exception carries one;
                # fall back to words, not to a bare "429" that could sit inside
                # a GUID or URL in the message.
                $status = $null
                try {
                    $raw = $_.Exception.Response.StatusCode
                    # A bare [int]$null would give 0 and mask the "no status
                    # code at all" case, so cast only when something is there.
                    if ($null -ne $raw) { $status = [int]$raw }
                } catch {}
                $msg = "$($_.Exception.Message)"
                $throttled = ($status -eq 429) -or
                             ($null -eq $status -and $msg -match 'throttl|too many requests|status(?: code)?:? 429')
                if ($throttled -and $try -lt 5) {
                    $wait = [math]::Pow(2, $try) * 5
                    Write-Warning "Throttled, waiting $wait s (attempt $try/5)"
                    Start-Sleep -Seconds $wait
                } else { throw }
            }
        }

        # Determine whether the response IS a collection, before reading its contents.
        # A value-reading helper returns $null for an empty array (it stringifies to ''),
        # so relying on one here made every empty collection look like a single-object
        # response and injected a phantom record. Test for the key itself instead.
        $hasValueKey = $false
        if ($resp -is [System.Collections.IDictionary]) { $hasValueKey = $resp.Contains('value') }
        elseif ($resp) { $hasValueKey = [bool]$resp.PSObject.Properties['value'] }

        if ($hasValueKey) {
            $vals = if ($resp -is [System.Collections.IDictionary]) { $resp['value'] } else { $resp.value }
            foreach ($v in @($vals)) { if ($null -ne $v) { [void]$out.Add($v) } }
        } elseif ($resp) {
            [void]$out.Add($resp)
        }

        if ($MaxItems -gt 0 -and $out.Count -ge $MaxItems) { break }
        $next = Get-GraphProp $resp @('@odata.nextLink', 'odata.nextLink')
    }
    if ($MaxItems -gt 0 -and $out.Count -gt $MaxItems) {
        $out.RemoveRange($MaxItems, $out.Count - $MaxItems)
    }
    return , $out
}

Export-ModuleMember -Function Invoke-GraphPaged
