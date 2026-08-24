<#
.SYNOPSIS
    Analyzes users holding a given license SKU and tiers them as candidates for
    reclamation, shared-mailbox conversion, review, or keep — using multiple
    activity/identity signals. READ-ONLY: makes no changes.

.DESCRIPTION
    Built to free up seats on a fully-consumed SKU (default: Office 365 E3 /
    ENTERPRISEPACK). Gathers signals per licensed user and assigns a confidence
    tier plus a recommended action and the lever to pull (direct license vs which
    group it came from).

    SIGNALS GATHERED (all from Microsoft Graph; EXO optional):
      - AccountEnabled                       (disabled = strong reclaim signal)
      - signInActivity: last interactive,    (dormancy — the primary signal)
        non-interactive, successful sign-in
      - Mailbox last activity + size + archive (Graph usage report, one bulk call)
      - lastPasswordChangeDateTime           (password staleness — supporting signal)
      - createdDateTime                      (protects new joiners from false positives)
      - userType                             (Guests shouldn't hold E3)
      - onPremisesSyncEnabled                (context)
      - licenseAssignmentStates              (direct vs group-assigned — the lever)
      - Resource/functional naming heuristic (recep./tm./rx./ultra. clinic accounts etc.)
      - [EXO, optional] RecipientTypeDetails + LitigationHoldEnabled

    TIERS (transparent, rule-based — every row shows which signals fired):
      KEEP            active recently — do not touch
      REVIEW          stale-ish or weak/conflicting signals — human review
      CONVERT_SHARED  dormant but mailbox has data worth keeping, <50GB, no hold
      RECLAIM         disabled, never-used, or clearly reclaimable
      EXCLUDE         new joiner / on hold — explicitly protected from reclaim

.PARAMETER SkuPartNumber
    SKU to analyze. Default 'ENTERPRISEPACK' (Office 365 E3).

.PARAMETER InactiveDaysStrong
    Dormancy threshold (days) for strong reclaim/convert signal. Default 365.

.PARAMETER InactiveDaysSoft
    Dormancy threshold (days) for the REVIEW band. Default 30.

.PARAMETER NewJoinerGraceDays
    Accounts created within this many days that have never signed in are treated
    as onboarding and EXCLUDED from reclaim. Default 30.

.PARAMETER MailboxReportPeriod
    Graph mailbox usage report window: D7 | D30 | D90 | D180. Default D180.

.PARAMETER EnrichFromExchange
    If set, connects to EXO (one bulk call) to add RecipientTypeDetails and
    LitigationHoldEnabled. Off by default to avoid the WAM broker issue; the
    Graph-only path still produces full tiering.

.PARAMETER SampleSize
    For fast iteration: analyze only the first N licensed users. 0 = all. Default 0.

.PARAMETER ResourcePattern
    Regex fragments matched against the local part of the UPN to spot shared /
    functional accounts (reception desks, scanners, service accounts, ...).
    Replace with the naming your own tenant uses.

.PARAMETER OutputFolder
    Where the report CSV + log are written. Default: script folder.

.EXAMPLE
    # Quick first look on 200 users, Graph only:
    .\Get-LicenseReclamationPlan.ps1 -SampleSize 200

.EXAMPLE
    # Full E3 analysis with Exchange enrichment:
    .\Get-LicenseReclamationPlan.ps1 -EnrichFromExchange

.NOTES
    When to use  : A SKU runs out of seats and procurement asks whether to buy more or whether seats can be recovered.
    Why it exists: Tiers every holder into KEEP / REVIEW / CONVERT_SHARED / RECLAIM / EXCLUDE from several independent signals and shows which signals fired on each row. It knows a shared mailbox under 50 GB needs no licence, that HQ and HS are entity codes rather than countries, and it detects an anonymised mailbox usage report instead of counting it as zero.
    READ-ONLY. Produces the CSV that Invoke-LicenseReclamation.ps1 executes.
    Scopes: User.Read.All, Directory.Read.All, AuditLog.Read.All, Reports.Read.All
    signInActivity requires Entra ID P1+ and AuditLog.Read.All.

    CAVEATS (read before acting on output):
      - signInActivity has 24-48h latency and a finite retention window; a null
        means "no sign-in in the retained window", not always "never".
      - The mailbox usage report is useless if the tenant has "conceal user
        details in reports" enabled — the script detects and warns.
      - Litigation/in-place hold is only known with -EnrichFromExchange; without
        it, CONVERT_SHARED rows carry a "verify hold" flag.
#>

[CmdletBinding()]
param(
    [string]$SkuPartNumber = 'ENTERPRISEPACK',
    [int]   $InactiveDaysStrong = 365,
    [int]   $InactiveDaysSoft = 30,
    [int]   $NewJoinerGraceDays = 30,
    [ValidateSet('D7', 'D30', 'D90', 'D180')]
    [string]$MailboxReportPeriod = 'D180',
    [switch]$EnrichFromExchange,
    [int]   $SampleSize = 0,
    # Resource / functional account naming heuristic. Adjust to your own conventions.
    [string[]]$ResourcePattern = @(
        'recep', 'tm\.', 'tm[0-9]', 'rx\.', 'rx[0-9]', 'ultra', 'electro', 'espiro', 'colpos',
        'tomo', 'tomamuestra', 'tomademuestra', 'mensajeria', 'ksk', 'kiosco', 'coord',
        'emedico', 'mdistancia', 'patologia', 'gcia', 'atn\.', 'atencion', 'sala', 'room',
        'shared', 'svc[_\.]', '^sa_', 'printer', 'scan', 'almacen', 'citomica',
        'robot', 'noreply', 'no-reply', 'alertas', '^info@', 'soporte', '^support', 'laboratorio'
    ),
    [string]$OutputFolder = $PSScriptRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptStart = Get-Date
$stamp = $scriptStart.ToString('yyyyMMdd_HHmmss')
$reportPath = Join-Path $OutputFolder "LicenseReclamation_${SkuPartNumber}_$stamp.csv"
$logPath = Join-Path $OutputFolder "LicenseReclamation_${SkuPartNumber}_$stamp.log"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message
    Write-Host $line -ForegroundColor $(switch ($Level) { 'WARN' { 'Yellow' }'ERROR' { 'Red' }'OK' { 'Green' }default { 'Cyan' } })
    Add-Content -Path $logPath -Value $line
}

function Get-DaysSince {
    param($DateValue)
    if (-not $DateValue) { return $null }
    try { return [int]((Get-Date) - [datetime]$DateValue).TotalDays } catch { return $null }
}

$resourcePatterns = $ResourcePattern

function Test-ResourceName {
    param([string]$Upn, [string]$DisplayName)
    $local = ($Upn -split '@')[0].ToLowerInvariant()
    foreach ($p in $resourcePatterns) {
        if ($local -match $p) { return $true }
    }
    # Display names like "Recepción 1 ...", "Toma de Muestras ...", "Rayos X ..."
    if ($DisplayName -match '(?i)recep|toma de muestra|rayos x|ultrasonido|electrocard|coordinaci|gerencia|kiosco|patolog') { return $true }
    return $false
}

#region ── Connect ───────────────────────────────────────────────────────────────
Write-Log "Connecting to Microsoft Graph..."
Connect-MgGraph -Scopes 'User.Read.All', 'Directory.Read.All', 'AuditLog.Read.All', 'Reports.Read.All' -NoWelcome
Write-Log "Graph connected." 'OK'

if ($EnrichFromExchange) {
    # Reuse an existing EXO session if one is already open (connect however your CA allows
    # BEFORE running this script — certificate, interactive, etc. -Device is CA-blocked here).
    $exoLive = $false
    try { Get-EXOMailbox -ResultSize 1 -ErrorAction Stop | Out-Null; $exoLive = $true } catch { $exoLive = $false }
    if ($exoLive) {
        Write-Log "Existing Exchange Online session detected — reusing it." 'OK'
    } else {
        Write-Log "No EXO session found. Attempting interactive Connect-ExchangeOnline..." 'WARN'
        try { Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop; Write-Log "EXO connected." 'OK' }
        catch {
            Write-Log "EXO connect failed — continuing Graph-only. Connect EXO manually first, then re-run. $_" 'WARN'
            $EnrichFromExchange = $false
        }
    }
}
#endregion

#region ── Resolve SKU ───────────────────────────────────────────────────────────
$sku = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq $SkuPartNumber
if (-not $sku) { Write-Log "SKU '$SkuPartNumber' not found in tenant." 'ERROR'; return }
$skuId = $sku.SkuId
Write-Log "SKU $SkuPartNumber = $skuId | Consumed $($sku.ConsumedUnits) / Enabled $($sku.PrepaidUnits.Enabled)"
#endregion

#region ── Pull licensed users (bulk, paged) ────────────────────────────────────
Write-Log "Querying users with $SkuPartNumber assigned..."
$props = @(
    'id', 'userPrincipalName', 'displayName', 'accountEnabled', 'userType',
    'createdDateTime', 'onPremisesSyncEnabled', 'signInActivity',
    'lastPasswordChangeDateTime', 'assignedLicenses', 'licenseAssignmentStates', 'passwordPolicies', 'usageLocation'
)
$users = Get-MgUser -All `
    -Filter "assignedLicenses/any(x:x/skuId eq $skuId)" `
    -Property $props -ConsistencyLevel eventual -CountVariable licCount
Write-Log "Users holding $SkuPartNumber : $($users.Count)"

if ($SampleSize -gt 0 -and $users.Count -gt $SampleSize) {
    $users = $users | Select-Object -First $SampleSize
    Write-Log "SampleSize active — analyzing first $SampleSize." 'WARN'
}
#endregion

#region ── Mailbox usage report (one bulk call) ─────────────────────────────────
$mbx = @{}   # upn(lower) -> @{ LastActivity; StorageBytes; ItemCount; HasArchive }
try {
    $tmp = Join-Path $env:TEMP "mbxusage_$stamp.csv"
    Write-Log "Pulling mailbox usage report ($MailboxReportPeriod)..."
    Invoke-MgGraphRequest -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='$MailboxReportPeriod')" `
        -OutputFilePath $tmp
    $rows = Import-Csv $tmp
    $upnCol = ($rows | Get-Member -Name '*Principal Name*' -MemberType NoteProperty | Select-Object -First 1).Name
    if (-not $upnCol) {
        Write-Log "Mailbox report has no UPN column — user details likely concealed in reports. Skipping mailbox signals." 'WARN'
    } else {
        $sample = ($rows | Select-Object -First 1).$upnCol
        if ($sample -notmatch '@') {
            Write-Log "Mailbox report UPNs look anonymized ('$sample') — reports concealment is ON. Skipping mailbox signals." 'WARN'
        } else {
            foreach ($r in $rows) {
                $key = ($r.$upnCol).ToLowerInvariant()
                $mbx[$key] = @{
                    LastActivity = $r.'Last Activity Date'
                    StorageBytes = [int64]($r.'Storage Used (Byte)' -as [int64])
                    ItemCount    = [int]($r.'Item Count' -as [int])
                    HasArchive   = $r.'Has Archive'
                }
            }
            Write-Log "Mailbox usage rows: $($mbx.Count)" 'OK'
        }
    }
    Remove-Item $tmp -ErrorAction SilentlyContinue
} catch {
    Write-Log "Mailbox usage report failed — continuing without mailbox signals. $_" 'WARN'
}
#endregion

#region ── Optional EXO enrichment (bulk) ───────────────────────────────────────
$exo = @{}   # upn(lower) -> @{ Type; Hold }
if ($EnrichFromExchange) {
    try {
        Write-Log "Pulling EXO recipient + hold data (bulk)..."
        Get-EXOMailbox -ResultSize Unlimited -Properties RecipientTypeDetails, LitigationHoldEnabled |
        ForEach-Object {
            $exo[$_.UserPrincipalName.ToLowerInvariant()] = @{
                Type = $_.RecipientTypeDetails
                Hold = [bool]$_.LitigationHoldEnabled
            }
        }
        Write-Log "EXO mailboxes: $($exo.Count)" 'OK'
    } catch { Write-Log "EXO enrichment failed — continuing. $_" 'WARN' }
}
#endregion

#region ── Group-name cache for license source ──────────────────────────────────
$groupNameCache = @{}
function Resolve-GroupName {
    param([string]$GroupId)
    if (-not $GroupId) { return $null }
    if ($groupNameCache.ContainsKey($GroupId)) { return $groupNameCache[$GroupId] }
    try { $n = (Get-MgGroup -GroupId $GroupId -Property DisplayName).DisplayName } catch { $n = $GroupId }
    $groupNameCache[$GroupId] = $n
    return $n
}
#endregion

#region ── Score & tier ─────────────────────────────────────────────────────────
$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$n = 0
foreach ($u in $users) {
    $n++
    if ($n % 250 -eq 0) { Write-Progress -Activity "Scoring" -Status "$n / $($users.Count)" -PercentComplete ($n / $users.Count * 100) }

    $upn = $u.UserPrincipalName
    $upnL = $upn.ToLowerInvariant()
    $signals = [System.Collections.Generic.List[string]]::new()

    # ── sign-in activity (take the most recent of all three) ─────────────────
    $si = $u.SignInActivity
    $lastInteractive = if ($si) { $si.LastSignInDateTime } else { $null }
    $lastNonInter = if ($si) { $si.LastNonInteractiveSignInDateTime } else { $null }
    $lastSuccess = if ($si -and ($si.PSObject.Properties.Name -contains 'LastSuccessfulSignInDateTime')) { $si.LastSuccessfulSignInDateTime } else { $null }
    $signInDates = @($lastInteractive, $lastNonInter, $lastSuccess) | Where-Object { $_ } | ForEach-Object { [datetime]$_ }
    $lastSignIn = if ($signInDates) { ($signInDates | Measure-Object -Maximum).Maximum } else { $null }

    # ── mailbox activity ─────────────────────────────────────────────────────
    $mb = if ($mbx.ContainsKey($upnL)) { $mbx[$upnL] } else { $null }
    $mbLastActivity = if ($mb -and $mb.LastActivity) { [datetime]$mb.LastActivity } else { $null }
    $mbStorageGB = if ($mb) { [math]::Round($mb.StorageBytes / 1GB, 2) } else { $null }
    $mbHasArchive = if ($mb) { $mb.HasArchive } else { $null }

    # ── effective last activity = newest of any sign-in or mailbox activity ──
    $activityDates = @($lastSignIn, $mbLastActivity) | Where-Object { $_ }
    $effectiveLast = if ($activityDates) { ($activityDates | Measure-Object -Maximum).Maximum } else { $null }
    $inactiveDays = Get-DaysSince $effectiveLast

    # ── other signals ────────────────────────────────────────────────────────
    $accountEnabled = $u.AccountEnabled
    $ageDays = Get-DaysSince $u.CreatedDateTime
    $pwdAgeDays = Get-DaysSince $u.LastPasswordChangeDateTime
    $isGuest = $u.UserType -eq 'Guest'
    $isResource = Test-ResourceName -Upn $upn -DisplayName $u.DisplayName
    $exoInfo = if ($exo.ContainsKey($upnL)) { $exo[$upnL] } else { $null }
    $recipType = if ($exoInfo) { $exoInfo.Type } else { $null }
    $onHold = if ($exoInfo) { $exoInfo.Hold } else { $null }
    $neverSignedIn = (-not $lastSignIn)
    $neverActive = (-not $effectiveLast)

    # ── license source: direct vs group ─────────────────────────────────────
    $assignState = @($u.LicenseAssignmentStates | Where-Object { $_.SkuId -eq $skuId })
    $viaGroups = @($assignState | Where-Object { $_.AssignedByGroup } | ForEach-Object { $_.AssignedByGroup }) | Select-Object -Unique
    $isDirect = [bool]($assignState | Where-Object { -not $_.AssignedByGroup })
    $sourceDesc = if ($viaGroups) { 'Group: ' + (($viaGroups | ForEach-Object { Resolve-GroupName $_ }) -join '; ') }
    elseif ($isDirect) { 'Direct' } else { 'Unknown' }

    # ── UsageLocation vs the country the licensing group implies ─────────────
    $usageLoc = $u.UsageLocation
    $grpCountry = $null
    if ($viaGroups) {
        $firstName = Resolve-GroupName ($viaGroups | Select-Object -First 1)
        if ($firstName -match '^([A-Z]{2})-\d{2}-') { $grpCountry = $matches[1] }
    }
    # HQ/HS are entity codes, not ISO country codes: HQ is the head-office entity and
    # HS a shared-services entity whose staff span several countries, so neither can be
    # compared against usageLocation — they are flagged for a human instead.
    # Some groups also use naming codes that differ from the ISO code (UK group = GB country).
    $groupCodeAlias = @{ 'UK' = 'GB' }
    $grpCountryIso = if ($grpCountry -and $groupCodeAlias.ContainsKey($grpCountry)) { $groupCodeAlias[$grpCountry] } else { $grpCountry }
    $countryMatch =
    if (-not $grpCountry) { 'Direct/Unknown' }
    elseif ($grpCountry -in @('HQ', 'HS')) { 'Entity-group (HQ/HS) — verify' }
    elseif (-not $usageLoc) { 'No UsageLocation set' }
    elseif ($usageLoc -eq $grpCountryIso) { 'Match' }
    else { "MISMATCH (user=$usageLoc / group=$grpCountry)" }

    # ── build signal list ────────────────────────────────────────────────────
    if (-not $accountEnabled) { $signals.Add('Disabled') }
    if ($isGuest) { $signals.Add('Guest') }
    if ($neverSignedIn) { $signals.Add('NeverSignedIn') }
    if ($neverActive) { $signals.Add('NoActivityRecord') }
    if ($inactiveDays -ne $null -and $inactiveDays -ge $InactiveDaysStrong) { $signals.Add("Dormant>${InactiveDaysStrong}d($inactiveDays)") }
    elseif ($inactiveDays -ne $null -and $inactiveDays -ge $InactiveDaysSoft) { $signals.Add("Stale($inactiveDays d)") }
    if ($isResource) { $signals.Add('ResourceNaming') }
    if ($recipType -and $recipType -ne 'UserMailbox') { $signals.Add("Mbx:$recipType") }
    if ($onHold -eq $true) { $signals.Add('LitigationHold') }
    if ($pwdAgeDays -ne $null -and $pwdAgeDays -ge 180) { $signals.Add("PwdAge:${pwdAgeDays}d") }
    if ($mbStorageGB -ne $null) { $signals.Add("Mbx:${mbStorageGB}GB") }

    # ── tier decision (top-down, first match wins) ──────────────────────────
    $tier = ''; $action = ''; $reason = ''

    $isNewJoiner = ($ageDays -ne $null -and $ageDays -le $NewJoinerGraceDays -and $neverSignedIn)
    $hasMailbox = ($mb -ne $null)
    $noMailbox = (-not $hasMailbox) -or ($mbStorageGB -ne $null -and $mbStorageGB -eq 0 -and $neverActive)
    $overSharedCap = ($mbStorageGB -ne $null -and $mbStorageGB -ge 50)       # shared mbx >50GB still needs a license
    $canConvert = ($mbStorageGB -ne $null -and $mbStorageGB -gt 1 -and $mbStorageGB -lt 50 -and ($mbHasArchive -ne 'True') -and ($onHold -ne $true))

    if ($onHold -eq $true) {
        $tier = 'EXCLUDE'; $action = 'Keep (on hold)'
        $reason = 'Litigation/in-place hold — cannot reclaim or convert without a retention plan.'
    } elseif ($isNewJoiner) {
        $tier = 'EXCLUDE'; $action = 'Keep (new joiner)'
        $reason = "Created ${ageDays}d ago, no sign-in yet — likely onboarding."
    } elseif ($isGuest) {
        $tier = 'RECLAIM'; $action = 'Remove license (guest)'
        $reason = 'Guest/B2B user holding E3.'
    }
    # ── Shared/Room/Equipment mailbox: needs NO license under 50GB, regardless ─
    #    of activity. Type overrides dormancy — an active shared mbx is still waste.
    elseif ($recipType -in @('SharedMailbox', 'RoomMailbox', 'EquipmentMailbox', 'SchedulingMailbox')) {
        if ($overSharedCap) {
            $tier = 'REVIEW'; $action = "Retention decision — $recipType >50GB"
            $reason = "$recipType at ${mbStorageGB}GB exceeds the 50GB free limit — a license is required above 50GB; review whether to archive down or keep licensed."
        } else {
            $tier = 'RECLAIM'; $action = "Remove license ($recipType needs none)"
            $reason = "$recipType under 50GB requires no license — mailbox keeps working after removal."
        }
    } elseif (-not $accountEnabled) {
        if ($overSharedCap) {
            $tier = 'REVIEW'; $action = 'Retention decision — disabled, >50GB mailbox'
            $reason = "Disabled but mailbox ${mbStorageGB}GB exceeds the 50GB shared-mailbox limit — needs an archive/retention decision before the license is pulled."
        } elseif ($canConvert) {
            $tier = 'CONVERT_SHARED'; $action = 'Convert to shared mailbox, then remove license'
            $reason = "Disabled leaver, mailbox ${mbStorageGB}GB worth retaining; under 50GB, no archive/hold — preserve as shared and free the seat."
        } elseif ($noMailbox) {
            $tier = 'RECLAIM'; $action = 'Remove license (disabled, no mailbox)'
            $reason = 'Disabled account with no mailbox — pure license waste.'
        } else {
            $tier = 'RECLAIM'; $action = 'Remove license (disabled, empty mailbox)'
            $reason = "Disabled account, mailbox ${mbStorageGB}GB (minimal/empty) — safe to reclaim."
        }
    } elseif ($noMailbox -and $ageDays -ne $null -and $ageDays -ge $InactiveDaysSoft) {
        $tier = 'RECLAIM'; $action = 'Remove license (no mailbox)'
        $reason = "Enabled but no mailbox provisioned, created ${ageDays}d ago — licensed seat doing nothing."
    } elseif ($neverActive -and $ageDays -ne $null -and $ageDays -ge $InactiveDaysStrong) {
        $tier = 'RECLAIM'; $action = 'Remove license (never used)'
        $reason = "No sign-in or mailbox activity on record, created ${ageDays}d ago."
    } elseif ($inactiveDays -ne $null -and $inactiveDays -ge $InactiveDaysStrong) {
        if ($overSharedCap) {
            $tier = 'REVIEW'; $action = 'Retention decision — dormant, >50GB mailbox'
            $reason = "Dormant ${inactiveDays}d, mailbox ${mbStorageGB}GB over the 50GB shared cap — retention decision needed."
        } elseif ($canConvert) {
            $tier = 'CONVERT_SHARED'; $action = 'Convert to shared mailbox, then remove license'
            $reason = "Dormant ${inactiveDays}d but mailbox ${mbStorageGB}GB has data worth retaining; under 50GB, no archive/hold."
        } elseif ($isResource) {
            $tier = 'CONVERT_SHARED'; $action = 'Review for shared conversion (resource account)'
            $reason = "Dormant ${inactiveDays}d resource/functional account."
        } else {
            $tier = 'RECLAIM'; $action = 'Remove license (dormant)'
            $reason = "Dormant ${inactiveDays}d; mailbox empty/minimal — safe to reclaim."
        }
    } elseif ($isResource -and ($inactiveDays -eq $null -or $inactiveDays -ge $InactiveDaysSoft)) {
        $tier = 'CONVERT_SHARED'; $action = 'Review for shared conversion (resource account)'
        $reason = 'Resource/functional naming with low activity — candidate for shared mailbox.'
    } elseif ($inactiveDays -ne $null -and $inactiveDays -ge $InactiveDaysSoft) {
        $tier = 'REVIEW'; $action = 'Human review'
        $reason = "Stale ${inactiveDays}d — between soft and strong thresholds."
    } elseif ($neverActive) {
        $tier = 'REVIEW'; $action = 'Human review'
        $reason = 'No activity record but account too new to auto-reclaim; check sign-in retention.'
    } else {
        $tier = 'KEEP'; $action = 'Keep (active)'
        $reason = "Active within ${InactiveDaysSoft}d."
    }

    # hold flag when EXO not run
    if (-not $EnrichFromExchange -and $tier -eq 'CONVERT_SHARED') {
        $reason += ' [verify litigation hold before converting — EXO enrichment not run]'
    }

    $results.Add([PSCustomObject]@{
            UPN                 = $upn
            DisplayName         = $u.DisplayName
            Tier                = $tier
            RecommendedAction   = $action
            Reason              = $reason
            Signals             = ($signals -join '; ')
            AccountEnabled      = $accountEnabled
            UserType            = $u.UserType
            InactiveDays        = $inactiveDays
            LastSignIn          = $lastSignIn
            LastInteractive     = $lastInteractive
            LastNonInteractive  = $lastNonInter
            MailboxLastActivity = $mbLastActivity
            MailboxGB           = $mbStorageGB
            HasArchive          = $mbHasArchive
            RecipientType       = $recipType
            OnHold              = $onHold
            PasswordAgeDays     = $pwdAgeDays
            CreatedDaysAgo      = $ageDays
            OnPremSynced        = $u.OnPremisesSyncEnabled
            LicenseSource       = $sourceDesc
            UsageLocation       = $usageLoc
            LicenseGroupCountry = $grpCountry
            CountryMatch        = $countryMatch
            ObjectId            = $u.Id
        })
}
Write-Progress -Activity "Scoring" -Completed
#endregion

#region ── Export + summary ─────────────────────────────────────────────────────
$results | Sort-Object Tier, InactiveDays -Descending | Export-Csv -Path $reportPath -NoTypeInformation -Encoding unicode
Write-Log "Report saved: $reportPath" 'OK'

Write-Host ""
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  LICENSE RECLAMATION — $SkuPartNumber" -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Cyan

$tierOrder = 'RECLAIM', 'CONVERT_SHARED', 'REVIEW', 'EXCLUDE', 'KEEP'
$byTier = $results | Group-Object Tier
foreach ($t in $tierOrder) {
    $g = $byTier | Where-Object Name -eq $t
    $c = if ($g) { $g.Count } else { 0 }
    $color = switch ($t) { 'RECLAIM' { 'Green' }'CONVERT_SHARED' { 'Cyan' }'REVIEW' { 'Yellow' }'EXCLUDE' { 'DarkGray' }'KEEP' { 'Gray' } }
    Write-Host ("  {0,-15} {1,6}" -f $t, $c) -ForegroundColor $color
}

$reclaimable = @($results | Where-Object Tier -in 'RECLAIM', 'CONVERT_SHARED').Count
Write-Host ""
Write-Host ("  Immediately reclaimable (RECLAIM): {0}" -f @($results | Where-Object Tier -eq 'RECLAIM').Count) -ForegroundColor Green
Write-Host ("  Convertible to shared (frees seat): {0}" -f @($results | Where-Object Tier -eq 'CONVERT_SHARED').Count) -ForegroundColor Cyan
Write-Host ("  Total potential seats recovered:    {0}" -f $reclaimable) -ForegroundColor Green
Write-Host ""
Write-Host ("  Analyzed : {0}" -f $results.Count)
Write-Host ("  Elapsed  : {0}" -f ((Get-Date) - $scriptStart).ToString('hh\:mm\:ss'))
Write-Host ("  Report   : {0}" -f $reportPath)
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Cyan
#endregion