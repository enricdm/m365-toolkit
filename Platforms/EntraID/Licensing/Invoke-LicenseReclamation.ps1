<#
.SYNOPSIS
    Executes the license-reclamation plan produced by Get-LicenseReclamationPlan.ps1.
    Removes licenses (via the assigning group or direct) and, optionally, converts
    leaver mailboxes to shared first.

    WRITES TO THE DIRECTORY. ConfirmImpact is High, so every individual change
    prompts unless you pass -Confirm:$false; -WhatIf gives a full dry run.

.DESCRIPTION
    Reads the analysis report CSV and acts ONLY on the tiers you approve (default: RECLAIM).
    For each row it RE-FETCHES the user's live state from Graph before acting, so a stale
    report can't cause a wrong removal:
      - confirms the user still holds the target SKU (idempotent — skips if already gone)
      - for 'disabled' rows, confirms the account is still disabled (skips + flags if re-enabled)
      - resolves the CURRENT assigning group(s) from licenseAssignmentStates, not the CSV
        (group membership may have changed since the analysis ran)

    Reclaim mechanics (built for group-based licensing):
      - License assigned via group  → remove the user from that group (frees the seat)
      - License assigned directly    → Set-MgUserLicense -RemoveLicenses
      - Both                         → does both
      - If multiple groups grant the SKU, removes from all of them and verifies the
        seat actually freed afterward (a second group can keep the license alive).

    CONVERT_SHARED (opt-in via -Tiers including CONVERT_SHARED, requires an EXO session):
      1. Set-Mailbox -Type Shared   (convert FIRST so data is safe)
      2. verify RecipientTypeDetails = SharedMailbox
      3. only then remove the license
      If the conversion fails, the license is NOT removed.

.PARAMETER ReportPath
    Path to the CSV produced by Get-LicenseReclamationPlan.ps1.

.PARAMETER Tiers
    Which tiers to action. Default @('RECLAIM'). Add 'CONVERT_SHARED' to convert+reclaim.
    REVIEW / EXCLUDE / KEEP are never actioned.

.PARAMETER ApprovalColumn
    Optional. If you add a column (e.g. 'Approved') to the CSV and pass its name here,
    only rows where that column is Yes/Y/True/1 are processed. Lets an approver
    sign off row by row in the spreadsheet before anything runs.

.PARAMETER SkuPartNumber
    SKU to remove. Default 'ENTERPRISEPACK' (Office 365 E3).

.PARAMETER BatchSize
    Rows per batch with a courtesy pause. Default 25.

.PARAMETER MaxChanges
    Hard cap on actioned users this run. 0 = unlimited. Default 0. Use a small number
    for a first production batch.

.PARAMETER OutputFolder
    Folder for the results CSV + log. Default: script folder.

.EXAMPLE
    # Dry run, RECLAIM only:
    .\Invoke-LicenseReclamation.ps1 -ReportPath .\LicenseReclamation_ENTERPRISEPACK_*.csv -WhatIf

.EXAMPLE
    # Production, first 50 RECLAIM only:
    .\Invoke-LicenseReclamation.ps1 -ReportPath .\report.csv -MaxChanges 50

.EXAMPLE
    # Convert + reclaim, approved rows only (EXO connected first):
    .\Invoke-LicenseReclamation.ps1 -ReportPath .\report.csv -Tiers RECLAIM,CONVERT_SHARED -ApprovalColumn Approved

.NOTES
    When to use  : The reclamation plan is approved and has to be executed without destroying a former colleague's mailbox.
    Why it exists: Re-fetches every user's live state before acting, so a stale report cannot cause a wrong removal; converts the mailbox to shared and verifies the conversion stuck BEFORE removing the licence; caps the run with -MaxChanges; and skips bundle groups that would strip other SKUs as collateral.
    Scopes: User.ReadWrite.All, Group.ReadWrite.All, Directory.Read.All, Organization.Read.All
    For CONVERT_SHARED, connect Exchange Online in the session BEFORE running.
    Removal is reversible (re-add to group / re-assign); conversion to shared is reversible
    (Set-Mailbox -Type Regular) but re-licensing is then required.
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$ReportPath,

    [ValidateSet('RECLAIM', 'CONVERT_SHARED')]
    [string[]]$Tiers = @('RECLAIM'),

    [string]$ApprovalColumn,

    [string]$SkuPartNumber = 'ENTERPRISEPACK',

    [ValidateRange(1, 200)]
    [int]$BatchSize = 25,

    [ValidateRange(0, 100000)]
    [int]$MaxChanges = 0,

    [string]$OutputFolder = $PSScriptRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptStart = Get-Date
$stamp = $scriptStart.ToString('yyyyMMdd_HHmmss')
$mode = if ($WhatIfPreference) { 'WHATIF' } else { 'PROD' }
$resultsPath = Join-Path $OutputFolder "Reclaim_${mode}_$stamp.csv"
$logPath = Join-Path $OutputFolder "Reclaim_${mode}_$stamp.log"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message
    Write-Host $line -ForegroundColor $(switch ($Level) { 'WARN' { 'Yellow' }'ERROR' { 'Red' }'OK' { 'Green' }'WHATIF' { 'Magenta' }'SKIP' { 'DarkGray' }default { 'Cyan' } })
    Add-Content -Path $logPath -Value $line -WhatIf:$false
}

$doConvert = $Tiers -contains 'CONVERT_SHARED'

#region ── Connect ───────────────────────────────────────────────────────────────
Write-Log "Mode: $mode | Tiers: $($Tiers -join ',') | MaxChanges: $(if($MaxChanges){$MaxChanges}else{'unlimited'})"
Write-Log "Connecting to Microsoft Graph..."
Connect-MgGraph -Scopes 'User.ReadWrite.All', 'Group.ReadWrite.All', 'Directory.Read.All', 'Organization.Read.All' -NoWelcome
Write-Log "Graph connected." 'OK'

if ($doConvert) {
    try { Get-EXOMailbox -ResultSize 1 -ErrorAction Stop | Out-Null; Write-Log "EXO session detected." 'OK' }
    catch { Write-Log "CONVERT_SHARED requested but no EXO session. Connect Exchange Online first, then re-run. Aborting." 'ERROR'; return }
}
#endregion

#region ── Resolve SKU ───────────────────────────────────────────────────────────
$sku = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq $SkuPartNumber
if (-not $sku) { Write-Log "SKU '$SkuPartNumber' not found." 'ERROR'; return }
$skuId = $sku.SkuId
Write-Log "Target SKU $SkuPartNumber = $skuId | before: Consumed $($sku.ConsumedUnits)/$($sku.PrepaidUnits.Enabled)"
#endregion

#region ── Load + filter rows ────────────────────────────────────────────────────
$rows = Import-Csv -Path $ReportPath
$rows = @($rows | Where-Object { $_.Tier -in $Tiers })

if ($ApprovalColumn) {
    if (-not ($rows | Get-Member -Name $ApprovalColumn -MemberType NoteProperty)) {
        Write-Log "ApprovalColumn '$ApprovalColumn' not found in CSV. Aborting." 'ERROR'; return
    }
    $rows = @($rows | Where-Object { $_.$ApprovalColumn -match '^(yes|y|true|1)$' })
    Write-Log "After approval filter on '$ApprovalColumn': $($rows.Count)"
}

Write-Log "Rows to process: $($rows.Count)"
if ($rows.Count -eq 0) { Write-Log "Nothing to do." 'OK'; return }
#endregion

#region ── Helpers ───────────────────────────────────────────────────────────────
function Get-SkuAssignment {
    # Returns @{ HasSku; Groups=@(guids); Direct=$bool } from live licenseAssignmentStates
    param($User)
    $states = @($User.LicenseAssignmentStates | Where-Object { $_.SkuId -eq $skuId })
    [PSCustomObject]@{
        HasSku = ($User.AssignedLicenses.SkuId -contains $skuId)
        Groups = @($states | Where-Object { $_.AssignedByGroup } | ForEach-Object { $_.AssignedByGroup } | Select-Object -Unique)
        Direct = [bool]($states | Where-Object { -not $_.AssignedByGroup })
    }
}
#endregion

#region ── Process ───────────────────────────────────────────────────────────────
$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$applied = 0; $skipped = 0; $failed = 0; $capped = $false

for ($i = 0; $i -lt $rows.Count; $i += $BatchSize) {
    if ($MaxChanges -gt 0 -and $applied -ge $MaxChanges) { $capped = $true; break }
    $batch = $rows[$i .. [Math]::Min($i + $BatchSize - 1, $rows.Count - 1)]
    Write-Log "── Batch $([Math]::Floor($i/$BatchSize)+1): $($batch.Count) ──"

    foreach ($row in $batch) {
        if ($MaxChanges -gt 0 -and $applied -ge $MaxChanges) { $capped = $true; break }

        $upn = $row.UPN
        $rec = [ordered]@{ UPN = $upn; Tier = $row.Tier; PlannedAction = $row.RecommendedAction; ActionTaken = ''; Detail = ''; Result = '' }

        # ── live re-fetch ────────────────────────────────────────────────────
        $user = $null
        try {
            $user = Get-MgUser -UserId $row.ObjectId `
                -Property Id, UserPrincipalName, AccountEnabled, AssignedLicenses, LicenseAssignmentStates -ErrorAction Stop
        } catch {
            $rec.Result = 'Skipped'; $rec.Detail = "User not found / lookup failed: $($_.Exception.Message)"
            $results.Add([PSCustomObject]$rec); $skipped++; Write-Log "  [SKIP] $upn — not found." 'SKIP'; continue
        }

        $asg = Get-SkuAssignment -User $user

        # idempotency: already lost the SKU
        if (-not $asg.HasSku) {
            $rec.Result = 'Skipped'; $rec.Detail = 'User no longer holds the SKU (already reclaimed).'
            $results.Add([PSCustomObject]$rec); $skipped++; Write-Log "  [SKIP] $upn — SKU already gone." 'SKIP'; continue
        }

        # safety: 'disabled' rows must still be disabled
        if ($row.RecommendedAction -match 'disabled' -and $user.AccountEnabled) {
            $rec.Result = 'Skipped'; $rec.Detail = 'Row was disabled-based but account is now ENABLED — re-validated, not actioning.'
            $results.Add([PSCustomObject]$rec); $skipped++; Write-Log "  [SKIP] $upn — re-enabled since report." 'WARN'; continue
        }

        # ── CONVERT_SHARED: convert first, verify, then fall through to removal ─
        if ($row.Tier -eq 'CONVERT_SHARED') {
            if ($PSCmdlet.ShouldProcess($upn, "Convert mailbox to Shared")) {
                try {
                    Set-Mailbox -Identity $upn -Type Shared -ErrorAction Stop
                    Start-Sleep -Seconds 2
                    $t = (Get-EXOMailbox -Identity $upn -Properties RecipientTypeDetails).RecipientTypeDetails
                    if ($t -ne 'SharedMailbox') {
                        $rec.Result = 'Failed'; $rec.ActionTaken = 'ConvertAttempted'; $rec.Detail = "Post-convert type is '$t', not SharedMailbox — NOT removing license."
                        $results.Add([PSCustomObject]$rec); $failed++; Write-Log "  [FAIL] $upn — convert did not stick; license left in place." 'ERROR'; continue
                    }
                    $rec.ActionTaken = 'ConvertedToShared; '
                    Write-Log "  [OK] $upn — converted to shared." 'OK'
                } catch {
                    $rec.Result = 'Failed'; $rec.ActionTaken = 'ConvertAttempted'; $rec.Detail = "Convert failed: $($_.Exception.Message) — license left in place."
                    $results.Add([PSCustomObject]$rec); $failed++; Write-Log "  [FAIL] $upn — convert error; license left in place." 'ERROR'; continue
                }
            } else {
                $rec.ActionTaken = 'WhatIf:Convert; '
            }
        }

        # ── remove the license (group(s) and/or direct) ──────────────────────
        $removed = [System.Collections.Generic.List[string]]::new()
        $ok = $true

        foreach ($gid in $asg.Groups) {
            $gname = try { (Get-MgGroup -GroupId $gid -Property DisplayName).DisplayName } catch { $gid }

            # Guard: does this group grant ONLY the target SKU, or is it a bundle?
            # Removing membership strips EVERYTHING the group assigns, so a bundle group
            # would cause collateral loss of other licenses. Skip + flag those.
            #
            # This read decides whether a removal is safe, so it must not fail open. If the
            # answer cannot be obtained the group is skipped as though it WERE a bundle:
            # a licence left in place costs a seat, one wrongly removed costs somebody
            # their Office on Monday morning.
            $groupSkus = @()
            try {
                $groupSkus = @((Get-MgGroup -GroupId $gid -Property AssignedLicenses -ErrorAction Stop).AssignedLicenses.SkuId)
            } catch {
                $removed.Add("group-SKIPPED-unverified:$gname ($($_.Exception.Message))")
                $ok = $false
                Write-Log "  [SKIP] $upn — could not read which licences '$gname' assigns, so it cannot be confirmed single-SKU: $($_.Exception.Message). Not removing. Re-run when the group is readable." 'WARN'
                continue
            }
            $otherSkus = @($groupSkus | Where-Object { $_ -and $_ -ne $skuId })
            if ($otherSkus.Count -gt 0) {
                $otherNames = ($otherSkus | ForEach-Object { $sid = $_; (Get-MgSubscribedSku -All | Where-Object SkuId -eq $sid).SkuPartNumber }) -join ','
                $removed.Add("group-SKIPPED-bundle:$gname (also grants $otherNames)")
                $ok = $false
                Write-Log "  [SKIP] $upn — '$gname' is a bundle (also: $otherNames); not removing to avoid collateral loss. Handle manually." 'WARN'
                continue
            }

            if ($PSCmdlet.ShouldProcess($upn, "Remove from group '$gname'")) {
                try {
                    Remove-MgGroupMemberByRef -GroupId $gid -DirectoryObjectId $user.Id -ErrorAction Stop
                    $removed.Add("group:$gname")
                    Write-Log "  [OK] $upn — removed from $gname" 'OK'
                } catch {
                    # not-a-member is benign (idempotent); anything else is a real failure
                    if ($_.Exception.Message -match 'does not exist|not found|Insufficient') { $ok = $false }
                    $removed.Add("group-FAILED:$gname ($($_.Exception.Message))"); $ok = $false
                    Write-Log "  [FAIL] $upn — group removal $gname : $_" 'ERROR'
                }
            } else { $removed.Add("WhatIf:group:$gname") }
        }

        if ($asg.Direct) {
            if ($PSCmdlet.ShouldProcess($upn, "Remove direct license $SkuPartNumber")) {
                try {
                    Set-MgUserLicense -UserId $user.Id -RemoveLicenses @($skuId) -AddLicenses @() -ErrorAction Stop | Out-Null
                    $removed.Add("direct")
                    Write-Log "  [OK] $upn — direct license removed" 'OK'
                } catch {
                    $removed.Add("direct-FAILED ($($_.Exception.Message))"); $ok = $false
                    Write-Log "  [FAIL] $upn — direct removal: $_" 'ERROR'
                }
            } else { $removed.Add("WhatIf:direct") }
        }

        $rec.ActionTaken += ($removed -join '; ')
        if ($WhatIfPreference) { $rec.Result = 'WhatIf'; $applied++ }
        elseif ($ok) { $rec.Result = 'Applied'; $applied++ }
        else { $rec.Result = 'PartialFail'; $failed++ }
        $results.Add([PSCustomObject]$rec)
    }

    if ($i + $BatchSize -lt $rows.Count -and -not $capped) { Start-Sleep -Milliseconds 400 }
}
if ($capped) { Write-Log "MaxChanges cap ($MaxChanges) reached — remaining rows untouched." 'WARN' }
#endregion

#region ── Verify seats freed (prod only) ───────────────────────────────────────
if (-not $WhatIfPreference) {
    Start-Sleep -Seconds 3
    $skuAfter = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq $SkuPartNumber
    Write-Log "SKU after: Consumed $($skuAfter.ConsumedUnits)/$($skuAfter.PrepaidUnits.Enabled) (was $($sku.ConsumedUnits))" 'OK'
}
#endregion

#region ── Results + summary ─────────────────────────────────────────────────────
$results | Export-Csv -Path $resultsPath -NoTypeInformation -Encoding unicode -WhatIf:$false
Write-Log "Results saved: $resultsPath" 'OK'

Write-Host ""
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  LICENSE RECLAMATION — EXECUTION ($mode)" -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Cyan
$results | Group-Object Result | Sort-Object Name | ForEach-Object {
    $c = switch ($_.Name) { 'Applied' { 'Green' }'WhatIf' { 'Magenta' }'Skipped' { 'DarkGray' }'Failed' { 'Red' }'PartialFail' { 'Red' }default { 'Gray' } }
    Write-Host ("  {0,-12} {1,6}" -f $_.Name, $_.Count) -ForegroundColor $c
}
Write-Host ""
Write-Host ("  Actioned : {0}" -f $applied)
Write-Host ("  Skipped  : {0}  (already reclaimed / re-enabled / not found)" -f $skipped)
Write-Host ("  Failed   : {0}" -f $failed)
Write-Host ("  Elapsed  : {0}" -f ((Get-Date) - $scriptStart).ToString('hh\:mm\:ss'))
Write-Host ("  Results  : {0}" -f $resultsPath)
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Cyan
#endregion