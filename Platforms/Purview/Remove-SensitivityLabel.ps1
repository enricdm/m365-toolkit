<#
.SYNOPSIS
    Removes an inadvertently applied sensitivity label (and its encryption) from
    Office files, using the Azure Rights Management super user feature plus the
    PurviewInformationProtection client cmdlets.

.DESCRIPTION
    Remediation for the case where a sensitivity label keeps applying to files after
    the labelling policy itself has been deleted: the files stay encrypted, and
    without super user rights not even an administrator can open them to fix it.

    *** READ THE DANGER NOTES BELOW BEFORE RUNNING ANY PHASE ***

    It is structured in distinct phases so it can be run step-by-step,
    reviewed between phases, and stopped at any point. Nothing destructive
    happens until Phase 4, and Phase 4 supports -WhatIf for a dry run.

    PHASES:
        Phase 1 - Preflight checks (modules, connectivity, current super user state)
        Phase 2 - Enable super user feature + add the executing admin as super user
        Phase 3 - Inventory: build a CSV of files actually carrying the Internal label
        Phase 4 - Remediation: remove the label (and protection) from inventoried files
        Phase 5 - Cleanup: remove super user assignment + disable feature

    DANGER - WHAT PHASES 2 AND 4 ACTUALLY DO
    ----------------------------------------
    PHASE 2 GRANTS TENANT-WIDE DECRYPTION. The Azure RMS super user feature lets the
    named account open and decrypt EVERY protected file and email in the tenant,
    regardless of who owns it or who the label was meant to exclude. It is not scoped
    to the files you are trying to fix. Treat enabling it as a temporary,
    time-boxed, logged break-glass action, not a routine step.

    PHASE 4 IS IRREVERSIBLE AND TAKES NO BACKUP. Remove-FileLabel strips the label and
    the encryption from each file in place. There is no undo, and this script does not
    copy the files anywhere first. If the label was legitimate on some of the files in
    your inventory, that protection is simply gone. Back the target files up yourself
    before running Phase 4, and always run it once with -WhatIf.

    NOTHING FORCES PHASE 5. If you stop after Phase 4 — or the run fails, or you close
    the window — the super user feature stays ENABLED and the account stays a super
    user indefinitely. No timer, no automatic revert, no warning. Phase 5 is what
    removes it, and running it is your responsibility. Verify afterwards with
    Get-AipServiceSuperUserFeature and Get-AipServiceSuperUser.

.NOTES
    When to use  : A sensitivity label keeps applying after its policy was deleted, files stay encrypted, and not even an administrator can open them to fix it.
    Why it exists: Five separately runnable phases with the warnings inside the script rather than in a README: phase 2 grants tenant-wide decryption, phase 4 is irreversible and takes no backup, and nothing forces phase 5 - if the run stops there, the super user assignment stays in place indefinitely with no timer and no alert.
    PREREQUISITES (READ BEFORE RUNNING)
    -----------------------------------
    1) Roles required on the executing admin account:
       - Microsoft Entra: Global Administrator OR Information Protection Administrator
         (needed to manage AipService / super user feature)
       - SharePoint Administrator (needed if you point Set-FileLabel at SPO URLs)
       - The same account must also be the one added as super user, because
         super user rights are evaluated at the user-token level when
         Set-FileLabel/Remove-FileLabel decrypt the file.

    2) Software required on the workstation that runs this script:
       - Windows 10/11 or Windows Server (64-bit). The PurviewInformationProtection
         module is Windows-only.
       - Microsoft Purview Information Protection client installed
         (this installs the PurviewInformationProtection PowerShell module).
         Download: https://www.microsoft.com/download/details.aspx?id=53018
       - PowerShell 5.1 (Windows PowerShell), running "as Administrator".
       - AIPService PowerShell module (PSGallery): Install-Module AIPService

    3) Service prerequisite:
       - Azure Rights Management must be activated on the tenant.
         Verify with: Get-AipService  (status must be "Enabled")
       - The super user feature is DISABLED by default and must be enabled
         only for the duration of this remediation, then disabled again.

    4) Scope:
       - This script targets FILES (Office docs, PDFs, etc.) on local paths,
         UNC paths, or SharePoint Online document URLs.
       - It does NOT remove the label from Exchange Online emails. Email
         remediation requires a different approach (eDiscovery export +
         Set-FileLabel on the resulting PST, per Microsoft docs).

    SAFETY MODEL
    ------------
    - Phase 4 runs against a CSV inventory you build in Phase 3, so you
      review the exact list of files BEFORE anything is modified.
    - Phase 4 supports -WhatIf to preview without changing anything.
    - All actions are logged to a transcript file.
    - Super user activity is auditable via Get-AipServiceAdminLog and the
      Purview audit log.

.PARAMETER Phase
    Which phase to run. Always start with 1. Re-running a phase is safe.

.PARAMETER SuperUserUpn
    UPN of the admin account that will be temporarily granted super user
    rights. Must be the same account currently signed in to PowerShell.

.PARAMETER LabelId
    The GUID of the sensitivity label to remove. Get it from the Purview portal or
    via Security & Compliance PowerShell:
        Connect-IPPSSession
        Get-Label | Where-Object DisplayName -eq '<label name>' | Select-Object DisplayName, Guid

.PARAMETER JustificationMessage
    Reason recorded in the audit log for each label removal. Use your change or
    incident reference.

.PARAMETER ScanPaths
    One or more paths to scan in Phase 3. Examples:
        - Local:  "C:\Temp\AffectedFiles"
        - UNC:    "\\fileserver\share\Colombia"
        - SPO:    "https://contoso.sharepoint.com/sites/Example/Shared Documents"

.PARAMETER InventoryCsv
    Path to the CSV produced by Phase 3 and consumed by Phase 4.

.PARAMETER LogFolder
    Folder where transcript and CSV outputs are written.

.EXAMPLE
    # Phase 1 - just check that the environment is ready
    .\Remove-SensitivityLabel.ps1 -Phase 1

.EXAMPLE
    # Phase 4 - dry run only (no changes)
    .\Remove-SensitivityLabel.ps1 -Phase 4 `
        -InventoryCsv '<script folder>\Exports\labeled-files.csv' `
        -LabelId '<label-guid>' -WhatIf

#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateRange(1,5)]
    [int]$Phase,

    [string]$SuperUserUpn,

    [string]$LabelId,

    [string]$JustificationMessage = 'Remove inadvertently applied sensitivity label',

    [string[]]$ScanPaths,

    [string]$InventoryCsv = (Join-Path $PSScriptRoot 'Exports\labeled-files.csv'),

    [string]$LogFolder = (Join-Path $PSScriptRoot 'Exports')
)

#region Helpers --------------------------------------------------------------

function Initialize-LogFolder {
    if (-not (Test-Path $LogFolder)) {
        New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
    }
    $stamp  = Get-Date -Format 'yyyyMMdd-HHmmss'
    $script:TranscriptPath = Join-Path $LogFolder "Phase$Phase-$stamp.log"
    Start-Transcript -Path $script:TranscriptPath -Append | Out-Null
    Write-Host "Transcript: $script:TranscriptPath" -ForegroundColor Cyan
}

function Stop-LogFolder {
    try { Stop-Transcript | Out-Null } catch { }
}

function Assert-Module {
    param([string]$Name, [string]$InstallHint)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw "Required module '$Name' is not installed. $InstallHint"
    }
}

#endregion -------------------------------------------------------------------

Initialize-LogFolder

try {

    switch ($Phase) {

        # =====================================================================
        # PHASE 1 — Preflight
        # =====================================================================
        1 {
            Write-Host "`n=== PHASE 1: Preflight checks ===" -ForegroundColor Yellow

            Write-Host "`n[1/5] Checking PowerShell version..."
            if ($PSVersionTable.PSVersion.Major -lt 5) {
                throw "Requires Windows PowerShell 5.1 or later."
            }
            Write-Host "  OK: $($PSVersionTable.PSVersion)"

            Write-Host "`n[2/5] Checking required modules..."
            Assert-Module -Name 'AIPService' -InstallHint "Run: Install-Module AIPService -Scope CurrentUser"
            Assert-Module -Name 'PurviewInformationProtection' `
                -InstallHint "Install the Microsoft Purview Information Protection client (https://aka.ms/MIPClient)."
            Write-Host "  OK: AIPService + PurviewInformationProtection present."

            Write-Host "`n[3/5] Connecting to AIPService (Azure Rights Management)..."
            Import-Module AIPService -ErrorAction Stop
            Connect-AipService -ErrorAction Stop | Out-Null
            $svc = Get-AipService
            Write-Host "  AipService status: $($svc.FunctionalState)"
            if ($svc.FunctionalState -ne 'Enabled') {
                throw "Azure Rights Management is not enabled on this tenant. Aborting."
            }

            Write-Host "`n[4/5] Checking super user feature state..."
            $featureState   = Get-AipServiceSuperUserFeature
            $existingUsers  = Get-AipServiceSuperUser
            $existingGroup  = Get-AipServiceSuperUserGroup -ErrorAction SilentlyContinue
            Write-Host "  SuperUserFeature : $featureState"
            Write-Host "  SuperUsers       : $($existingUsers -join ', ')"
            Write-Host "  SuperUserGroup   : $existingGroup"

            Write-Host "`n[5/5] Reminder of what Phase 2 will do:"
            Write-Host "  - Enable-AipServiceSuperUserFeature   (idempotent)"
            Write-Host "  - Add-AipServiceSuperUser -EmailAddress <SuperUserUpn>"
            Write-Host "  Both actions are reversed by Phase 5."
            Write-Host "`nPhase 1 completed. Review the output above before running Phase 2." -ForegroundColor Green
        }

        # =====================================================================
        # PHASE 2 — Enable super user + assign the executing admin
        # =====================================================================
        2 {
            if (-not $SuperUserUpn) { throw "-SuperUserUpn is required for Phase 2." }
            Write-Host "`n=== PHASE 2: Enable super user feature ===" -ForegroundColor Yellow

            Import-Module AIPService -ErrorAction Stop
            Connect-AipService -ErrorAction Stop | Out-Null

            if ($PSCmdlet.ShouldProcess("AIP tenant", "Enable super user feature")) {
                Enable-AipServiceSuperUserFeature
            }
            if ($PSCmdlet.ShouldProcess($SuperUserUpn, "Add as Azure RMS super user")) {
                Add-AipServiceSuperUser -EmailAddress $SuperUserUpn
            }

            Write-Host "`nCurrent super users:" -ForegroundColor Cyan
            Get-AipServiceSuperUser

            Write-Host "`nPhase 2 completed. Test on ONE file before Phase 3 if possible." -ForegroundColor Green
            Write-Host "Test command:  Get-FileStatus -Path '<single test file>'"
        }

        # =====================================================================
        # PHASE 3 — Inventory files actually carrying the Internal label
        # =====================================================================
        3 {
            if (-not $ScanPaths) { throw "-ScanPaths is required for Phase 3." }
            if (-not $LabelId)   { throw "-LabelId is required for Phase 3." }

            Write-Host "`n=== PHASE 3: Inventory ===" -ForegroundColor Yellow
            Write-Host "Looking for files where MainLabelId or SubLabelId = $LabelId"

            Import-Module PurviewInformationProtection -ErrorAction Stop

            $results = New-Object System.Collections.Generic.List[object]
            foreach ($path in $ScanPaths) {
                Write-Host "`nScanning: $path"
                # Get-FileStatus accepts local paths, UNC paths, and SPO URLs.
                # It does NOT recurse on its own; -Path supports a folder and
                # the cmdlet enumerates files under it.
                Get-FileStatus -Path $path -ErrorAction Continue |
                    Where-Object {
                        $_.IsLabeled -and
                        ($_.MainLabelId -eq $LabelId -or
                         $_.SubLabelId  -eq $LabelId)
                    } |
                    ForEach-Object { $results.Add($_) }
            }

            if (-not (Test-Path (Split-Path $InventoryCsv))) {
                New-Item -ItemType Directory -Path (Split-Path $InventoryCsv) -Force | Out-Null
            }
            $results | Select-Object FileName, IsLabeled, MainLabelName, SubLabelName,
                                     IsRMSProtected, RMSOwner, LabelDate |
                Export-Csv -Path $InventoryCsv -NoTypeInformation -Encoding UTF8

            Write-Host "`nInventory written: $InventoryCsv"
            Write-Host "Total files matched: $($results.Count)" -ForegroundColor Green
            Write-Host "`nREVIEW the CSV before running Phase 4." -ForegroundColor Yellow
        }

        # =====================================================================
        # PHASE 4 — Remove label + protection from inventoried files
        # =====================================================================
        4 {
            if (-not $LabelId) { throw "-LabelId is required for Phase 4." }
            if (-not (Test-Path $InventoryCsv)) {
                throw "Inventory CSV not found: $InventoryCsv. Run Phase 3 first."
            }

            Write-Host "`n=== PHASE 4: Remediation ===" -ForegroundColor Yellow
            Import-Module PurviewInformationProtection -ErrorAction Stop

            $files = Import-Csv $InventoryCsv
            Write-Host "Files to process: $($files.Count)"

            $reportPath = Join-Path $LogFolder ("Phase4-results-{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
            $report     = New-Object System.Collections.Generic.List[object]

            foreach ($f in $files) {
                $target = $f.FileName
                if ($PSCmdlet.ShouldProcess($target, "Remove-FileLabel -RemoveLabel -RemoveProtection")) {
                    try {
                        $r = Remove-FileLabel -Path $target `
                                              -RemoveLabel `
                                              -RemoveProtection `
                                              -JustificationMessage $JustificationMessage `
                                              -PreserveFileDetails `
                                              -ErrorAction Stop
                        $report.Add([pscustomobject]@{
                            FileName = $target
                            Status   = $r.Status
                            Comment  = $r.Comment
                        })
                    } catch {
                        $report.Add([pscustomobject]@{
                            FileName = $target
                            Status   = 'Error'
                            Comment  = $_.Exception.Message
                        })
                    }
                }
            }

            $report | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8
            Write-Host "`nResults: $reportPath" -ForegroundColor Green
            $report | Group-Object Status | Format-Table Name, Count
        }

        # =====================================================================
        # PHASE 5 — Cleanup: remove super user, disable feature
        # =====================================================================
        5 {
            if (-not $SuperUserUpn) { throw "-SuperUserUpn is required for Phase 5." }
            Write-Host "`n=== PHASE 5: Cleanup ===" -ForegroundColor Yellow

            Import-Module AIPService -ErrorAction Stop
            Connect-AipService -ErrorAction Stop | Out-Null

            if ($PSCmdlet.ShouldProcess($SuperUserUpn, "Remove from Azure RMS super users")) {
                Remove-AipServiceSuperUser -EmailAddress $SuperUserUpn
            }
            if ($PSCmdlet.ShouldProcess("AIP tenant", "Disable super user feature")) {
                Disable-AipServiceSuperUserFeature
            }

            Write-Host "`nFinal state:" -ForegroundColor Cyan
            Write-Host "SuperUserFeature: $(Get-AipServiceSuperUserFeature)"
            Write-Host "SuperUsers     : $(Get-AipServiceSuperUser -join ', ')"
            Write-Host "`nPhase 5 completed. Confirm both values above are cleared before closing out." -ForegroundColor Green
        }
    }

} finally {
    Stop-LogFolder
}
