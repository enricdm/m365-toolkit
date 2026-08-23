# Purview

One remediation script for Microsoft Purview Information Protection: removing a sensitivity label,
and the encryption it applied, from files that should never have been labelled.

## Index

| Script | What it does | Changes state | Auth |
|---|---|:---:|---|
| [`Remove-SensitivityLabel.ps1`](Remove-SensitivityLabel.ps1) | Five-phase removal of a sensitivity label and its encryption from Office files | **Yes** — phases 2, 4 and 5 | interactive admin |

## Requirements

- **Windows PowerShell 5.1**, run as Administrator, on **Windows 10/11 or Windows Server (64-bit)**.
  The `PurviewInformationProtection` module is Windows-only and does not work in PowerShell 7.
- **Modules**
  - `AIPService` — `Install-Module AIPService`
  - `PurviewInformationProtection` — comes with the
    [Microsoft Purview Information Protection client](https://aka.ms/MIPClient)
- **Roles:** Global Administrator *or* Information Protection Administrator to manage the super user
  feature; SharePoint Administrator as well if you point `-ScanPaths` at SharePoint Online URLs.
  The account you run as must be the same one added as super user — super user rights are evaluated
  on the user token when the file is decrypted.
- **Service:** Azure Rights Management must be activated on the tenant (`Get-AipService` reports `Enabled`).

Handles files (Office documents, PDFs) on local paths, UNC paths, or SharePoint Online document
URLs. It does **not** remediate labels on Exchange Online email — that needs an eDiscovery export
and `Set-FileLabel` against the resulting PST.

## Usage

### `Remove-SensitivityLabel.ps1`

> ## ⚠ This is the most dangerous script in this repository. Read this section fully.
>
> **Phase 2 grants tenant-wide decryption.** The Azure RMS super user feature lets the named account
> open and decrypt **every protected file and email in the entire tenant** — not just the files you
> are fixing. It cannot be scoped to a subset. It is a break-glass capability, and while it is on,
> one account can read everything the organisation has ever encrypted.
>
> **Phase 4 is irreversible and takes no backup.** `Remove-FileLabel` strips the label *and* the
> encryption from each file in place. There is no undo, and this script does **not** copy the files
> anywhere first. If the label was legitimate on any file in your inventory, that protection is
> simply gone. **Back the target files up yourself before running Phase 4.**
>
> **Nothing forces Phase 5.** If you stop after Phase 4 — or the run errors, or you close the window —
> the super user feature stays **enabled** and the account stays a super user **indefinitely**. There
> is no timer, no automatic revert, and no warning. Running Phase 5 is entirely your responsibility.
> Verify afterwards with `Get-AipServiceSuperUserFeature` and `Get-AipServiceSuperUser`.
>
> Treat the whole operation as time-boxed and logged: announce it, do it, close it out the same day.
> Super user activity is auditable via `Get-AipServiceAdminLog` and the Purview audit log — and it
> will be audited, so keep the window short and narrow.

A sensitivity label with encryption does not just mark a file — it encrypts it with Azure RMS. So
when a label lands on documents that have to stay readable, or a policy changes and the label has to
come off a large set of files, removing the label is not enough on its own: the content has to be
decrypted. And decrypting content that is not yours requires the Azure RMS **super user** feature,
which gives one identity the ability to decrypt any protected content in the tenant. There is no
smaller version of that privilege.

The same thing happens from the other direction: when a labelling policy is deleted, files already
labelled keep their label and stay encrypted, and if that label's encryption excludes
administrators, nobody can open those files to clean them up — including the person who has to fix
it. Super user is the only supported way back in.

That is why this is five phases and not one command. Phase 2 grants the privilege, phase 3
inventories what actually carries the label, **phase 4 decrypts irreversibly and with no backup**,
and phase 5 hands the privilege back. The split exists to force a human to read the inventory
between 3 and 4. The gap worth knowing about is that nothing in the script forces phase 5: if nobody
runs it, that identity keeps tenant-wide decryption indefinitely.

| Phase | Does | Changes state |
|:---:|---|:---:|
| 1 | Preflight: modules, connectivity, current super user state | No |
| 2 | Enables the super user feature and adds the executing admin | **Yes** |
| 3 | Inventory: builds a CSV of files actually carrying the label | No |
| 4 | Remediation: removes the label and protection from inventoried files | **Yes, irreversibly** |
| 5 | Cleanup: removes the super user assignment and disables the feature | **Yes** |

```powershell
# Phase 1 — check the environment is ready. Always start here.
.\Remove-SensitivityLabel.ps1 -Phase 1

# Phase 4 — dry run. Never skip this.
.\Remove-SensitivityLabel.ps1 -Phase 4 -InventoryCsv .\Exports\labeled-files.csv `
    -LabelId '<label-guid>' -WhatIf
```

Full sequence:

```powershell
.\Remove-SensitivityLabel.ps1 -Phase 2 -SuperUserUpn 'admin@contoso.com'
.\Remove-SensitivityLabel.ps1 -Phase 3 -LabelId '<label-guid>' `
    -ScanPaths 'C:\Temp\Affected','https://contoso.sharepoint.com/sites/Example/Shared Documents'
# review the CSV, back up the files, then:
.\Remove-SensitivityLabel.ps1 -Phase 4 -LabelId '<label-guid>' -InventoryCsv .\Exports\labeled-files.csv -WhatIf
.\Remove-SensitivityLabel.ps1 -Phase 4 -LabelId '<label-guid>' -InventoryCsv .\Exports\labeled-files.csv
.\Remove-SensitivityLabel.ps1 -Phase 5 -SuperUserUpn 'admin@contoso.com'   # do not skip
```

Get the label GUID from the Purview portal, or:

```powershell
Connect-IPPSSession
Get-Label | Where-Object DisplayName -eq '<label name>' | Select-Object DisplayName, Guid
```

**Input:** `-LabelId` (the label GUID), `-ScanPaths` for Phase 3, and the Phase 3 CSV for Phase 4.
**Output:** a per-phase transcript log, the Phase 3 inventory CSV, and a Phase 4 results CSV — all in `-LogFolder`.
**Permissions:** Global Admin or Information Protection Admin, plus SharePoint Admin for SPO paths.

Safeguards that do exist, so you can weigh them against the risks above: Phase 4 only ever touches
files listed in the CSV you reviewed in Phase 3, so the blast radius is something you chose rather
than something the script discovered; it supports `-WhatIf`; `-PreserveFileDetails` keeps
`Modified`/`Modified By` intact; every phase is transcript-logged; and per-file failures are caught
and recorded rather than aborting the batch. Use `-JustificationMessage` to record your change
reference in the audit log.

## Known rough edges

- **Nothing forces Phase 5, and there is no timer.** If the run errors, or the window is closed, or
  somebody simply stops after Phase 4, the super user feature stays enabled and the account keeps
  tenant-wide decryption indefinitely. This is the single most consequential gap in the repository.
  A wrapper that guaranteed cleanup in a `finally` block would be the right fix; it does not exist
  yet, so the control is a person remembering.
- **Phase 4 takes no backup and cannot be undone.** `Remove-FileLabel` strips the label and the
  encryption in place. Back the target files up yourself first — the script will not do it and will
  not ask whether you did.
- **The super user privilege cannot be scoped.** It is all protected content in the tenant or none.
  That is a property of Azure RMS, not of this script, but it means the blast radius of Phase 2 is
  never proportionate to the job you are doing.
- **Windows PowerShell 5.1 only, on Windows, as Administrator.** The `PurviewInformationProtection`
  module does not work in PowerShell 7, which makes this the one script here that cannot run
  alongside the rest of the toolkit in the same session.
- **Email is out of scope.** Labels on Exchange Online messages need an eDiscovery export and
  `Set-FileLabel` against the resulting PST. This script only handles files.
- **There is no revocation script for the grant beyond Phase 5 itself**, and no check that Phase 5
  actually ran. Verify by hand with `Get-AipServiceSuperUserFeature` and `Get-AipServiceSuperUser`.
- **Rehearse the whole sequence on a throwaway file set** before pointing it at anything real.
  Given what Phase 4 does, a dry run of Phase 4 is the minimum, not a courtesy.
