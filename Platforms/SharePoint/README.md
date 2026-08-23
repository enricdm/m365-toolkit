# SharePoint Online

Administration scripts for SharePoint Online and OneDrive for Business: connection helpers,
permission auditing, storage governance, and lifecycle/recovery work. Most are read-only
reporting tools; the three that write are marked below and each has a dry-run mode.

## Index

| Script | What it does | Changes state | Auth |
|---|---|:---:|---|
| [`Connect-PnPSite.ps1`](Connect-PnPSite.ps1) | Opens an app-only PnP connection to a site, admin center, or OneDrive | No | app-only + cert |
| [`Clear-SpoTokenCache.ps1`](Clear-SpoTokenCache.ps1) | Purges local MSAL/WAM token caches, then reconnects — fixes "stuck identity" sign-ins | No (local caches only) | interactive |
| [`Permissions/Get-SiteOwnerStatus.ps1`](Permissions/Get-SiteOwnerStatus.ps1) | Lists each site's owners and whether their accounts are still enabled | No | app-only + cert |
| [`Permissions/Get-EveryoneExceptExternalGrant.ps1`](Permissions/Get-EveryoneExceptExternalGrant.ps1) | Finds "Everyone except external users" grants at site and list level | No | interactive |
| [`Storage/Get-SiteStorage.ps1`](Storage/Get-SiteStorage.ps1) | Reports storage usage and quota per site collection | No | interactive |
| [`Storage/Set-SiteStorageQuota.ps1`](Storage/Set-SiteStorageQuota.ps1) | Sets the storage quota on one or more sites | **Yes** | interactive |
| [`Storage/Send-StorageNotification.ps1`](Storage/Send-StorageNotification.ps1) | Emails site owners whose sites are over their storage threshold | **Yes** (sends mail) | app-only + cert |
| [`Lifecycle/Restore-OneDriveFolder.ps1`](Lifecycle/Restore-OneDriveFolder.ps1) | Restores a deleted folder subtree from a recycle bin | **Yes** | app-only + cert |
| [`Lifecycle/analyze_site_cleanup.py`](Lifecycle/analyze_site_cleanup.py) | Attributes inactive sites to a country/business unit and builds a cleanup workbook | No (reads exports) | n/a — offline |

## Requirements

- **PowerShell 7.2+** (tested on 7.6.x). `Get-SiteStorage.ps1`, `Set-SiteStorageQuota.ps1` and
  `Clear-SpoTokenCache.ps1` use the SPO module, which also works on Windows PowerShell 5.1.
- **Modules**
  - `PnP.PowerShell` — the connection, permission and lifecycle scripts
  - `Microsoft.Online.SharePoint.PowerShell` — the storage scripts
  - `ImportExcel` — `Send-StorageNotification.ps1` only
- **Python 3.9+** with `pandas` and `openpyxl` for `analyze_site_cleanup.py`
- **Roles:** SharePoint Administrator for tenant-wide operations; Site Collection Admin on each
  site for the direct permission scan and for restoring into someone's OneDrive.

### App registration (for the app-only scripts)

Create an Entra app registration with a certificate credential and grant admin consent for:

| API | Permission | Used by |
|---|---|---|
| SharePoint | `Sites.FullControl.All` (Application) | owner status, restore |
| Microsoft Graph | `Group.Read.All`, `User.Read.All` (Application) | owner status |
| Microsoft Graph | `Mail.Send` (Application) | storage notifications |

Upload the certificate's public key to the app and keep the private key in
`Cert:\CurrentUser\My`. The scripts resolve it by subject (`-CertSubject`) or accept an explicit
`-Thumbprint`.

> **Scope `Mail.Send`.** Granted tenant-wide it lets the app send as *anyone*. Restrict it to the
> one sender mailbox with an [application access policy](https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access):
> ```powershell
> New-ApplicationAccessPolicy -AppId '<client-id>' `
>     -PolicyScopeGroupId '<mail-enabled security group holding the sender mailbox>' `
>     -AccessRight RestrictAccess -Description 'Restrict reporting app to shared mailbox'
> ```

## Usage

### Read-only scripts

```powershell
# Open an app-only session (-Url is mandatory: no default target, by design)
.\Connect-PnPSite.ps1 -Url 'https://contoso.sharepoint.com/sites/Example' `
    -ClientId '<client-id>' -Tenant 'contoso.onmicrosoft.com'

# Fix a sign-in stuck on the wrong account, then reconnect
.\Clear-SpoTokenCache.ps1 -AdminUrl 'https://contoso-admin.sharepoint.com'

# Storage usage for every site, or just the ones matching a name
.\Storage\Get-SiteStorage.ps1 -AdminUrl 'https://contoso-admin.sharepoint.com'
.\Storage\Get-SiteStorage.ps1 -AdminUrl 'https://contoso-admin.sharepoint.com' -TitleFilter 'Finance'
```

**Output:** each writes a timestamped CSV to the current folder or `-OutputPath`.

### `Permissions/Get-SiteOwnerStatus.ps1`

Answers "who actually owns this site, and do they still work here?" — the question you hit the
moment you try to clean up storage or decommission sites. It resolves owners differently by site
type: group-connected sites via the Microsoft 365 group's owners in Graph, non-group sites via the
site's Owners group and Site Collection Admins read per-site through PnP (those are SharePoint
groups, invisible to Graph `/groups`).

```powershell
# Targeted — resolve owners only for the sites in a CSV (fast)
.\Permissions\Get-SiteOwnerStatus.ps1 -AdminUrl 'https://contoso-admin.sharepoint.com' `
    -ClientId '<client-id>' -Tenant 'contoso.onmicrosoft.com' -InputCsv .\sites.csv

# Whole tenant
.\Permissions\Get-SiteOwnerStatus.ps1 -AdminUrl 'https://contoso-admin.sharepoint.com' `
    -ClientId '<client-id>' -Tenant 'contoso.onmicrosoft.com'
```

**Input:** optionally a CSV with a `Url` / `SiteUrl` / `Site URL` column, or explicit `-SiteUrl` values.
**Output:** one row per site — owners, enabled status, storage, and a `HasEnabledOwner` flag.
**Permissions:** `Group.Read.All`, `User.Read.All`, `Sites.FullControl.All` (all Application).

Two behaviours worth knowing:

- **`accountEnabled` is tri-state:** `True`, `False`, or `Unknown`. A lookup that fails (throttling,
  missing Graph permission, deleted user) is reported as `Unknown`, never as `False`. Otherwise a
  permissions problem would silently manufacture a list of "orphaned" sites and someone would act on it.
- **Do not filter on `PercentUsed`.** It is derived from `StorageMaximumLevel`, which in a
  pooled-storage tenant is a large shared cap rather than a per-site quota. Get the threshold from a
  quota-based export and join on `SiteUrl`; this script supplies the owners.

### `Permissions/Get-EveryoneExceptExternalGrant.ps1`

Finds direct grants to **"Everyone except external users"** (EEEU) — the claim that quietly makes
content readable by every licensed user in the tenant. Scans site and list/library level; per-item
scanning is opt-in with `-ScanItems` because it is slow on large sites.

```powershell
# Whole tenant
.\Permissions\Get-EveryoneExceptExternalGrant.ps1 -ClientId '<client-id>' `
    -AdminCenterUrl 'https://contoso-admin.sharepoint.com'

# One group of sites, including item-level grants
.\Permissions\Get-EveryoneExceptExternalGrant.ps1 -ClientId '<client-id>' `
    -AdminCenterUrl 'https://contoso-admin.sharepoint.com' -UrlPattern '*/sites/Example*' -ScanItems
```

**Input:** `-UrlPattern` wildcard against site URL (default `*` = every site).
**Output:** CSV of findings — site, object, principal, login claim, permission levels.
**Permissions:** SharePoint Admin, plus Site Collection Admin on each scanned site.

> **Why this script matches on the login claim, not the display name.**
> EEEU's display name is localised: `Everyone except external users` in English,
> `Jeder, außer externen Benutzern` in German, and so on for every tenant locale. A scanner that
> compares display names only finds the grants whose language happens to be in its list — in any
> other tenant it walks the whole estate, matches nothing, and reports **"no EEEU grants found"**.
>
> That is the worst possible failure mode for a permissions tool: a silent false negative that
> looks exactly like a clean result. Nobody re-runs a scan that came back green.
>
> So detection is primarily **claim-based**, on the well-known identifier
> `c:0-.f|rolemanager|spo-grid-all-users`, which is byte-identical in every locale. The display-name
> list is kept only as a secondary signal. If you fork this for another built-in principal, match
> its claim, not its label.

### `Lifecycle/analyze_site_cleanup.py`

Takes a tenant "inactive sites" policy report — which tells you a site is idle but not who owns it —
and attributes each site to a country or business unit, so the list can actually be actioned.

Attribution runs in priority order: an **Entra user export** (owner email → `usageLocation`) first,
a **keyword match** on site name/URL only when no owner can be resolved, and otherwise `Unattributed`.
Every row records *which* method produced it, so a reviewer can tell a hard match from a guess.
Sites owned solely by an exempt entity are preserved and listed separately.

All business rules live outside the script in a JSON file — copy
[`Lifecycle/site-cleanup-rules.example.json`](Lifecycle/site-cleanup-rules.example.json) and edit it
for your tenant.

```bash
cp site-cleanup-rules.example.json site-cleanup-rules.json

python analyze_site_cleanup.py \
    --inactive-report ./inactive-sites.csv \
    --entra-export    ./entra-users.csv \
    --sites-report    ./sites.csv \
    --rules           ./site-cleanup-rules.json \
    --outdir          ./Exports
```

**Input:** CSV exports (inactive-site policy report, Entra user export with `usageLocation` /
`companyName` / `accountEnabled`, current site inventory) plus the rules file.
**Output:** an `.xlsx` with a dashboard, per-group tabs, `Unattributed` and `Exempt` tabs, plus a master CSV.
**Permissions:** none — it reads files you have already exported and never contacts SharePoint.

> **Deletion candidates are not a deletion list.** Everything this produces is a *proposal* for human
> review. Keyword-attributed rows in particular are informed guesses. Never feed the output
> straight into a deletion script.

## Scripts that change state

### `Storage/Set-SiteStorageQuota.ps1`

Sets the storage quota on one or more site collections. Run `Get-SiteStorage.ps1` first and confirm
the exact URLs from its output.

```powershell
# Always dry-run first
.\Storage\Set-SiteStorageQuota.ps1 -AdminUrl 'https://contoso-admin.sharepoint.com' `
    -SiteUrl 'https://contoso.sharepoint.com/sites/Example' -WhatIf

# Apply
.\Storage\Set-SiteStorageQuota.ps1 -AdminUrl 'https://contoso-admin.sharepoint.com' `
    -SiteUrl 'https://contoso.sharepoint.com/sites/Example' -QuotaMB 102400
```

**Input:** admin URL and one or more site URLs.
**Output:** a pre-change CSV of every target's current quota, written to `-ExportDir` before anything is modified.
**Permissions:** SharePoint Administrator.

> **Writes.** Safeguards: `SupportsShouldProcess` with `ConfirmImpact = 'High'`, so it prompts unless
> you pass `-Confirm:$false` and supports `-WhatIf`; and it records the previous quota for every site
> before changing it, giving you a rollback record. A site whose current quota cannot be read is
> logged as `unknown` — never as `0`, which would look like a valid value to roll back to.
> Raising quota consumes tenant-pooled storage, so get whatever approval your storage policy requires.

### `Storage/Send-StorageNotification.ps1`

Emails the owners of sites still over their storage threshold, driven by a tracker workbook.
Authenticates app-only by signing a JWT client assertion with the certificate and calling the Graph
REST endpoint directly — no `Microsoft.Graph` SDK, which sidesteps the MSAL assembly conflicts you
get when PnP or Az is loaded in the same session.

```powershell
# Dry run: still acquires a token (proving cert auth and Mail.Send work) but sends nothing
.\Storage\Send-StorageNotification.ps1 -ExcelPath .\StorageTracker.xlsx `
    -SenderMailbox 'admin@contoso.com' -TenantId '<tenant-id>' -ClientId '<client-id>' `
    -CertificateThumbprint '<cert-thumbprint>' -WhatIf

# Live, capped at 10 messages on the first real run
.\Storage\Send-StorageNotification.ps1 -ExcelPath .\StorageTracker.xlsx `
    -SenderMailbox 'admin@contoso.com' -TenantId '<tenant-id>' -ClientId '<client-id>' `
    -CertificateThumbprint '<cert-thumbprint>' -MaxSends 10
```

**Input:** an `.xlsx` tracker with site title, URL, resolution status, and owner email columns. The
header row is auto-detected, so decorative banner rows above it do not break parsing.
**Output:** a timestamped CSV run log recording every send, skip, and failure.
**Permissions:** Graph `Mail.Send` (Application), scoped by application access policy.

> **Sends real email to real people.** Safeguards: `-WhatIf` (the token is still acquired, so a dry
> run genuinely validates auth and permissions rather than just skipping the work); `-MaxSends` caps a
> run; `-DelayMs` throttles; rows without a valid recipient are skipped and logged rather than guessed at.
> Use `-WhatIf` and read the log before every live run — there is no unsend.

### `Lifecycle/Restore-OneDriveFolder.ps1`

Bulk-restores a deleted folder subtree from a site's or OneDrive's first-stage recycle bin, for
mass-delete recovery. Reports first: item count, total size, deletion date range, and who deleted
what. Restores folders before files so child paths exist, and retries with exponential back-off on
throttling.

```powershell
# Report only — no changes
.\Lifecycle\Restore-OneDriveFolder.ps1 -SiteUrl 'https://contoso-my.sharepoint.com/personal/user_contoso_com' `
    -PathFilter '*Documents/Reports*' -ClientId '<client-id>' -Tenant 'contoso.onmicrosoft.com'

# Restore
.\Lifecycle\Restore-OneDriveFolder.ps1 -SiteUrl 'https://contoso-my.sharepoint.com/personal/user_contoso_com' `
    -PathFilter '*Documents/Reports*' -ClientId '<client-id>' -Tenant 'contoso.onmicrosoft.com' -Execute
```

**Input:** target site/OneDrive URL and a wildcard path filter. Both mandatory.
**Output:** a CSV snapshot of the targeted set (written even on a dry run) and, on failures, a separate failures CSV.
**Permissions:** Site Collection Admin on the target site —
`Set-SPOUser -Site <url> -LoginName <admin> -IsSiteCollectionAdmin $true`.

> **Writes.** Safeguards: dry-run by default — nothing is restored without `-Execute`; the target set
> is snapshotted to CSV before any change; restores are idempotent, so a partial run can simply be
> re-run. `-SiteUrl` is mandatory with no default, because this script operates on another person's
> content and a wrong default target would be silent and invasive.
>
> Note it restores *into a live site*: if a name now collides with a newer file, resolve that first.
> Only first-stage recycle bin items are eligible by default (`-FirstStageOnly`).
