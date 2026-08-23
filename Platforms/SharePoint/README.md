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

## Requirements

- **PowerShell 7.2+** (tested on 7.6.x). `Get-SiteStorage.ps1`, `Set-SiteStorageQuota.ps1` and
  `Clear-SpoTokenCache.ps1` use the SPO module, which also works on Windows PowerShell 5.1.
- **Modules**
  - `PnP.PowerShell` — the connection, permission and lifecycle scripts
  - `Microsoft.Online.SharePoint.PowerShell` — the storage scripts
  - `ImportExcel` — `Send-StorageNotification.ps1` only
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

### `Connect-PnPSite.ps1`

Every app-only script in this folder needs the same four steps before it can do
anything: find the certificate, pair it with the app registration, open the PnP
connection, fail clearly when the certificate is not there. This is that, factored
out, so the certificate-resolution logic exists once instead of four times.

Its parameters are mandatory with no defaults, and that is the point. An earlier
version defaulted `-Url` to one person's OneDrive, so running it with no arguments
connected you to a colleague's files and gave no indication that anything unusual
had happened. A wrong default on a connection cmdlet is worse than no default: it
turns a forgotten parameter into a silent, invasive success. `-Thumbprint` is the
one genuinely optional value — omit it and the newest non-expired certificate
matching `-CertSubject` is used.

```powershell
# Open an app-only session (-Url is mandatory: no default target, by design)
.\Connect-PnPSite.ps1 -Url 'https://contoso.sharepoint.com/sites/Example' `
    -ClientId '<client-id>' -Tenant 'contoso.onmicrosoft.com'
```

### `Clear-SpoTokenCache.ps1`

The symptom is specific and maddening: `Connect-SPOService` keeps signing you in
as the wrong account, or into a tenant you no longer work with, and
`Disconnect-SPOService` does nothing about it. The token is not in the module. It
is in the local identity caches — MSAL, the WAM token broker, `.IdentityService` —
and nothing in the SPO module reaches them.

Eight lines, with a side effect considerably larger than the name suggests: those
caches are shared across the machine. Clearing them signs you out of *every*
Microsoft 365 tool on that profile, not just SharePoint, so expect fresh sign-ins
in Teams, Outlook and the Graph SDK afterwards. Close the PowerShell window and
open a new one before running it — a session that has already loaded the identity
assemblies can write the caches back on exit.

```powershell
# Fix a sign-in stuck on the wrong account, then reconnect
.\Clear-SpoTokenCache.ps1 -AdminUrl 'https://contoso-admin.sharepoint.com'
```

### `Storage/Get-SiteStorage.ps1`

The first half of the storage chain, and the reason it is a separate script from
the second half: reading quota and writing quota deserve different levels of care.
Run this, read the output, confirm the exact URLs — SharePoint site URLs are long,
similar, and easy to copy one line off — and only then hand the confirmed ones to
`Set-SiteStorageQuota.ps1`. Its export doubles as the "before" snapshot for that
change.

One caveat about the percentage it reports: it is computed against whatever cap is
set on the site, and in a pooled-storage tenant that cap is often a large shared
ceiling rather than a real per-site allocation. Useful for spotting sites that
have been given an explicit quota; not a reliable ranking of who is actually full.
The same trap is documented in more detail under `Get-SiteOwnerStatus.ps1` below.

```powershell
# Storage usage for every site, or just the ones matching a name
.\Storage\Get-SiteStorage.ps1 -AdminUrl 'https://contoso-admin.sharepoint.com'
.\Storage\Get-SiteStorage.ps1 -AdminUrl 'https://contoso-admin.sharepoint.com' -TitleFilter 'Finance'
```

**Output:** a timestamped CSV in the current folder or `-OutputPath`.

### `Permissions/Get-SiteOwnerStatus.ps1`

A site whose only owner has left the company is not an edge case in a tenant this
size; it is a steady background rate. Nothing breaks when it happens, which is
exactly the problem — the site keeps serving its content, keeps consuming pooled
storage, and there is no longer anybody with the standing to answer a question
about it or approve its deletion. No standard report surfaces this, because the
three facts you need live in three different places: group-connected sites keep
their owners in the Microsoft 365 group, non-group sites keep theirs in a
SharePoint group that Graph `/groups` cannot see at all, and `accountEnabled`
lives in Entra.

Joining those three is most of what this script does. The care went into the
failure cases, because this output is what decommissioning decisions get made on,
and a report that is wrong in a plausible-looking way is more dangerous than one
that fails outright.

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

"Everyone except external users" is SharePoint's way of saying "the whole
company". Granting it is sometimes exactly right — an intranet, a policy library,
anything meant to be readable by all staff — and sometimes it is what happens when
a site gets shared in a hurry, after which a finance folder is readable by every
licensed user in the tenant and no error is ever raised. Nobody stumbles onto that
by browsing. You find it by scanning for the principal, or you do not find it.

The design decision that carries this script — matching on the login claim rather
than the display name — is written up below, and it is a lesson learned rather
than foresight. A previous version lost the claim check and compared display
titles only, and it returned clean reports for sites that did have EEEU grants
sitting on them. A permissions scanner with false negatives is worse than no
scanner at all, because a green result is precisely what stops anyone from looking
again.

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

## Scripts that change state

### `Storage/Set-SiteStorageQuota.ps1`

Raising a quota is a one-line change, and the reason it deserves a script is what
happens afterwards. The original version changed the quota and did not record what
it had been, so the only route back was somebody's recollection of a number in
megabytes. It also had no `-WhatIf`, which meant this one script broke the
convention every other script here follows. Both were added, and the safeguards
below are the result — none of them clever, all of them missing at first.

The `unknown` in the note below is the same principle as the tri-state
`accountEnabled` above, and it is the thread that runs through this repository: a
quota that could not be read is recorded as unread, because `0` looks like a
value somebody could roll back to, and eventually somebody would.

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

The end of the storage chain. `Get-SiteStorage.ps1` finds the sites over
threshold, `Get-SiteOwnerStatus.ps1` says who owns them, and this is the part that
asks those owners to do something about it — the weekly follow-up, driven from the
tracker workbook so the state of each conversation stays with the people having it
rather than inside a script.

The decision worth pausing on is the permission, not the code. `Mail.Send` as an
application permission, granted the ordinary way, lets the app send as *any*
mailbox in the tenant — a reporting job with the standing to impersonate the
finance director. Scoping it with an Exchange Application Access Policy to the
single sender mailbox is what turns that into something proportionate, and it is
the sort of constraint nobody notices is missing, because the script works
identically either way. The other decision, signing the token by hand instead of
using the Graph SDK, is explained below; it is a workaround for an assembly
conflict, not a preference.

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

This is recovery work, and it arrives urgent: someone synced a library, tidied up
what they thought was a local copy, and took a folder tree with them. The
first-stage recycle bin still holds everything, so the data is not gone — but
"restore everything" is the wrong answer when the accidental deletion happened in
the middle of a legitimate cleanup, and restoring a few thousand items by hand
through the web UI is not an answer at all.

So the unit of work is a path subtree, and it is worth being clear that
`-PathFilter` is the only real safeguard here: it is the whole of what stands
between restoring one folder and dumping half a library back into a live site.
That is why it reports before it acts, and why neither it nor `-SiteUrl` has a
default — the script is usually pointed at somebody else's content, under time
pressure, by someone who has already had a bad morning.

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