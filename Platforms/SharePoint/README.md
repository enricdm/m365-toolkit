# SharePoint Online

Two scripts for SharePoint Online and OneDrive for Business: one that finds tenant-wide
permission grants nobody meant to make, and one that gets a deleted folder tree back.

This folder used to hold a storage-governance chain as well — quota reporting, owner
resolution, threshold notifications. Those four scripts only mean anything together, as a
sequence, and publishing them one by one said nothing that the admin centre does not already
say. They were pulled out to be published as a project or not at all.

## Index

| Script | What it does | Changes state | Auth |
|---|---|:---:|---|
| [`Permissions/Get-EveryoneExceptExternalGrant.ps1`](Permissions/Get-EveryoneExceptExternalGrant.ps1) | Finds "Everyone except external users" grants at site and list level | No | interactive |
| [`Lifecycle/Restore-DeletedFolder.ps1`](Lifecycle/Restore-DeletedFolder.ps1) | Restores a deleted folder subtree from a recycle bin | **Yes** | app-only + cert |

## Requirements

- **PowerShell 7.2+** (tested on 7.6.x)
- **Modules:** `PnP.PowerShell`
- **Roles:** SharePoint Administrator for tenant-wide operations; Site Collection Admin on each
  site for the permission scan and for restoring into someone's OneDrive.

### App registration (for `Restore-DeletedFolder.ps1`)

Create an Entra app registration with a certificate credential and grant admin consent for
SharePoint `Sites.FullControl.All` (Application).

Upload the certificate's public key to the app and keep the private key in
`Cert:\CurrentUser\My`. The script resolves it by subject (`-CertSubject`) or accepts an explicit
`-Thumbprint`.

## Usage

### Stuck sign-ins: the token cache fix lives elsewhere

The symptom shows up here more than anywhere else — `Connect-SPOService` keeps
signing you in as the wrong account, or into a tenant you no longer work with, and
`Disconnect-SPOService` does nothing about it, because the token is not in the
module. It is in the local identity caches, and nothing in the SPO module reaches
them.

The fix is [`_Shared/Tools/Clear-M365TokenCache.ps1`](../_Shared/Tools/Clear-M365TokenCache.ps1),
and it is not in this folder because it is not a SharePoint tool. It clears the
machine-wide MSAL, WAM and `.IdentityService` caches, which signs you out of
*every* Microsoft 365 tool on that profile — Teams, Outlook, the Graph SDK. It
used to be called `Clear-SpoTokenCache.ps1` and sat here, and that name told you
it was SharePoint-scoped when its blast radius is the whole machine.

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

### `Lifecycle/Restore-DeletedFolder.ps1`

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
.\Lifecycle\Restore-DeletedFolder.ps1 -SiteUrl 'https://contoso-my.sharepoint.com/personal/user_contoso_com' `
    -PathFilter '*Documents/Reports*' -ClientId '<client-id>' -Tenant 'contoso.onmicrosoft.com'

# Restore
.\Lifecycle\Restore-DeletedFolder.ps1 -SiteUrl 'https://contoso-my.sharepoint.com/personal/user_contoso_com' `
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

## Known rough edges

- **A wrong default on a connection cmdlet is worse than no default.** There used to be a
  `Connect-PnPSite.ps1` here that claimed to factor out certificate resolution, but nothing
  called it, so it was an extra copy rather than a shared one; it has been removed. An early
  version of it defaulted `-Url` to one person's OneDrive, so running it with no arguments
  connected you to a colleague's files and looked like success. Every `-Url`/`-SiteUrl` in this
  folder is mandatory for that reason.
- **`Get-EveryoneExceptExternalGrant.ps1` reconnects once per site.** `Connect-PnPOnline` is
  called inside the site loop, because PnP binds a connection to a single site. Across a large
  tenant that dominates the runtime — a full scan is measured in hours, not minutes. Use
  `-UrlPattern` to scope it wherever you can.
- **`-ScanItems` is slow enough to be a separate decision.** Item-level scanning walks every list
  item with unique permissions and is opt-in for that reason. Do not add it to a tenant-wide run
  without knowing how long you are prepared to wait.
- **The permission scan needs Site Collection Admin on every site it reads.** The app registration
  having `Sites.FullControl.All` is not always sufficient for the per-site PnP calls, and a site
  that cannot be opened is reported as an error row rather than silently skipped — but it is still
  a gap in the scan, and a scan with gaps is what this script exists to avoid.
- **`Restore-DeletedFolder.ps1` restores into a live site.** Names that have since been reused will
  collide, and only first-stage recycle bin items are eligible by default. Neither is a defect, but
  both surprise people mid-recovery.
