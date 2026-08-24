# M365 Toolkit

PowerShell tooling for administering Microsoft 365, Entra ID and Intune at scale — identity,
licensing, mail flow, SharePoint permissions, endpoint management and data governance.

These are working scripts, not demos. They were written to solve real problems in production
tenants — most of them in a multi-country tenant of roughly 40,000 users, and the Intune tooling in
a managed device estate of around 2,000 endpoints — then audited, parameterised and stripped of
anything specific to those environments. Every script runs against your own tenant with your own
credentials.

Several of these scripts have a write-up behind them — the problem, the design and what came
of it — at [enricdiaz.com/projects](https://enricdiaz.com/projects).

---

## What's here

| Platform | Scripts | What it covers |
|---|---:|---|
| [**Entra ID**](Platforms/EntraID/) | 16 | App registrations, federated credentials, Conditional Access, MFA exclusions, usage location, licence reclamation, device object cleanup |
| [**Endpoint**](Platforms/Endpoint/) | 10 | Intune assignments and group membership, device and configuration exports, primary user and category maintenance, stale-device cleanup, out-of-support Windows, desktop Visio/Project usage |
| [**Exchange Online**](Platforms/Exchange/) | 11 | Mail groups, shared mailboxes and calendars, forwarding, mail-flow protection and verification, resource rooms |
| [**SharePoint Online**](Platforms/SharePoint/) | 6 | Site permissions, storage quotas and notifications, OneDrive recovery |
| [**Active Directory**](Platforms/ActiveDirectory/) | 2 | proxyAddresses repair for hybrid identity, SMBv1 network probe |
| [**Power Platform**](Platforms/PowerPlatform/) | 2 | Environment role audit, Environment Maker grants |
| [**Purview**](Platforms/Purview/) | 1 | Sensitivity label and encryption removal |

Each platform folder has its own README with an index, prerequisites and examples.

---

## Two things worth knowing before you run anything

### Scripts that change state say so

Every README carries a **Changes state** column. Close to half of these scripts change something —
they remove licences, delete group members, grant privileges, send mail, or decrypt files. Those
are marked, and each one documents its safeguards.

The convention throughout:

```powershell
# Preview. Nothing is modified.
.\Invoke-Something.ps1 -TenantId '<tenant-id>'

# Apply.
.\Invoke-Something.ps1 -TenantId '<tenant-id>' -Execute
```

`-Execute` defaults to off. Where a script supports `-WhatIf`, that works too.

### Absence of data is not evidence of absence

Several scripts here cross-reference sources that can fail independently — a usage report that
returns no rows, a group that can't be enumerated, a DNS resolver that's firewalled. When that
happens they report `unknown` or `not-measured`, never the negative value.

That distinction is deliberate, and it comes from experience. A report that quietly treats a
failed lookup as "no activity" doesn't fail — it produces a confident, wrong answer. Scripts that
had that flaw were fixed; the pattern is documented where it matters.

---

## Requirements

- **PowerShell 7.2+** (a few scripts declare `#Requires -Version 7.0`) — with two exceptions:
  `Get-VisioProjectDesktopUsage.ps1` runs on the managed device under Windows PowerShell 5.1, and
  `Remove-SensitivityLabel.ps1` requires it, because the `PurviewInformationProtection` module is
  Windows PowerShell only.
- Modules, per platform: `Microsoft.Graph`, `ExchangeOnlineManagement`, `PnP.PowerShell`,
  `ActiveDirectory` (RSAT), `Microsoft.PowerApps.Administration.PowerShell`, `ImportExcel`
- Permissions vary per script and are documented in each README — most are read-only Graph
  scopes; the ones that write say which role they need

Nothing is hardcoded. Tenant, application and certificate identifiers are parameters.

### Shared configuration

`Platforms/_Shared/Data/domain-country-map.psd1` maps email domains to ISO country codes. It
ships with `contoso.*` examples — **replace it with your own** before using anything that
resolves country from a domain.

`Platforms/_Shared/Modules/M365.Common.psm1` holds helpers shared across scripts, notably
`Invoke-GraphPaged`, which handles `@odata.nextLink` correctly. It exists because three scripts
each carried their own copy of that function, one copy got a bug fix and the other two didn't —
and the resulting phantom records silently disabled a safety check on a bulk delete.

`Platforms/_Shared/Tools/Clear-M365TokenCache.ps1` clears the machine-wide MSAL, WAM and
`.IdentityService` token caches, which is the fix for a PowerShell session stuck signing in as the
wrong account. It lives here rather than under a platform because it is not scoped to one: clearing
those caches signs you out of **every** Microsoft 365 tool on that profile — Teams, Outlook, the
Graph SDK — not just the one that was misbehaving. It was called `Clear-SpoTokenCache.ps1` and sat
under `SharePoint/`, which described where the symptom usually shows up rather than what the script
touches.

---

## Getting started

```powershell
git clone https://github.com/enricdm/m365-toolkit.git
cd m365-toolkit

# Read the platform README first
code Platforms/EntraID/README.md

# Most read-only scripts need nothing but a tenant and a Graph sign-in
.\Platforms\EntraID\Applications\Export-AppRegistration.ps1 -TenantId '<tenant-id>'
```

---

## Conventions

- `Verb-Noun.ps1`, PowerShell approved verbs throughout
- Every script documents itself in its header. Most use comment-based help, so
  `Get-Help .\Script.ps1 -Full` works; four older ones carry a banner comment block instead, which
  reads fine in an editor but returns nothing useful to `Get-Help`
- Most scripts write timestamped output to `Exports/` next to the script, anchored to
  `$PSScriptRoot`, never overwriting. Seven default their output path to the **current directory**
  instead — check `-OutputPath` before running those from an arbitrary working directory
- Where a script grants a privilege, it documents how to revoke it. Only two grant one
  (`Grant-EnvironmentMaker.ps1` and the Purview super user), and neither revocation is scripted —
  both are documented manual steps

---

## Honesty about the state of things

Each platform README has a **Known rough edges** section listing defects that haven't been fixed.
They're documented rather than hidden: an interactive prompt that blocks unattended runs, a scan
that reconnects per site and is slow across a large tenant, a script that needs a workbook you'd
have to build yourself.

If something here doesn't work in your environment, the README probably already says why.

---

## Licence

MIT — see [LICENSE](LICENSE).
