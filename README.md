# M365 Toolkit

PowerShell tooling for administering Microsoft 365, Entra ID and Intune at scale — identity,
licensing, mail flow, SharePoint permissions, endpoint management and data governance.

These are working scripts, not demos. They were written to solve real problems in production
tenants — most of them in a multi-country tenant of roughly 40,000 users, and the Intune tooling in
a managed device estate of around 2,000 endpoints — then audited, parameterised and stripped of
anything specific to those environments. Every script runs against your own tenant with your own
credentials.

---

## What's here

| Platform | Scripts | What it covers |
|---|---:|---|
| [**Entra ID**](Platforms/EntraID/) | 18 | App registrations, federated credentials, Conditional Access, MFA exclusions, usage location, licence reclamation, device object cleanup |
| [**Endpoint**](Platforms/Endpoint/) | 14 | Intune assignments and group membership, device and configuration exports, primary user and category maintenance, stale-device cleanup, device-side remediation pairs, out-of-support Windows |
| [**Exchange Online**](Platforms/Exchange/) | 11 | Mail groups, shared mailboxes and calendars, forwarding, TLS and mail-flow inspection, resource rooms |
| [**SharePoint Online**](Platforms/SharePoint/) | 8 | Site permissions, storage quotas and notifications, OneDrive recovery |
| [**Active Directory**](Platforms/ActiveDirectory/) | 2 | Privileged-account discovery, SMBv1 network probe, proxy address repair |
| [**Power Platform**](Platforms/PowerPlatform/) | 2 | Environment role audit, Environment Maker grants |
| [**Purview**](Platforms/Purview/) | 1 | Sensitivity label and encryption removal |

Each platform folder has its own README with an index, prerequisites and examples.

---

## Two things worth knowing before you run anything

### Scripts that change state say so

Every README carries a **Changes state** column. Roughly a third of these scripts write to the
tenant — they remove licences, delete group members, grant privileges, send mail, or decrypt
files. Those are marked, and each one documents its safeguards.

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

- **PowerShell 7.2+** (a few scripts declare `#Requires -Version 7.0`) — with one deliberate
  exception: everything under `Platforms/Endpoint/Remediations/` runs **on the managed device**, as
  SYSTEM, under **Windows PowerShell 5.1**, because that is what Intune Remediations executes. Those
  scripts use no modules and make no network calls.
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

- `Verb-Noun.ps1`, PowerShell approved verbs
- Comment-based help on every script: `Get-Help .\Script.ps1 -Full`
- Output to `Exports/` next to the script, timestamped, never overwriting
- No relative `.\` paths — everything anchors to `$PSScriptRoot`
- Where a script grants a privilege, it documents how to revoke it

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
