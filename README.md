# M365 Toolkit

[![CI](https://github.com/enricdm/m365-toolkit/actions/workflows/ci.yml/badge.svg)](https://github.com/enricdm/m365-toolkit/actions/workflows/ci.yml)

PowerShell tooling for administering Microsoft 365, Entra ID and Intune at scale — identity,
licensing, mail flow, SharePoint permissions, endpoint management and data governance.

These are working scripts, not demos. They were written to solve real problems in production
tenants — a multi-country tenant, and a managed device estate large enough that doing any of this
by hand was not an option — then audited, parameterised and stripped of anything specific to those
environments. Every script runs against your own tenant with your own credentials.

Several of these scripts have a write-up behind them — the problem, the design and what came
of it — at [enricdiaz.com/projects](https://enricdiaz.com/projects).

---

## What's here

| Platform | Scripts | What it covers |
|---|---:|---|
| [**Entra ID**](Platforms/EntraID/) | 12 | App registrations, Conditional Access, MFA exemptions, VPN sign-in risk review, usage location, licence reclamation, admin and security group provisioning |
| [**Endpoint**](Platforms/Endpoint/) | 8 | Intune assignments and group membership, device exports, primary user and category maintenance, stale-device cleanup, out-of-support Windows |
| [**Exchange Online**](Platforms/Exchange/) | 9 | Mail groups, shared mailboxes and calendars, forwarding, mail-flow protection, resource rooms |
| [**SharePoint Online**](Platforms/SharePoint/) | 2 | Everyone-Except-External grant audit, deleted folder recovery |
| [**Active Directory**](Platforms/ActiveDirectory/) | 2 | proxyAddresses repair for hybrid identity, SMBv1 network probe |
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

- **PowerShell 7.2+** (a few scripts declare `#Requires -Version 7.0`) — with one exception:
  `Remove-SensitivityLabel.ps1` requires Windows PowerShell 5.1, because the
  `PurviewInformationProtection` module is Windows PowerShell only.
- Modules, per platform: `Microsoft.Graph`, `ExchangeOnlineManagement`, `PnP.PowerShell`,
  `ActiveDirectory` (RSAT), `ImportExcel`
- Permissions vary per script and are documented in each README — most are read-only Graph
  scopes; the ones that write say which role they need

Nothing is hardcoded. Tenant, application and certificate identifiers are parameters.

### Shared configuration

`Platforms/_Shared/Data/domain-country-map.psd1` maps email domains to ISO country codes. It
ships with `contoso.*` examples — **replace it with your own** before using anything that
resolves country from a domain.

It is one file rather than a constant inside a script because that mapping previously lived
in **six** places across one estate with values that had drifted apart, and a cost-allocation
report ended up treating `GB` (2,775 users) and `UK` (2) as two different countries.

`Platforms/_Shared/Modules/M365.Common.psm1` holds helpers shared across scripts, notably
`Invoke-GraphPaged`, which handles `@odata.nextLink` correctly. It exists because three scripts
each carried their own copy of that function, one copy got a bug fix and the other two didn't —
and the resulting phantom records silently disabled a safety check on a bulk delete.

`Platforms/_Shared/Tools/Test-EndpointConnectivity.ps1` tests a list of `host:port`
destinations in four layers — DNS, timed TCP connect, TLS handshake, HTTP request — and
names the layer that failed. It exists for the argument that follows a firewall request,
and the layer that earns its keep is the TLS one: an inspecting proxy lets the connection
through and hands you a certificate it signed itself, which passes any test that stops at
"did it connect", and then breaks anything that pins or validates the issuer. The blocked
destinations are printed at the end in a form you can paste into the ticket. It is not
Microsoft 365 tooling at all, but it is the tool that gets a deployment unblocked.

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

## Tests

The shared module carries a Pester suite in [`Tests/`](Tests/) — the first case pins down
the phantom-record regression the module exists to prevent (an empty Graph collection must
return an empty list, never a one-element list holding the raw response). CI runs the suite
and error-level PSScriptAnalyzer on every push; nothing in it touches a tenant.

## Conventions

- `Verb-Noun.ps1`, PowerShell approved verbs throughout
- Every script documents itself in its header. Most use comment-based help, so
  `Get-Help .\Script.ps1 -Full` works; four older ones carry a banner comment block instead, which
  reads fine in an editor but returns nothing useful to `Get-Help`
- Most scripts write timestamped output to `Exports/` next to the script, anchored to
  `$PSScriptRoot`, never overwriting. Seven default their output path to the **current directory**
  instead — check `-OutputPath` before running those from an arbitrary working directory
- Where a script grants a privilege, it documents how to revoke it. Only one grants one
  (the Purview super user), and the revocation is not scripted — it is a documented manual step

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
