# Power Platform

Two scripts for Power Platform environment governance: a tenant-wide audit of privileged role
assignments, and a least-privilege grant for a single environment.

## Index

| Script | What it does | Changes state | Auth |
|---|---|:---:|---|
| [`Get-PowerPlatformAdminAudit.ps1`](Get-PowerPlatformAdminAudit.ps1) | Lists every System Administrator / Environment Maker assignment across all environments | No | interactive admin |
| [`Grant-EnvironmentMaker.ps1`](Grant-EnvironmentMaker.ps1) | Grants one user the Environment Maker role on one environment | **Yes** (with `-Execute`) | interactive admin |

## Requirements

- **Modules**
  - `Microsoft.PowerApps.Administration.PowerShell` —
    `Install-Module Microsoft.PowerApps.Administration.PowerShell -Force -AllowClobber`
  - `Microsoft.Graph.Users` — `Grant-EnvironmentMaker.ps1` only, to resolve a UPN to an object ID
- **Roles:** Power Platform Administrator, or Global Administrator / Dynamics 365 Service Admin.

Both scripts sign in interactively via `Add-PowerAppsAccount`.

## Usage

### `Get-PowerPlatformAdminAudit.ps1`

Power Platform makes environments easy to create, and their administrator roles live **outside**
Entra ID's role model — so they never show up in the privileged access reviews that cover the rest
of the tenant. Someone holding environment administrator can build flows that run under their own
credentials and reach corporate data, and a PIM review will not see any of it. The
standing-privilege question has to be asked separately here, and environments tend to accumulate
makers nobody remembers granting.

So this enumerates every environment in the tenant and reports who holds System Administrator or
Environment Maker on each. One easy mistake it documents in a comment:
`Get-AdminPowerAppRoleAssignment` is an **app**-level cmdlet — it wants an `-AppName` and answers a
different question. Environment-level roles come from `Get-AdminPowerAppEnvironmentRoleAssignment`.

```powershell
.\Get-PowerPlatformAdminAudit.ps1

# Deep-dive specific environments when the main scan returns nothing
.\Get-PowerPlatformAdminAudit.ps1 -DiagnosticEnvironmentName '<environment-guid>','<environment-guid>'
```

**Input:** none; optionally `-DiagnosticEnvironmentName` with one or more environment GUIDs.
**Output:** an on-screen table, a total count, and a diagnostic list of every distinct `RoleType` seen.
**Permissions:** Power Platform Administrator. Read-only.

> **An empty result is ambiguous, and the script says so rather than hiding it.** Zero assignments
> can mean the environment genuinely has no explicit role records — access coming instead from
> implicit tenant-admin rights or an environment security-group restriction — or it can mean the
> signed-in account cannot see them. Those are very different findings, and a bare "0" would let you
> file a clean audit over a permissions failure.
>
> `-DiagnosticEnvironmentName` tells them apart by dumping the raw assignment records for specific
> environments. When the main scan comes back empty the script prints the environment list ready to
> paste into that switch. Role matching is substring-based (`Admin`, `Maker`) because the exact
> `RoleType` strings vary between module versions.

## Scripts that change state

### `Grant-EnvironmentMaker.ps1`

The common request is "I need to import my solution into the team environment", and the common
answer is to make the person an environment administrator, which grants a great deal more than the
request needs. Environment Maker is the smaller answer: it lets someone create and import apps and
flows in **one** environment, and nothing beyond that — not administration of the environment, not
anything at tenant level.

What it does not do is worth stating plainly, because that is where half an hour of confusion goes:
in an environment with Dataverse, importing a solution also needs the `System Customizer` Dataverse
role, and this script grants no Dataverse role at all — that step stays manual, in the portal
(details below). Worth knowing too: there is no revocation script in this repository, for this grant
or for the Purview super user one. Both privileges are granted by script and taken back by hand.

```powershell
# Dry run — prints the intended change and exits without connecting to anything
.\Grant-EnvironmentMaker.ps1 -EnvironmentId '<environment-guid>' -UserPrincipalName 'user@contoso.com'

# Apply
.\Grant-EnvironmentMaker.ps1 -EnvironmentId '<environment-guid>' -UserPrincipalName 'user@contoso.com' -Execute
```

**Input:** the target environment's GUID (`EnvironmentName`, e.g. from the audit script's output) and
the user's UPN.
**Output:** the resolved object ID, the assignment result, and a verification table read back from the API.
**Permissions:** Power Platform Administrator, plus Graph `User.Read.All` to resolve the UPN.

> **Writes a role assignment when `-Execute` is passed.** Safeguards: dry-run by default, and without
> `-Execute` it does not even connect; the user is resolved with `Get-MgUser -UserId` rather than a
> `-Filter` query, so a UPN either resolves to exactly one object or the script aborts — a filter can
> return a collection and silently assign the role to an array of principals; the environment is
> confirmed to exist before assigning; and the result is verified by reading the assignment back.
>
> One non-obvious failure it handles: `Set-AdminPowerAppEnvironmentRoleAssignment` can return an
> error object (a 403, for instance) as **normal output** instead of throwing. A naive script prints
> "Success" and moves on. This one inspects the returned object and fails loudly.

> **Dataverse environments need a second, manual step.** In an environment with a Dataverse
> database, Environment Maker alone grants access to the Power Apps/Flow authoring surfaces but
> **not** solution import rights. Those come from a Dataverse security role such as **System
> Customizer**, assigned in the Power Platform Admin Center under *Environment → Settings → Users +
> permissions → Security roles*. The script prints this reminder after a successful grant, but it
> cannot do it for you — so a request is not finished when this script exits.

## Known rough edges

- **There is no revocation script.** `Grant-EnvironmentMaker.ps1` gives the role and nothing here
  takes it back. Revoke in the Power Platform Admin Center, or with
  `Remove-AdminPowerAppEnvironmentRoleAssignment`. The same asymmetry applies to the Purview super
  user grant one folder over — both privileges are granted by script and returned by hand, which is
  the wrong way round for anything you want reliably closed out.
- **A Dataverse environment needs a manual second step.** Environment Maker alone does not grant
  solution import rights; the `System Customizer` Dataverse role does, and it is assigned in the
  portal. The script prints the reminder, but a request is not finished when the script exits.
- **Role matching is substring-based.** `Get-PowerPlatformAdminAudit.ps1` matches on `Admin` and
  `Maker` because the exact `RoleType` strings have changed between module versions. A future
  version that renames them again will silently match less, so the diagnostic list of every distinct
  `RoleType` seen is worth reading rather than skipping.
- **An empty audit result is ambiguous by nature.** Zero assignments can mean no explicit role
  records exist, or that the signed-in account cannot see them. The script says so and gives you
  `-DiagnosticEnvironmentName` to tell the two apart, but it cannot resolve it for you — and a bare
  zero is exactly the kind of result that gets filed as a clean audit.
