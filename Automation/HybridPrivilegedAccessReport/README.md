# Hybrid privileged access report

Who holds privilege in your tenant, emailed to you once a month, with the accounts that
break your naming convention called out separately.

The question sounds simple until you try to answer it. Directory roles can be held
permanently or through PIM eligibility, they can be assigned to a person or to a group,
and a role assigned to a group tells you nothing until you expand it. The portal will
show you any one of those; putting them side by side is the work.

This is the cloud half of a hybrid report. The other half walks on-premises Active
Directory and is not published here yet — see [What is missing](#what-is-missing).

---

## What it does

A Logic App that, on demand or on a schedule:

1. Reads every directory role definition, every PIM **eligibility** schedule and every
   **active assignment** schedule.
2. Resolves each principal. If it is a group, it expands the members — a role granted to
   a group is a role held by everyone in it, and reports that stop at the group name
   understate the answer.
3. Classifies each account against your naming convention: it either matches, or it goes
   in the review list.
4. Sorts the roles into tiers you define — highest privilege, administrative, read-only,
   and everything else.
5. Emails an HTML summary with a CSV attached.

The accounts that match the convention are the boring part. The point of the report is
the callout at the top with the ones that do not: a human-looking account holding a
directory role that nobody named according to the rules is either an oversight or
something worth asking about, and either way somebody should look at it.

## What you need

- An Azure subscription and a resource group to hold the Logic App (Consumption is fine)
- A **system-assigned managed identity** on the Logic App
- An Office 365 API connection for sending the mail
- Graph application permissions, granted to that managed identity:

| Permission | Why |
|---|---|
| `RoleManagement.Read.Directory` | the role definitions and both schedule types |
| `User.Read.All` | resolving principals that are users |
| `Group.Read.All` | resolving principals that are groups |
| `GroupMember.Read.All` | expanding those groups to their members |

All read-only. Nothing here writes to the directory.

Granting Graph application permissions to a managed identity is not something the portal
does for you — there is no consent button. It is a Graph call, run once, by someone who
can consent:

```powershell
Connect-MgGraph -Scopes 'AppRoleAssignment.ReadWrite.All','Application.Read.All'

$miObjectId = '<the Logic App managed identity object id>'
$graph      = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"

foreach ($perm in 'RoleManagement.Read.Directory','User.Read.All','Group.Read.All','GroupMember.Read.All') {
    $role = $graph.AppRole | Where-Object { $_.Value -eq $perm -and $_.AllowedMemberTypes -contains 'Application' }
    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $miObjectId `
        -PrincipalId $miObjectId -ResourceId $graph.Id -AppRoleId $role.Id
}
```

## Deploying it

1. Create a Consumption Logic App in your resource group.
2. Turn on the system-assigned managed identity (Settings → Identity) and copy its
   **object id**.
3. Run the permission grant above with that object id. Give it a few minutes.
4. Create an Office 365 Outlook API connection in the same resource group and authorise
   it as the mailbox you want the report to come from.
5. Open the Logic App in code view, paste in
   [`LogicApp/PrivilegedRoleReport.logicapp.json`](LogicApp/PrivilegedRoleReport.logicapp.json),
   and save.
6. Work through the table below before you run it.
7. Run it once by hand and read the output before you put it on a schedule.

The trigger ships as **When an HTTP request is received**, because that is how it was
run in practice — fired by the on-premises half so both halves of the report land
together. If you only want the cloud half, swap the trigger for a Recurrence and set
whatever cadence suits you. Nothing else in the flow depends on the trigger type.

## What you must change

Everything in the shipped JSON is a placeholder. It will run untouched and email a
report to nobody useful, so go through all of it:

| What | Where in the JSON | Notes |
|---|---|---|
| Recipient address | `Send_Email_With_Attachments` → `To` | ships as `security-reports@contoso.com` |
| Subscription id | `parameters.$connections` | ships as all zeroes |
| Resource group | `parameters.$connections` | ships as `RSG-AUTOMATION` |
| Region | `parameters.$connections` | ships as `westeurope`; must match where you made the connection |
| Admin account domain | the `accountClassification` expressions, **four places** | ships as `@admin.contoso.com` |
| Admin account prefix | same four expressions | ships as `adm.` |
| `AdminRoles` | `Initialize_AdminRoles` | see below |
| `ReaderRoles` | `Initialize_ReaderRoles` | see below |
| `HighestPrivRoles` | `Initialize_HighestPrivRoles` | see below |

**The role lists are the one thing you should not simply accept.** They are Entra
built-in role template ids, which are the same in every tenant, so they will work
as they are — but they encode somebody else's opinion about which roles matter.
`HighestPrivRoles` in particular is what drives the top section of the email, and if
your organisation treats a role as critical that is not in that list, the report will
quietly rank it as ordinary. Read the three lists against your own tiering model
before the first real run. The ids are documented under
[Microsoft Entra built-in roles](https://learn.microsoft.com/entra/identity/role-based-access-control/permissions-reference).

**The classification convention is two string tests, not a policy engine.** An account
counts as nominal if the UPN starts with the admin prefix or ends with the admin
domain. That is deliberately crude, and it works because a naming convention that
cannot be checked with two string tests is a naming convention nobody follows. If yours
is more complicated, this is the part you rewrite.

## Things that are not obvious until they break

These cost real time to find. They are in no particular order because failures are not
in any particular order either.

- **Managed identity permissions do not appear in the portal.** After the grant above,
  the Logic App's identity blade still shows nothing useful and the app registration
  view does not list the permissions the way it would for a normal app. Verify with
  `Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId <object id>`, not by
  looking.
- **They also take a few minutes to propagate.** A run immediately after the grant will
  fail with `Authorization_RequestDenied` and look exactly like a permission you got
  wrong. Wait, then retry once, before you start debugging.
- **The sender must be a mailbox.** The Office 365 connector authenticates as whatever
  account you signed the connection in as. If you point the report at a distribution
  list, sending fails; a DL has no mailbox and no Sent Items folder.
- **A role assigned to a group is invisible if you do not expand it.** This is the
  failure mode the report exists to prevent, and it is worth stating plainly because
  every hand-rolled version of this report starts by listing assignments and stops
  there.
- **Eligible is not the same as active, and both matter.** Somebody eligible for Global
  Administrator holds that privilege in every sense that matters to an auditor, even if
  they have not activated it this month. The report reads both schedules for that
  reason.

## What is missing

Stated plainly, because the alternative is you finding out yourself:

- **The on-premises half is not here.** A scheduled PowerShell script walks AD, classifies
  privileged accounts against the same convention, and emails its own report. It is
  documented internally and needs a sanitisation pass before it can be published. Until
  then this covers Entra only, and "who is privileged" in a hybrid estate is not fully
  answered by either half alone.
- **The HTML is a hand-built table.** It renders fine in Outlook, which is what it was
  built for, and it has not been tested anywhere else.
- **There is no state between runs.** Every run is a fresh snapshot, so the report tells
  you who holds privilege today and not what changed since last month. Diffing the CSVs
  is left to you.
- **Nothing handles Graph throttling.** In a tenant with a lot of assignments the paged
  reads could hit a 429 and the run would fail rather than back off. It has not happened,
  which is not the same as it not being possible.
