# Entra ID

Operational tooling for Microsoft Entra ID: app-registration hygiene, Conditional
Access exception analysis, identity lifecycle and license reclamation. Everything
here talks to Microsoft Graph and is written to run against *any* tenant — tenant
ids, domains, group names and repository names are parameters, never constants.

Most scripts are read-only and produce CSV/XLSX you can hand to someone else. The
six that change directory state are called out explicitly below and each one has
its own section.

These were written against a multi-country tenant of roughly 40,000 users, with
on-prem AD synchronised by Entra Connect from several forests with different sync
rules. That shape matters more than the headcount: attributes are populated
inconsistently depending on which forest an account came from, so a report that
treats a blank field as a negative answer will be wrong for thousands of people
and will look perfectly healthy while it is.

That is the thread running through this folder. An audit of the earlier version of
these tools found eight separate reports drawing false conclusions without ever
failing — an empty read was being counted as data. One declared 340 Visio and 179
Project licenses unused because two optional parameters were missing; it finished
green, with no errors. Several scripts here therefore keep `unknown` and
`not-measured` distinct from the negative value, and say so in their output. Where
one does, this README points it out.

## Index

| Script | What it does | Changes state | Auth |
|---|---|:---:|---|
| **Applications** | | | |
| [`Export-AppRegistration.ps1`](Applications/Export-AppRegistration.ps1) | Full app-registration inventory: owners, credentials, API permissions, sign-in activity, hygiene flags | No | interactive |
| [`Export-SamlCertificateExpiry.ps1`](Applications/Export-SamlCertificateExpiry.ps1) | Lists SAML apps with their certificate expiry date and notification addresses | No | interactive |
| [`New-M2MAppRegistration.ps1`](Applications/New-M2MAppRegistration.ps1) | Bulk-creates client app registrations + service principals for cert-based M2M auth | **Yes** | interactive |
| **ConditionalAccess** | | | |
| [`Get-MfaExemption.ps1`](ConditionalAccess/Get-MfaExemption.ps1) | Users genuinely exempt from MFA — exception-group members and direct exclusions, nested groups expanded. `-AllExclusions` for the full dump | No | app-only + cert, or interactive |
| [`Get-ExemptionSignInActivity.ps1`](ConditionalAccess/Get-ExemptionSignInActivity.ps1) | Adds 30-day sign-in telemetry to those exemptions and says which can be dropped or narrowed | No | app-only + cert, or interactive |
| [`Export-ConditionalAccessPolicy.ps1`](ConditionalAccess/Export-ConditionalAccessPolicy.ps1) | Exports all CA policies with every GUID resolved to a display name | No | interactive |
| **Identity** | | | |
| [`New-SecurityGroup.ps1`](Identity/New-SecurityGroup.ps1) | Creates security groups from a definition file and adds members | **Yes** (with `-Execute`) | interactive |
| [`New-AdminAccount.ps1`](Identity/New-AdminAccount.ps1) | Creates a cloud-only admin account derived from a person's ordinary account, with TAP and per-user MFA | **Yes** | interactive |
| **Licensing** | | | |
| [`Get-LicenseReclamationPlan.ps1`](Licensing/Get-LicenseReclamationPlan.ps1) | Tiers every holder of a SKU into KEEP / REVIEW / CONVERT_SHARED / RECLAIM / EXCLUDE with the reasoning shown | No | interactive |
| [`Invoke-LicenseReclamation.ps1`](Licensing/Invoke-LicenseReclamation.ps1) | Executes that plan: removes licenses, removes group membership, optionally converts mailboxes to shared | **Yes** | interactive (+ EXO) |
| **Devices** | | | |

## Requirements

**PowerShell**

- PowerShell 7.2 or later.
- `Microsoft.Graph` 2.x. Individual scripts need only some sub-modules —
  `.Authentication`, `.Users`, `.Groups`, `.Applications`, `.Identity.SignIns`,
  `.Reports` — and each script lists its own in the header.
- `ExchangeOnlineManagement` only for `Invoke-LicenseReclamation.ps1 -Tiers CONVERT_SHARED`
  and `Get-LicenseReclamationPlan.ps1 -EnrichFromExchange`.

**Graph permissions**

Each script documents its own scopes in the comment-based help. Broadly:

| Task | Scopes |
|---|---|
| Read-only reporting | `Directory.Read.All`, `User.Read.All`, `Group.Read.All`, `Policy.Read.All`, `Application.Read.All`, `AuditLog.Read.All`, `Reports.Read.All` |
| Identity Protection data | `IdentityRiskyUser.Read.All`, `IdentityRiskEvent.Read.All` |
| Creating accounts / groups | `User.ReadWrite.All`, `Group.ReadWrite.All`, `GroupMember.ReadWrite.All`, `UserAuthenticationMethod.ReadWrite.All`, `Policy.ReadWrite.AuthenticationMethod` |
| Creating app registrations | `Application.ReadWrite.All` |
| License reclamation | `User.ReadWrite.All`, `Group.ReadWrite.All`, `Directory.Read.All`, `Organization.Read.All` |

`signInActivity` and sign-in logs require Entra ID **P1 or higher**, and Graph
retains interactive sign-in logs for roughly 30 days.

**App-only certificate auth** (optional, used by the ConditionalAccess exports)

1. Create an app registration, grant it the application permissions above and
   admin-consent them.
2. Upload a certificate to that app registration and install the private key in
   `Cert:\CurrentUser\My`.
3. Pass `-ClientId '<client-id>' -CertThumbprint '<cert-thumbprint>'`. Omit both
   and the script falls back to interactive delegated auth, printing a warning.


## Usage

### Read-only reporting

#### `Export-AppRegistration.ps1`

App registrations accumulate. A tenant that has been running for years ends up with
hundreds of them, and the hard question is never how to create another one — it is
which of the existing ones are still alive, which are carrying credentials about to
expire, and who is supposed to care. None of that is visible from the portal
without opening every app in turn.

So this dumps everything about every app in one pass: owners, secrets and
certificates with their expiry dates, federated credentials, API permissions,
directory roles held by the service principal, redirect URIs. Then it adds
seventeen boolean `Flag_*` columns — no owner, expired secret, expiring secret,
long-lifetime secret, no activity, unverified publisher, disabled by Microsoft,
personal-account audience, and so on — because the export exists to be filtered in
Excel by whoever is running the clean-up, not read top to bottom.

Activity is collected from two independent places on purpose. Audit-log sign-ins
only go back 30 days, so an app that is called once a quarter looks dead there; the
service principal's own `signInActivity` persists beyond the log window and catches
exactly that case. Neither source alone is enough to call an app abandoned.

```powershell
# Everything about every app registration, with hygiene flags for Excel filtering
.\Applications\Export-AppRegistration.ps1 -TenantId '<tenant-id>' -OutputPath .\Exports
```

#### `Export-SamlCertificateExpiry.ps1`

A SAML signing certificate expires and the application stops authenticating.
Entra will warn about it, but only to the addresses in
`notificationEmailAddresses`, which by default is whoever created the app — often
someone who left. Knowing which certificates expire when, and who is actually going
to be told, is a list worth having before the first outage rather than after.

The detail worth knowing here is why the filter is written the way it is: Graph
does not return `null` for an app with no SAML certificate. It returns
`0001-01-01T00:00:00Z`. Parse that as a date and every non-SAML service principal
in the tenant turns into a SAML app whose certificate expired two thousand years
ago. The sentinel is filtered out explicitly, and the script tells you when it
found zero SAML apps instead of quietly writing an empty CSV.

```powershell
# SAML signing certificates and who gets warned when they expire
.\Applications\Export-SamlCertificateExpiry.ps1 -OutputPath .\Exports\saml.csv
```

#### `Export-ConditionalAccessPolicy.ps1`

The raw policy export from Graph is close to unreadable. A policy that says
`excludeGroups: ["20f65788-…"]` is useless in a review: nobody in the room knows
what that group is, and the reviewer who does know has to leave and go look it up,
for every GUID, in every policy.

So this resolves every identifier to a display name before writing anything. That
turns out to be harder than one lookup, because a GUID in a CA policy can be a
user, a group, a directory role, a service principal, an application id, or a named
location, and the policy body does not tell you which. The resolver tries five
directory endpoints in order, then falls back to a service-principal lookup by
`appId`, and caches what it finds. Both files are kept: the raw JSON is lossless,
the resolved JSON is the one you can hand to an auditor, and the flat CSV is for
diffing two exports against each other.

```powershell
# All CA policies with GUIDs resolved: raw JSON, resolved JSON, flat CSV, named locations
.\ConditionalAccess\Export-ConditionalAccessPolicy.ps1 -TenantId '<tenant-id>'
```

#### Password age, without a script

There was one here and it has been removed: scoped to a country, it was a single
Graph query plus a computed column, which is not worth a file. The query, if you
need it:

```powershell
Get-MgUser -All -ConsistencyLevel eventual -CountVariable c `
    -Filter "usageLocation eq 'ES' and accountEnabled eq true" `
    -Property displayName,userPrincipalName,lastPasswordChangeDateTime,signInActivity |
    Export-Csv .\password-age-ES.csv -NoTypeInformation
```

`-ConsistencyLevel eventual -CountVariable` is not optional there: combining a
`$filter` with `signInActivity` needs an advanced query, and without it the call
fails rather than degrading.

The part worth keeping is the classification, not the query. Report three states, not
two: a user whose `lastPasswordChangeDateTime` is absent is `Unknown`, never `OK`.
That field is empty often enough on synchronised accounts that folding it into the
compliant bucket quietly inflates the pass rate — which is the same mistake, in a
smaller place, as counting an empty read as a negative result.

**Input:** parameters only.
**Output:** timestamped CSV/JSON under each script's `Exports` folder (or `-OutputPath`).
**Permissions:** the read-only scopes in the table above.

### The MFA exemption chain

Every MFA exemption was granted for a reason, and almost none of them are ever
revisited. The list grows, the original reason expires, and the exemption stays —
usually inside a group that somebody added a nested group to, so the number of
exempt people is larger than anyone believes. Asking "how many users are excluded
from MFA?" and reading the count of direct members off the CA policy blade gives an
answer that is always too low.

`Get-MfaExemption.ps1` answers the first question properly, and it took two goes to
get there. It expands excluded groups transitively, including groups nested inside
them, so the count is the real population and not just the names that happen to be
listed on the policy.

The second correction is the one worth reading, because the script was wrong in the
direction that does not look wrong. It used to report every exclusion it found, and
an exclusion is not an exemption: the largest exclusion lists in this tenant belong
to policies whose grant control is **block**, and excluding somebody from a policy
that blocks is the opposite of exempting them from MFA. The run that prompted the fix
produced **17,682 rows where the real answer was 192** — a factor of 92 — and the
report was filtered by hand before anyone used it. A number that has to be corrected
by hand before use is not a report, and the number this produces is the one somebody
quotes to security.

So the narrowing is now the default and the script is named for the question rather
than the mechanism. `-AllExclusions` still gives the full dump; that is a legitimate
question — *where does this user appear in any exclusion list* — but it is a
different one, and it says so.

`Get-ExemptionSignInActivity.ps1` exists because that list, on its own, cannot be
acted on. Knowing 400 people are exempt tells you nothing about which exemptions to
remove first. So the second script takes the CSV the first produced and answers the
question that follows it: **of the people who are exempt, who actually uses the
exemption?** An exemption nobody exercises is security debt that can be withdrawn
without breaking anything, and withdrawing it is a conversation you can win with
data rather than intuition. That is why this is two scripts and not one — the first
is a fact about configuration, the second is a fact about behaviour, and they come
from completely different places.

They are meant to be run one after the other — the second consumes the CSV the
first produces. The first says *who* is exempt, the second says *whether the
exemption is still deserved*.

```powershell
# 1. Who is exempt from MFA today? Scope it to the baseline policy; without that the
#    input set includes app-scoped policies too.
.\ConditionalAccess\Get-MfaExemption.ps1 -TenantId '<tenant-id>' `
    -PolicyName '*Require MFA*' -MasterExceptionGroup 'CA-Exception-MFA'
#    -> Exports\CA-MFA-Exemptions_<stamp>.csv
#    -> Exports\CA-MFA-ExemptionSources_<stamp>.csv   (per-group breakdown)

# 2. Feed that CSV back in to see 30 days of sign-in behaviour per exempt user.
.\ConditionalAccess\Get-ExemptionSignInActivity.ps1 -TenantId '<tenant-id>' `
    -InputCsv .\ConditionalAccess\Exports\CA-MFA-Exemptions_<stamp>.csv -Detailed
#    -> Exports\Exemption-SignInActivity_<stamp>.csv
```

Step 2 sorts the output so the best candidates come first: no sign-ins at all
(drop the exemption), a single application (scope the exemption to that app),
non-interactive only (move it to a certificate or managed identity), then broad
usage. It also flags users who are *already completing MFA* on some sign-ins,
which makes their exemption moot.

Step 2 narrows the CSV to true exemptions by default; groups that *enforce* MFA on
their members are usually excluded from block policies too, and counting them
inflates the total. Pass `-AllRows` to keep every row.

Once the exemption list exists, somebody has to own each row, and in a multi-country
tenant that means splitting it by country before it can be sent anywhere. That is a
separate, offline step, and deliberately so: attribution is guesswork over three
imperfect signals, and I would rather it happened in a script that produces a
reviewable spreadsheet than inside the enumeration that produces the security
number.

Resolving `usageLocation` for the resulting UPN list needs no script. The Entra admin
center exports it directly — **Identity → Users → All users → Download users** returns
`userPrincipalName`, `usageLocation`, `accountEnabled` and `userType` — and for a
filtered list, one pipeline does it:

```powershell
Import-Csv .\exempt-users.csv | ForEach-Object {
    Get-MgUser -UserId $_.UserPrincipalName `
               -Property UserPrincipalName,UsageLocation,AccountEnabled,UserType `
               -ErrorAction SilentlyContinue
} | Export-Csv .\usage-location.csv -NoTypeInformation
```

Whatever you use, keep the UPNs that fail to resolve in the output instead of dropping
them, so the row count going in matches the row count coming out.

### `New-M2MAppRegistration.ps1`

> **Warning: creates directory objects.** One app registration plus one service
> principal per name, with the supplied owner attached to both.

Creating one app registration by hand is a two-minute job. Creating fifteen of them
consistently is not, and the inconsistency is what causes trouble later: an app
created in a hurry gets no owner, no service principal, a redirect URI copied from
another app, and a client secret that will outlive the project. Every one of those
shows up as a flag in `Export-AppRegistration.ps1` a year afterwards.

The decision worth pointing out is that this creates **no credentials at all**. It
would be easy to generate a client secret and print it, and that is precisely the
path that ends with secrets in chat history and in PowerShell transcripts. The
owner uploads their own certificate afterwards, which keeps the private key on the
machine that will use it and keeps this script entirely out of the secret's
lifecycle.

Safeguards: an existing application with the same `displayName` is skipped with a
warning rather than duplicated, so re-running after a partial failure is safe.
Governance metadata is stamped into the Notes field of every app it creates.

```powershell
.\Applications\New-M2MAppRegistration.ps1 -Name 'APP-Client-01','APP-Client-02' `
    -OwnerUpn 'admin@contoso.com'
```

**Input:** the names to create and an owner UPN that must already exist.
**Output:** a summary table of display name, client id, object id and SP object id.
**Permissions:** `Application.ReadWrite.All`, `User.Read.All`.

### `New-SecurityGroup.ps1`

> **Warning: creates groups and adds members** — but only with `-Execute`. Without it
> the script previews every action and touches nothing.

Group creation requests arrive as a list — a new team, a new project, a new
workspace, with the members already named in the ticket. Doing it through the portal
is clicking, and clicking is where the typo in the group name comes from and where
the member that got skipped goes unnoticed. The request is data; it should be
applied as data.

The two decisions here are the preview default and idempotency, and they exist for
the same reason: this is the kind of script you end up running more than once.
Nothing is written without `-Execute`, so the first run is always a review of what
would happen. And a second run after a partial failure is safe rather than
destructive — an existing group with the same display name is reused rather than
duplicated, members already present are skipped, and UPNs that do not resolve are
reported as `NOT FOUND` instead of aborting the run. A bad UPN in row 40 should not
prevent rows 41 to 80 from being processed; it should appear in the results table
where somebody can fix it.

Groups are created as plain security groups (`mailEnabled=$false`,
`securityEnabled=$true`, no `groupTypes`), which is what tools like Microsoft Fabric
require for workspace access — a Microsoft 365 group looks similar in the portal and
will not work there.

```powershell
# Preview
.\Identity\New-SecurityGroup.ps1 -DefinitionsPath .\groups.psd1

# Apply
.\Identity\New-SecurityGroup.ps1 -DefinitionsPath .\groups.psd1 -Execute
```

`groups.psd1` looks like this:

```powershell
@{ Groups = @(
    @{ Name        = 'App_Dev_ES'
       Description = 'Application developers - Spain'
       Members     = @('user1@contoso.com', 'user2@contoso.com') }
) }
```

**Input:** a definitions file, or the CONFIG block at the top of the script.
**Output:** an on-screen result table and `SecurityGroups_<stamp>.csv` next to the script.
**Permissions:** `Group.ReadWrite.All`, `User.Read.All`, `GroupMember.ReadWrite.All`.

### `New-AdminAccount.ps1`

> **Warning: creates a privileged account.** It creates a cloud-only admin user, sets
> its manager and `ExtensionAttribute1`, issues a Temporary Access Pass, enforces
> per-user MFA and optionally adds it to a Conditional Access group. There is no
> `-WhatIf`; the safeguard is an explicit confirmation prompt showing the exact
> account it is about to create.

Separating administrative privilege from the account somebody reads mail and opens
attachments with is standard Microsoft guidance, and it is one of the least
automated things in most tenants. It gets done by hand, inconsistently, which is why
so many tenants end up with admin accounts that follow three different naming
conventions and cannot be traced back to a person at all. An admin account nobody
can attribute is an admin account nobody dares disable.

You supply only the person's ordinary account. Everything else — UPN, display
name, mail nickname, usage location, manager — is derived from it, which is what
keeps admin accounts consistently named and reliably linked back to a human. The
manager and `ExtensionAttribute1` are set for the same reason: six months later,
someone auditing privileged accounts needs to know whose this is without asking
around.

The handover is done with a Temporary Access Pass rather than a password. That is
the part I would keep even if nothing else survived: a TAP lets the person register
their own MFA method on first use, so no password ever has to be typed into a chat
window or an email to reach them. The credential that gets shared is short-lived by
construction instead of by policy.

```powershell
# Internal-only admin account
.\Identity\New-AdminAccount.ps1 -NominalUpn 'user@contoso.com' `
    -AdminUpnSuffix 'contoso.onmicrosoft.com'

# External access: offer the CA group, and print the credentials for handover
.\Identity\New-AdminAccount.ps1 -NominalUpn 'user@contoso.com' `
    -ConditionalAccessGroupId '<group-object-id>' -AssignedRoles 'User Administrator' `
    -ShowCredentials
```

**Input:** the nominal UPN. The nominal account must have `surname`,
`givenName` and `usageLocation` populated; missing values are warned about, and a
missing manager leaves the admin account without one rather than failing.
**Output:** the created account, a summary, and — only with `-ShowCredentials` —
a ready-to-send message for the end user.
**Permissions:** `User.ReadWrite.All`, `Directory.ReadWrite.All`,
`UserAuthenticationMethod.ReadWrite.All`, `Policy.ReadWrite.AuthenticationMethod`,
`Group.ReadWrite.All`. Activate User Administrator, Authentication Administrator,
Authentication Policy Administrator and Groups Administrator in PIM first.

Two things worth knowing:

- **The TAP lifetime is a security setting, not a convenience one.** It defaults to
  720 minutes (12 h) and is capped at 24 h by `[ValidateRange]`. A Temporary Access
  Pass is a fully valid credential for its entire lifetime, so a multi-day TAP is a
  multi-day bypass of your sign-in controls. Your tenant's TAP policy caps it too,
  and the lower of the two wins.

  The previous version of this script issued the TAP with a lifetime of 14,000
  minutes — just under ten days — while the comment directly above it said 12 hours.
  Nobody noticed, because a TAP that works is indistinguishable from a TAP that
  works for too long. Hence the 720 default and the `[ValidateRange]`: a value that
  is a security decision should not be a number somebody can quietly get wrong.
- **Credentials are not printed by default.** The temporary password and the TAP
  are masked unless you pass `-ShowCredentials`, so an ordinary run leaves nothing
  usable in console scrollback or a PowerShell transcript. The confirmation prompt
  warns you when the switch is missing, before the account is created.

PIM eligible role assignment stays manual and out of scope; pass the role you
assigned in `-AssignedRoles` so it shows up in the end-user message.

### License reclamation: plan, then execute

An E3 SKU sitting at 100% consumption is a purchase request unless somebody can show
which of those seats are doing nothing. The evidence is scattered across places that
disagree with each other — sign-in activity, mailbox usage reports, password age,
recipient type — and none of them is conclusive on its own. A user with no
interactive sign-ins in a year might be dormant, or might be a service account, or
might be on parental leave, or might simply be outside the log retention window.

These two scripts are the pair I would show first if someone asked what good looks
like in this repository, and the reason is the split: **the one that analyses does
not execute, and the one that executes does not decide.**

`Get-LicenseReclamationPlan.ps1` reads everything and writes a CSV in which every
user is sorted into one of five tiers — KEEP, REVIEW, CONVERT_SHARED, RECLAIM,
EXCLUDE — with the signals that produced that verdict printed next to them. It
changes nothing. `Invoke-LicenseReclamation.ps1` reads that CSV and acts, but it
never reinterprets it: REVIEW, EXCLUDE and KEEP rows are not actionable no matter
what you pass on the command line.

The reason the boundary is drawn there is time. Reports get circulated, approved,
sat on, and executed three weeks later. In three weeks people come back from leave,
accounts get re-enabled, and licenses get reassigned. A script that trusted the CSV
would strip licenses from people who returned — and it would report success while
doing it. So the executing half re-checks the live state of every user before it
touches them, which is described in its own section below.

Tiering deliberately does not collapse "no evidence of activity" into "inactive".
EXCLUDE protects new joiners who have never signed in because they started last
Tuesday, and every caveat that could turn a blank into a false positive is printed
with the run rather than buried in the code.

```powershell
# 1. Analyse. Read-only. Start with a sample to see the shape of the output.
.\Licensing\Get-LicenseReclamationPlan.ps1 -SkuPartNumber 'ENTERPRISEPACK' -SampleSize 200

# Full run with Exchange enrichment (adds mailbox type and litigation hold)
.\Licensing\Get-LicenseReclamationPlan.ps1 -EnrichFromExchange
#    -> LicenseReclamation_<SKU>_<stamp>.csv
```

Every row carries its tier, the recommended action, the signals that fired and the
*lever* — whether the license is direct or granted by a group, and which group.
Tiers are rule-based and printed in full, so the output can be argued with rather
than merely trusted.

**Known caveats, stated in the script itself:** `signInActivity` lags 24-48 h and
has a finite retention window, so a null means "nothing in the retained window",
not always "never". The mailbox usage report is useless if the tenant conceals user
details in reports — the script detects this and warns. Without
`-EnrichFromExchange`, litigation hold is unknown and every `CONVERT_SHARED` row is
tagged "verify hold before converting".

Those three caveats are the whole reason this script is careful. The audit that
prompted the rewrite found a license report that declared 340 Visio and 179 Project
seats unused, in green, with no errors, because two optional parameters were missing
and the resulting empty reads were counted as zero usage. "Not measured" and "not
used" are different findings, and only one of them justifies removing a license from
somebody.

#### `Invoke-LicenseReclamation.ps1`

> **Warning: this is the one script here that removes things.** It removes licenses,
> removes users from licensing groups, and — with `-Tiers CONVERT_SHARED` — converts
> user mailboxes to shared. `ConfirmImpact` is `High`, so every individual change
> prompts unless you pass `-Confirm:$false`, and `-WhatIf` gives a complete dry run.

Removing a license is not a quiet operation. The user loses Office, and if the
license was granting the mailbox, the mailbox starts a deletion clock. Getting this
wrong for one person is a bad afternoon; getting it wrong in a batch of 300 is an
incident. Everything below exists because this script is executed against a decision
somebody else made, some time ago, from data that has since moved.

So the rule this script is built around is that it never re-decides anything and
never assumes the plan still holds. It reads a verdict somebody else approved, then
checks against live Graph state, immediately before each change, whether that verdict
still describes reality — and where it does not, the row is skipped and flagged
instead of forced through. Somebody who came back from leave keeps their license, and
the log says why they were skipped. Every safeguard below is a version of that one
rule.

It is also the best-built script in this folder, and the safeguards are the reason:

- **It re-validates live state before every action.** The CSV is a plan, not a
  source of truth. For each row it re-fetches the user from Graph and confirms they
  still hold the SKU (skip if already reclaimed), confirms a "disabled" row is
  *still* disabled (skip and flag if the account was re-enabled since the analysis),
  and resolves the current assigning groups from `licenseAssignmentStates` rather
  than trusting the group named in the CSV.
- **`-MaxChanges` caps the blast radius.** Run the first production batch with
  `-MaxChanges 50` and read the results before going further.
- **Bundle groups are detected and skipped.** Removing someone from a group strips
  *everything* that group assigns. If the assigning group grants any SKU besides the
  target, the row is skipped and flagged with the other SKUs named, instead of
  quietly costing the user their other licenses.
- **Conversion happens before removal, and only sticks if it worked.** For
  `CONVERT_SHARED` the mailbox is converted first, the resulting
  `RecipientTypeDetails` is verified, and the license is removed only if it reads
  `SharedMailbox`. A failed conversion leaves the license in place.
- **`-ApprovalColumn` supports per-row sign-off.** Add a column to the CSV, have an
  approver mark rows Yes, and only those are processed.
- **It verifies the outcome.** After a production run it re-reads the SKU and logs
  consumed-before versus consumed-after, so you can see the seats actually freed.

```powershell
# Full dry run
.\Licensing\Invoke-LicenseReclamation.ps1 -ReportPath .\LicenseReclamation_ENTERPRISEPACK_<stamp>.csv -WhatIf

# First production batch, RECLAIM tier only, capped at 50 users
.\Licensing\Invoke-LicenseReclamation.ps1 -ReportPath .\report.csv -MaxChanges 50

# Convert and reclaim, approved rows only (connect Exchange Online first)
.\Licensing\Invoke-LicenseReclamation.ps1 -ReportPath .\report.csv `
    -Tiers RECLAIM,CONVERT_SHARED -ApprovalColumn Approved
```

**Input:** the plan CSV. `REVIEW`, `EXCLUDE` and `KEEP` rows are never actioned,
whatever you pass. `CONVERT_SHARED` additionally requires a live Exchange Online
session; the script aborts up front if there isn't one.
**Output:** `Reclaim_<PROD|WHATIF>_<stamp>.csv` and a matching `.log`, one row per
user with what was attempted and what happened.
**Permissions:** `User.ReadWrite.All`, `Group.ReadWrite.All`, `Directory.Read.All`,
`Organization.Read.All`.

**Reversibility:** removing a license is reversible (re-add to the group, or
re-assign directly). Converting a mailbox to shared is reversible with
`Set-Mailbox -Type Regular`, but the mailbox then needs a license again.

## Known rough edges

Stated plainly, because pretending otherwise would be worse:

- **`Get-MfaExemption.ps1` still evaluates every enabled MFA-enforcing policy unless
  you scope it.** The exemption criterion now filters the output, so the 92x
  over-count described above is gone, but the *input* set is still every policy
  without `-PolicyName` / `-PolicyId`. It warns when more than five are in scope and
  writes a per-group breakdown. Scope it to the baseline policy when the number is
  going to be quoted.
- **The reason those block-policy rows were ever selected is not established.** The
  policy filter requires `builtInControls` to contain `mfa`, or an
  `authenticationStrength`, and the policies contributing the bulk of the inflated
  count had neither in the exports available. The narrowing fixes the output either
  way, because it filters on what an exemption *is* rather than on how the policy was
  selected — but if you are relying on this in anger, run it once against your own
  tenant and check the per-policy diagnostics rather than trusting that the input set
  is what you expect.
- **`Get-MfaExemption.ps1` does not expand excluded directory roles.** Only users
  and groups are expanded; a policy that excludes a role prints a warning and its
  role members are not counted.
- **`New-AdminAccount.ps1` has no `-WhatIf`.** It creates a privileged account after
  a single confirmation prompt. Read the proposal before typing `Y`.
- **`New-SecurityGroup.ps1` mail nickname derivation is naive.** It strips
  everything outside `[a-zA-Z0-9._-]` from the display name and does not check for a
  collision with an existing nickname.
- **The domain map ships with placeholder domains.** Any script that reads
  `domain-country-map.psd1` will resolve nothing useful until you replace them.
