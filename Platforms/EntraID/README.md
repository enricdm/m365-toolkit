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
| [`Export-SamlNotificationEmail.ps1`](Applications/Export-SamlNotificationEmail.ps1) | Lists SAML apps with their certificate expiry date and notification addresses | No | interactive |
| [`New-ClientApp.ps1`](Applications/New-ClientApp.ps1) | Bulk-creates client app registrations + service principals for cert-based M2M auth | **Yes** | interactive |
| [`Update-AppContact.ps1`](Applications/Update-AppContact.ps1) | Resolves a responsible human per app through five ranked signals, with a confidence score, and writes them into the tracking workbook | No (writes a file) | interactive |
| [`Update-FederatedCredential.ps1`](Applications/Update-FederatedCredential.ps1) | Replaces an app's GitHub OIDC credentials with one all-branches FFIC per repository | **Yes** (destructive) | interactive |
| **ConditionalAccess** | | | |
| [`Get-MfaExclusion.ps1`](ConditionalAccess/Get-MfaExclusion.ps1) | Every user excluded from MFA-enforcing CA policies, with nested groups expanded | No | app-only + cert, or interactive |
| [`Get-ExemptionSignInActivity.ps1`](ConditionalAccess/Get-ExemptionSignInActivity.ps1) | Adds 30-day sign-in telemetry to those exemptions and says which can be dropped or narrowed | No | app-only + cert, or interactive |
| [`Get-ExemptionUsageLocation.ps1`](ConditionalAccess/Get-ExemptionUsageLocation.ps1) | Looks up `usageLocation` for a list of UPNs | No | interactive |
| [`Export-ConditionalAccessPolicy.ps1`](ConditionalAccess/Export-ConditionalAccessPolicy.ps1) | Exports all CA policies with every GUID resolved to a display name | No | interactive |
| [`Export-VpnSignInLog.ps1`](ConditionalAccess/Export-VpnSignInLog.ps1) | Sign-ins for a VPN application plus Identity Protection risk detections and risky users | No | app-only + cert, or interactive |
| **Identity** | | | |
| [`Get-PasswordAge.ps1`](Identity/Get-PasswordAge.ps1) | Password age per enabled user in one country, flagged against a maximum-age policy | No | interactive |
| [`Get-MissingLocationReport.ps1`](Identity/Get-MissingLocationReport.ps1) | Categorises accounts with no `usageLocation` into eight actionable buckets | No | interactive |
| [`Export-CountryUserReport.ps1`](Identity/Export-CountryUserReport.ps1) | Country-scoped users/groups/licenses as a three-sheet workbook, delivered to SharePoint | **Yes** (uploads a file with `-Execute`) | managed identity |
| [`New-SecurityGroup.ps1`](Identity/New-SecurityGroup.ps1) | Creates security groups from a definition file and adds members | **Yes** (with `-Execute`) | interactive |
| [`New-AdminAccount.ps1`](Identity/New-AdminAccount.ps1) | Creates a cloud-only admin account derived from a person's ordinary account, with TAP and per-user MFA | **Yes** | interactive |
| **Licensing** | | | |
| [`Get-LicenseReclamationPlan.ps1`](Licensing/Get-LicenseReclamationPlan.ps1) | Tiers every holder of a SKU into KEEP / REVIEW / CONVERT_SHARED / RECLAIM / EXCLUDE with the reasoning shown | No | interactive |
| [`Invoke-LicenseReclamation.ps1`](Licensing/Invoke-LicenseReclamation.ps1) | Executes that plan: removes licenses, removes group membership, optionally converts mailboxes to shared | **Yes** | interactive (+ EXO) |
| **Devices** | | | |
| [`Remove-EntraDevice.ps1`](Devices/Remove-EntraDevice.ps1) | Removes device objects from the directory, from an explicit list or by staleness | **Yes** (destructive) | interactive |

## Requirements

**PowerShell**

- PowerShell 7.2 or later.
- `Microsoft.Graph` 2.x. Individual scripts need only some sub-modules —
  `.Authentication`, `.Users`, `.Groups`, `.Applications`, `.Identity.SignIns`,
  `.Reports` — and each script lists its own in the header.
- `ImportExcel` for the two scripts that read or write `.xlsx`
  (`Update-AppContact.ps1`, `Export-CountryUserReport.ps1`).
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

**Domain / country data**

`Update-AppContact.ps1` and `Get-MissingLocationReport.ps1` read their
domain-to-country mapping from a shared data file:

```
Platforms/_Shared/Data/domain-country-map.psd1
```

It ships with **placeholder `contoso.*` domains**. Replace them with your own
before relying on those scripts, or keep your own copy elsewhere and pass
`-DomainMapPath`. Nothing in the file is secret — it is a list of your public
e-mail domains and the country each one belongs to.

The file exists because this mapping previously lived in **six** places across
the estate with values that had drifted apart. One cost-allocation report ended
up treating `GB` (2,775 users) and `UK` (2) as two different countries.

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

#### `Export-SamlNotificationEmail.ps1`

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
.\Applications\Export-SamlNotificationEmail.ps1 -OutputPath .\Exports\saml.csv
```

The CSV it produces is also the Tier 1 input for `Update-AppContact.ps1` below —
a SAML notification address is the strongest available signal for who owns an app.

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

#### `Export-VpnSignInLog.ps1`

This one came out of a credential-exposure review of the VPN. The sign-in log on
its own answers a narrow question — this account signed in, from this IP, and it
worked. What it does not tell you is whether the credential that worked was already
known to be circulating.

That is why the script pulls Identity Protection risk detections and the current
risky-user list in the same run, over the same window, and writes them alongside
the sign-ins. Leaked credentials and password-spray detections only mean something
next to the sign-ins they correspond to; joining them afterwards from two separate
exports is where the analysis usually stalls. The app is discovered by display-name
search (`-AppFilter 'Forti'` by default) so it is not pinned to one vendor, and
`-AppId` skips discovery entirely when you already know the id.

The 30-day retention limit is real and stated up front: `-DaysBack 90` will not
return 90 days. For a longer window the sign-ins have to come from Log Analytics or
Sentinel instead, if they are archived there.

```powershell
# VPN sign-ins + Identity Protection risk, last 30 days
.\ConditionalAccess\Export-VpnSignInLog.ps1 -TenantId '<tenant-id>' -AppFilter 'Forti' -Interactive
```

#### `Get-MissingLocationReport.ps1`

`usageLocation` looks like a minor attribute and is not one. Without it you cannot
assign a license at all, and in a multi-country tenant it determines which services
are legally available to that user. Accounts arriving from different forests get it
populated inconsistently, so the gap is never a clean list.

The point of this script is that "missing" is not one problem. It sorts every
affected account into eight categories, because the answers are genuinely
different: a shared or room mailbox does not need one (B), a guest does not need one
(F), a licensed user mailbox without one is a real defect that must be fixed (C), an
account that is missing from Entra altogether is a stale record in the source export
and a different investigation entirely (X). Category A is the one that saves the
most time — the attribute *is* set in Entra and the source system simply had not
caught up, which is not a fix, it is a sync lag.

The Exchange Online dependency was removed deliberately: mailbox type is resolved
through Graph `mailboxSettings.userPurpose` instead, which avoids the WAM/MSAL
broker failures that make EXO connections unreliable inside Windows Terminal and VS
Code. The bulk remediation block at the bottom of the file is left commented out on
purpose.

```powershell
# Accounts with no usageLocation, sorted into eight buckets (A-G, X)
.\Identity\Get-MissingLocationReport.ps1 -CsvPath .\proofpoint_export.csv
```

#### `Get-PasswordAge.ps1`

A password-age report for a whole 40,000-user tenant is a spreadsheet nobody owns.
The people who chase stale passwords work per country, and the maximum age they are
measured against is not the same everywhere, so a single global list is both too
long to act on and wrong for part of its rows. This is scoped on `usageLocation`
server-side and takes the policy maximum as a parameter, which makes the output a
list one team can actually work through.

The `Status` column has three values, not two. A user whose
`lastPasswordChangeDateTime` is absent is reported as `Unknown`, never as `OK` —
that field is empty often enough on synchronised accounts that folding it into the
compliant bucket would quietly inflate the pass rate. `Over policy` means measured
and over; `OK` means measured and under; `Unknown` means the report could not tell
and somebody has to look.

```powershell
# Password age for enabled users in one country, against a policy maximum
.\Identity\Get-PasswordAge.ps1 -UsageLocation 'ES' -MaxAgeDays 90
```

**Input:** parameters only, except `Get-MissingLocationReport.ps1`, which needs a
directory-export CSV containing `Email`, `SSO ID`, `Location` columns.
**Output:** timestamped CSV/JSON under each script's `Exports` folder (or `-OutputPath`).
**Permissions:** the read-only scopes in the table above.

### The MFA exemption chain

Every MFA exemption was granted for a reason, and almost none of them are ever
revisited. The list grows, the original reason expires, and the exemption stays —
usually inside a group that somebody added a nested group to, so the number of
exempt people is larger than anyone believes. Asking "how many users are excluded
from MFA?" and reading the count of direct members off the CA policy blade gives an
answer that is always too low.

`Get-MfaExclusion.ps1` answers the first question properly: it expands excluded
groups transitively, including groups nested inside them, so the count is the real
population and not just the names that happen to be listed on the policy.

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
# 1. Who is excluded from MFA today? Scope it to the baseline policy, or the
#    number will include app-scoped exclusions and be meaninglessly large.
.\ConditionalAccess\Get-MfaExclusion.ps1 -TenantId '<tenant-id>' `
    -PolicyName '*Require MFA*' -MasterExceptionGroup 'CA-Exception-MFA'
#    -> Exports\CA-MFA-Exclusions_<stamp>.csv
#    -> Exports\CA-MFA-ExclusionSources_<stamp>.csv   (which group inflates the count)

# 2. Feed that CSV back in to see 30 days of sign-in behaviour per exempt user.
.\ConditionalAccess\Get-ExemptionSignInActivity.ps1 -TenantId '<tenant-id>' `
    -InputCsv .\ConditionalAccess\Exports\CA-MFA-Exclusions_<stamp>.csv -Detailed
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

`Get-ExemptionUsageLocation.ps1` does the Graph half — one lookup per UPN. UPNs it
cannot resolve are kept in the output with `Found = "No - <reason>"` rather than
dropped, so the row count going in matches the row count coming out and nothing
disappears from the list without saying why.

```powershell
# Resolve usageLocation for a UPN list, then classify each exception by country
.\ConditionalAccess\Get-ExemptionUsageLocation.ps1 -TenantId '<tenant-id>' -InputCsv .\users.csv
```

### `Update-AppContact.ps1`

The inventory from `Export-AppRegistration.ps1` produces a list of apps with
expiring credentials. The next question is always the same and always the hard one:
who do I email about this? The `owner` field is empty on a large share of apps, and
where it is populated it frequently points at an admin account with no mailbox, or
at somebody who left two reorganisations ago.

No single signal is reliable, so this does not pick one. It resolves a contact
through five ranked signals — SAML notification addresses, the owner's real
corporate mailbox, the owner's admin account resolved back to a human through
Graph, the country-team alias, the homepage TLD — falling through to a sixth tier
that simply means "nobody could be determined" rather than inventing an answer. Each
row records the source and a confidence score (95 down to 35) alongside the contact,
so a reviewer can see *why* a name was chosen and where to challenge it. A lower
tier that independently agrees on the same country adds five points, which is the
only place where two weak signals are allowed to reinforce each other.

The unresolved apps are written to their own `_manual_review.csv`. That file is the
actual deliverable of the run: it is the list of applications in the tenant that
nobody owns.

```powershell
# Dry run: full resolution, summary printed, nothing written
.\Applications\Update-AppContact.ps1 -ExcelPath .\AppTracking.xlsx `
    -SamlCsvPath .\SAML_Notification_Emails.csv -WhatIf

# Real run, all signals enabled
.\Applications\Update-AppContact.ps1 -ExcelPath .\AppTracking.xlsx `
    -SamlCsvPath .\SAML_Notification_Emails.csv -SpCsvPath .\servicePrincipals.csv `
    -DomainMapPath .\my-domains.psd1
```

**Input:** a tracking workbook with an `App Tracking` sheet, the CSV from
`Export-SamlNotificationEmail.ps1`, and optionally a service-principal export
(without it Tier 5 is skipped). Also `-AdminAccountDomain` and
`-CorporateDomainPattern`, which must match your tenant's naming.
**Output:** `<input>_enriched.xlsx` plus a `_manual_review.csv` listing everything
it could not resolve.
**Permissions:** `Directory.Read.All` (Global Reader is enough).

> **Note:** it writes only files, never the directory. `-WhatIf` skips the write;
> existing values in the contact column are preserved unless you pass `-Force`.

### `New-ClientApp.ps1`

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
.\Applications\New-ClientApp.ps1 -Name 'APP-Client-01','APP-Client-02' `
    -OwnerUpn 'admin@contoso.com'
```

**Input:** the names to create and an owner UPN that must already exist.
**Output:** a summary table of display name, client id, object id and SP object id.
**Permissions:** `Application.ReadWrite.All`, `User.Read.All`.

### `Update-FederatedCredential.ps1`

> **Warning: destructive.** It deletes **every** federated identity credential on the
> target app registration before recreating them. Anything not covered by
> `-Repository` will not come back.

GitHub Actions authenticating to Azure by OIDC is the right pattern — no secret is
stored anywhere — but the way the trust is usually configured makes it a recurring
chore. A subject-mode federated credential pins one exact `sub` claim, for example
`repo:org/repo:ref:refs/heads/main`. That means every new branch that needs to
deploy requires a new credential, added by hand, by someone with
`Application.ReadWrite.All`. With a 20-credential ceiling per app, a handful of
repositories with a few long-lived branches each hits the wall, and the workflow
failure that announces it looks like an authentication bug rather than a quota.

Flexible federated identity credentials (FFIC) use a `claimsMatchingExpression`
instead of a fixed subject, so a single credential matches `refs/heads/*` for a
whole repository. One credential per repository, and no further work when branches
come and go. That is the whole reason this script exists: it is a migration, run
once per app, that removes a class of repeated manual work permanently.

The trade is that a wildcard is broader than a pinned subject — any branch in that
repository can now assume the identity. That is the correct scope when the
repository is the trust boundary, which is the usual case; it is the wrong scope if
you were relying on the subject pin to stop non-`main` branches from deploying to
production, and in that case the pinned credential is what you want to keep.

Safeguards: the current credentials are exported to timestamped JSON *before*
anything is deleted, printed to screen, and the deletion only proceeds after you
type `yes` at the prompt. The final state is printed for verification.

```powershell
.\Applications\Update-FederatedCredential.ps1 -TenantId '<tenant-id>' -AppId '<client-id>' `
    -GitHubOrg 'contoso' -Repository 'platform-infra','platform-shared','platform-app'
```

**Input:** tenant, app id, GitHub organisation and the repository list.
**Output:** `Exports\federatedCreds-backup-<stamp>.json`, plus the new credentials.
**Permissions:** `Application.ReadWrite.All`.

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

### `Export-CountryUserReport.ps1`

> **Warning: uploads a file to SharePoint** when `-Execute` is passed. It makes no
> directory changes. Without `-Execute` it builds the workbook locally and prints
> the destination it would have used.

Country IT leads keep asking the same question — who are my users, which groups are
they in, what are they licensed for — and the honest answer is that it changes every
week. Answering it by hand once is fine; answering it monthly for every country is
the definition of work that should not involve a person. So this is written to run
unattended and deliver the result where the people who asked already look.

Two decisions follow from "unattended". There is no credential to store or rotate,
because it runs as a managed identity. And `-Execute` gates only the upload, never
the build — a dry run still produces the complete workbook locally, so the content
can be checked before it lands on a site other people read. The rewrite from CSV to
`.xlsx` came from the same practical direction: non-ASCII surnames plus whatever
encoding Excel decided to assume produced a report several country teams could not
read at all.

It is written as an Azure Automation runbook authenticating with the Automation
Account's system-assigned managed identity, so no secret is stored anywhere. The
output is a genuine three-sheet `.xlsx` (Users / GroupMemberships / Licenses),
normalised one row per user-group and user-license pair so it can be filtered and
pivoted — and, incidentally, immune to the UTF-8 mojibake that the earlier CSV
version suffered with non-ASCII names.

```powershell
# Dry run
.\Identity\Export-CountryUserReport.ps1 -FilterValue 'ES'

# Build and upload
.\Identity\Export-CountryUserReport.ps1 -FilterValue 'ES' `
    -SpHostName 'contoso.sharepoint.com' -SpSitePath '/sites/Example' -Execute
```

**Input:** `-FilterStrategy` (UsageLocation | Country | Group | Domain) and its
value; the SharePoint destination.
**Output:** `<prefix>_<yyyy-MM-dd>.xlsx` in the target library.
**Permissions:** managed-identity app roles `User.Read.All`, `Group.Read.All`,
`Sites.Selected` (with write granted on the target site).

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

### `Devices/Remove-EntraDevice.ps1`

Deleting the Entra device object is the last step of decommissioning, after the Intune record has
gone. It is also the step people get wrong, because an Entra device object and an Intune managed
device are two different things: the Intune record is the MDM enrollment, the Entra object is the
directory identity that Conditional Access, device-based licensing, BitLocker key escrow and
Windows Hello for Business all evaluate against.

```powershell
# What would be removed, by staleness. Deletes nothing.
.\Remove-EntraDevice.ps1 -StaleDays 180

# Remove a reviewed list
.\Remove-EntraDevice.ps1 -InputCsv .\decommissioned.csv -Execute
```

**Input:** `-DeviceId`, `-InputCsv` (a `DeviceId` column, object ID or `deviceId`), or `-StaleDays`.
**Output:** one CSV row per device with `WOULD-DELETE`, `DELETED`, `SKIPPED`, `NOT-FOUND`,
`NO-SIGNIN-DATA` or `FAILED`.
**Permissions:** `Device.ReadWrite.All`.

> **Order matters.** Retire or delete in Intune **first**, then remove the Entra object. Removing
> the Entra object while the device is still enrolled and in use leaves an orphan that will usually
> re-register on next sign-in — so the cleanup looks like it failed, when in fact it was done in the
> wrong order.
>
> **BitLocker recovery keys are escrowed against the Entra device object.** Deleting it can make
> those keys unrecoverable. If the hardware is being reused rather than destroyed, export the keys
> first. The script warns but cannot check for you.
>
> `-StaleDays` rejects values below 90 — an Entra device object going quiet for a few weeks is
> normal — and `-MaxDevices` defaults to 25. Objects with no `approximateLastSignInDateTime` at all
> are reported as `NO-SIGNIN-DATA` and never acted on: no timestamp is not an old timestamp.

## Known rough edges

Stated plainly, because pretending otherwise would be worse:

- **`Get-MfaExclusion.ps1` over-counts by default.** With no `-PolicyName` /
  `-PolicyId` it evaluates every enabled MFA-enforcing policy, so being excluded
  from one app-scoped policy counts as an exclusion. It warns when more than five
  policies are in scope and writes a per-group breakdown, but the headline number is
  only meaningful once you scope it to the baseline policy.
- **`Get-MfaExclusion.ps1` does not expand excluded directory roles.** Only users
  and groups are expanded; a policy that excludes a role prints a warning and its
  role members are not counted.
- **`New-AdminAccount.ps1` has no `-WhatIf`.** It creates a privileged account after
  a single confirmation prompt. Read the proposal before typing `Y`.
- **`Update-AppContact.ps1` Tier 4/5 are inference, not fact.** A country-team alias
  or a homepage TLD is a hint; the confidence score (55 and 35) says so. Treat
  anything below 75 as a suggestion to verify.
- **`New-SecurityGroup.ps1` mail nickname derivation is naive.** It strips
  everything outside `[a-zA-Z0-9._-]` from the display name and does not check for a
  collision with an existing nickname.
- **The domain map ships with placeholder domains.** Any script that reads
  `domain-country-map.psd1` will resolve nothing useful until you replace them.