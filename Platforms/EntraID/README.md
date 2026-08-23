# Entra ID

Operational tooling for Microsoft Entra ID: app-registration hygiene, Conditional
Access exception analysis, identity lifecycle and license reclamation. Everything
here talks to Microsoft Graph and is written to run against *any* tenant — tenant
ids, domains, group names and repository names are parameters, never constants.

Most scripts are read-only and produce CSV/XLSX you can hand to someone else. The
five that change directory state are called out explicitly below and each one has
its own section.

## Index

| Script | What it does | Modifies state | Auth |
|---|---|:---:|---|
| **Applications** | | | |
| [`Export-AppRegistration.ps1`](Applications/Export-AppRegistration.ps1) | Full app-registration inventory: owners, credentials, API permissions, sign-in activity, hygiene flags | No | interactive |
| [`Export-SamlNotificationEmail.ps1`](Applications/Export-SamlNotificationEmail.ps1) | Lists SAML apps with their certificate expiry date and notification addresses | No | interactive |
| [`New-ClientApp.ps1`](Applications/New-ClientApp.ps1) | Bulk-creates client app registrations + service principals for cert-based M2M auth | **Yes** | interactive |
| [`Update-AppContact.ps1`](Applications/Update-AppContact.ps1) | Resolves a responsible human per app through six ranked signals and writes them into the tracking workbook | No (writes a file) | interactive |
| [`Update-FederatedCredential.ps1`](Applications/Update-FederatedCredential.ps1) | Replaces an app's GitHub OIDC credentials with one all-branches FFIC per repository | **Yes** (destructive) | interactive |
| **ConditionalAccess** | | | |
| [`Get-MfaExclusion.ps1`](ConditionalAccess/Get-MfaExclusion.ps1) | Every user excluded from MFA-enforcing CA policies, with nested groups expanded | No | app-only + cert, or interactive |
| [`Get-ExemptionSignInActivity.ps1`](ConditionalAccess/Get-ExemptionSignInActivity.ps1) | Adds 30-day sign-in telemetry to those exemptions and says which can be dropped or narrowed | No | app-only + cert, or interactive |
| [`Get-ExemptionUsageLocation.ps1`](ConditionalAccess/Get-ExemptionUsageLocation.ps1) | Looks up `usageLocation` for a list of UPNs | No | interactive |
| [`classify_exceptions_by_country.py`](ConditionalAccess/classify_exceptions_by_country.py) | Assigns a country to each MFA exception by cross-checking UPN domain, sign-in geography and usageLocation | No | none (offline) |
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

## Requirements

**PowerShell**

- PowerShell 7.2 or later.
- `Microsoft.Graph` 2.x. Individual scripts need only some sub-modules —
  `.Authentication`, `.Users`, `.Groups`, `.Applications`, `.Identity.SignIns`,
  `.Reports` — and each script lists its own in the header.
- `ImportExcel` for the three scripts that read or write `.xlsx`
  (`Update-AppContact.ps1`, `Export-CountryUserReport.ps1`).
- `ExchangeOnlineManagement` only for `Invoke-LicenseReclamation.ps1 -Tiers CONVERT_SHARED`
  and `Get-LicenseReclamationPlan.ps1 -EnrichFromExchange`.

**Python** (one script)

- Python 3.9+, `pandas` and `openpyxl`.

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

`Update-AppContact.ps1`, `Get-MissingLocationReport.ps1` and
`classify_exceptions_by_country.py` all read their domain-to-country mapping from
a shared data file:

```
Platforms/_Shared/Data/domain-country-map.psd1
```

It ships with **placeholder `contoso.*` domains**. Replace them with your own
before relying on those scripts, or keep your own copy elsewhere and pass
`-DomainMapPath` (PowerShell) / `--domain-map-path` (Python). Nothing in the file
is secret — it is a list of your public e-mail domains and the country each one
belongs to.

## Usage

### Read-only reporting

```powershell
# Everything about every app registration, with hygiene flags for Excel filtering
.\Applications\Export-AppRegistration.ps1 -TenantId '<tenant-id>' -OutputPath .\Exports

# SAML signing certificates and who gets warned when they expire
.\Applications\Export-SamlNotificationEmail.ps1 -OutputPath .\Exports\saml.csv

# All CA policies with GUIDs resolved: raw JSON, resolved JSON, flat CSV, named locations
.\ConditionalAccess\Export-ConditionalAccessPolicy.ps1 -TenantId '<tenant-id>'

# VPN sign-ins + Identity Protection risk, last 30 days
.\ConditionalAccess\Export-VpnSignInLog.ps1 -TenantId '<tenant-id>' -AppFilter 'Forti' -Interactive

# Accounts with no usageLocation, sorted into eight buckets (A-G, X)
.\Identity\Get-MissingLocationReport.ps1 -CsvPath .\proofpoint_export.csv

# Password age for enabled users in one country, against a policy maximum
.\Identity\Get-PasswordAge.ps1 -UsageLocation 'ES' -MaxAgeDays 90
```

**Input:** parameters only, except `Get-MissingLocationReport.ps1`, which needs a
directory-export CSV containing `Email`, `SSO ID`, `Location` columns.
**Output:** timestamped CSV/JSON under each script's `Exports` folder (or `-OutputPath`).
**Permissions:** the read-only scopes in the table above.

### The MFA exemption chain

The two exemption scripts are meant to be run one after the other — the second
consumes the CSV the first produces. That is the whole point: the first says *who*
is exempt, the second says *whether the exemption is still deserved*.

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

Country attribution for the same population is a separate, offline step:

```powershell
# Resolve usageLocation for a UPN list, then classify each exception by country
.\ConditionalAccess\Get-ExemptionUsageLocation.ps1 -TenantId '<tenant-id>' -InputCsv .\users.csv
```

```bash
python ConditionalAccess/classify_exceptions_by_country.py exceptions.xlsx \
    --usage-location UsageLocation_<stamp>.csv \
    --domain-map-path ../_Shared/Data/domain-country-map.psd1
```

The Python script cross-checks three independent signals (UPN domain, sign-in
geography, usageLocation) rather than trusting `usageLocation` alone, because many
exceptions are service or shared accounts where it was never set. Rows it cannot
settle land on a `Needs Review` sheet instead of being guessed.

### `Update-AppContact.ps1`

Answers "who owns this app registration?" for a whole tenant. It resolves a
contact through six ranked signals — SAML notification addresses, the owner's real
mailbox, the owner's admin account resolved back to a human, the country-team
alias, the homepage TLD — and records both the source and a confidence score
alongside each contact, so a reviewer can see *why* a name was chosen and where to
challenge it.

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

> **Aviso:** it writes only files, never the directory. `-WhatIf` skips the write;
> existing values in the contact column are preserved unless you pass `-Force`.

### `New-ClientApp.ps1`

> **Aviso: creates directory objects.** One app registration plus one service
> principal per name, with the supplied owner attached to both.

Safeguards: an existing application with the same `displayName` is skipped with a
warning rather than duplicated, so re-running after a partial failure is safe. No
credentials are created — the owner uploads their own certificate afterwards,
which keeps secrets out of this script entirely.

```powershell
.\Applications\New-ClientApp.ps1 -Name 'APP-Client-01','APP-Client-02' `
    -OwnerUpn 'admin@contoso.com'
```

**Input:** the names to create and an owner UPN that must already exist.
**Output:** a summary table of display name, client id, object id and SP object id.
**Permissions:** `Application.ReadWrite.All`, `User.Read.All`.

### `Update-FederatedCredential.ps1`

> **Aviso: destructive.** It deletes **every** federated identity credential on the
> target app registration before recreating them. Anything not covered by
> `-Repository` will not come back.

Why it exists: subject-mode federated credentials pin one exact `sub` claim, so a
repository needs a separate credential per branch and the 20-credential ceiling
arrives fast. Flexible federated identity credentials (FFIC) use a claims-matching
expression instead, so one credential covers `refs/heads/*` for a whole repository.

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

> **Aviso: creates groups and adds members** — but only with `-Execute`. Without it
> the script previews every action and touches nothing.

Idempotent: an existing group with the same display name is reused rather than
duplicated, members already present are skipped, and UPNs that do not resolve are
reported as `NOT FOUND` instead of aborting the run.

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

> **Aviso: creates a privileged account.** It creates a cloud-only admin user, sets
> its manager and `ExtensionAttribute1`, issues a Temporary Access Pass, enforces
> per-user MFA and optionally adds it to a Conditional Access group. There is no
> `-WhatIf`; the safeguard is an explicit confirmation prompt showing the exact
> account it is about to create.

You supply only the person's ordinary account. Everything else — UPN, display
name, mail nickname, usage location, manager — is derived from it, which is what
keeps admin accounts consistently named and reliably linked back to a human.

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
- **Credentials are not printed by default.** The temporary password and the TAP
  are masked unless you pass `-ShowCredentials`, so an ordinary run leaves nothing
  usable in console scrollback or a PowerShell transcript. The confirmation prompt
  warns you when the switch is missing, before the account is created.

PIM eligible role assignment stays manual and out of scope; pass the role you
assigned in `-AssignedRoles` so it shows up in the end-user message.

### `Export-CountryUserReport.ps1`

> **Aviso: uploads a file to SharePoint** when `-Execute` is passed. It makes no
> directory changes. Without `-Execute` it builds the workbook locally and prints
> the destination it would have used.

Written as an Azure Automation runbook authenticating with the Automation
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

Two scripts, deliberately split. The first only looks; the second only acts on what
the first wrote, and re-checks everything before it does.

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

#### `Invoke-LicenseReclamation.ps1`

> **Aviso: this is the one script here that removes things.** It removes licenses,
> removes users from licensing groups, and — with `-Tiers CONVERT_SHARED` — converts
> user mailboxes to shared. `ConfirmImpact` is `High`, so every individual change
> prompts unless you pass `-Confirm:$false`, and `-WhatIf` gives a complete dry run.

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
