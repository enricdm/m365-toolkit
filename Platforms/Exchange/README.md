# Exchange Online

Operational PowerShell for Exchange Online recipient management: distribution and
security groups, shared mailboxes and calendars, mail flow protection, and room
resources. These are the scripts behind everyday service-desk requests — "create
this fax list", "who can book this room", "block this domain" — written so the
work is repeatable, previewable, and leaves an artefact behind.

Two conventions run through all of them:

- **Dry run is the default.** Every script that writes previews what it would do
  and changes nothing until you pass `-Execute`.
- **Everything exports.** Results go to a timestamped CSV under `Exports/` (or
  `Evidence/`, `logs/`) next to the script, so there is a record of what was done
  and to whom.

Example values are placeholders (`contoso.com`, `<tenant-id>`, `<client-id>`).
Nothing needs editing inside a file to run it in your own tenant.

## Index

| Script | What it does | Changes state | Auth |
|---|---|:---:|---|
| [`Groups/Edit-MailGroupMember.ps1`](Groups/Edit-MailGroupMember.ps1) | Adds/removes members on existing groups; auto-creates MailContacts for external addresses | **Yes** | interactive (pre-existing EXO session) |
| [`Groups/New-FaxDistributionList.ps1`](Groups/New-FaxDistributionList.ps1) | Provisions a fax-to-mail DL with unauthenticated relay allowed, and verifies it | **Yes** | interactive (pre-existing EXO session) |
| [`Groups/New-MailGroup.ps1`](Groups/New-MailGroup.ps1) | Creates a DL, mail-enabled security group, or M365 group from a member CSV | **Yes** | interactive (+ Graph for M365 groups) |
| [`Mailboxes/New-SharedCalendar.ps1`](Mailboxes/New-SharedCalendar.ps1) | Provisions a shared team calendar with group-based Editor access | **Yes** | app-only cert **or** interactive |
| [`Mailboxes/New-SharedMailbox.ps1`](Mailboxes/New-SharedMailbox.ps1) | Creates a shared mailbox and grants FullAccess + SendAs from a CSV | **Yes** | interactive (pre-existing EXO session) |
| [`Mailboxes/Set-MailboxForwarding.ps1`](Mailboxes/Set-MailboxForwarding.ps1) | Sets server-side forwarding on one or more mailboxes | **Yes** | interactive (pre-existing EXO session) |
| [`MailFlow/Block-MaliciousDomain.ps1`](MailFlow/Block-MaliciousDomain.ps1) | Blocks a domain in the Tenant Allow/Block List as Sender and URL | **Yes** | interactive (reuses an existing session) |
| [`MailFlow/Test-PhishingSimulationRule.ps1`](MailFlow/Test-PhishingSimulationRule.ps1) | Traces both legs of a phishing-report transport rule and writes an evidence report | No | interactive |
| [`MailFlow/Get-MailboxReceiveVolume.ps1`](MailFlow/Get-MailboxReceiveVolume.ps1) | Splits mailboxes above/below a mail-volume threshold for per-mailbox licensing | No | Graph, `Reports.Read.All` |
| [`Resources/Get-CountryResource.ps1`](Resources/Get-CountryResource.ps1) | Finds room/equipment mailboxes for a country across several weak signals | No | app-only cert **or** interactive |
| [`Resources/Set-RoomMailbox.ps1`](Resources/Set-RoomMailbox.ps1) | Configures a room mailbox: booking policy, place metadata, who may reserve it | **Yes** | interactive |

## Requirements

- **PowerShell 7.0+** — required by `Get-CountryResource.ps1` and
  `Set-RoomMailbox.ps1`; recommended for the rest.
- **ExchangeOnlineManagement 3.0+** (`Install-Module ExchangeOnlineManagement`).
  `Test-PhishingSimulationRule.ps1` needs **3.7+** for `Get-MessageTraceV2`.
- **Microsoft.Graph.Reports** — only for `Get-MailboxReceiveVolume.ps1`
  (`Install-Module Microsoft.Graph.Reports -Scope CurrentUser`).

Roles and permissions:

| Need | Used by |
|---|---|
| Exchange **Recipient Administrator** | the group, shared mailbox and forwarding scripts |
| Exchange **Administrator** | `New-SharedCalendar.ps1`, `Set-RoomMailbox.ps1` (`New-Mailbox`, `Set-CalendarProcessing`, `Set-Place`) |
| **Security Administrator** (or Organization Management) | `Block-MaliciousDomain.ps1` — Tenant Allow/Block List writes |
| **View-Only Recipients** / message trace read | the two read-only MailFlow and Resources scripts |
| Graph `Reports.Read.All` (admin consent) | `Get-MailboxReceiveVolume.ps1` |

### App-only certificate auth

`New-SharedCalendar.ps1` and `Get-CountryResource.ps1` can run unattended. Register
an app in Entra ID, give it the `Exchange.ManageAsApp` application permission, grant
it the **Exchange Administrator** directory role, and upload a certificate. Then
either pass the three values or set the environment variables the parameters
default to:

```powershell
$env:EXO_APPID     = '<client-id>'
$env:EXO_CERTTHUMB = '<cert-thumbprint>'
$env:EXO_ORG       = 'contoso.onmicrosoft.com'
```

Both scripts require **all three** or **none**. A partial configuration is rejected
rather than silently falling back to interactive — that is how you end up believing
a scheduled run was unattended when it was not. Every other script authenticates
interactively with the signed-in account.

---

## `Groups/`

Group creation is the most common write a service desk makes in Exchange Online,
and it is where the platform is least honest about timing. `New-DistributionGroup`
returns, the group exists, and the next cmdlet cannot see it yet. Add members
immediately and the call fails intermittently, with an error that reads like a
permissions problem and sends you looking in the wrong place. The `Start-Sleep`
in step 3 below is not superstition; it is the only handle any of these scripts
has on directory replication.

The three group scripts share one pattern, learned the hard way:

1. **Validate before writing** — resolve every member and check the address is not
   already taken, because a half-created group is worse than no group.
2. **Create the group.**
3. **Wait for directory replication** (`Start-Sleep`) before touching it again.
   The follow-up cmdlet will fail against a group that exists but has not
   propagated yet.
4. **Add members with `-BypassSecurityGroupManagerCheck`.** Without it, an admin
   who is not listed in `ManagedBy` cannot modify the group they just created.
5. **Verify** the resulting object against what was asked for.

[`New-FaxDistributionList.ps1`](Groups/New-FaxDistributionList.ps1) is the
**reference implementation** of that pattern — it is the only one that does all
five steps, including an exportable preflight and a post-provisioning check of the
live object. Read that one first; the other two are lighter variants.

### `Groups/New-FaxDistributionList.ps1`

Fax has not gone away in a tenant this size; it has moved behind a gateway that
receives the fax and emails it on to a list of people. The handoff is where it
breaks, and it breaks quietly: the list is created correctly, the gateway
reports success, and nothing ever arrives.

The fix — `RequireSenderAuthenticationEnabled = $false` — is worth being plain
about, because it means the address now accepts mail from anyone on the
internet. That is a justified exception for a fax gateway, not a free one, and
it is real spam surface, so the script sets it explicitly rather than as a side
effect and says so. This is also the only one of the three group scripts that
checks for an address collision *before* it writes anything, which is why it is
the reference implementation above and why the other two are the ones with
catching up to do.

Provisions a fax-to-mail distribution list. Fax gateways are the awkward case:
they relay **unauthenticated**, so the Exchange default
`RequireSenderAuthenticationEnabled = $true` silently drops every inbound fax with
no NDR the requester ever sees. This script disables that explicitly and refuses to
finish if the live object comes back with a different value than was asked for.

It also checks address collisions before creating anything, resolves every member
in a preflight pass exported to CSV (so the requester can confirm the list before
it exists), creates MailContacts for external members, and diffs the final
membership against what was requested.

```powershell
# Preview: resolves owner and every member, writes the preflight CSV, changes nothing
.\Groups\New-FaxDistributionList.ps1 -GroupName 'DE.CTY.Example-Fax' `
    -DisplayName 'Fax, Example Site' -PrimarySmtp 'Example-Fax@contoso.com' `
    -Owner 'owner@contoso.com' -Members 'first.user@contoso.com','second.user@contoso.com'

# Apply, with a legacy alias a pre-staged gateway still delivers to
.\Groups\New-FaxDistributionList.ps1 -GroupName 'DE.CTY.Example-Fax' `
    -DisplayName 'Fax, Example Site' -PrimarySmtp 'Example-Fax@contoso.com' `
    -Owner 'owner@contoso.com' -AliasAddresses 'Example.Fax@contoso.com' `
    -Reference 'CHG-0001' -Execute
```

**Input:** parameters only. An existing `Connect-ExchangeOnline` session.
**Output:** `Exports/FaxDL-Preflight-*.csv` on a dry run; `FaxDL-Object-*.csv` and
`FaxDL-Members-*.csv` after `-Execute`.
**Permissions:** Exchange Recipient Administrator.

> **Writes:** creates a distribution group, MailContacts for external members, and
> adds members. Safeguards: dry run by default; aborts if the primary SMTP, any
> alias, or the group name is already in use; aborts if the owner does not resolve;
> post-run verification fails loudly if `RequireSenderAuthenticationEnabled` does
> not match what was requested. Rollback is a single
> `Remove-DistributionGroup -Identity '<GroupName>' -Confirm:$false`.
>
> **Known wart:** the final "missing members" diff compares the addresses you
> passed against each member's *primary* SMTP. Request someone by an alias and they
> are reported `MISSING` even though they were added correctly.

### `Groups/New-MailGroup.ps1`

"Create a group for this team" sounds like one request, but it is three
different objects and the choice is expensive to reverse. A distribution list
delivers mail and grants nothing. A mail-enabled security group can also be used
to grant access — which is what a shared calendar or a document library needs.
A Microsoft 365 group brings a mailbox, a SharePoint site and a Teams surface
along with it, whether or not anyone asked for those. Once people have started
sending to the address, changing your mind is a migration rather than an edit,
so the type is asked for up front instead of inferred.

The known defects listed below are worth reading before running it. There is no
pre-write collision check, and the type comes from a `Read-Host`, which rules
out scheduling it. They are documented rather than quietly patched because they
are the concrete evidence for the consolidation argument at the end of this
file: the same fix exists one directory over and never travelled.

One script for the three group types — distribution list, mail-enabled security
group, or Microsoft 365 group — chosen from a prompt at runtime. External addresses
get a MailContact created first (DL and SG only; M365 groups do not take them), and
sender restrictions are applied after the members are in.

```powershell
# Preview from a member CSV (column 'Email' or 'Member')
.\Groups\New-MailGroup.ps1 -GroupName 'CC.CITY.Example' `
    -PrimarySmtp 'CC.CITY.Example@contoso.com' -Owner 'owner@contoso.com' `
    -CsvPath .\members.csv

# Apply, restricting who may send to the group
.\Groups\New-MailGroup.ps1 -GroupName 'CC.CITY.Example' `
    -PrimarySmtp 'CC.CITY.Example@contoso.com' -Owner 'owner@contoso.com' `
    -CsvPath .\members.csv -AllowedSenders 'owner@contoso.com' -Execute
```

**Input:** a one-column CSV of addresses; an existing EXO session; for M365 groups
also `Connect-MgGraph -Scopes "Group.ReadWrite.All"`.
**Output:** console only — verification listing of the group and its members.
**Permissions:** Exchange Recipient Administrator.

> **Writes:** creates a group, creates MailContacts, adds members, applies sender
> restrictions. Safeguards: dry run by default; existing contacts and guest/mail
> users are reused instead of duplicated; per-member failures are reported and do
> not stop the run.
>
> **Known defects, not fixed:** the group type comes from an interactive
> `Read-Host`, so the script cannot run unattended. It does **not** check for an
> address collision before creating (unlike `New-FaxDistributionList.ps1`). After
> creation it sleeps a fixed 10 seconds and hopes replication has caught up.
> `Ensure-MailContact` uses an unapproved PowerShell verb.

### `Groups/Edit-MailGroupMember.ps1`

Most group work is not creation. It is the steady stream of "add these four,
remove these two" that follows every reorganisation — and by the time it reaches
you, somebody has usually done half of it by hand already. A script that treats
"remove a person who is not a member" as an error is useless for that: it throws
on the second job and leaves the run half applied, which is the worst state to
hand back.

So absence counts as success here, and the run is built to be repeatable. The
other decision worth naming is that an internal address which does not resolve
is reported as a probable typo and skipped, rather than having a MailContact
created for it: auto-creating a contact for `jonh.smith@contoso.com` would make
the typo permanent and quietly route that person's mail out of the tenant.

Membership changes on groups that already exist — no creation, no sender-restriction
logic. Group type is auto-detected per group (`Add/Remove-DistributionGroupMember`
vs `Add/Remove-UnifiedGroupLinks`), external addresses that do not exist yet get a
MailContact, and removing someone who is already absent is treated as success, so
the script is safe to re-run and tolerates changes the requester already made by
hand.

Edit the `$Jobs` list at the top of the file: one entry per group, each with its
own `Add` and `Remove` sets.

```powershell
# Preview every job
.\Groups\Edit-MailGroupMember.ps1

# Apply, treating two domains as internal (no MailContact auto-creation for them)
.\Groups\Edit-MailGroupMember.ps1 -InternalDomains 'contoso.com','contoso.co.uk' -Execute
```

**Input:** the `$Jobs` list inside the script; an existing EXO session.
**Output:** `logs/Edit-MailGroupMember-<timestamp>.csv` — one row per action, always
written, including on a dry run.
**Permissions:** Exchange Recipient Administrator.

> **Writes:** adds and removes group members, and creates MailContacts for unknown
> external addresses. Safeguards: dry run by default; current membership is read
> first so no-op adds and removes are skipped rather than attempted; an internal
> address that does not resolve is reported as a probable typo and skipped instead
> of having a contact created for it; a group that cannot be resolved is skipped,
> not fatal.
>
> **Known wart:** the job list lives in the file rather than in a parameter or CSV,
> so running it means editing it.

---

## `Mailboxes/`

### `Mailboxes/New-SharedCalendar.ps1`

A team asks for a shared calendar, and what they usually end up with is a mailbox
whose calendar folder carries a dozen individual permissions — one per person,
stamped by hand, never removed when somebody moves on. Six months later nobody
can say who has access without dumping the folder ACL, and the answer is always
"more people than you think".

Two things here are wider in scope than the script's name suggests, and both are
deliberate. The `Default` permission on the calendar folder is set to
`AvailabilityOnly`, and `Default` means the whole organisation, not the team:
free/busy visible to everyone, contents only to the access group. And the alias
and group naming (`shared.<CC>.<CITY>.<name>`, `MB_<alias>_Kalender_Editor`)
follows one subsidiary's convention — it is a convention, not a requirement, so
adapt it before running this in a tenant that names things differently.

Provisions a shared team calendar. The access model is the point: instead of
stamping folder permissions per person, one mail-enabled **security** group
(`MB_<alias>_Kalender_Editor`) is granted Editor once on the calendar folder, and
people are managed as members of that group. Day-to-day membership changes then
never touch the mailbox.

It resolves the real calendar folder name rather than assuming `\Calendar` — a
mailbox with German regional settings has `\Kalender`, and that is the single most
common reason a calendar permission script fails. It also verifies the access group
is actually security-enabled, because folder permissions granted to a plain
distribution list reach nobody.

```powershell
# Preview, and dump a known-good calendar's config to compare against
.\Mailboxes\New-SharedCalendar.ps1 -ReferenceMailbox shared.DE.CTY.Other-Team@contoso.com

# Apply, cloud-only, stamping each person directly instead of via a group
.\Mailboxes\New-SharedCalendar.ps1 -Alias 'shared.DE.CTY.Example-Team' `
    -DisplayName 'DE.CTY.SHARED Example-Team' `
    -Members 'first.user@contoso.com','second.user@contoso.com' `
    -CreateMailbox -DirectGrant -Execute
```

**Input:** parameters. App-only credentials or an interactive sign-in.
**Output:** `Exports/SharedCalendar_<alias>_<timestamp>.csv` — every action with its
status, plus the resulting folder permissions on screen.
**Permissions:** Exchange Administrator.

> **Writes:** optionally creates the shared mailbox, sets regional configuration,
> adds members to the access group, and sets calendar folder permissions (Default =
> `AvailabilityOnly`, the group or each user = Editor). Optionally grants
> FullAccess with AutoMapping. Safeguards: dry run by default; every action goes
> through one wrapper that logs a `DRY-RUN`/`OK`/`FAIL` status; refuses to use an
> access group that is not security-enabled and says why; warns when the group is
> directory-synced and membership must be changed on-prem instead.
>
> **Known limitation:** mail-enabled security groups can no longer be created in
> Exchange Online, so `-CreateAccessGroup` will usually fail. In a hybrid tenant,
> create the group in on-prem AD and let it sync; in a cloud-only tenant, use
> `-DirectGrant`.

### `Mailboxes/New-SharedMailbox.ps1`

A shared mailbox is the standard answer to "we need a team address", and the
provisioning is trivial. The access grants are where the tickets come from. Two
separate rights are involved and requesters conflate them: FullAccess lets you
read the mailbox, SendAs lets you send from it, and someone who asks for "access"
means both. Grant one and not the other and you get a team that can read the
inbox but cannot answer from it, which comes straight back as a second ticket.
So the two are granted independently and one failing does not hide the other.

The other recurring surprise is the partner address in the member CSV, which
cannot be given access at all — and the useful outcome there is the eight
internal people provisioned plus a clear line about the one that could not be,
rather than a failed run.

Creates a shared mailbox and grants FullAccess + SendAs to everyone in a CSV.
External addresses are skipped for mailbox permissions with an explanation rather
than failing silently — Exchange cannot grant mailbox rights to a non-tenant
identity, and that surprises requesters every time.

```powershell
# Preview
.\Mailboxes\New-SharedMailbox.ps1 -MailboxName 'Team.Inbox' `
    -PrimarySmtp 'team.inbox@contoso.com' -CsvPath .\members.csv

# Apply without AutoMapping, so users add the mailbox themselves
.\Mailboxes\New-SharedMailbox.ps1 -MailboxName 'Team.Inbox' `
    -DisplayName 'Team Inbox' -PrimarySmtp 'team.inbox@contoso.com' `
    -CsvPath .\members.csv -AutoMapping $false -Execute
```

**Input:** a one-column CSV (`Email` or `Member`); an existing EXO session.
**Output:** console summary of granted, skipped and failed, plus a verification
listing of the resulting permissions.
**Permissions:** Exchange Recipient Administrator.

> **Writes:** creates a shared mailbox and grants FullAccess + SendAs. Safeguards:
> dry run by default; skips creation if the mailbox already exists, so a re-run only
> tops up permissions; FullAccess and SendAs are granted independently so one
> failing does not hide the other; external addresses are skipped with a warning.

### `Mailboxes/Set-MailboxForwarding.ps1`

Sets server-side forwarding on one or more mailboxes.

The request behind this is ordinary — somebody leaves and their mail has to reach
whoever is covering the role. What makes it unusual is that an attacker who
compromises a mailbox does exactly the same thing, with the same cmdlet, to read
someone's mail without ever signing in again. Automating this means automating
both, so the design assumes misuse as well as use: the previous value of every
setting it touches is recorded first, and the external-target guard below was
added afterwards, not shipped with the original.

> ### ⚠ Read this before using it
>
> Server-side forwarding to an external address is **the exact mechanism attackers
> configure after a Business Email Compromise**, and it is the first thing an
> incident responder looks for. Setting it deliberately means creating the artefact
> that a BEC investigation hunts. Only run it against a written, approved request,
> keep the exported CSV as the record that IT configured it, prefer an internal
> target, and remove the forward when the reason for it ends.

The script enforces that in code: if the target domain is **not an accepted domain
of the tenant**, it refuses to apply anything unless you also pass
`-AllowExternalTarget`. The dry run tells you in advance that the switch will be
required. If the accepted-domain list cannot be read at all, the target is treated
as external — failing safe rather than assuming.

```powershell
# Preview - internal target, no extra switch needed
.\Mailboxes\Set-MailboxForwarding.ps1 `
    -SourceMailboxes 'user.one@contoso.com','user.two@contoso.com' `
    -ForwardTo 'shared.inbox@contoso.com'

# Apply to an external target - requires the explicit confirmation switch
.\Mailboxes\Set-MailboxForwarding.ps1 -SourceMailboxes 'user.one@contoso.com' `
    -ForwardTo 'someone@fabrikam.com' -AllowExternalTarget -Execute
```

**Input:** parameters; an existing EXO session.
**Output:** `Exports/MailboxForwarding_<timestamp>.csv` — previous and new forward
per mailbox, and whether the target was external. Keep it.
**Permissions:** Exchange Recipient Administrator.

> **Writes:** sets `ForwardingSMTPAddress` and `DeliverToMailboxAndForward`.
> Safeguards: dry run by default; external targets are blocked behind
> `-AllowExternalTarget`; the previous value of both settings is captured before
> the change so the CSV is a rollback plan; every change is read back and verified;
> `-KeepCopyInMailbox` defaults to `$true` so mail stays discoverable in the source
> mailbox, and a warning fires if it is turned off for an external target. To undo:
> `Set-Mailbox -Identity '<mailbox>' -ForwardingSMTPAddress $null -DeliverToMailboxAndForward $false`.
>
> Note that many tenants block external auto-forwarding at the outbound spam policy.
> This script will report success while the policy quietly drops the forwarded mail —
> check the policy, not just the mailbox attribute.

---

## `MailFlow/`

### `MailFlow/Block-MaliciousDomain.ps1`

This is incident-response tooling. An indicator arrives — a user report, a vendor
feed, another company in the group that got hit first — and the useful window for
acting on it is measured in minutes, not change windows. The work itself is two
separate journeys through the Tenant Allow/Block List in the portal, done under
time pressure, which is precisely the situation in which one of the two gets
forgotten and nobody notices until the same campaign lands again through the
other axis.

The no-expiry choice noted below is deliberate for the same reason: an indicator
does not stop being one after thirty days, and a block that lapses on a timer is
worse than one that was never set, because by then everyone assumes the
protection is still there.

Blocks a domain in the Tenant Allow/Block List as **both** a Sender and a URL entry.
Both matter: a Sender block stops mail claiming to come from the domain, but does
nothing about a link to it inside a message arriving from somewhere else. Most
phishing needs both.

`-Domain` is mandatory and has deliberately no default — a default here would mean
shipping someone's indicator of compromise inside the script.

```powershell
# Dry run
.\MailFlow\Block-MaliciousDomain.ps1 -Domain 'malicious.example'

# Apply the URL entry only, with a change reference on the entry
.\MailFlow\Block-MaliciousDomain.ps1 -Domain 'malicious.example' -ListType Url `
    -Notes 'Phishing campaign, ref 12345' -Execute
```

**Input:** parameters. Connects on its own, reusing an existing session if there is
one.
**Output:** `Exports/TABL-Block_<domain>_<timestamp>.csv`.
**Permissions:** Security Administrator (or Organization Management).

> **Writes:** creates Tenant Allow/Block List block entries with no expiration.
> Safeguards: dry run by default; checks each list for an existing entry first and
> skips rather than duplicating; re-reads both entries after writing to confirm
> they exist; per-list failures are recorded and do not abort the other list.
>
> Entries are created with `-NoExpiration` — they stay until someone removes them.
> Remove with `Remove-TenantAllowBlockListItems`.

### `MailFlow/Test-PhishingSimulationRule.ps1`

A phishing simulation platform only works if the mail users report actually
reaches it, which is normally arranged with a transport rule that adds the
vendor as a recipient. The rule gets configured once and then nobody looks at it
again — until the vendor dashboard shows suspiciously few reports and somebody
has to answer whether the pipeline is broken or the users simply stopped
reporting.

That is a verification question rather than a configuration one, and it needs a
different kind of evidence. A rule can exist, be enabled, and match on paper
while still not firing — a condition that no longer fits the report button's
format, a higher-priority rule consuming the message first. Reading the rule
definition back tells you nothing you did not already know, so this goes after
the traffic instead and leaves a report behind that can be attached to a ticket.

Read-only. Proves whether the transport rule that forwards user-reported phishing to
the simulation vendor is actually firing, by tracing **both legs**: inbound to the
trap mailbox, and outbound to the vendor address the rule adds as a recipient. A
rule that exists in the policy is not a rule that works — only the outbound leg
proves it. Produces a CSV plus a short text evidence report you can attach to a
change record.

```powershell
.\MailFlow\Test-PhishingSimulationRule.ps1 `
    -TrapMailbox 'phishing.report@contoso.com' `
    -SimulationReportAddress 'reports@simulation-vendor.example' -DaysBack 7
```

**Input:** parameters; connects on its own.
**Output:** `Evidence/MessageTrace_<timestamp>.csv` and
`Evidence/EvidenceReport_<timestamp>.txt`.
**Permissions:** message trace read (View-Only Recipients / Security Reader).

> **Known limitations:** message trace only retains roughly 10 days, so a larger
> `-DaysBack` silently returns nothing for the older part of the range. The script
> calls `Disconnect-ExchangeOnline` when it finishes, which will also close a
> session you had open before running it.

### `MailFlow/Get-MailboxReceiveVolume.ps1`

Per-mailbox mail-security licensing turns on one number: how much mail each
mailbox actually receives. Get it wrong in one direction and you buy protection
for mailboxes that receive nothing; wrong in the other and real users go
uncovered. It is also a number the vendor and the finance team will both check
against the admin center, so it has to survive being compared.

The earlier version rebuilt that figure by hand: thirty days of message trace in
twelve-hour windows, sixty-odd calls, filtered on `Status = Delivered`. It
undercounted, consistently, against the same report in the admin center — and
losing an argument about whose count is right is not a good use of a licensing
review. So the decision was to stop reconstructing something Microsoft already
publishes and read the report's own source instead, described below. The lesson
generalises well past this script: when you rebuild a metric the provider
already calculates, your number and theirs will diverge, and the one quoted in
the meeting is theirs.

Read-only. Counts mail received per mailbox over a period and splits the tenant into
"under the threshold" and "at or over it" — the input a per-mailbox mail-security
licensing decision needs.

The interesting part is the data source. The obvious approach is one
`Get-MessageTraceV2` per mailbox, which means thousands of calls, pagination,
rate-limit handling, and a ceiling of about 10 days of history. This uses the Graph
*Email Activity User Detail* report instead: **one** API call returns per-mailbox
receive counts aggregated server-side for D7/D30/D90/D180 — seconds instead of
minutes, reaching back 180 days, and matching the M365 admin center exactly, so the
numbers can be reconciled against the portal instead of argued about.

```powershell
# Default: 30-day window, threshold 150
.\MailFlow\Get-MailboxReceiveVolume.ps1

# Longer window and a different threshold, to a chosen file
.\MailFlow\Get-MailboxReceiveVolume.ps1 -Period D180 -MailThreshold 100 `
    -ExportCsvPath .\Exports\eligibility-D180.csv
```

**Input:** a Graph session (it connects and requests the scope if needed).
**Output:** a CSV of every mailbox with its receive count and eligibility flag, plus
an on-screen summary and the top 10 highest-volume mailboxes.
**Permissions:** Graph `Reports.Read.All` (admin consent required).

> **Known limitations:** Graph usage data lags real time by about two days — the
> admin center report has the same lag, which is fine for a licensing decision.
> If the tenant has *Concealed names* enabled, UPNs come back hashed; counts stay
> valid but cannot be mapped to users. The script detects this and tells you which
> setting to turn off.

---

## `Resources/`

### `Resources/Get-CountryResource.ps1`

Ask which meeting rooms belong to a given country and there is no attribute that
answers it. Rooms were created across many years and several conventions: the
origin marker in `CustomAttribute12` only exists on the ones provisioned after
that convention did, some have a `UsageLocation`, some announce their country
only in a display-name token or in the shape of their SMTP address. Filter on any
single signal and you get a short, confident, wrong list — and the rooms you
missed are invisible, because nothing failed.

Hence five weak signals ORed together rather than one strong one, and the
`MatchedBy` column, which is the honest part of the output: "40 rooms, 9 of them
found only by a name pattern" is a visibly different answer from "40 rooms, all
marked", and it lets you judge how much to trust the total instead of taking it
on faith.

Read-only. Finds room and equipment mailboxes belonging to a country.

In a tenant that grew by acquisition, resource mailboxes are named every possible
way and the origin marker in `extensionAttribute12` is only populated on some of
them. Filtering on any single signal misses rooms. This ORs several weak signals
together — marker, usage location, country token anywhere in the display name, full
country name, address convention, plus any extra pattern you pass — and reports a
`MatchedBy` column showing which ones fired, so you can see how sparse the marker
actually is instead of trusting it.

`-CheckAddress` additionally reports whether proposed addresses are already in use
across **all** recipient types, not just resources. That is the check you run before
proposing new room addresses.

```powershell
# All rooms and equipment for a country
.\Resources\Get-CountryResource.ps1 -Country HU

# Rooms only, plus an in-use check on proposed addresses, in one pass
.\Resources\Get-CountryResource.ps1 -Country HU -ResourceType Room -CheckAddress `
    'RESHU.MR.RoomOne@contoso.com','RESHU.MR.RoomTwo@contoso.com'
```

**Input:** parameters. `-MarkerPrefix`, `-AddressPrefix` and `-Domain` are the
organisation-specific tokens; the defaults are placeholders, so set them to your own
conventions.
**Output:** `Exports/Resources_<CC>_<timestamp>.csv`, and
`Exports/AddressCheck_<CC>_<timestamp>.csv` when `-CheckAddress` is used.
**Permissions:** View-Only Recipients.

### `Resources/Set-RoomMailbox.ps1`

Room booking looks trivial until the room is a boardroom, and then the question
becomes who may reserve it. Exchange answers that through `BookInPolicy` and
`AllBookInPolicy`, two settings that are easy to get exactly backwards: an empty
`BookInPolicy` with restrictions enabled does not mean "anyone", it means the
room auto-declines every request from everybody, silently, and the first person
to find out is the one whose meeting disappeared. The script refuses to apply
that state at all.

The second thing to get right is which of the two access models you actually
want, because requesters ask for both with the same sentence — "give the team
access to the room" can mean restricting who may book it or letting people amend
other people's bookings, and the two are implemented in entirely different
places. They are set out below, along with the trap in the second one.

Configures a room mailbox: place metadata for Room Finder, an auto-accept booking
policy, and who is allowed to use it. Two access models:

- **Booking** (default) — restricts *who may reserve* the room via `BookInPolicy`.
  Groups go in as-is, because Exchange evaluates their membership at request time;
  the restriction then survives staff changes without re-running anything.
- **Manage** — grants Editor on the room's calendar folder, for people who need to
  fix up other people's bookings.

The trap in Manage mode is that folder permissions do **not** reach the members of a
distribution list. The script detects it, refuses to pretend it worked, and offers
`-ExpandGroups`, which recursively flattens groups and shared mailboxes to
individuals with a cycle guard.

```powershell
# Preview: who would be allowed to book the room
.\Resources\Set-RoomMailbox.ps1 -RoomName 'Room 1.01' `
    -Managers 'reception@contoso.com','Example Team Leads'

# Manage mode, flattening the DL to its members, applied
.\Resources\Set-RoomMailbox.ps1 -RoomName 'Room 1.01' -CityCode 'CTY' -SiteCode 'SITE' `
    -Managers 'Example Team Leads' -Mode Manage -ExpandGroups -Execute
```

**Input:** parameters; connects interactively. `-PrimarySmtp` is derived from the
country/city/site codes and room name unless you pass it explicitly — and if what
you pass disagrees with the derived alias, the script warns rather than provisioning
a room whose address contradicts its name.
**Output:** `Exports/RoomMailbox_<alias>_<timestamp>.csv` and a full transcript
`.log` beside it.
**Permissions:** Exchange Administrator.

> **Writes:** optionally creates the room mailbox (`-CreateInCloud`), sets place
> metadata, sets the calendar processing policy, and either restricts `BookInPolicy`
> or grants Editor on the calendar folder. Safeguards: dry run by default, with every
> action routed through one wrapper that prints `would:` instead of acting; refuses
> to apply an **empty** `BookInPolicy`, which would auto-decline every request from
> everyone; run-level deduplication so a person reached through two groups is granted
> once; group expansion is cycle-guarded; dynamic distribution groups are reported as
> not expandable rather than silently ignored.

---

## A note on architecture

The group scripts in `Groups/` are largely the same script more than once. They
differ in which validations they perform and how much they verify afterwards, not
in what they fundamentally do: resolve members, create a group, wait for
replication, add members, verify.

They should be one script with a `-Type Distribution|Security|M365` parameter,
built on `New-FaxDistributionList.ps1`'s validation and verification, with the
fax-specific mail flow handling behind a switch. That would remove the real problem
here, which is not duplication but **drift**: `New-MailGroup.ps1` has no address
collision check, and `Edit-MailGroupMember.ps1` keeps its work list inside the file.
Fixes applied to one have not reached the others.

They are published as they are because that is how they were written under service
requests — one script per request, each improving on the last. The consolidation is
the obvious next step, and knowing that is worth more than pretending it is already
done.

## Known rough edges

The consolidation debt in `Groups/` is described in the note above; the rest are individual
defects, all of them reachable in normal use.

- **`New-MailGroup.ps1` cannot run unattended.** The group type comes from a `Read-Host`, so it
  cannot be scheduled or driven from a queue.
- **`New-MailGroup.ps1` has no address collision check.** It will attempt to create a group whose
  primary SMTP is already taken and fail partway. `New-FaxDistributionList.ps1` checks first; that
  fix never travelled.
- **`New-MailGroup.ps1` sleeps a fixed 10 seconds** and assumes replication has caught up. On a slow
  day it has not, and the member-add step fails with an error that reads like a permissions problem.
- **`Ensure-MailContact` in `New-MailGroup.ps1` uses an unapproved verb.** Harmless at runtime,
  flagged by PSScriptAnalyzer.
- **`Edit-MailGroupMember.ps1` keeps its work list inside the file.** Running it means editing it,
  which rules out scheduling and makes the run history harder to reconstruct than the CSV log
  suggests.
- **`New-FaxDistributionList.ps1` reports false `MISSING` members.** The closing diff compares the
  addresses you passed against each member's *primary* SMTP, so anyone requested by an alias is
  reported missing even though they were added correctly.
- **`New-SharedCalendar.ps1 -CreateAccessGroup` will usually fail.** Mail-enabled security groups
  can no longer be created in Exchange Online. Create the group on-prem and let it sync, or use
  `-DirectGrant`.
- **`Test-PhishingSimulationRule.ps1` disconnects Exchange Online when it finishes** — including a
  session you had open before you ran it.
- **Message trace retains roughly 10 days.** A larger `-DaysBack` returns nothing for the older part
  of the range rather than warning that it cannot cover it.
- **Graph usage data lags about two days**, so `Get-MailboxReceiveVolume.ps1` is never quite
  current. Fine for a licensing decision, wrong for anything operational.
- **Six of these scripts carry a banner header rather than comment-based help** —
  `Edit-MailGroupMember`, `New-FaxDistributionList`, `New-MailGroup`, `New-SharedMailbox`,
  `Get-MailboxReceiveVolume` and `Test-PhishingSimulationRule`. They document themselves
  perfectly well in an editor, but `Get-Help -Full` returns almost nothing for them.
- **`Set-RoomMailbox.ps1` and `Get-CountryResource.ps1` ship with placeholder naming tokens.**
  `-MarkerPrefix`, `-AddressPrefix`, the city/site codes and the calendar naming convention follow
  one subsidiary's scheme. They are parameters, but the defaults are not yours.
