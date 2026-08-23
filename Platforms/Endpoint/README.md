# Endpoint

Intune device management: reporting, bulk maintenance, configuration snapshots and device-side
detection/remediation pairs — plus two standalone endpoint reports.

Everything under `Intune/` talks to Microsoft Graph from an admin workstation. Everything under
`Remediations/` runs **on the device**, as SYSTEM, deployed through Intune Remediations.

## Index

### Reporting — read-only

| Script | What it does | Changes state | Auth |
|---|---|:---:|---|
| [`Intune/Get-IntuneGroupAssignment.ps1`](Intune/Get-IntuneGroupAssignment.ps1) | Everything assigned to an Entra ID group across 14 object types, separating include from exclude | No | delegated Graph |
| [`Intune/Get-IntuneDeviceGroupMembership.ps1`](Intune/Get-IntuneDeviceGroupMembership.ps1) | Which groups a device is in, or which devices are in a group — direct and via primary user | No | delegated Graph |
| [`Intune/Get-IntuneDeviceUser.ps1`](Intune/Get-IntuneDeviceUser.ps1) | Last logged-on user vs assigned primary user, with mismatch and staleness flags | No | delegated Graph |
| [`Intune/Export-IntuneDevice.ps1`](Intune/Export-IntuneDevice.ps1) | Device inventory with platform, ownership and Android enrollment-type filters; optional delta | No | delegated Graph |
| [`Intune/Export-IntuneConfiguration.ps1`](Intune/Export-IntuneConfiguration.ps1) | Point-in-time JSON snapshot of tenant configuration, with script bodies decoded to `.ps1` | No | delegated Graph |
| [`Get-OutOfSupportDevice.ps1`](Get-OutOfSupportDevice.ps1) | Reports every Windows device in Intune and flags those past end-of-servicing | No | delegated Graph |
| [`Get-VisioProjectDesktopUsage.ps1`](Get-VisioProjectDesktopUsage.ps1) | Runs on a device and reports whether Visio/Project are installed and when they were last used | No | none — runs locally |

### Maintenance — writes to the tenant

| Script | What it does | Changes state | Auth |
|---|---|:---:|---|
| [`Intune/Set-IntuneDevicePrimaryUser.ps1`](Intune/Set-IntuneDevicePrimaryUser.ps1) | Aligns primary user with the user who actually signs in | **Yes** — `-Execute` | delegated Graph |
| [`Intune/Set-IntuneDeviceCategory.ps1`](Intune/Set-IntuneDeviceCategory.ps1) | Assigns device categories from a mapping, by subnet or device name | **Yes** — `-Execute` | delegated Graph |
| [`Intune/Invoke-IntuneStaleDeviceCleanup.ps1`](Intune/Invoke-IntuneStaleDeviceCleanup.ps1) | Finds devices that stopped checking in and retires or deletes them | **Yes, destructive** — `-Execute` | delegated Graph |

### Remediations — run on the device

Detection and repair are separate scripts, as Intune Remediations expects. Detection exits `0` for
compliant and `1` for "remediate"; nothing else.

| Pair | What it checks and fixes | Changes state | Auth |
|---|---|:---:|---|
| [`Remediations/LocalAdmin/`](Remediations/LocalAdmin/) | A named local administrator account exists, is enabled, and is in the Administrators group | **Yes** — repair half | none — SYSTEM |
| [`Remediations/Bookmarks/`](Remediations/Bookmarks/) | Managed-bookmarks policy for Edge or Chrome matches the expected set | **Yes** — repair half | none — SYSTEM |

## Requirements

**`Intune/` and `Get-OutOfSupportDevice.ps1`**

- PowerShell 7.2+
- Module: `Microsoft.Graph.Authentication` (`Install-Module Microsoft.Graph`)
- `Platforms/_Shared/Modules/M365.Common.psm1` — imported for `Invoke-GraphPaged`. **The `Intune/`
  scripts do not work outside this repo layout**; they resolve it as `..\..\_Shared\` from
  `$PSScriptRoot`.
- Scopes are documented per script in `.NOTES`. Read-only scripts ask only for `.Read.All`; the
  three that write ask for `DeviceManagementManagedDevices.ReadWrite.All`.
- Intune Administrator, or a custom role with the equivalent read/write on managed devices

**`Get-VisioProjectDesktopUsage.ps1`**

- Windows PowerShell 5.1 or PowerShell 7, 64-bit
- No modules and no credentials. Must run **as the signed-in user** — it reads that user's `HKCU`.

**`Remediations/`**

- Windows PowerShell 5.1, 64-bit, running as SYSTEM
- No modules, no credentials, no network calls
- Deploy through **Devices → Remediations**. Set *Run in 64-bit PowerShell* = **Yes**.
  Leave *Run using logged-on credentials* = **No** — these need SYSTEM.

> Intune Remediations passes **no arguments**. The parameters exist so the scripts can be tested and
> reviewed; before deploying, edit the defaults in the `param()` block to match your environment.

## Usage

### `Intune/Get-IntuneGroupAssignment.ps1`

The question this answers is "what actually lands on this group?" — the one you need before changing
or deleting it. It walks 14 collections: compliance policies, device configurations, settings-catalog
policies, ADMX configurations, applications, app configuration policies, three kinds of app
protection policy, WIP policies, remediation scripts, platform scripts and Autopilot profiles.

```powershell
.\Get-IntuneGroupAssignment.ps1 -GroupName 'GRP-Devices-Classrooms'

# Several groups, exported, showing object types with no assignment too
.\Get-IntuneGroupAssignment.ps1 -GroupName 'GRP-Staff','GRP-Kiosk' `
    -IncludeEmpty -OutputPath .\Exports\assignments.csv
```

**Input:** `-GroupName` (exact match) or `-GroupId`.
**Output:** one row per assignment; optional CSV.
**Permissions:** `Group.Read.All`, `DeviceManagementConfiguration.Read.All`,
`DeviceManagementApps.Read.All`, `DeviceManagementManagedDevices.Read.All`,
`DeviceManagementServiceConfig.Read.All`.

> **Include and exclude are reported separately, and that matters.** A group used purely as an
> *exclusion* is not receiving the policy — it is being kept away from it. Reading one as the other
> inverts the meaning of the whole report.

> **Assignments to the built-in "All devices" / "All users" targets are not group assignments and do
> not appear here, by design.** They still apply to the devices in your group. An empty result means
> "nothing is assigned *to this group*", not "nothing applies to these devices". `-IncludeEmpty`
> prints a row for every object type queried, so the report shows what was checked rather than only
> what was found.

### `Intune/Get-IntuneDeviceGroupMembership.ps1`

The complement of the previous script: that one goes group → policy, this one goes device ↔ group.
Together they close the loop from a device to the policy that reached it.

```powershell
# Which groups is this device in, directly and through its user
.\Get-IntuneDeviceGroupMembership.ps1 -DeviceName 'LAP-0042' -IncludeUserGroups

# Which Intune devices are in these groups
.\Get-IntuneDeviceGroupMembership.ps1 -GroupName 'GRP-Devices-Kiosk'
```

**Input:** `-DeviceName` (wildcards accepted) or `-GroupName`.
**Output:** one row per device/group pair, with a `MembershipVia` column of `Device` or `PrimaryUser`.

> **Device groups and user groups are not the same thing**, and conflating them is the usual reason a
> policy appears to apply for no reason. A policy assigned to a *user* group reaches a device through
> whoever is its primary user. `-IncludeUserGroups` reports both and labels which is which.

> **Dynamic group membership is evaluated asynchronously by Entra.** A device just enrolled, renamed
> or recategorised may not yet appear in a dynamic group whose rule it already satisfies. An absent
> row means "not a member right now", not "the rule does not match it". Nested groups are reported
> one level deep; transitive membership is not expanded.

### `Intune/Get-IntuneDeviceUser.ps1`

The drift between who the device says used it last and who Intune thinks owns it is what makes
licence attribution, retirement and support routing go wrong.

```powershell
# Windows only, with Autopilot, treating two profiles as shared devices
.\Get-IntuneDeviceUser.ps1 -Platform Windows -IncludeAutopilot `
    -SharedProfileName 'AP-Classrooms','AP-MeetingRooms' -OutputPath .\Exports\device-users.csv

# Only the rows that need attention
.\Get-IntuneDeviceUser.ps1 -Platform Windows | Where-Object Status -eq 'Mismatch'
```

**Output:** a `Status` column of `Match`, `Mismatch`, `Shared`, `NoPrimaryUser`, `NoLoginData` or
`NoUserData`.

> **Shared devices mismatch by design.** Many people sign in to a classroom or meeting-room PC and
> the primary user is unset or a resource account. `-SharedProfileName` labels those `Shared` instead
> of drowning the real mismatches in expected noise.

> **A blank last-logged-on user is not proof that nobody used the device.** `usersLoggedOn` is
> populated by the Intune agent and is routinely empty on freshly enrolled devices, on devices that
> have not checked in since the property was introduced, and on most non-Windows platforms. Those
> rows report `NoLoginData`, which means *not known*.

### `Intune/Export-IntuneDevice.ps1`

One export with filters, replacing the pattern of keeping a separate copy per filter.

```powershell
.\Export-IntuneDevice.ps1 -Platform Android -EnrollmentType BYOD
.\Export-IntuneDevice.ps1 -Platform Windows -Ownership company -ComplianceState noncompliant

# Incremental: the first run seeds the token, later runs return only changes
.\Export-IntuneDevice.ps1 -DeltaStatePath .\State\devices-delta.json
```

`-EnrollmentType` takes the friendly Android Enterprise names — `BYOD`, `COPE`, `FullyManaged`,
`Dedicated` — and maps them to the values Graph reports.

> **The delta has a real limit, worth understanding before relying on it.** Intune's `managedDevices`
> collection has no delta endpoint; the delta lives on the *Entra* devices collection. So the script
> uses the Entra delta to learn which devices changed, then fetches those from Intune. A change that
> happens only on the Intune side and never touches the Entra device object — a compliance state
> flip, a new last-sync timestamp — may not surface. **Delta mode is for tracking devices appearing
> and disappearing, not for tracking every attribute.** For current compliance across the estate, run
> a full export.

### `Intune/Export-IntuneConfiguration.ps1`

A readable, diffable record of how the tenant is configured. Run it on a schedule into a git working
tree and `git diff` answers "what changed in Intune last week?", which the console cannot.

```powershell
# Into a git repo, same paths every run, so the diff is the history
.\Export-IntuneConfiguration.ps1 -OutputRoot C:\IntuneSnapshot -NoTimestampFolder -IncludeAssignments

# Just the scripts
.\Export-IntuneConfiguration.ps1 -Include PlatformScripts,RemediationScripts
```

Script bodies — platform, remediation, Win32 detection and requirement rules — are returned by Graph
base64-encoded. They are decoded and written as `.ps1` beside the JSON, so the actual code is
diffable rather than an opaque blob.

> **This is an export, not a backup product, and the difference is not pedantic.**
> There is no restore. Secrets are never returned by Graph, so certificate payloads, VPN pre-shared
> keys, Wi-Fi passwords and `.intunewin` contents are **absent** — a profile rebuilt from this JSON
> would be missing them. Assignments reference group object IDs, which are meaningless in another
> tenant. Applications export their *definition*, not their installer.
>
> Every run writes `manifest.json` recording which types were queried and which failed, so a type
> with zero objects reads as "none found" rather than "not checked".

### `Intune/Set-IntuneDevicePrimaryUser.ps1`

```powershell
# Dry run from a reviewed CSV — reports only
.\Set-IntuneDevicePrimaryUser.ps1 -InputCsv .\changes.csv

# Live drift, excluding shared devices
.\Set-IntuneDevicePrimaryUser.ps1 -FromDrift `
    -ExcludeAutopilotProfile 'AP-Classrooms','AP-MeetingRooms' -Execute
```

**Permissions:** `DeviceManagementManagedDevices.ReadWrite.All`, `User.Read.All`.

> There is no bulk undo. Export the current state first with
> `.\Get-IntuneDeviceUser.ps1 -OutputPath .\before.csv`.
>
> A device with no last logged-on user is **skipped, never cleared**. Absence of sign-in data means
> "not known", not "no user". Use `-ExcludeAutopilotProfile` with `-FromDrift`: on a shared device
> the "last user" is whoever signed in last, which is not an owner.

### `Intune/Set-IntuneDeviceCategory.ps1`

Device category drives dynamic groups, reporting and support routing. Set by hand it drifts
immediately; this keeps it derived from something factual.

```powershell
# Dry run — what would change, and why
.\Set-IntuneDeviceCategory.ps1 -MappingCsv .\subnet-to-site.csv

# Match on naming convention instead — no per-device calls
.\Set-IntuneDeviceCategory.ps1 -MappingCsv .\name-to-category.csv -MatchOn DeviceName -Execute
```

Mapping format:

```csv
Key,Category
10.20.30.0,Site-North
10.20.40.0,Site-South
```

> **Subnet mode costs one Graph call per device.** `hardwareInformation.subnetAddress` is not
> returned by the list endpoint — only by a single-device GET. A 2,000-device fleet is 2,000 calls
> and takes minutes. That is a property of the API, not of the script; progress is reported every
> 100 devices. `DeviceName` mode needs no extra calls.
>
> **A device whose subnet is blank is reported and skipped, never given a default category.** A blank
> subnet means the device has not sent hardware inventory yet — missing information, not membership
> of a fallback.
>
> Changing a category can move a device between dynamic groups, and therefore change which policies
> and apps target it. Run without `-Execute` and read the CSV.

### `Intune/Invoke-IntuneStaleDeviceCleanup.ps1`

```powershell
# See what is stale. Changes nothing.
.\Invoke-IntuneStaleDeviceCleanup.ps1 -StaleDays 120

# Delete stale corporate Windows records, capped at 25
.\Invoke-IntuneStaleDeviceCleanup.ps1 -StaleDays 180 -Platform Windows `
    -Action Delete -MaxDevices 25 -Execute
```

`Retire` removes company data and unenrols; the hardware keeps working. `Delete` removes the Intune
record only and does not touch the device — if it ever checks in again it may simply re-enrol.

> **Wipe is deliberately not offered.** Factory-resetting a fleet from a staleness query is not a
> cleanup operation. A stale record is very often a device that is merely switched off — a laptop in
> a drawer, someone on extended leave, a machine awaiting reassignment. Wipe destroys user data and
> cannot be undone. Wipe a specific device deliberately, one at a time, from the portal.
>
> `-StaleDays` rejects values below 30, and `-MaxDevices` defaults to 50. Devices past the cap are
> still written to the log as `SKIPPED`, so the CSV shows the full picture.
>
> **A stale `lastSyncDateTime` is evidence the device has not contacted Intune. It is not evidence
> the device is gone.** Confirm against your CMDB before acting on anything you cannot re-enrol.
> Devices with no usable timestamp at all are reported and never acted on.

### `Remediations/LocalAdmin/`

Detection checks that the account exists, is enabled, and is in the local Administrators group —
resolved by well-known SID `S-1-5-32-544`, so it works on non-English Windows. Repair creates or
fixes it.

> **No password is embedded, and that is the entire point.**
>
> The scripts this replaces each carried the same plaintext local-admin password in source, deployed
> to every device in the fleet. A shared static local administrator password is a lateral-movement
> primitive: recover it from one device and you have administrative access to every device that
> shares it. It also cannot be rotated without editing and redeploying.
>
> Instead the repair script generates a random password per device with a cryptographic RNG, never
> writes it to disk and never returns it. That password is therefore unknown to everyone — including
> you — which is only useful if something else manages it. **Deploy Windows LAPS against this
> account** (Endpoint security → Account protection → Local admin password solution), setting
> `AdministratorAccountName` to the same value.
>
> **Without LAPS this creates an account nobody can log in to.** That is a deliberate trade-off: an
> unusable break-glass account is a smaller problem than a fleet-wide shared one. Set up LAPS first.

`-RequirePasswordNeverExpires` is **off** by default: if LAPS manages the account, LAPS owns expiry,
and pinning it here fights that.

> **A fixed defect worth naming.** The detection script this replaces used `return 1` rather than
> `exit 1` on the account-missing branch. `return` does not set the process exit code, so a device
> **missing** the account exited `0` — reported compliant, never remediated. The check inverted its
> result in precisely the case it existed for.

### `Remediations/Bookmarks/`

Writes the browser **policy** — `ManagedFavorites` for Edge, `ManagedBookmarks` for Chrome — under
`HKLM\SOFTWARE\Policies\`. Managed bookmarks appear in a read-only folder on the bookmarks bar,
survive profile resets and cannot be deleted by the user, which is what you want on shared devices.

Detection compares the **set of URLs**, order- and case-insensitively; titles are ignored, because
renaming a folder label is not a compliance problem but a missing destination is.

> **Why the policy and not the bookmarks file.** Editing
> `%LOCALAPPDATA%\...\User Data\Default\Bookmarks` directly is destructive — the scripts this
> replaces cleared `bookmark_bar.children` before inserting, discarding everything the user had — and
> unreliable: the browser holds the file open and rewrites it on exit, silently reverting the change;
> the file carries a checksum, so an externally edited copy can be discarded; and it only affects the
> Default profile of whoever ran it.
>
> `-SeedUserFile` is still available for the narrow case where users must be able to *edit* the
> seeded bookmarks. It is off by default, **appends only**, backs the file up first, and skips itself
> with a message if the browser is running.

The policy applies at next browser launch. A device still showing the old bookmarks immediately after
remediation has not necessarily failed.

### `Get-OutOfSupportDevice.ps1`

Windows 10 and 11 ship as versions, and each version has its own end-of-servicing date. Past that
date a device stops receiving security updates while continuing to report as managed and compliant
in Intune — compliance policy and servicing lifecycle are unrelated things, so nothing in the
console marks it. There is no standard report that crosses the build number against the lifecycle
calendar.

So this one does. It pulls every Windows device from Intune, converts each `osVersion` build number
into its feature-update name (`19045` → `Windows 10 22H2`), compares that against a table of
end-of-servicing dates, and enriches each row with the primary user's `usageLocation`, country and
department, plus the last logged-on user — the answer to "how much of the Windows estate is past end
of servicing, and whose devices are they?", which is the report you need before an upgrade campaign.

```powershell
.\Get-OutOfSupportDevice.ps1

# Custom output location
.\Get-OutOfSupportDevice.ps1 -OutputPath .\Exports\devices.csv
```

**Input:** none beyond the optional `-OutputPath`.
**Output:** two CSVs — the full device report, and a `_OutOfSupportOnly.csv` filtered to the flagged
devices — plus an on-screen summary grouped by usage location.
**Permissions:** `DeviceManagementManagedDevices.Read.All`, `User.Read.All` (delegated).

> **Version note — build-parsing fix.** An earlier revision reported **zero** out-of-support devices
> on a fleet that actually had roughly 2,100 of them. `osVersion` is `<major>.<minor>.<build>.<revision>`
> (e.g. `10.0.19045.4046`), and the feature-update build is the **third** segment; the previous code
> read the first, which is always the literal `10`. Every lookup missed, every device fell through to
> "unknown version", and nothing was ever flagged — a clean, confident, entirely wrong report.
>
> This is fixed, and the parsing is now covered by a comment in the script so it does not regress.
> Worth knowing for two reasons: if you have output from an older copy, re-run it; and a zero result
> from any lifecycle report deserves suspicion before it deserves trust.

Two behaviours to be aware of:

- **The device query uses the Graph `/beta` endpoint,** solely because `usersLoggedOn` does not exist
  on `managedDevice` in `/v1.0` (asking for it there returns HTTP 400). Per-user lookups stay on
  `/v1.0`. If your devices do not populate that collection, the last-logged-on columns are simply blank.
- **The EOL dates assume Enterprise/Education servicing timelines,** standard for corporate-managed
  devices. Home/Pro editions reach end of servicing earlier, so adjust `$eolMap` if a meaningful part
  of the fleet runs those.

### `Get-VisioProjectDesktopUsage.ps1`

Desktop Visio and Project are among the few M365 products whose usage Graph does not report at all.
To decide whether one of those licences is being used, you have to ask the endpoint — an attempt to
measure it from the tenant side ended in the conclusion that it was not measurable there, which is
why the answer had to come down to the device.

So this runs on the device and reads the two local signals that do exist, both from the signed-in
user's registry: the **Office File MRU** (last file opened in the app, with a timestamp) and
**UserAssist** (last `VISIO.EXE` / `WINPROJ.EXE` launch plus a run count — ROT13-decoded, with the
launch time read out of the FILETIME inside the value blob). It keeps the later of the two, along
with install state, edition and executable version.

Read-only — `HKCU` and `HKLM` reads only, no writes to the device. It emits a single line of JSON to
stdout, shaped for Intune Remediations, which captures the last stdout line.

```powershell
# Intune Remediations: upload as the detection script
.\Get-VisioProjectDesktopUsage.ps1

# GPO logon script: also append a CSV row to a share
.\Get-VisioProjectDesktopUsage.ps1 -OutputShare '\\server\share\VPUsage'
```

**Input:** optional `-OutputShare` UNC path.
**Output:** one compact JSON line on stdout; optionally one CSV row per computer/user in the share.
**Permissions:** none, but it must run in the user's context.

For Intune Remediations, set **"Run this script using the logged-on credentials" = Yes**, **"Enforce
script signature check" = No**, and **"Run in 64-bit PowerShell" = Yes**. Collect results from the
Pre-remediation detection output column, or via Graph `deviceHealthScripts` run states.

> **These signals are a floor, not a measurement.** MRU only populates when a file is opened or
> saved; a user who only ever creates blank documents shows up through UserAssist instead, and a
> freshly reimaged device has neither. **"No signal" is not proof of non-use.** Read it alongside
> device coverage — and do not reclaim a licence on the strength of one empty result.
>
> If it runs as SYSTEM it reports that in the `Context` field and skips the per-user signals rather
> than reporting a misleading "never used".

## Known rough edges

- **The `Intune/` and `Remediations/` scripts have not been run against a live tenant from this
  repo.** They are verified statically — all parse, PSScriptAnalyzer is clean, comment-based help
  renders. Test in a non-production tenant first, especially the three that write.
- **`Set-IntuneDeviceCategory.ps1` in `Subnet` mode is slow by construction** (one Graph call per
  device). There is no batch endpoint that returns `subnetAddress`.
- **Several collections only exist on Graph `/beta`** — settings catalog, remediation scripts,
  `usersLoggedOn`, `hardwareInformation`. Those calls target `/beta` deliberately, and beta contracts
  can change without notice. If a collection starts returning a different shape, that is the first
  place to look.
- **`Get-IntuneDeviceGroupMembership.ps1` does not expand transitive membership.** If your assignment
  model relies on nested groups, treat the output as a starting point rather than the final answer.
- **`Export-IntuneConfiguration.ps1` has no restore counterpart** and is not planned to have one.
