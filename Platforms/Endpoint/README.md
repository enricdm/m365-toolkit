# Endpoint

Intune device management: reporting, bulk maintenance and device hygiene.

Everything here talks to Microsoft Graph from an admin workstation.

## Index

### Reporting — read-only

| Script | What it does | Changes state | Auth |
|---|---|:---:|---|
| [`Intune/Get-IntuneGroupAssignment.ps1`](Intune/Get-IntuneGroupAssignment.ps1) | Everything assigned to an Entra ID group across 14 object types, separating include from exclude | No | delegated Graph |
| [`Intune/Get-IntuneDeviceGroupMembership.ps1`](Intune/Get-IntuneDeviceGroupMembership.ps1) | Which groups a device is in, or which devices are in a group — direct and via primary user | No | delegated Graph |
| [`Intune/Get-IntuneDeviceUserDrift.ps1`](Intune/Get-IntuneDeviceUserDrift.ps1) | Last logged-on user vs assigned primary user, with mismatch and staleness flags | No | delegated Graph |
| [`Intune/Export-IntuneDevice.ps1`](Intune/Export-IntuneDevice.ps1) | Device inventory with platform, ownership and Android enrollment-type filters; optional delta | No | delegated Graph |
| [`Intune/Get-OutOfSupportDevice.ps1`](Intune/Get-OutOfSupportDevice.ps1) | Reports every Windows device in Intune and flags those past end-of-servicing | No | delegated Graph |

### Maintenance — writes to the tenant

| Script | What it does | Changes state | Auth |
|---|---|:---:|---|
| [`Intune/Set-IntuneDevicePrimaryUser.ps1`](Intune/Set-IntuneDevicePrimaryUser.ps1) | Aligns primary user with the user who actually signs in | **Yes** — `-Execute` | delegated Graph |
| [`Intune/Set-IntuneDeviceCategory.ps1`](Intune/Set-IntuneDeviceCategory.ps1) | Assigns device categories from a mapping, by subnet or device name | **Yes** — `-Execute` | delegated Graph |
| [`Intune/Invoke-IntuneStaleDeviceCleanup.ps1`](Intune/Invoke-IntuneStaleDeviceCleanup.ps1) | Finds devices that stopped checking in and retires or deletes them | **Yes, destructive** — `-Execute` | delegated Graph |

### Not here: the local administrator remediation pair

There used to be an Intune Remediations detection/repair pair here that created a named local
administrator account with a random per-device password, on the understanding that Windows LAPS
would then take ownership of it and rotate it.

That is no longer worth shipping. **Windows LAPS is the answer to this problem**, and current
versions manage the account itself — creation included — rather than only its password. A script
that creates the account so LAPS can adopt it solves a problem LAPS no longer has, and publishing
one invites somebody to deploy it instead of configuring LAPS properly. Use Windows LAPS.

Two things it found are worth keeping even though the code is gone.

**A shared local-admin password is a lateral-movement primitive.** The four scripts it replaced each
carried the same plaintext password in source, deployed to every device in the fleet. Recover it
from one device and you have administrative access to every device that shares it, and it cannot be
rotated without editing and redeploying. If you are looking at scripts that do that, treat the
password as compromised and move to LAPS.

**`return 1` is not `exit 1`, and in a detection script the difference inverts the result.** The
detection half it replaced ended its account-missing branch with `return`. `return` does not set the
process exit code, so a device that was **missing** the account exited `0` — reported compliant, and
never remediated. The check failed in exactly the case it existed for, and nothing surfaced it,
because a detection script that reports compliant looks identical to one that found nothing wrong.
Anything you deploy through Intune Remediations should be tested for the exit code it actually
returns, not the branch you think it took.

## Requirements

**`Intune/`**

- PowerShell 7.2+
- Module: `Microsoft.Graph.Authentication` (`Install-Module Microsoft.Graph`)
- `Platforms/_Shared/Modules/M365.Common.psm1` — imported for `Invoke-GraphPaged`. **The `Intune/`
  scripts do not work outside this repo layout**; they resolve it as `..\..\_Shared\` from
  `$PSScriptRoot`.
- Scopes are documented per script in `.NOTES`. Read-only scripts ask only for `.Read.All`; the
  three that write ask for `DeviceManagementManagedDevices.ReadWrite.All`.
- Intune Administrator, or a custom role with the equivalent read/write on managed devices

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

### `Intune/Get-IntuneDeviceUserDrift.ps1`

The drift between who the device says used it last and who Intune thinks owns it is what makes
licence attribution, retirement and support routing go wrong.

```powershell
# Windows only, with Autopilot, treating two profiles as shared devices
.\Get-IntuneDeviceUserDrift.ps1 -Platform Windows -IncludeAutopilot `
    -SharedProfileName 'AP-Classrooms','AP-MeetingRooms' -OutputPath .\Exports\device-users.csv

# Only the rows that need attention
.\Get-IntuneDeviceUserDrift.ps1 -Platform Windows | Where-Object Status -eq 'Mismatch'
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
> `.\Get-IntuneDeviceUserDrift.ps1 -OutputPath .\before.csv`.
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

### `Intune/Get-OutOfSupportDevice.ps1`

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

## Known rough edges

- **The `Intune/` scripts have not been run against a live tenant from this
  repo.** They are verified statically — all parse, parameter sets bind, comment-based help
  renders, and PSScriptAnalyzer reports nothing beyond `PSAvoidUsingWriteHost`, which these
  scripts trigger deliberately: they are interactive tools whose progress output is the point.
  Test in a non-production tenant first, especially the three that write.
- **`Set-IntuneDeviceCategory.ps1` in `Subnet` mode is slow by construction** (one Graph call per
  device). There is no batch endpoint that returns `subnetAddress`.
- **Several collections only exist on Graph `/beta`** — settings catalog, remediation scripts,
  `usersLoggedOn`, `hardwareInformation`. Those calls target `/beta` deliberately, and beta contracts
  can change without notice. If a collection starts returning a different shape, that is the first
  place to look.
- **`Get-IntuneDeviceGroupMembership.ps1` does not expand transitive membership.** If your assignment
  model relies on nested groups, treat the output as a starting point rather than the final answer.
