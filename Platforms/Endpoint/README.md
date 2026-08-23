# Endpoint

Two read-only endpoint reporting scripts: a tenant-wide Windows support-lifecycle report from
Intune, and a per-device detection script for desktop Visio/Project usage.

## Index

| Script | What it does | Changes state | Auth |
|---|---|:---:|---|
| [`Get-OutOfSupportDevice.ps1`](Get-OutOfSupportDevice.ps1) | Reports every Windows device in Intune and flags those past end-of-servicing | No | delegated Graph |
| [`Get-VisioProjectDesktopUsage.ps1`](Get-VisioProjectDesktopUsage.ps1) | Runs on a device and reports whether Visio/Project are installed and when they were last used | No | none — runs locally |

## Requirements

- `Get-OutOfSupportDevice.ps1`
  - Module: `Microsoft.Graph.Authentication` (`Install-Module Microsoft.Graph`)
  - Scopes: `DeviceManagementManagedDevices.Read.All`, `User.Read.All`
  - Intune Administrator or an equivalent role that can read managed devices
- `Get-VisioProjectDesktopUsage.ps1`
  - Windows PowerShell 5.1 or PowerShell 7, 64-bit
  - No modules and no credentials. Must run **as the signed-in user** — it reads that user's `HKCU`.

## Usage

### `Get-OutOfSupportDevice.ps1`

Answers "how much of the Windows estate is past end of servicing, and whose devices are they?" —
the report you need before an upgrade campaign. It pulls every Windows device from Intune, converts
each `osVersion` build number into its feature-update name (`19045` → `Windows 10 22H2`), compares
that against a table of end-of-servicing dates, and enriches each row with the primary user's
`usageLocation`, country and department, plus the last logged-on user.

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

There is no Microsoft cloud API that reports desktop Visio or Project usage, which makes licence
reclamation for those two products guesswork. This script runs on the device instead and combines
the two local signals that do exist: the **Office File MRU** (last file opened in the app, with a
timestamp) and **UserAssist** (last `VISIO.EXE` / `WINPROJ.EXE` launch plus a run count). It reports
the later of the two, along with install state, edition and executable version.

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
