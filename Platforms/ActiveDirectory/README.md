# Active Directory

Two on-premises Active Directory tools: a proxyAddresses hygiene auditor for hybrid identity, and
an unauthenticated SMBv1 network probe.

## Index

| Script | What it does | Changes state | Auth |
|---|---|:---:|---|
| [`Repair-ProxyAddressPrimary.ps1`](Repair-ProxyAddressPrimary.ps1) | Finds AD objects with a broken primary SMTP address, and optionally repairs the safe cases | **Yes** (with `-Execute`) | current AD session |
| [`Test-SmbV1.ps1`](Test-SmbV1.ps1) | Probes hosts over the network to determine whether SMBv1 is enabled | No | none — pre-auth probe |

## Requirements

- **Windows PowerShell 5.1 or PowerShell 7+**
- `Repair-ProxyAddressPrimary.ps1`: the **ActiveDirectory** module (RSAT), line of sight to a domain
  controller, and write permission on the user objects to use `-Execute`.
- `Test-SmbV1.ps1`: no module and no credentials — only TCP 445 reachability to the targets. It uses
  raw .NET sockets and nothing else.

## Usage

### `Repair-ProxyAddressPrimary.ps1`

In hybrid identity, `proxyAddresses` is the attribute that decides a user's mail address, and the
primary one is marked by an **uppercase** `SMTP:` prefix — exactly one value, exactly that casing.
A `STMP:` or an `SMPT:`, a typo at the keyboard and nothing more, leaves the object with **no**
primary address at all. ADUC draws both entries the same way, so nothing looks wrong at the place
the mistake was made; the symptom turns up later and somewhere else, as mail that does not deliver
or an object Entra Connect refuses to sync.

This script finds those objects and classifies each one: `NoPrimary`, `BadPrefix`,
`MultiplePrimary`, `PrimaryIsMOERA` (primary is the `*.onmicrosoft.com` routing address), and
`MailMismatch` (the `mail` attribute disagrees with the primary). It detects five patterns and
repairs exactly one — the only one with a single correct answer. An object with two primaries needs
a person to decide which address wins, and guessing there would quietly change someone's reply-to
address.

```powershell
# Dry-run: report and export CSV, change nothing
.\Repair-ProxyAddressPrimary.ps1 -Server 'DC01.corp.local' -AcceptedRoot 'contoso.com'

# Apply the safe repairs, scoped to one OU
.\Repair-ProxyAddressPrimary.ps1 -Server 'DC01.corp.local' -AcceptedRoot 'contoso.com' `
    -SearchBase 'OU=Users,DC=corp,DC=local' -Execute
```

**Input:** a domain controller (`-Server`) and the domain the primary address should use
(`-AcceptedRoot`). `-SearchBase` narrows the scope; omit it for the whole domain.
**Output:** a `;`-delimited audit CSV of every finding, plus a second change-log CSV when `-Execute`
is used.
**Permissions:** read access to `proxyAddresses` for the audit; write access to the user objects to repair.

> **Writes when `-Execute` is passed.** Safeguards: dry-run by default, and it repairs exactly one
> narrow case — an object with **zero** primaries and **exactly one** malformed value whose address
> is well-formed and not already present as a secondary alias. Everything else (multiple primaries,
> several bad values, MOERA primaries) is reported and deliberately left alone, because those need a
> human decision about which address should win. The `mail` attribute is only realigned when the
> repaired address belongs to `-AcceptedRoot`.
>
> Run a delta sync on Entra Connect afterwards and verify in Exchange Online. Script output is in Spanish.

### `Test-SmbV1.ps1`

> #### ⚠ Authorisation required before use
>
> **This is a network scanner.** It opens raw TCP connections to hosts you name and sends a
> hand-built SMB protocol packet. Running it against infrastructure you do not own, or do not have
> written permission to test, may be treated as hostile activity, may breach computer-misuse law in
> your jurisdiction, and will very likely trip IDS/EDR alerts and start an incident.
>
> Only run it against hosts you are explicitly authorised to test, and tell your security team
> before you do. The same warning is in the script header.

A vulnerability scan flagged SMBv1 on several domain controllers. SMBv1 is the protocol EternalBlue
exploited — WannaCry and NotPetya both came in that way. Microsoft deprecated it in 2014 and it
hasn't shipped enabled since Windows 10 1709, but it keeps reappearing: an old NAS turns it back on,
a network printer needs it, a legacy scanner nobody wants to touch.

The question wasn't whether it was enabled. It was how to **prove** it. Checking the registry tells
you what's *configured* — and on machines where a GPO didn't apply, or the service never restarted
after the change, that diverges from what the server actually answers on the wire. If you're going
to close an audit finding, you want the second answer.

So this asks the way the scanner asks: open 445, send an `SMB_COM_NEGOTIATE` offering **only** SMB1
dialects — deliberately omitting `SMB 2.???` so an SMB2-capable server can't upgrade its way out of
the question — and read which one it picks.

Because SMB dialect negotiation happens *before* authentication, this needs no credentials and no
admin rights on the target — only reachability. It authenticates nothing, exploits nothing, and
changes nothing on the target. It is the honest way to check a host you cannot log into, and to
verify that a remediation actually took effect rather than trusting a registry value elsewhere.

```powershell
# One host
.\Test-SmbV1.ps1 -ComputerName DC01

# Several, with a longer timeout over a slow link
.\Test-SmbV1.ps1 -ComputerName DC01,DC02,FS01 -TimeoutMs 10000
```

**Input:** `-ComputerName` — one or more hostnames or IPs. Mandatory, with no default target list, by design.
**Output:** a table plus a timestamped CSV in `-ExportDir`.
**Permissions:** none on the target. Just network reachability to port 445.

It reports **four distinct outcomes**, and keeping them apart is the whole point:

| Result | Meaning |
|---|---|
| `ENABLED` | The server selected an SMB1 dialect. Finding confirmed. |
| `DISABLED` | Refused the dialects, answered from an SMB2-only stack, or reset the connection after negotiation. |
| `NO_TCP` | Port 445 unreachable or filtered. **Not a verdict.** |
| `UNCLEAR` | Answered, but not in a shape this probe understands. Inspect by hand. |

Two details that make the results trustworthy:

- **A post-negotiate RST counts as `DISABLED`, not as an error.** A host with the SMB1 server driver
  stopped typically accepts the connection and then resets it once it sees an SMB1-only dialect list.
  A naive probe records that as a connection failure and loses the answer; here it is treated as the
  valid negative signal it is.
- **`NO_TCP` is never folded into `DISABLED`.** A firewalled host proves nothing about whether SMB1
  is enabled on it. Collapsing "I could not reach it" into "it is safe" turns a blocked scan into a
  false all-clear, which is exactly the kind of result that survives into a compliance report unchallenged.

If `nmap` is available, `nmap -p445 --script smb-protocols <host>` is a reasonable cross-check.
