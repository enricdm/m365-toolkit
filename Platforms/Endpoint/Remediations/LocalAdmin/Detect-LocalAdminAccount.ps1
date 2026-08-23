<#
.SYNOPSIS
    Intune remediation DETECTION: verifies that a named local administrator account
    exists, is enabled, and belongs to the local Administrators group.

.DESCRIPTION
    Detection half of a device-side remediation pair. Pair with
    Repair-LocalAdminAccount.ps1.

    Exit 0 = compliant, exit 1 = remediate. Nothing else. Intune reads the exit
    code; anything written to stdout is shown in the remediation output column.

    The Administrators group is resolved by its well-known SID (S-1-5-32-544), not
    by the name "Administrators", so this works on non-English Windows.

.PARAMETER AccountName
    Local account to check. Edit the default before deploying, or wrap this script
    with your own value - Intune remediation scripts run with no arguments.

.PARAMETER RequirePasswordNeverExpires
    Require PasswordNeverExpires to be set. OFF by default.

    Deliberate: a local admin password that never expires is only acceptable when
    nothing else rotates it. If Windows LAPS manages this account - which is the
    recommended configuration - LAPS owns expiry, and forcing PasswordNeverExpires
    fights it. Only enable this if you are certain no rotation mechanism is in play.

.EXAMPLE
    # Deploy as-is after editing the default AccountName
    .\Detect-LocalAdminAccount.ps1

.EXAMPLE
    # Test locally against a different account
    .\Detect-LocalAdminAccount.ps1 -AccountName 'svc-local'

.NOTES
    Context  : run as SYSTEM (Intune remediation default), 64-bit
    Requires : Windows PowerShell 5.1, Microsoft.PowerShell.LocalAccounts
    Rights   : read-only. This script never creates or modifies an account.

    Not applicable to Entra-joined devices where the account is managed by an
    external tool that may create it lazily; a "missing" result there can be a
    timing artefact rather than a real gap.

    Replaces (merged): Detect_itadm.ps1, Get-adminavUser.ps1
    Fixes: the original Detect_itadm.ps1 used 'return 1' rather than 'exit 1' on
    the account-missing branch. 'return' does not set the process exit code, so a
    device MISSING the account reported exit 0 = compliant, and never remediated.
    That inverted the result in exactly the case the check existed for.
#>

[CmdletBinding()]
param(
    [string]$AccountName = 'itadm',
    [switch]$RequirePasswordNeverExpires
)

# No $ErrorActionPreference = 'Stop' here: this script must always reach an
# explicit exit code, never die with an unhandled terminating error (Intune would
# record that as a script failure rather than "needs remediation").

try {
    $user = Get-LocalUser -Name $AccountName -ErrorAction Stop
}
catch {
    Write-Output "NON-COMPLIANT: local account '$AccountName' does not exist."
    exit 1
}

if (-not $user.Enabled) {
    Write-Output "NON-COMPLIANT: local account '$AccountName' exists but is disabled."
    exit 1
}

if ($RequirePasswordNeverExpires -and $user.PasswordNeverExpires -ne $true) {
    Write-Output "NON-COMPLIANT: '$AccountName' exists but PasswordNeverExpires is not set."
    exit 1
}

# --- local Administrators membership, resolved by SID ---
try {
    $adminSid   = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')
    $adminGroup = ($adminSid.Translate([System.Security.Principal.NTAccount]).Value -split '\\')[-1]
    $members    = @(Get-LocalGroupMember -Group $adminGroup -ErrorAction Stop)
}
catch {
    # Enumeration can fail on domain-joined devices when a member SID no longer
    # resolves. Treat that as unknown, not as compliant: report and remediate.
    Write-Output "NON-COMPLIANT: could not enumerate local Administrators - $($_.Exception.Message)"
    exit 1
}

$isAdmin = $false
foreach ($m in $members) {
    if ($m.ObjectClass -ne 'User') { continue }
    $leaf = ($m.Name -split '\\')[-1]
    if ($leaf -ieq $AccountName) { $isAdmin = $true; break }
}

if (-not $isAdmin) {
    Write-Output "NON-COMPLIANT: '$AccountName' exists but is not a member of local Administrators."
    exit 1
}

Write-Output "COMPLIANT: '$AccountName' exists, is enabled and is a local administrator."
exit 0
