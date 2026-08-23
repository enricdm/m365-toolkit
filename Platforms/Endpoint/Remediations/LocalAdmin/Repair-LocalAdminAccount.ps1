<#
.SYNOPSIS
    Intune remediation REPAIR: creates or repairs a named local administrator
    account with a randomly generated password, intended to be handed over to
    Windows LAPS for rotation.

.DESCRIPTION
    Repair half of a device-side remediation pair. Pair with
    Detect-LocalAdminAccount.ps1.

    Handles three states:
      - account missing        -> created with a random password
      - account exists, disabled -> enabled
      - account not an admin   -> added to local Administrators (resolved by SID,
                                  so it works on non-English Windows)

    NO PASSWORD IS EMBEDDED IN THIS SCRIPT, and that is the whole point.

    The versions this replaces each carried the same plaintext local-admin password
    in source, deployed to every device in the fleet. A shared, static local
    administrator password is a lateral-movement primitive: recover it from one
    device and you have administrative access to every device that shares it. It
    also cannot be rotated without editing and redeploying the script.

    Instead this generates a random password per device, which is never written to
    disk, never logged and never returned. That password is therefore unknown to
    everyone - including you - which is only useful if something else manages it.
    Deploy Windows LAPS against this account so LAPS takes ownership, rotates it on
    a schedule, and escrows the current value in Entra ID or Active Directory where
    it can be retrieved by an authorised admin.

    WITHOUT LAPS (OR AN EQUIVALENT) IN PLACE, THIS SCRIPT CREATES AN ACCOUNT YOU
    CANNOT LOG IN TO. That is a deliberate trade-off, not an oversight: an
    unusable break-glass account is a smaller problem than a fleet-wide shared one.
    Set up LAPS first.

.PARAMETER AccountName
    Local account to create or repair. Edit the default before deploying - Intune
    remediation scripts run with no arguments.

.PARAMETER FullName
    Display name applied when the account is created.

.PARAMETER Description
    Description applied when the account is created.

.PARAMETER PasswordLength
    Length of the generated password. Default 24.

.EXAMPLE
    # Deploy as-is after editing the default AccountName
    .\Repair-LocalAdminAccount.ps1

.NOTES
    Context  : run as SYSTEM (Intune remediation default), 64-bit
    Requires : Windows PowerShell 5.1, Microsoft.PowerShell.LocalAccounts
    Rights   : WRITES - creates a local account and grants it administrator rights.

    Follow-up, required:
      1. Configure Windows LAPS to manage this account
         (Intune > Endpoint security > Account protection > Local admin password solution),
         setting AdministratorAccountName to the same value as -AccountName.
      2. Confirm the password is being escrowed and can be retrieved.
      3. Only then rely on this account for break-glass access.

    This script does not set PasswordNeverExpires. LAPS owns password lifetime;
    pinning expiry here would work against it.

    Replaces (merged): Create-ITAdm.ps1, Create-AdminAV.ps1,
    Create-WIN-LocalAdmUsr.ps1, Set_itadm.ps1 (and its .txt duplicate)
    All four carried the same plaintext password in source. See _SECRETS-REPORT.md:
    that credential should be rotated regardless of whether this script is adopted.
#>

[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute(
    'PSAvoidUsingConvertToSecureStringWithPlainText', '',
    Justification = 'New-LocalUser -Password only accepts a SecureString. The plaintext is generated in-process by a cryptographic RNG, is never read from source, disk, or a parameter, and is discarded immediately after conversion. There is no stored credential to protect.')]
param(
    [string]$AccountName    = 'itadm',
    [string]$FullName       = 'IT Admin',
    [string]$Description    = 'Managed local administrator (password rotated by Windows LAPS)',
    [ValidateRange(14, 120)]
    [int]$PasswordLength    = 24
)

function New-RandomPassword {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        'PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Pure function: returns a string, changes no system state.')]
    [CmdletBinding()]
    <#
      Cryptographic RNG, not Get-Random: Get-Random is seeded from a predictable
      source and is not suitable for credentials.
      Guarantees at least one character from each class so the result satisfies
      complexity policy regardless of draw.
    #>
    param([int]$Length)

    $sets = @(
        'ABCDEFGHJKLMNPQRSTUVWXYZ'      # no I or O
        'abcdefghijkmnopqrstuvwxyz'     # no l
        '23456789'                      # no 0 or 1
        '!#%&*+-:=?@_'                  # avoids quotes/backtick that break parsing
    )
    $all = -join $sets

    $rng   = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        $bytes = [byte[]]::new(4)
        $pick  = {
            param([string]$Pool)
            $rng.GetBytes($bytes)
            $v = [BitConverter]::ToUInt32($bytes, 0)
            $Pool[[int]($v % [uint32]$Pool.Length)]
        }

        $chars = New-Object System.Collections.Generic.List[char]
        foreach ($s in $sets) { $chars.Add((& $pick $s)) }
        while ($chars.Count -lt $Length) { $chars.Add((& $pick $all)) }

        # Fisher-Yates shuffle so the guaranteed characters are not always first
        for ($i = $chars.Count - 1; $i -gt 0; $i--) {
            $rng.GetBytes($bytes)
            $j = [int]([BitConverter]::ToUInt32($bytes, 0) % [uint32]($i + 1))
            $tmp = $chars[$i]; $chars[$i] = $chars[$j]; $chars[$j] = $tmp
        }
        return (-join $chars)
    }
    finally { $rng.Dispose() }
}

$changed = @()

try {
    # ---- account exists? ----
    $user = $null
    try { $user = Get-LocalUser -Name $AccountName -ErrorAction Stop } catch { $user = $null }

    if (-not $user) {
        $plain  = New-RandomPassword -Length $PasswordLength
        $secure = ConvertTo-SecureString -String $plain -AsPlainText -Force
        # overwrite the plaintext as soon as the SecureString exists
        $plain  = $null
        [System.GC]::Collect()

        New-LocalUser -Name $AccountName -Password $secure -FullName $FullName `
                      -Description $Description -ErrorAction Stop | Out-Null
        $changed += 'created'
        $user = Get-LocalUser -Name $AccountName -ErrorAction Stop
        Write-Output "Created local account '$AccountName' with a random password (not recorded)."
    }

    # ---- enabled? ----
    if (-not $user.Enabled) {
        Enable-LocalUser -Name $AccountName -ErrorAction Stop
        $changed += 'enabled'
        Write-Output "Enabled local account '$AccountName'."
    }

    # ---- administrator? ----
    $adminSid   = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')
    $adminGroup = ($adminSid.Translate([System.Security.Principal.NTAccount]).Value -split '\\')[-1]

    $isAdmin = $false
    foreach ($m in @(Get-LocalGroupMember -Group $adminGroup -ErrorAction Stop)) {
        if ($m.ObjectClass -ne 'User') { continue }
        if ((($m.Name -split '\\')[-1]) -ieq $AccountName) { $isAdmin = $true; break }
    }

    if (-not $isAdmin) {
        Add-LocalGroupMember -Group $adminGroup -Member $AccountName -ErrorAction Stop
        $changed += 'granted admin'
        Write-Output "Added '$AccountName' to local group '$adminGroup'."
    }

    if ($changed.Count -eq 0) {
        Write-Output "No change needed: '$AccountName' already compliant."
    } else {
        Write-Output "Remediated '$AccountName': $($changed -join ', ')."
    }
    Write-Output 'Reminder: Windows LAPS must manage this account for the password to be retrievable.'
    exit 0
}
catch {
    # Report the failure honestly. Intune shows a non-zero exit as failed
    # remediation, which is the correct signal - do not swallow this.
    Write-Output "ERROR: $($_.Exception.Message)"
    exit 1
}
