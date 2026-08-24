<#
.SYNOPSIS
    Creates a cloud-only administrator (ADM) account derived from a person's
    ordinary ("nominal") account.

.DESCRIPTION
    Input ONLY the nominal account. The script derives the rest from it,
    proposes the ADM UPN for confirmation, then:
      - creates the cloud-only ADM account
      - sets Manager + ExtensionAttribute1 (link to the human)
      - issues a Temporary Access Pass (TAP)
      - enforces per-user MFA
      - optionally adds it to a Conditional Access group (prompted)
      - prints a ready-to-send message for the end user

    STILL MANUAL (out of scope): PIM eligible role assignment. Pass the role you
    assign by hand in -AssignedRoles so it appears in the end-user message.

    WRITES TO THE DIRECTORY. It creates a user, changes group membership and
    enables per-user MFA. It asks for confirmation before creating the account,
    but there is no -WhatIf mode.

    --- ROLES REQUIRED (activate in PIM before running) ---------------------
      User Administrator                 -> create account, manager, ext.attribute
      Authentication Administrator       -> issue the TAP
      Authentication Policy Administrator-> enforce per-user MFA
      Groups Administrator               -> add to Conditional Access group
      Privileged Role Administrator      -> assign PIM role (the manual step)

    Prerequisite connection:
      Connect-MgGraph -Scopes "User.ReadWrite.All","Directory.ReadWrite.All",
        "UserAuthenticationMethod.ReadWrite.All","Policy.ReadWrite.AuthenticationMethod",
        "Group.ReadWrite.All"

.PARAMETER NominalUpn
    UPN of the person's ordinary account. Everything else is derived from it.

.PARAMETER AdminUpnSuffix
    Domain the ADM account is created on, e.g. contoso.onmicrosoft.com.

.PARAMETER AdminNameOverride
    Force the name token instead of deriving it from the surname.

.PARAMETER AssignedRoles
    Free text listing the PIM role(s) assigned manually; shown in the end-user message.

.PARAMETER TapLifetimeMinutes
    Temporary Access Pass validity, in minutes. Default 720 (12 h). Must also be
    within the tenant's TAP policy maximum. Keep this short: a TAP is a valid
    credential for its whole lifetime.

.PARAMETER ConditionalAccessGroupId
    Object ID of the Conditional Access group offered for accounts that sign in
    from outside. Omit to skip that step entirely.

.PARAMETER ShowCredentials
    Print the temporary password and TAP in clear, plus the ready-to-send end-user
    message. OFF by default: without it both values are masked, so running the
    script does not leave credentials in console scrollback or a PowerShell
    transcript. Turn it on only when you are about to hand the message over.

.EXAMPLE
    .\New-AdminAccount.ps1 -NominalUpn 'user@contoso.com' -AdminUpnSuffix 'contoso.onmicrosoft.com'

.EXAMPLE
    .\New-AdminAccount.ps1 -NominalUpn 'user@contoso.com' `
        -ConditionalAccessGroupId '<group-object-id>' -AssignedRoles 'User Administrator' -ShowCredentials

.NOTES
    When to use  : Someone joins the administration team and needs their ADM account without forgetting the TAP, per-user MFA or the link back to their ordinary account.
    Why it exists: Derives everything from the nominal account - ASCII-safe name, country, line manager, ExtensionAttribute1 link - and retries the TAP because a freshly created user lags in replication. Nine portal steps in one, and credentials are masked by default so a normal run leaves nothing usable in console scrollback or a transcript.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$NominalUpn,

    [string]$AdminUpnSuffix = 'contoso.onmicrosoft.com',

    [string]$AdminNameOverride = '',

    [string]$AssignedRoles = '',

    # 720 minutes = 12 h. Anything longer is a credential that stays valid for days.
    [ValidateRange(10, 1440)]
    [int]$TapLifetimeMinutes = 720,

    [string]$ConditionalAccessGroupId,

    [string]$ConditionalAccessGroupName = 'Conditional Access',

    [switch]$ShowCredentials
)

$nominalUPN      = $NominalUpn
$admNameOverride = $AdminNameOverride
$assignedRoles   = $AssignedRoles
$tapLifetimeMin  = $TapLifetimeMinutes
$caGroupId       = $ConditionalAccessGroupId

# --- Strip accents/diacritics for an ASCII-safe UPN (Zuñiga -> Zuniga) ----
function Remove-Diacritics {
    param([string]$Text)
    $n = $Text.Normalize([Text.NormalizationForm]::FormD)
    -join ($n.ToCharArray() | Where-Object {
            [Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne [Globalization.UnicodeCategory]::NonSpacingMark
        })
}

# --- Generate a compliant temp password (16 chars, all categories) --------
function New-TempPassword {
    $u = (65..90)  | Get-Random -Count 4 | ForEach-Object { [char]$_ }
    $l = (97..122) | Get-Random -Count 6 | ForEach-Object { [char]$_ }
    $d = (48..57)  | Get-Random -Count 4 | ForEach-Object { [char]$_ }
    $s = ('!', '#', '$', '%', '&', '*', '-', '_') | Get-Random -Count 2
    -join (($u + $l + $d + $s) | Get-Random -Count 16)
}

# 1. Look up the nominal account and derive everything
$nominal = Get-MgUser -UserId $nominalUPN -Property "givenName,surname,usageLocation,mail,userPrincipalName,id"
if (-not $nominal) { Write-Host "Nominal account not found: $nominalUPN" -ForegroundColor Red; return }

$firstNameReal = $nominal.GivenName
$lastNameReal = $nominal.Surname
$usageLocation = $nominal.UsageLocation
$countryCode = $nominal.UsageLocation
$nominalEmail = if ($nominal.Mail) { $nominal.Mail } else { $nominal.UserPrincipalName }

# Manager of the ADM account = the nominal account's own line manager (org chart)
$mgr = $null; $mgrUPN = $null
try {
    $mgr = Get-MgUserManager -UserId $nominalUPN -ErrorAction Stop
    $mgrUPN = (Get-MgUser -UserId $mgr.Id -Property userPrincipalName).UserPrincipalName
} catch {
    Write-Host "WARNING: nominal account has no manager set — the ADM account will be created WITHOUT a manager. Set one in the org chart, or assign it manually after." -ForegroundColor Yellow
}

foreach ($f in @{Surname = $lastNameReal; "Usage location" = $usageLocation; "First name" = $firstNameReal }.GetEnumerator()) {
    if ([string]::IsNullOrWhiteSpace($f.Value)) {
        Write-Host "WARNING: nominal account has no $($f.Key) set — fix it or use the override." -ForegroundColor Yellow
    }
}

$nameToken = if ($admNameOverride) { $admNameOverride } else { (Remove-Diacritics $lastNameReal) -replace '\s', '' }
$upn = "$countryCode-ADM-$nameToken@$AdminUpnSuffix"
$displayName = "$countryCode-ADM-$nameToken"
$mailNick = "$countryCode-ADM-$nameToken"

# 2. Show proposal and confirm
Write-Host "`n--- Proposed ADM account (from $nominalUPN) ---" -ForegroundColor Cyan
Write-Host "  UPN / Display:  $upn"
Write-Host "  First / Last:   $firstNameReal / $lastNameReal"
Write-Host "  Usage location: $usageLocation"
Write-Host "  Manager:        $(if ($mgrUPN) { $mgrUPN } else { '(none — nominal account has no manager)' })"
Write-Host "  Ext.Attribute1: $nominalEmail"
Write-Host "  TAP lifetime:   $tapLifetimeMin min"
if (-not $ShowCredentials) {
    Write-Host "  NOTE: the temporary password and TAP will NOT be displayed. Cancel and re-run" -ForegroundColor Yellow
    Write-Host "        with -ShowCredentials if you need to hand them to the user now." -ForegroundColor Yellow
}
if ((Read-Host "`nCreate this account? (Y/N)") -ne 'Y') { Write-Host "Cancelled." -ForegroundColor Yellow; return }

# 3. Create the cloud-only ADM account
$tempPassword = New-TempPassword
$adm = New-MgUser `
    -AccountEnabled `
    -DisplayName       $displayName `
    -MailNickname      $mailNick `
    -UserPrincipalName $upn `
    -GivenName         $firstNameReal `
    -Surname           $lastNameReal `
    -UsageLocation     $usageLocation `
    -PasswordProfile   @{ Password = $tempPassword; ForceChangePasswordNextSignIn = $true }
Write-Host "Created: $upn" -ForegroundColor Green

# 4. Manager = the nominal account's line manager (from the org chart)
if ($mgr) {
    Set-MgUserManagerByRef -UserId $adm.Id -BodyParameter @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($mgr.Id)" }
    Write-Host "Manager set: $mgrUPN" -ForegroundColor Green
} else {
    Write-Host "Manager NOT set — nominal account had no manager. Assign one manually." -ForegroundColor Yellow
}

# 5. ExtensionAttribute1 = nominal email
Update-MgUser -UserId $adm.Id -OnPremisesExtensionAttributes @{ ExtensionAttribute1 = $nominalEmail }
Write-Host "ExtensionAttribute1 set: $nominalEmail" -ForegroundColor Green

# 6. Issue the TAP (one-time use) — retry, as a freshly-created user can lag in replication
$tapCode = $null
$maxTries = 6
for ($i = 1; $i -le $maxTries; $i++) {
    try {
        $tap = New-MgUserAuthenticationTemporaryAccessPassMethod -UserId $adm.Id -IsUsableOnce:$true -LifetimeInMinutes $tapLifetimeMin -ErrorAction Stop
        $tapCode = $tap.TemporaryAccessPass
        Write-Host "TAP generated (valid $tapLifetimeMin min)" -ForegroundColor Green
        break
    } catch {
        if ($i -eq $maxTries) {
            Write-Host "TAP FAILED after $maxTries tries -> $($_.Exception.Message)" -ForegroundColor Red
        } else {
            Write-Host "TAP not ready (try $i/$maxTries) — waiting for replication..." -ForegroundColor DarkGray
            Start-Sleep -Seconds 10
        }
    }
}

# 7. Enforce per-user MFA (beta endpoint via the existing session — no extra module)
try {
    Invoke-MgGraphRequest -Method PATCH `
        -Uri "https://graph.microsoft.com/beta/users/$($adm.Id)/authentication/requirements" `
        -Body (@{ perUserMfaState = "enforced" } | ConvertTo-Json) -ContentType "application/json" -ErrorAction Stop
    Write-Host "Per-user MFA: ENFORCED" -ForegroundColor Green
} catch {
    Write-Host "Per-user MFA enforce FAILED -> $($_.Exception.Message)" -ForegroundColor Red
}

# 8. Conditional Access group — only if the account needs external access
if (-not $caGroupId) {
    Write-Host "No -ConditionalAccessGroupId supplied — skipping the Conditional Access group step." -ForegroundColor DarkGray
} elseif ((Read-Host "Will this account access from OUTSIDE (external)? (Y/N)") -eq 'Y') {
    try {
        New-MgGroupMember -GroupId $caGroupId -DirectoryObjectId $adm.Id -ErrorAction Stop
        Write-Host "Added to Conditional Access group ($ConditionalAccessGroupName)" -ForegroundColor Green
    } catch {
        Write-Host "CA group add FAILED -> $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "Skipped Conditional Access group (internal only)." -ForegroundColor DarkGray
}

# 9. Summary + ready-to-send end-user message
#    Credentials are masked unless -ShowCredentials is passed, so a normal run
#    leaves nothing usable in console scrollback or a PowerShell transcript.
$maskedPassword = if ($ShowCredentials) { $tempPassword } else { '(not displayed — see -ShowCredentials)' }
$maskedTap      = if ($ShowCredentials) { $tapCode }      else { '(not displayed — see -ShowCredentials)' }

Write-Host "`n=================== ADM ACCOUNT READY ===================" -ForegroundColor Cyan
Write-Host "UPN:            $upn"
Write-Host "Temp password:  $maskedPassword"
Write-Host "TAP:            $maskedTap  (valid $tapLifetimeMin min)"
Write-Host "Reminder: assign the approved PIM role as ELIGIBLE in the portal." -ForegroundColor Yellow

$endUserMessage = @"
Dear $firstNameReal,

We have successfully created a new Azure ADM Account for you. Please see the details below:

    - Username:  $upn
    - Temporary Password:  $tempPassword

Role Assignment: $assignedRoles

    1. Please set up the passkey for this account first: open the Microsoft Authenticator app on your mobile device, select "Add account", and sign in with your $upn account.
    2. It will ask for the TAP/Temporary Password (one-time use). Please use this TAP:    $tapCode
    3. Once done, log in to https://portal.azure.com/ to change the temporary password to one of your choosing, and use MFA / Microsoft Authenticator / Passkey for your ADM account.

Thank you,
Identity & Collaboration Team
"@

if ($ShowCredentials) {
    Write-Host "`n----------------- COPY-PASTE TO END USER -----------------" -ForegroundColor Magenta
    Write-Host $endUserMessage
    Write-Host "----------------------------------------------------------" -ForegroundColor Magenta
    Write-Host "The password and TAP above are live credentials. Send them over an agreed channel and clear your console." -ForegroundColor Yellow
} else {
    Write-Host "`nEnd-user message not printed: it contains the temporary password and TAP." -ForegroundColor DarkGray
    Write-Host "The account exists and the TAP is live. Reset the password and issue a fresh TAP" -ForegroundColor DarkGray
    Write-Host "from the portal, or delete the account and re-run with -ShowCredentials." -ForegroundColor DarkGray
}