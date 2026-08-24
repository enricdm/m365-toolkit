<#
.SYNOPSIS
    Provisions a shared team calendar in Exchange Online, following a
    group-based access convention.

.DESCRIPTION
    Convention (adapt the defaults to your own standard):
      DisplayName : <ISO2>.<CITY>.SHARED <Name>          e.g. DE.CTY.SHARED Example-Team
      Alias       : shared.<ISO2>.<CITY>.<NameNoSpaces>  e.g. shared.DE.CTY.Example-Team
      Access      : mail-enabled SECURITY group MB_<alias>_Kalender_Editor,
                    stamped Editor once on the calendar folder; people are managed
                    as members of that group, not as individual folder ACEs.

    The point of the group model is that day-to-day membership changes never
    touch the mailbox: you add or remove someone from one group instead of
    re-stamping folder ACEs.

    This script:
      - optionally creates the shared mailbox (-CreateMailbox)
      - optionally applies regional config (localises folder names, e.g. Kalender)
      - resolves the REAL calendar folder name (never assumes \Calendar)
      - resolves the access group, verifies it is SECURITY-enabled, syncs its members
      - stamps Editor for the group on the calendar folder
      - falls back to direct per-user stamping with -DirectGrant
      - optionally dumps a reference mailbox's config for comparison
      - writes a timestamped CSV of everything it did / would do

    Dry-run is the default. Nothing changes until -Execute is passed.

.NOTES
    When to use  : Provisioning a shared team calendar, when you want the next membership change to require running nothing at all.
    Why it exists: Access is stamped once for a mail-enabled security group instead of per user, so day-to-day membership changes never touch the mailbox. It resolves the REAL calendar folder name rather than assuming \Calendar, which is where scripts break in a multi-language tenant, and documents the hybrid sequence for creating the access group on-prem.
    AUTHENTICATION
    Two supported modes, chosen by what you pass:
      - app-only certificate: supply -AppId, -CertificateThumbprint and
        -Organization (or set EXO_APPID / EXO_CERTTHUMB / EXO_ORG, which are the
        parameter defaults). Use this for unattended runs.
      - interactive / delegated: supply none of the three. The signed-in account
        needs Exchange Administrator (or Recipient Administrator + the folder
        permission rights).
    The banner states which mode was actually used - it is not assumed.

    HYBRID TENANTS
    Mail-enabled security groups can no longer be created in Exchange Online, so
    in a hybrid tenant the MB_<alias>_Kalender_Editor group MUST be created in
    on-prem AD. Recommended sequence:
      1. create the shared mailbox on-prem (remote mailbox) -> let it sync
      2. create MB_<alias>_Kalender_Editor in on-prem AD, add its members
      3. wait for the directory sync cycle
      4. run this script (no -CreateMailbox) to stamp the calendar permission
    Use -CreateMailbox / -CreateAccessGroup only for a cloud-only object.

.EXAMPLE
    .\New-SharedCalendar.ps1
    .\New-SharedCalendar.ps1 -ReferenceMailbox shared.DE.CTY.Other-Team@contoso.com
    .\New-SharedCalendar.ps1 -Execute
    .\New-SharedCalendar.ps1 -Execute -DirectGrant

.EXAMPLE
    # Unattended, app-only
    .\New-SharedCalendar.ps1 -Execute -AppId '<client-id>' `
        -CertificateThumbprint '<cert-thumbprint>' -Organization 'contoso.onmicrosoft.com'
#>

[CmdletBinding()]
param(
    [switch]$Execute,
    [switch]$CreateMailbox,
    [switch]$CreateAccessGroup,
    [switch]$DirectGrant,
    [switch]$GrantFullAccess,
    [string]$ReferenceMailbox,
    [ValidateSet('Reviewer','Author','Editor','PublishingEditor','Owner')]
    [string]$AccessLevel = 'Editor',

    # ---- the calendar ----
    [string]$Domain      = 'contoso.com',
    [string]$DisplayName = 'DE.CTY.SHARED Example-Team',
    [string]$Alias       = 'shared.DE.CTY.Example-Team',
    [string]$PrimarySmtp,                                   # defaults to <Alias>@<Domain>

    # People. In group mode these become members of the access group; with
    # -DirectGrant they are stamped individually on the folder instead.
    [string[]]$Members = @(
        'first.user@contoso.com'
        'second.user@contoso.com'
    ),

    # ---- EXO app-only cert auth (leave empty for interactive/delegated) ----
    [string]$AppId                 = $env:EXO_APPID,
    [string]$CertificateThumbprint = $env:EXO_CERTTHUMB,
    [string]$Organization          = $env:EXO_ORG
)

# ============================== CONFIG ==============================

if (-not $PrimarySmtp) { $PrimarySmtp = "$Alias@$Domain" }

# Mail-enabled SECURITY group carrying the calendar permission.
$AccessGroupName = "MB_${Alias}_Kalender_Editor"
$AccessGroupSmtp = "$AccessGroupName@$Domain"

# Optional explicit Owner ACE on top of the group. Off by default: editing rights
# come purely via the MB_*_Kalender_Editor group, which is the whole point of the
# group model. Set to an address to also give one person Owner on the folder.
$FolderOwner = $null

# AvailabilityOnly, not None: the rest of the org sees free/busy but no details.
# That is usually what a shared team calendar wants.
$DefaultFolderAccess = 'AvailabilityOnly'
$HideFromGAL         = $false      # keep visible so users can self-add the calendar
$SetRegional         = $true
$RegionalLanguage    = 'de-DE'
$RegionalTimeZone    = 'W. Europe Standard Time'
$OrganizationalUnit  = $null

$ExportDir = Join-Path $PSScriptRoot 'Exports'

# ============================== HELPERS =============================

$script:runMode = if ($Execute) { 'EXECUTE' } else { 'DRY-RUN' }
$script:model   = if ($DirectGrant) { 'Direct ACE' } else { 'Access group' }
$script:appOnly = -not ([string]::IsNullOrWhiteSpace($AppId) -or
                        [string]::IsNullOrWhiteSpace($CertificateThumbprint) -or
                        [string]::IsNullOrWhiteSpace($Organization))

function Write-Step { param([string]$m) Write-Host "[STEP] $m" -ForegroundColor Cyan }
function Write-OK   { param([string]$m) Write-Host "[ OK ] $m" -ForegroundColor Green }
function Write-Warn { param([string]$m) Write-Host "[WARN] $m" -ForegroundColor Yellow }
function Write-Dry  { param([string]$m) Write-Host "[DRY ] would: $m" -ForegroundColor DarkGray }
function Write-Die  { param([string]$m) Write-Host "[FAIL] $m" -ForegroundColor Red; exit 1 }

function Invoke-Action {
    param([string]$Description, [scriptblock]$Action)
    if (-not $Execute) { Write-Dry $Description; return 'DRY-RUN' }
    try   { & $Action | Out-Null; Write-OK $Description; return 'OK' }
    catch { Write-Warn "$Description -> $($_.Exception.Message)"; return "FAIL: $($_.Exception.Message)" }
}

$results = New-Object System.Collections.Generic.List[object]
function Add-Result {
    param($Principal, $Target, $Action, $Status)
    $results.Add([pscustomobject]@{
        Timestamp = (Get-Date).ToString('dd/MM/yyyy HH:mm:ss')
        Principal = $Principal; Target = $Target; Action = $Action; Status = $Status
    })
}

function Resolve-CalendarPath {
    param([string]$Smtp)
    try {
        $f = Get-MailboxFolderStatistics -Identity $Smtp -FolderScope Calendar -ErrorAction Stop |
             Where-Object { $_.FolderType -eq 'Calendar' } | Select-Object -First 1
        if ($f) { return "$Smtp`:" + ($f.FolderPath -replace '/', '\') }
    } catch { }
    return $null
}

# ============================== BANNER ==============================

Write-Host ""
Write-Host "===== New-SharedCalendar  [$script:runMode] [$script:model] =====" -ForegroundColor White
Write-Host "  Display name : $DisplayName"
Write-Host "  Alias        : $Alias"
Write-Host "  Primary SMTP : $PrimarySmtp"
if (-not $DirectGrant) { Write-Host "  Access group : $AccessGroupName" }
Write-Host "  Access level : $AccessLevel  (Default = $DefaultFolderAccess)"
Write-Host "  Auth mode    : $(if ($script:appOnly) { "app-only (AppId $AppId)" } else { 'interactive / delegated' })"
Write-Host ""

# ============================== CONNECT =============================
# App-only when all three of -AppId / -CertificateThumbprint / -Organization are
# supplied; interactive otherwise. Partial app-only config is rejected rather
# than silently falling back, because that is how you end up believing you ran
# unattended when you did not.

if (-not $script:appOnly) {
    $partial = @($AppId, $CertificateThumbprint, $Organization) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    if ($partial.Count -gt 0) {
        Write-Die 'Incomplete app-only configuration: -AppId, -CertificateThumbprint and -Organization must all be set (or all be empty for interactive auth).'
    }
}

Write-Step "Connecting to Exchange Online ($(if ($script:appOnly) { 'app-only' } else { 'interactive' }))"
try {
    if ($script:appOnly) {
        Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $CertificateThumbprint `
            -Organization $Organization -ShowBanner:$false -ErrorAction Stop
        Write-OK "Connected app-only as $AppId against $Organization"
    } else {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-OK 'Connected interactively (delegated permissions of the signed-in account)'
    }
} catch { Write-Die "Connection failed: $($_.Exception.Message)" }

# ========================= REFERENCE DUMP ===========================

if ($ReferenceMailbox) {
    Write-Step "Reference object: $ReferenceMailbox"
    $ref = Get-Mailbox -Identity $ReferenceMailbox -ErrorAction SilentlyContinue
    if ($ref) {
        $ref | Format-List Name,Alias,DisplayName,PrimarySmtpAddress,RecipientTypeDetails,
                           IsDirSynced,HiddenFromAddressListsEnabled,OrganizationalUnit
        $refCal = Resolve-CalendarPath $ref.PrimarySmtpAddress.ToString()
        if ($refCal) {
            Write-OK "Reference calendar folder: $refCal"
            Get-MailboxFolderPermission -Identity $refCal -ErrorAction SilentlyContinue |
                Select-Object User,AccessRights,SharingPermissionFlags | Format-Table -AutoSize
        }
    } else { Write-Warn "Reference mailbox $ReferenceMailbox not found." }
    Write-Host ""
}

# ========================= MAILBOX PROVISION ========================

Write-Step "Checking for mailbox $PrimarySmtp"
$mbx = Get-Mailbox -Identity $PrimarySmtp -ErrorAction SilentlyContinue

if (-not $mbx) {
    if (-not $CreateMailbox) {
        Write-Warn 'Mailbox not found. In a hybrid tenant, create it on-prem and let it sync,'
        Write-Warn 'or re-run with -CreateMailbox for a cloud-only object. Preview only below.'
    } else {
        $newParams = @{ Name = $Alias; DisplayName = $DisplayName; Alias = $Alias
                        PrimarySmtpAddress = $PrimarySmtp; Shared = $true }
        if ($OrganizationalUnit) { $newParams['OrganizationalUnit'] = $OrganizationalUnit }

        $st = Invoke-Action "New-Mailbox -Shared '$DisplayName'" { New-Mailbox @newParams -ErrorAction Stop }
        Add-Result $Alias $PrimarySmtp 'New-Mailbox -Shared' $st

        if ($Execute -and $st -eq 'OK') {
            Write-Step 'Waiting for mailbox to become addressable (up to 120s)'
            $wait = 0
            do { Start-Sleep -Seconds 10; $wait += 10
                 $mbx = Get-Mailbox -Identity $PrimarySmtp -ErrorAction SilentlyContinue
            } while (-not $mbx -and $wait -lt 120)
            if ($mbx) { Write-OK "Mailbox available after ${wait}s" }
            else      { Write-Die 'Not addressable yet. Re-run without -CreateMailbox shortly.' }
        }
    }
} else {
    Write-OK "Mailbox exists ($($mbx.RecipientTypeDetails), IsDirSynced=$($mbx.IsDirSynced))"
    if ($mbx.RecipientTypeDetails -ne 'SharedMailbox') {
        Write-Warn "Expected SharedMailbox, found $($mbx.RecipientTypeDetails)."
    }
    if ($mbx.Alias -ne $Alias) {
        Write-Warn "Alias drifted: expected '$Alias', found '$($mbx.Alias)'. On-prem sAMAccountName truncation?"
    }
}

# ========================== MAILBOX SETTINGS ========================

if ($mbx -or -not $Execute) {

    if ($SetRegional) {
        Write-Step "Applying regional config ($RegionalLanguage / $RegionalTimeZone)"
        $st = Invoke-Action "Set-MailboxRegionalConfiguration Language=$RegionalLanguage, TZ=$RegionalTimeZone, localise folder names" {
            Set-MailboxRegionalConfiguration -Identity $PrimarySmtp -Language $RegionalLanguage `
                -TimeZone $RegionalTimeZone -LocalizeDefaultFolderName:$true -ErrorAction Stop
        }
        Add-Result $Alias $PrimarySmtp 'Set-MailboxRegionalConfiguration' $st
    }

    if ($HideFromGAL) {
        $st = Invoke-Action 'Set-Mailbox HiddenFromAddressListsEnabled=$true' {
            Set-Mailbox -Identity $PrimarySmtp -HiddenFromAddressListsEnabled $true -ErrorAction Stop
        }
        Add-Result $Alias $PrimarySmtp 'HideFromGAL' $st
    }
}

# ======================= RESOLVE CALENDAR FOLDER ====================

Write-Step 'Resolving calendar folder (localised name aware)'
$calPath = $null
if ($mbx) { $calPath = Resolve-CalendarPath $PrimarySmtp }

if ($calPath) { Write-OK "Calendar folder: $calPath" }
else {
    $calPath = "$PrimarySmtp`:\Kalender"
    Write-Warn "Could not resolve folder - assuming $calPath for preview"
}

# ========================== ACCESS GROUP ============================

$groupPrincipal = $null

if (-not $DirectGrant) {
    Write-Step "Resolving access group $AccessGroupName"
    $grp = Get-Recipient -Identity $AccessGroupSmtp -ErrorAction SilentlyContinue
    if (-not $grp) { $grp = Get-Recipient -Identity $AccessGroupName -ErrorAction SilentlyContinue }

    if (-not $grp) {
        if ($CreateAccessGroup) {
            Write-Warn 'Mail-enabled security groups can no longer be created in Exchange Online.'
            Write-Warn 'Create MB_..._Kalender_Editor in on-prem AD and let directory sync bring it in.'
            $st = Invoke-Action "New-DistributionGroup -Type Security '$AccessGroupName' (may be blocked by EXO)" {
                New-DistributionGroup -Name $AccessGroupName -Alias $AccessGroupName `
                    -PrimarySmtpAddress $AccessGroupSmtp -Type Security `
                    -MemberJoinRestriction Closed -MemberDepartRestriction Closed -ErrorAction Stop
            }
            Add-Result $AccessGroupName '-' 'New-DistributionGroup -Type Security' $st
            if ($Execute -and $st -eq 'OK') {
                Start-Sleep -Seconds 15
                $grp = Get-Recipient -Identity $AccessGroupSmtp -ErrorAction SilentlyContinue
            }
        } else {
            Write-Warn "$AccessGroupName not found. Create it on-prem, or use -DirectGrant."
        }
    }

    if ($grp) {
        $groupPrincipal = $grp.PrimarySmtpAddress.ToString()
        Write-OK "Found $groupPrincipal [$($grp.RecipientTypeDetails)]"

        if ($grp.RecipientTypeDetails -notmatch 'SecurityGroup') {
            Write-Warn "$groupPrincipal is NOT security-enabled - folder permissions will NOT reach its members."
            Write-Warn 'Convert it in on-prem AD, or re-run with -DirectGrant.'
            $groupPrincipal = $null
        }
    }

    # --- sync membership -------------------------------------------------
    if ($groupPrincipal) {
        Write-Step 'Syncing access group membership'
        $current = @()
        try {
            $current = Get-DistributionGroupMember -Identity $groupPrincipal -ResultSize Unlimited -ErrorAction Stop |
                       ForEach-Object { $_.PrimarySmtpAddress.ToString().ToLower() }
        } catch { Write-Warn "Could not read members: $($_.Exception.Message)" }

        foreach ($upn in $Members) {
            $rcpt = Get-Recipient -Identity $upn -ErrorAction SilentlyContinue
            if (-not $rcpt) {
                Write-Warn "$upn -> not found, skipped"
                Add-Result $upn $groupPrincipal 'GroupMember' 'FAIL: recipient not found'
                continue
            }
            $resolved = $rcpt.PrimarySmtpAddress.ToString()

            if ($current -contains $resolved.ToLower()) {
                Write-OK "$resolved already a member"
                Add-Result $resolved $groupPrincipal 'GroupMember' 'OK (already present)'
                continue
            }
            $st = Invoke-Action "Add-DistributionGroupMember $resolved -> $AccessGroupName" {
                Add-DistributionGroupMember -Identity $groupPrincipal -Member $resolved -ErrorAction Stop
            }
            Add-Result $resolved $groupPrincipal 'GroupMember' $st
        }

        if ($grp.IsDirSynced) {
            Write-Warn 'Group is dir-synced: membership must be changed in on-prem AD, not here.'
        }
    }
}

# ========================= DEFAULT PERMISSION =======================

Write-Step "Setting Default calendar permission to $DefaultFolderAccess"
$st = Invoke-Action "Set-MailboxFolderPermission Default=$DefaultFolderAccess on $calPath" {
    Set-MailboxFolderPermission -Identity $calPath -User Default -AccessRights $DefaultFolderAccess -ErrorAction Stop
}
Add-Result 'Default' $calPath "FolderPermission=$DefaultFolderAccess" $st

# ========================= FOLDER PERMISSIONS =======================

function Grant-CalendarAccess {
    param([string]$Principal, [string]$Level)

    if (-not $Execute) {
        Write-Dry "$Level -> $Principal on $calPath"
        Add-Result $Principal $calPath "FolderPermission=$Level" 'DRY-RUN'
        return
    }
    try {
        Add-MailboxFolderPermission -Identity $calPath -User $Principal -AccessRights $Level -ErrorAction Stop | Out-Null
        Write-OK "$Level -> $Principal (added)"
        Add-Result $Principal $calPath "FolderPermission=$Level" 'OK (added)'
    } catch {
        if ($_.Exception.Message -match 'existing permission|already exist') {
            try {
                Set-MailboxFolderPermission -Identity $calPath -User $Principal -AccessRights $Level -ErrorAction Stop | Out-Null
                Write-OK "$Level -> $Principal (updated existing)"
                Add-Result $Principal $calPath "FolderPermission=$Level" 'OK (updated)'
            } catch {
                Write-Warn "$Principal -> $($_.Exception.Message)"
                Add-Result $Principal $calPath "FolderPermission=$Level" "FAIL: $($_.Exception.Message)"
            }
        } else {
            Write-Warn "$Principal -> $($_.Exception.Message)"
            Add-Result $Principal $calPath "FolderPermission=$Level" "FAIL: $($_.Exception.Message)"
        }
    }
}

if ($DirectGrant) {
    Write-Step 'Granting calendar permissions per user (direct ACE mode)'
    foreach ($upn in $Members) {
        $rcpt = Get-Recipient -Identity $upn -ErrorAction SilentlyContinue
        if (-not $rcpt) {
            Write-Warn "$upn -> not found, skipped"
            Add-Result $upn $calPath "FolderPermission=$AccessLevel" 'FAIL: recipient not found'
            continue
        }
        Grant-CalendarAccess $rcpt.PrimarySmtpAddress.ToString() $AccessLevel
    }
} elseif ($groupPrincipal) {
    Write-Step 'Granting calendar permission to the access group'
    Grant-CalendarAccess $groupPrincipal $AccessLevel
} else {
    Write-Warn 'No usable access group and -DirectGrant not set - no member permissions applied.'
}

# Explicit owner ACE on top, so the requester can manage sharing himself.
if ($FolderOwner) {
    Write-Step "Granting Owner to $FolderOwner"
    $ownerRcpt = Get-Recipient -Identity $FolderOwner -ErrorAction SilentlyContinue
    if ($ownerRcpt) { Grant-CalendarAccess $ownerRcpt.PrimarySmtpAddress.ToString() 'Owner' }
    else { Write-Warn "$FolderOwner not found, skipped" }
}

# ============================ FULL ACCESS ===========================

if ($GrantFullAccess) {
    Write-Step 'Granting FullAccess + AutoMapping'
    $faTargets = if ($DirectGrant -or -not $groupPrincipal) { $Members } else { @($groupPrincipal) }
    foreach ($t in $faTargets) {
        $st = Invoke-Action "FullAccess (AutoMapping) -> $t" {
            Add-MailboxPermission -Identity $PrimarySmtp -User $t -AccessRights FullAccess `
                -InheritanceType All -AutoMapping $true -ErrorAction Stop
        }
        Add-Result $t $PrimarySmtp 'FullAccess+AutoMapping' $st
    }
}

# ============================== SUMMARY =============================

Write-Step 'Summary'
$results | Format-Table Principal, Target, Action, Status -AutoSize

if ($Execute -and $mbx) {
    Write-Step 'Resulting calendar permissions'
    Get-MailboxFolderPermission -Identity $calPath -ErrorAction SilentlyContinue |
        Select-Object User,AccessRights | Format-Table -AutoSize
}

if (-not (Test-Path $ExportDir)) { New-Item -Path $ExportDir -ItemType Directory -Force | Out-Null }
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$csv   = Join-Path $ExportDir "SharedCalendar_${Alias}_${stamp}.csv"
$results | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
Write-OK "Result written to $csv"

if (-not $Execute) {
    Write-Host ""
    Write-Host 'Dry-run only. Re-run with -Execute to apply.' -ForegroundColor Yellow
}