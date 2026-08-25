#Requires -Version 7.0

<#
.SYNOPSIS
    Configure (and optionally create) a resource room mailbox following a
    CC.CITY.SITE.<Room name> naming standard.

.DESCRIPTION
    Dry-run by default. Nothing is changed unless you pass -Execute.

    Two access models, chosen with -Mode:
      Booking : restrict WHO MAY RESERVE the room (BookInPolicy). Groups are
                accepted as-is - Exchange evaluates their membership at request
                time - so this survives staff changes without re-running.
      Manage  : grant Editor on the room's calendar folder instead, for people
                who need to fix up other people's bookings.
    Folder permissions do NOT reach the members of a distribution list, which is
    the trap in Manage mode; the script warns and offers -ExpandGroups, which
    flattens groups and shared mailboxes to individuals (cycle-guarded).

    PROVISIONING NOTE (hybrid):
    If your room estate is provisioned ON-PREM as remote room mailboxes
    (on-prem legacyExchangeDN + <alias>@<tenant>.mail.onmicrosoft.com routing),
    create the object on-prem first, let it sync, then run this script to
    configure it. On-prem one-liner (run in on-prem EMS, single line):

      New-RemoteMailbox -Room -Name "DE.CTY.SITE.Room 1.01" -DisplayName "DE.CTY.SITE.Room 1.01" -Alias "room.DE.CTY.SITE.Room101" -PrimarySmtpAddress "room.DE.CTY.SITE.Room101@contoso.com"

    If instead you want it cloud-native, run this with -CreateInCloud and it will
    New-Mailbox -Room directly in EXO.

    AUTHENTICATION: interactive / delegated. The signed-in account needs the
    Exchange Administrator role for New-Mailbox / Set-CalendarProcessing /
    Set-Place. For an unattended app-only variant see New-SharedCalendar.ps1,
    which takes -AppId / -CertificateThumbprint / -Organization.

.EXAMPLE
    # 1) Preview - Booking mode (default): restrict who can reserve the room
    .\Set-RoomMailbox.ps1 -RoomName 'Room 1.01' -Managers 'team.lead@contoso.com'

.EXAMPLE
    # 2) Apply Booking restriction (groups go into BookInPolicy as-is; shared mbx expanded)
    .\Set-RoomMailbox.ps1 -RoomName 'Room 1.01' -Managers 'reception@contoso.com' -Execute

.EXAMPLE
    # 3) Manage mode: grant Editor on the calendar folder instead (use -ExpandGroups for DLs)
    .\Set-RoomMailbox.ps1 -RoomName 'Room 1.01' -Mode Manage -ExpandGroups -Execute

.EXAMPLE
    # 4) Also create the mailbox cloud-native in the same run
    .\Set-RoomMailbox.ps1 -RoomName 'Room 1.01' -CreateInCloud -Execute

.NOTES
    When to use  : Provisioning or reconfiguring a room, and especially when someone reports that 'I made the reception group an editor of the room calendar and it does not work for them'.
    Why it exists: Separates restricting who may BOOK the room (BookInPolicy, which accepts groups and evaluates membership at request time, so it survives staff changes) from granting Editor on the calendar folder. Folder permissions do not reach the members of a distribution list, which is the trap in the second mode - hence -ExpandGroups, cycle-guarded.
#>
[CmdletBinding()]
param(
    # ---- target room (CC.CITY.SITE.Room) ----
    [string]$CountryCode = 'DE',
    [string]$CityCode = 'CTY',
    [string]$SiteCode = 'SITE',
    [string]$RoomName = 'Room 1.01',   # bare room name only - the CC.CITY.SITE prefix is added below
    [string]$Domain = 'contoso.com',
    # Leave empty to derive it from the codes above as room.<CC>.<CITY>.<SITE>.<RoomNoSpaces>@<Domain>.
    # Deriving it is the safer default: a hand-typed address that disagrees with
    # the derived alias is exactly how a room ends up with a mismatched SMTP.
    [string]$PrimarySmtp = '',
    [string]$Building = 'Example Building',
    [string]$City = 'Example City',

    # ---- who gets access (booking rights in Booking mode / Editor in Manage mode) ----
    # Accepts users, shared mailboxes, distribution lists and security groups.
    [string[]]$Managers = @(
        'first.user@contoso.com',
        'reception@contoso.com',
        'Example Team Leads'
    ),

    # ---- booking policy ----
    [int]$BookingWindowInDays = 180,
    [int]$MaxDurationInMinutes = 1440,

    # ---- behaviour ----
    [ValidateSet('Booking', 'Manage')]
    [string]$Mode = 'Booking',   # Booking = restrict who can reserve (BookInPolicy); Manage = Editor on the calendar folder
    [switch]$CreateInCloud,
    [switch]$ExpandGroups,   # flatten DLs/security groups to members + shared mailboxes to their delegates, grant each individually
    [switch]$Execute,
    [string]$ExportDir = (Join-Path $PSScriptRoot 'Exports')
)

# ---------- helpers ----------
function Write-Step($m) { Write-Host "`n[STEP] $m" -ForegroundColor Cyan }
function Write-OK  ($m) { Write-Host "[ OK ] $m"   -ForegroundColor Green }
function Write-Warn($m) { Write-Host "[WARN] $m"   -ForegroundColor Yellow }
function Die       ($m) { Write-Host "[FAIL] $m"   -ForegroundColor Red; try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue }catch {}; exit 1 }
function Do-Action([string]$desc, [scriptblock]$action) {
    if ($Execute) {
        try { & $action; Write-OK $desc }
        catch { Write-Warn "FAILED: $desc  ->  $($_.Exception.Message)" }
    } else { Write-Host "[DRY ] would: $desc" -ForegroundColor DarkGray }
}

# Recursively flatten a group (DL / security / dynamic) or shared mailbox into individual
# mail-bearing recipients. Cycle-guarded via $Seen (lowercased SMTPs already visited).
function Expand-ToIndividuals {
    param([string]$Identity, [System.Collections.Generic.HashSet[string]]$Seen)
    $r = Get-Recipient -Identity $Identity -ErrorAction SilentlyContinue
    if (-not $r) { return @() }
    $smtp = [string]$r.PrimarySmtpAddress
    if ($smtp -and -not $Seen.Add($smtp.ToLower())) { return @() }   # already visited -> stop (cycle/dup)
    $type = [string]$r.RecipientTypeDetails
    switch -Regex ($type) {
        'UserMailbox|MailUser|GuestMailUser|LinkedMailbox' { return , $r }
        'SharedMailbox|RoomMailbox|EquipmentMailbox' {
            $out = @()
            Get-MailboxPermission -Identity $smtp -ErrorAction SilentlyContinue |
            Where-Object { ($_.AccessRights -contains 'FullAccess') -and (-not $_.IsInherited) -and ($_.User -notmatch 'NT AUTHORITY\\SELF') } |
            ForEach-Object { $out += Expand-ToIndividuals -Identity ([string]$_.User) -Seen $Seen }
            return $out
        }
        'MailUniversalDistributionGroup|MailUniversalSecurityGroup|RoomList' {
            $out = @()
            # A group that cannot be enumerated must not come back as an empty group.
            # Swallowed, the members simply vanish from the expansion and the permissions
            # are then calculated against a list that is short without anything saying so.
            $members = $null
            try   { $members = @(Get-DistributionGroupMember -Identity $smtp -ResultSize Unlimited -ErrorAction Stop) }
            catch {
                Write-Warn "Could not enumerate '$Identity': $($_.Exception.Message)"
                Write-Warn "  Its members are NOT included below. Resolve this before relying on the result."
                return @()
            }
            foreach ($m in $members) {
                $out += Expand-ToIndividuals -Identity ([string]$m.PrimarySmtpAddress) -Seen $Seen
            }
            return $out
        }
        'DynamicDistributionGroup' {
            Write-Warn "'$Identity' is a dynamic distribution group - not expanded (resolve its filter manually)."
            return @()
        }
        default { return , $r }
    }
}

# ---------- derived names ----------
$DisplayName = "$CountryCode.$CityCode.$SiteCode.$RoomName"          # e.g. DE.CTY.SITE.Room 1.01
$Alias = "room.$CountryCode.$CityCode.$SiteCode." + ($RoomName -replace '[^A-Za-z0-9]', '')  # e.g. room.DE.CTY.SITE.Room101
if (-not $PrimarySmtp) { $PrimarySmtp = "$Alias@$Domain" }

# An explicitly passed -PrimarySmtp that does not match the derived alias is
# almost always a typo in one of the two, so say so rather than provisioning a
# room whose address disagrees with its name.
if ($PrimarySmtp -ne "$Alias@$Domain") {
    Write-Warn "PrimarySmtp '$PrimarySmtp' does not match the derived '$Alias@$Domain'. Confirm this is intentional."
}

$runMode = if ($Execute) { 'EXECUTE' }else { 'DRY-RUN' }
Write-Host "===== Set-RoomMailbox  [$Mode / $runMode] =====" -ForegroundColor Magenta
Write-Host "  Display name : $DisplayName"
Write-Host "  Alias        : $Alias"
Write-Host "  Primary SMTP : $PrimarySmtp"

if (-not (Test-Path $ExportDir)) { New-Item -ItemType Directory -Path $ExportDir | Out-Null }
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
Start-Transcript -Path (Join-Path $ExportDir "RoomMailbox_$($Alias)_$stamp.log") | Out-Null

# ---------- connect ----------
# Interactive / delegated auth by design: the account used needs the Exchange
# Administrator role for New-Mailbox / Set-CalendarProcessing / Set-Place.
# For an unattended app-only variant, see New-SharedCalendar.ps1.
Write-Step "Connecting to Exchange Online (interactive)"
try { Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null }
catch { Die "EXO connect failed: $($_.Exception.Message)" }
Write-OK "Connected to Exchange Online"

# ---------- create / locate mailbox ----------
Write-Step "Checking for mailbox $PrimarySmtp"
$mbx = Get-Mailbox -Identity $PrimarySmtp -ErrorAction SilentlyContinue
if (-not $mbx) {
    if ($CreateInCloud) {
        Do-Action "Create cloud room mailbox '$DisplayName' <$PrimarySmtp>" {
            New-Mailbox -Room -Name $DisplayName -DisplayName $DisplayName -Alias $Alias -PrimarySmtpAddress $PrimarySmtp -ErrorAction Stop | Out-Null
            Start-Sleep -Seconds 5
        }
        if ($Execute) { $mbx = Get-Mailbox -Identity $PrimarySmtp -ErrorAction SilentlyContinue }
    } elseif ($Execute) {
        Die "Mailbox not found. Provision it on-prem first (New-RemoteMailbox -Room, see header) and let it sync, or re-run with -CreateInCloud."
    } else {
        Write-Warn "Mailbox not found - showing planned config only. Provision on-prem first or use -CreateInCloud."
    }
} else { Write-OK "Found: $($mbx.DisplayName)" }

# ---------- room metadata (Room Finder) ----------
Write-Step "Setting place metadata"
Do-Action "Set-Place Building='$Building', City='$City', Country=$CountryCode" {
    Set-Place -Identity $PrimarySmtp -Building $Building -City $City -CountryOrRegion $CountryCode -ErrorAction Stop
}

# ---------- booking policy ----------
Write-Step "Setting calendar processing (AutoAccept, no conflicts, keep subject/organiser)"
Do-Action "Set-CalendarProcessing AutoAccept (window=$BookingWindowInDays d, max=$MaxDurationInMinutes min)" {
    Set-CalendarProcessing -Identity $PrimarySmtp `
        -AutomateProcessing AutoAccept `
        -AllowConflicts $false `
        -AllowRecurringMeetings $true `
        -BookingWindowInDays $BookingWindowInDays `
        -MaximumDurationInMinutes $MaxDurationInMinutes `
        -DeleteSubject $false -DeleteComments $false `
        -AddOrganizerToSubject $true -RemovePrivateProperty $false `
        -ErrorAction Stop
}

# ---------- resolve Calendar folder (handles Kalender localisation) ----------
$cal = $null
if ($mbx) {
    $cal = (Get-MailboxFolderStatistics -Identity $PrimarySmtp -FolderScope Calendar -ErrorAction SilentlyContinue |
        Where-Object FolderType -eq 'Calendar' | Select-Object -First 1).Name
}
if (-not $cal) { $cal = 'Calendar' }
$calId = "$($PrimarySmtp):\$cal"
Write-Host "  Calendar folder: $calId"

# ---------- access: booking rights (Booking) or calendar management (Manage) ----------
Write-Step ("Resolving access [$Mode]" + $(if ($ExpandGroups) { ' (groups expanded)' }else { '' }))

function Grant-CalEditor([string]$user) {
    try { Add-MailboxFolderPermission -Identity $calId -User $user -AccessRights Editor -ErrorAction Stop | Out-Null }
    catch {
        if ($_.Exception.Message -match 'existing permission') { Set-MailboxFolderPermission -Identity $calId -User $user -AccessRights Editor -ErrorAction Stop | Out-Null }
        else { throw }
    }
}

$granted = [System.Collections.Generic.HashSet[string]]::new()   # run-level dedup
$bookList = [System.Collections.Generic.List[string]]::new()      # Booking mode: BookInPolicy entries
$results = foreach ($p in $Managers) {
    $rcpt = Get-Recipient -Identity $p -ErrorAction SilentlyContinue
    if (-not $rcpt) {
        Write-Warn "'$p' not found - skipped"
        [pscustomobject]@{ Principal = $p; Via = ''; Resolved = ''; Type = 'NOT FOUND'; Action = 'skipped' }; continue
    }
    $type = [string]$rcpt.RecipientTypeDetails
    $smtp0 = [string]$rcpt.PrimarySmtpAddress
    $isGroup = $type -match 'Group|RoomList'
    $isShared = $type -eq 'SharedMailbox'
    $isUser = $type -match 'UserMailbox|MailUser|GuestMailUser|LinkedMailbox'

    # ======================= BOOKING MODE =======================
    if ($Mode -eq 'Booking') {
        # groups go in per-se (Exchange evaluates membership at request time) unless -ExpandGroups
        if ($isGroup -and -not $ExpandGroups) {
            if ($granted.Add($smtp0.ToLower())) { $bookList.Add($smtp0) }
            [pscustomobject]@{ Principal = $p; Via = ''; Resolved = $smtp0; Type = $type; Action = 'BookInPolicy (group per-se)' }; continue
        }
        # shared mailboxes (and, with -ExpandGroups, groups) flatten to individuals
        if ($isShared -or ($isGroup -and $ExpandGroups)) {
            $seen = [System.Collections.Generic.HashSet[string]]::new()
            $people = @(Expand-ToIndividuals -Identity $smtp0 -Seen $seen) | Sort-Object PrimarySmtpAddress -Unique
            if (-not $people) { Write-Warn "'$p' [$type] expanded to 0 individuals."; [pscustomobject]@{ Principal = $p; Via = $smtp0; Resolved = ''; Type = $type; Action = '0 members' }; continue }
            $why = if ($isShared) { 'shared->delegate' }else { 'member' }
            foreach ($u in $people) {
                $s = [string]$u.PrimarySmtpAddress
                if (-not $granted.Add($s.ToLower())) { [pscustomobject]@{ Principal = $p; Via = $smtp0; Resolved = $s; Type = [string]$u.RecipientTypeDetails; Action = 'dup (already listed)' }; continue }
                $bookList.Add($s)
                [pscustomobject]@{ Principal = $p; Via = $smtp0; Resolved = $s; Type = [string]$u.RecipientTypeDetails; Action = "BookInPolicy ($why)" }
            }
            continue
        }
        # individual users
        if ($granted.Add($smtp0.ToLower())) { $bookList.Add($smtp0) }
        [pscustomobject]@{ Principal = $p; Via = ''; Resolved = $smtp0; Type = $type; Action = 'BookInPolicy (user)' }; continue
    }

    # ======================= MANAGE MODE (Editor) =======================
    if ($ExpandGroups -and ($isGroup -or $isShared)) {
        $seen = [System.Collections.Generic.HashSet[string]]::new()
        $people = @(Expand-ToIndividuals -Identity $smtp0 -Seen $seen) | Sort-Object PrimarySmtpAddress -Unique
        if (-not $people) { Write-Warn "'$p' [$type] expanded to 0 individuals."; [pscustomobject]@{ Principal = $p; Via = $smtp0; Resolved = ''; Type = $type; Action = '0 members' }; continue }
        foreach ($u in $people) {
            $s = [string]$u.PrimarySmtpAddress
            if (-not $granted.Add($s.ToLower())) { [pscustomobject]@{ Principal = $p; Via = $smtp0; Resolved = $s; Type = [string]$u.RecipientTypeDetails; Action = 'dup (already granted)' }; continue }
            Do-Action "Editor -> $s  (via $p)" { Grant-CalEditor $s }
            [pscustomobject]@{ Principal = $p; Via = $smtp0; Resolved = $s; Type = [string]$u.RecipientTypeDetails; Action = $(if ($Execute) { 'Editor' }else { 'would grant Editor' }) }
        }
        continue
    }
    if ($type -eq 'MailUniversalDistributionGroup') {
        Write-Warn "'$p' is a distribution list - permission won't reach members. Re-run with -ExpandGroups, convert to a security group, or use -Mode Booking."
        [pscustomobject]@{ Principal = $p; Via = ''; Resolved = $smtp0; Type = $type; Action = 'SKIPPED (DL)' }; continue
    }
    if (-not $granted.Add($smtp0.ToLower())) { [pscustomobject]@{ Principal = $p; Via = ''; Resolved = $smtp0; Type = $type; Action = 'dup (already granted)' }; continue }
    Do-Action "Editor -> $smtp0 [$type]" { Grant-CalEditor $smtp0 }
    [pscustomobject]@{ Principal = $p; Via = ''; Resolved = $smtp0; Type = $type; Action = $(if ($Execute) { 'Editor' }else { 'would grant Editor' }) }
}

# ---------- Booking mode: apply the restriction ----------
if ($Mode -eq 'Booking') {
    if ($bookList.Count -gt 0) {
        Do-Action "Set-CalendarProcessing AllBookInPolicy=`$false, BookInPolicy ($($bookList.Count) entries); all others auto-declined" {
            Set-CalendarProcessing -Identity $PrimarySmtp -AllBookInPolicy $false -BookInPolicy ([string[]]$bookList) -ErrorAction Stop
        }
    } else {
        Write-Warn "BookInPolicy list is empty - NOT restricting (would decline everyone). Check the -Managers entries."
    }
}

# ---------- summary ----------
Write-Step "Summary"
$results | Format-Table -AutoSize
$csv = Join-Path $ExportDir "RoomMailbox_$($Alias)_$stamp.csv"
$results | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
Write-OK "Result written to $csv"
if (-not $Execute) { Write-Warn "DRY-RUN only - re-run with -Execute to apply." }

Stop-Transcript | Out-Null
# Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
