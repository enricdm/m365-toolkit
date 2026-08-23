<#
================================================================================
    New-FaxDistributionList.ps1

    Purpose : Provision a Fax-to-Mail distribution list in Exchange Online.

    Fax gateways are the awkward case for a distribution list: they relay
    unauthenticated, so the default RequireSenderAuthenticationEnabled = $true
    silently drops every inbound fax with no NDR the requester ever sees. This
    script provisions the DL with that setting explicitly disabled and refuses
    to finish if it ends up enabled anyway.
================================================================================

    NAMING CONVENTION
    -----------------
        <ISO2>.<name>                        country scope
        <ISO2>.<City3>.<name>                city scope
        <ISO2>.<City3>.<Dept3>.<name>        department scope
        No DL-/DG-/SG- prefix. Dots separate scope parts.
        Underscore reserved for function suffixes only (_MBX, _SendAs, _E, _RW, _f).

    Adapt the convention to your own standard; nothing below depends on it
    beyond the default parameter values.

    KEY BEHAVIOUR
    -------------
    - DRY RUN by default. Nothing is written until you pass -Execute.
    - Address collision is checked BEFORE anything is created.
    - Every member is resolved in a preflight pass and exported to CSV, so the
      requester can confirm the list before it is provisioned.
    - External (non-tenant) members are auto-created as MailContacts.
    - RequireSenderAuthentication is DISABLED by default (see above).
    - The provisioned object is verified against what was asked for, and both
      the object state and the member list are exported to <script folder>\Exports\

    PREREQUISITE:  Connect-ExchangeOnline  (account needs Exchange Recipient
                   Administrator or equivalent)

================================================================================
#>

[CmdletBinding()]
param(
    # ---- the group ----
    [Parameter(Mandatory)][string]$GroupName,          # e.g. DE.NUE.Example-Fax
    [Parameter(Mandatory)][string]$DisplayName,        # GAL display name
    [Parameter(Mandatory)][string]$PrimarySmtp,        # e.g. Example-Fax@contoso.com
    [Parameter(Mandatory)][string]$Owner,              # ManagedBy - usually the requester

    # Additional smtp aliases, e.g. a legacy address a pre-staged gateway still uses.
    [string[]]$AliasAddresses = @(),

    # Initial members. Addresses outside -InternalDomain are created as MailContacts.
    [string[]]$Members = @(
        'first.user@contoso.com',
        'second.user@contoso.com',
        'shared.inbox@contoso.com'
    ),

    # ---- mail flow ----
    # $true will silently drop inbound faxes on gateways that relay
    # unauthenticated. Leave $false unless you have confirmed the gateway
    # authenticates.
    [bool]$RequireSenderAuth = $false,

    # Restrict who may send TO the list. Empty = anyone (required for most fax
    # gateways). Populate only if the gateway sender address is known.
    [string[]]$AllowedSenders = @(),

    [bool]$HideFromGAL = $false,                       # legacy fax objects are usually GAL-visible
    [string]$InternalDomain = 'contoso.com',

    # ---- bookkeeping ----
    [string]$Notes = '',                               # free text stored on the group
    [string]$Reference = '',                           # change reference, recorded in the export only
    [string]$ExportDir = (Join-Path $PSScriptRoot 'Exports'),

    [switch]$Execute
)

# ==============================================================================
# HELPERS
# ==============================================================================

function Write-Step { param($m) Write-Host "`n[>] $m" -ForegroundColor Cyan }
function Write-OK { param($m) Write-Host "    [OK]   $m" -ForegroundColor Green }
function Write-Warn { param($m) Write-Host "    [WARN] $m" -ForegroundColor Yellow }
function Write-Die { param($m) Write-Host "`n[X] $m" -ForegroundColor Red; exit 1 }

$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$dateStr = Get-Date -Format "dd/MM/yyyy HH:mm"
$results = [System.Collections.Generic.List[object]]::new()

if (-not (Test-Path $ExportDir)) { New-Item -ItemType Directory -Path $ExportDir -Force | Out-Null }

Write-Host "`n================================================================" -ForegroundColor White
Write-Host "  Fax-to-Mail DL Provisioning  -  $dateStr" -ForegroundColor White
Write-Host "  Group : $GroupName" -ForegroundColor White
Write-Host "  Mode  : $(if ($Execute) { 'EXECUTE' } else { 'DRY RUN (use -Execute to apply)' })" -ForegroundColor $(if ($Execute) { 'Red' } else { 'Yellow' })
Write-Host "================================================================" -ForegroundColor White

# ==============================================================================
# 1. PRE-FLIGHT
# ==============================================================================

Write-Step "Verifying Exchange Online session"
try { $null = Get-OrganizationConfig -ErrorAction Stop; Write-OK "Session active" }
catch { Write-Die "No EXO session. Run Connect-ExchangeOnline first." }

Write-Step "Checking for address collisions"
$allAddresses = @($PrimarySmtp) + $AliasAddresses
foreach ($addr in $allAddresses) {
    $existing = Get-Recipient -Identity $addr -ErrorAction SilentlyContinue
    if ($existing) { Write-Die "Address '$addr' already in use by '$($existing.DisplayName)' [$($existing.RecipientType)]. Resolve before proceeding." }
    Write-OK "Free: $addr"
}

if (Get-DistributionGroup -Identity $GroupName -ErrorAction SilentlyContinue) { Write-Die "A distribution group named '$GroupName' already exists." }
Write-OK "Group name available"

Write-Step "Resolving owner"
$ownerObj = Get-Recipient -Identity $Owner -ErrorAction SilentlyContinue
if (-not $ownerObj) { Write-Die "Owner '$Owner' not found. Look it up with: Get-Recipient -Filter `"DisplayName -like '*<surname>*'`" | Select DisplayName,PrimarySmtpAddress" }
Write-OK "Owner: $($ownerObj.DisplayName) <$($ownerObj.PrimarySmtpAddress)>"

Write-Step "Validating members"
$resolved = [System.Collections.Generic.List[string]]::new()
$external = [System.Collections.Generic.List[string]]::new()

foreach ($m in $Members) {
    if ([string]::IsNullOrWhiteSpace($m)) { continue }
    $rcpt = Get-Recipient -Identity $m -ErrorAction SilentlyContinue
    if ($rcpt) {
        Write-OK "$m  ->  $($rcpt.DisplayName) [$($rcpt.RecipientType)]"
        $resolved.Add($rcpt.PrimarySmtpAddress)
        $results.Add([PSCustomObject]@{ Address = $m; Status = 'Resolved'; Type = $rcpt.RecipientType; Detail = $rcpt.DisplayName })
    } elseif ($m -notlike "*@$InternalDomain") {
        Write-Warn "$m  ->  external, MailContact will be created"
        $external.Add($m)
        $results.Add([PSCustomObject]@{ Address = $m; Status = 'External'; Type = 'MailContact (pending)'; Detail = '' })
    } else {
        Write-Warn "$m  ->  NOT FOUND in tenant despite internal domain. Check spelling."
        $results.Add([PSCustomObject]@{ Address = $m; Status = 'NotFound'; Type = ''; Detail = 'Internal domain but no recipient object' })
    }
}

$notFound = ($results | Where-Object Status -eq 'NotFound').Count
if ($notFound -gt 0) { Write-Warn "$notFound member(s) unresolved. Review before running with -Execute." }

if (-not $Execute) {
    $results | Export-Csv "$ExportDir\FaxDL-Preflight-$stamp.csv" -NoTypeInformation -Encoding UTF8
    Write-Host "`n[DRY RUN] No changes made. Preflight: $ExportDir\FaxDL-Preflight-$stamp.csv" -ForegroundColor Yellow
    Write-Host "[DRY RUN] Re-run with -Execute to provision.`n" -ForegroundColor Yellow
    return
}

# ==============================================================================
# 2. CREATE EXTERNAL CONTACTS
# ==============================================================================

foreach ($ext in $external) {
    Write-Step "Creating MailContact: $ext"
    try {
        $cn = ($ext -split '@')[0] -replace '[^\w\.\-]', ''
        $null = New-MailContact -Name $cn -ExternalEmailAddress $ext -ErrorAction Stop
        Write-OK "Created"
        $resolved.Add($ext)
        ($results | Where-Object Address -eq $ext).Type = 'MailContact (created)'
    } catch {
        # Commonly already exists as a B2B Guest / MailUser object.
        $fallback = Get-Recipient -Identity $ext -ErrorAction SilentlyContinue
        if ($fallback) { Write-OK "Already exists as $($fallback.RecipientType) - reusing"; $resolved.Add($fallback.PrimarySmtpAddress) }
        else { Write-Warn "Failed: $($_.Exception.Message)" }
    }
}

# ==============================================================================
# 3. CREATE THE DISTRIBUTION LIST
# ==============================================================================

Write-Step "Creating distribution list"
try {
    $dl = New-DistributionGroup -Name $GroupName -DisplayName $DisplayName -PrimarySmtpAddress $PrimarySmtp -Type Distribution -ManagedBy $Owner -MemberJoinRestriction Closed -Notes $Notes -ErrorAction Stop
    Write-OK "Created: $($dl.PrimarySmtpAddress)"
} catch { Write-Die "Creation failed: $($_.Exception.Message)" }

Start-Sleep -Seconds 5   # allow directory replication before follow-up cmdlets

# ==============================================================================
# 4. MAIL FLOW SETTINGS
# ==============================================================================

Write-Step "Applying mail flow settings"

try {
    Set-DistributionGroup -Identity $GroupName -RequireSenderAuthenticationEnabled $RequireSenderAuth -HiddenFromAddressListsEnabled $HideFromGAL -BypassSecurityGroupManagerCheck -ErrorAction Stop
    Write-OK "RequireSenderAuthentication = $RequireSenderAuth  |  HiddenFromGAL = $HideFromGAL"
} catch { Write-Die "Mail flow settings failed: $($_.Exception.Message)" }

if ($AllowedSenders.Count -gt 0) {
    Set-DistributionGroup -Identity $GroupName -AcceptMessagesOnlyFromSendersOrMembers $AllowedSenders -BypassSecurityGroupManagerCheck -ErrorAction Stop
    Write-OK "Sender restriction applied to $($AllowedSenders.Count) address(es)"
} else {
    Write-OK "No sender restriction (open) - required for most fax gateways"
}

foreach ($alias in $AliasAddresses) {
    try { Set-DistributionGroup -Identity $GroupName -EmailAddresses @{ Add = "smtp:$alias" } -BypassSecurityGroupManagerCheck -ErrorAction Stop; Write-OK "Alias added: $alias" }
    catch { Write-Warn "Alias '$alias' failed: $($_.Exception.Message)" }
}

# ==============================================================================
# 5. ADD MEMBERS
# ==============================================================================

Write-Step "Adding members"
$added = 0
foreach ($m in ($resolved | Select-Object -Unique)) {
    try { Add-DistributionGroupMember -Identity $GroupName -Member $m -BypassSecurityGroupManagerCheck -ErrorAction Stop; Write-OK "Added: $m"; $added++ }
    catch { Write-Warn "Failed: $m  ->  $($_.Exception.Message)" }
}

# ==============================================================================
# 6. VERIFY  (only reached with -Execute; the dry run returns in section 1)
# ==============================================================================

Write-Step "Verifying provisioned state"

$dl = Get-DistributionGroup -Identity $GroupName -ErrorAction Stop

$objectState = [PSCustomObject]@{
    Name            = $dl.Name
    DisplayName     = $dl.DisplayName
    PrimarySmtp     = $dl.PrimarySmtpAddress
    Aliases         = ($dl.EmailAddresses -join ';')
    SenderAuthReq   = $dl.RequireSenderAuthenticationEnabled
    JoinRestriction = $dl.MemberJoinRestriction
    HiddenFromGAL   = $dl.HiddenFromAddressListsEnabled
    Owner           = (($dl.ManagedBy | ForEach-Object { (Get-Recipient $_ -ErrorAction SilentlyContinue).PrimarySmtpAddress }) -join ';')
    Reference       = $Reference
    MembersAdded    = $added
    VerifiedOn      = (Get-Date -Format 'dd/MM/yyyy HH:mm')
}

if ($objectState.SenderAuthReq -ne $RequireSenderAuth) {
    Write-Die "RequireSenderAuthenticationEnabled is $($objectState.SenderAuthReq), expected $RequireSenderAuth. With `$true a fax gateway relay will bounce. Fix before closing."
}
Write-OK "RequireSenderAuthentication verified as $RequireSenderAuth"

$liveMembers = Get-DistributionGroupMember -Identity $GroupName -ResultSize Unlimited |
ForEach-Object {
    [PSCustomObject]@{
        Address     = [string]$_.PrimarySmtpAddress
        DisplayName = [string]$_.DisplayName
        Type        = [string]$_.RecipientTypeDetails
        Status      = 'Present in DL'
    }
}

$liveAddresses = @($liveMembers.Address)
$missing = $Members | Where-Object { $liveAddresses -notcontains $_ } | ForEach-Object {
    [PSCustomObject]@{ Address = $_; DisplayName = ''; Type = ''; Status = 'MISSING' }
}

if ($missing) { Write-Warn "$($missing.Count) requested member(s) not present in DL" }
else { Write-OK "All $($liveMembers.Count) members confirmed present" }

$objectState | Export-Csv (Join-Path $ExportDir "FaxDL-Object-$stamp.csv")  -NoTypeInformation -Encoding UTF8
@($liveMembers + $missing) | Export-Csv (Join-Path $ExportDir "FaxDL-Members-$stamp.csv") -NoTypeInformation -Encoding UTF8
Write-OK "Verification exports written to $ExportDir"

# ==============================================================================
# ROLLBACK (manual, single line) - removes the group created by this run:
#   Remove-DistributionGroup -Identity "<GroupName>" -Confirm:$false
# ==============================================================================
