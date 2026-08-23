<#
.SYNOPSIS
    Scheduled country-scoped user / group / license export -> multi-sheet XLSX ->
    SharePoint. Written as an Azure Automation runbook using the Automation
    Account's system-assigned managed identity.

.DESCRIPTION
    Rebuilt from an earlier CSV version:
      - Output is a real .xlsx (three sheets) instead of CSV. This removes the
        UTF-8/BOM mojibake entirely (xlsx stores text as UTF-8 XML), so non-ASCII
        characters render correctly with no encoding configuration.
      - Sheets:  Users | GroupMemberships | Licenses  (normalized, one row per pair,
        autofilter enabled) so group/license data is filterable and pivotable.
      - License SKUs are mapped to friendly product names (fallback = raw SkuPartNumber).
      - Group memberships are classified by type and flagged when their name matches
        the group-based-licensing naming convention.
      - -EnabledOnly to drop disabled accounts.

    The workbook is always built locally; it is only uploaded to SharePoint when
    -Execute is passed. Without it the runbook reports the destination and stops.

    Required managed-identity Graph app roles: User.Read.All, Group.Read.All,
    Sites.Selected (with 'write' granted on the target site).

    Required modules in the Automation Account (PS 7.2):
      Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Groups, ImportExcel

.PARAMETER FilterStrategy
    How the in-scope population is selected: UsageLocation | Country | Group | Domain.

.PARAMETER FilterValue
    Value for the chosen strategy: an ISO country code, a group object id, or a
    domain suffix.

.PARAMETER Execute
    Upload the workbook to SharePoint. Omit for a dry run.

.EXAMPLE
    .\Export-CountryUserReport.ps1 -FilterValue 'ES'
    Dry run: builds the workbook, prints where it would be uploaded.

.EXAMPLE
    .\Export-CountryUserReport.ps1 -FilterValue 'ES' -SpHostName 'contoso.sharepoint.com' `
        -SpSitePath '/sites/Example' -Execute
#>

param(
    [ValidateSet('UsageLocation', 'Country', 'Group', 'Domain')]
    [string]$FilterStrategy = 'UsageLocation',

    [Parameter(Mandatory)]
    [string]$FilterValue,

    [switch]$EnabledOnly,

    # --- SharePoint delivery target ---
    [string]$SpHostName     = 'contoso.sharepoint.com',
    [string]$SpSitePath     = '/sites/Example',
    [string]$SpLibraryName  = 'Documents',
    [string]$SpFolderPath   = 'Exports',
    [string]$FileNamePrefix = 'CountryUserExport',

    [switch]$Execute
)

# ============================== CONFIG ==============================
# --- License friendly-name map (fallback = raw SkuPartNumber) ---
$SkuFriendly = @{
    'SPE_E3'                          = 'Microsoft 365 E3'
    'SPE_F1'                          = 'Microsoft 365 F3'
    'ENTERPRISEPACK'                  = 'Office 365 E3'
    'ENTERPRISEPREMIUM'               = 'Office 365 E5'
    'STANDARDPACK'                    = 'Office 365 E1'
    'Microsoft_365_Copilot'           = 'Microsoft 365 Copilot'
    'Microsoft_365_E3_Extra_Features' = 'Microsoft 365 E3 Extra Features'
    'MCOMEETADV'                      = 'Microsoft 365 Audio Conferencing'
    'FLOW_FREE'                       = 'Microsoft Power Automate Free'
    'POWER_BI_STANDARD'               = 'Power BI (free)'
    'POWER_BI_PRO'                    = 'Power BI Pro'
    'POWERAPPS_VIRAL'                 = 'Microsoft Power Apps Plan 2 Trial'
    'POWERAPPS_PER_USER'              = 'Power Apps Premium'
    'POWERAPPS_DEV'                   = 'Microsoft Power Apps for Developer'
    'POWERAUTOMATE_ATTENDED_RPA'      = 'Power Automate Premium'
    'STREAM'                          = 'Microsoft Stream'
    'VISIOCLIENT'                     = 'Visio Plan 2'
    'PROJECTPROFESSIONAL'             = 'Project Plan 3'
    'PROJECTPREMIUM'                  = 'Project Plan 5'
    'Microsoft_Teams_Rooms_Pro'       = 'Microsoft Teams Rooms Pro'
    'CCIBOTS_PRIVPREV_VIRAL'          = 'Microsoft Copilot Studio Viral Trial'
    'PROJECT_MADEIRA_PREVIEW_IW_SKU'  = 'Dynamics 365 Business Central for IWs'
    'EMS'                             = 'Enterprise Mobility + Security E3'
}

# Group names matching this pattern are treated as group-based-licensing groups.
# Adjust to your own naming convention, or swap for an accurate per-group
# assignedLicenses lookup if you prefer precision over speed.
$LicenseGroupPattern = '-(M365|O365)-[LF]-'
# ===================================================================

# ------------------------------ Helpers ------------------------------
function Write-Step($m) { Write-Output "[*]  $m" }
function Write-Ok  ($m) { Write-Output "[OK] $m" }
function Write-Warn($m) { Write-Warning $m }
function Write-Die ($m) { throw $m }

function Get-FriendlyLicense($sku) {
    if ([string]::IsNullOrEmpty($sku)) { return $sku }
    if ($SkuFriendly.ContainsKey($sku)) { $SkuFriendly[$sku] } else { $sku }
}

function Get-GroupType($props) {
    $gt = @($props['groupTypes'])
    if ($gt -contains 'Unified') { 'Microsoft 365' }
    elseif ($props['securityEnabled'] -and -not $props['mailEnabled']) { 'Security' }
    elseif ($props['mailEnabled'] -and $props['securityEnabled']) { 'Mail-enabled security' }
    elseif ($props['mailEnabled'] -and -not $props['securityEnabled']) { 'Distribution' }
    else { 'Other' }
}

# ------------------------------ Connect ------------------------------
Write-Step 'Connecting to Microsoft Graph (managed identity)...'
Connect-MgGraph -Identity -NoWelcome -ErrorAction Stop
Write-Ok 'Connected.'

# --------------------------- Resolve users ---------------------------
Write-Step "Resolving in-scope users (strategy: $FilterStrategy = '$FilterValue', EnabledOnly=$EnabledOnly)..."

$props = 'Id', 'DisplayName', 'UserPrincipalName', 'Mail', 'Department', 'JobTitle', 'AccountEnabled', 'CreatedDateTime'

switch ($FilterStrategy) {
    'UsageLocation' {
        $f = "usageLocation eq '$FilterValue'"; if ($EnabledOnly) { $f += ' and accountEnabled eq true' }
        $users = Get-MgUser -All -Property $props -Filter $f
    }
    'Country' {
        $f = "country eq '$FilterValue'"; if ($EnabledOnly) { $f += ' and accountEnabled eq true' }
        $users = Get-MgUser -All -Property $props -Filter $f
    }
    'Group' {
        $members = Get-MgGroupMember -GroupId $FilterValue -All -ErrorAction Stop
        $userIds = $members | Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.user' } | Select-Object -ExpandProperty Id
        $users = foreach ($id in $userIds) { Get-MgUser -UserId $id -Property $props }
        if ($EnabledOnly) { $users = $users | Where-Object { $_.AccountEnabled } }
    }
    'Domain' {
        $users = Get-MgUser -All -Property $props | Where-Object { $_.UserPrincipalName -like "*@$FilterValue" }
        if ($EnabledOnly) { $users = $users | Where-Object { $_.AccountEnabled } }
    }
    default { Write-Die "Unknown FilterStrategy '$FilterStrategy'." }
}

if (-not $users) { Write-Die 'No users matched the filter - aborting before writing an empty export.' }
Write-Ok "$($users.Count) user(s) matched."

# ------------------------------ Enrich -------------------------------
$userRows = [System.Collections.Generic.List[object]]::new()
$groupRows = [System.Collections.Generic.List[object]]::new()
$licRows = [System.Collections.Generic.List[object]]::new()

foreach ($u in $users) {

    Write-Step "Processing $($u.UserPrincipalName)"

    # --- Licenses ---
    $lic = Get-MgUserLicenseDetail -UserId $u.Id -ErrorAction SilentlyContinue
    $skuList = @($lic.SkuPartNumber | Where-Object { $_ })
    $friendlyAll = foreach ($s in $skuList) { Get-FriendlyLicense $s }
    foreach ($s in $skuList) {
        $licRows.Add([PSCustomObject]@{
                UserPrincipalName = $u.UserPrincipalName
                DisplayName       = $u.DisplayName
                SkuPartNumber     = $s
                FriendlyName      = (Get-FriendlyLicense $s)
            })
    }

    # --- Groups ---
    $memberOf = Get-MgUserMemberOf -UserId $u.Id -All -ErrorAction SilentlyContinue
    $grpNames = @()
    foreach ($it in $memberOf) {
        if ($it.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.group') {
            $gn = $it.AdditionalProperties['displayName']
            $grpNames += $gn
            $groupRows.Add([PSCustomObject]@{
                    UserPrincipalName = $u.UserPrincipalName
                    DisplayName       = $u.DisplayName
                    GroupName         = $gn
                    GroupId           = $it.Id
                    GroupType         = (Get-GroupType $it.AdditionalProperties)
                    IsLicenseGroup    = [bool]($gn -match $LicenseGroupPattern)
                })
        }
    }

    # --- Manager ---
    $mgrName = ''; $mgrUpn = ''
    try {
        $mgr = Get-MgUserManager -UserId $u.Id -ErrorAction Stop
        $mgrName = $mgr.AdditionalProperties['displayName']
        $mgrUpn = $mgr.AdditionalProperties['userPrincipalName']
    } catch { }

    # --- Users summary row ---
    $userRows.Add([PSCustomObject]@{
            DisplayName       = $u.DisplayName
            UserPrincipalName = $u.UserPrincipalName
            Mail              = $u.Mail
            Department        = $u.Department
            JobTitle          = $u.JobTitle
            ManagerName       = $mgrName
            ManagerUPN        = $mgrUpn
            AccountEnabled    = $u.AccountEnabled
            CreatedDateTime   = if ($u.CreatedDateTime) { (Get-Date $u.CreatedDateTime).ToString('dd/MM/yyyy HH:mm:ss') } else { '' }
            GroupCount        = $grpNames.Count
            LicenseCount      = $skuList.Count
            Licenses          = ($friendlyAll -join '; ')
        })
}

# ------------------------------ Write XLSX ---------------------------
$stamp = (Get-Date).ToString('yyyy-MM-dd')
$fileName = "${FileNamePrefix}_${stamp}.xlsx"
$localPath = Join-Path $env:TEMP $fileName
if (Test-Path $localPath) { Remove-Item $localPath -Force }

$xl = @{ AutoSize = $true; AutoFilter = $true; FreezeTopRow = $true; BoldTopRow = $true }
$userRows  | Sort-Object DisplayName       | Export-Excel -Path $localPath -WorksheetName 'Users'            @xl
$groupRows | Sort-Object UserPrincipalName | Export-Excel -Path $localPath -WorksheetName 'GroupMemberships' @xl
$licRows   | Sort-Object UserPrincipalName | Export-Excel -Path $localPath -WorksheetName 'Licenses'         @xl
Write-Ok "Workbook built: $localPath (Users:$($userRows.Count)  Groups:$($groupRows.Count)  Licenses:$($licRows.Count))"

# ------------------------------ Dry run ------------------------------
if (-not $Execute) {
    Write-Warn "DRY RUN - not uploaded. Destination would be https://$SpHostName$SpSitePath -> $SpLibraryName/$SpFolderPath/$fileName. Re-run with -Execute to upload."
    Disconnect-MgGraph | Out-Null
    return
}

# --------------------- Upload to SharePoint (Graph) ------------------
Write-Step 'Resolving SharePoint site and library...'
$site = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/${SpHostName}:${SpSitePath}" -OutputType PSObject
$drives = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives" -OutputType PSObject
$drive = $drives.value | Where-Object { $_.name -eq $SpLibraryName } | Select-Object -First 1
if (-not $drive) { Write-Die "Library '$SpLibraryName' not found on the target site." }

Write-Step 'Uploading workbook to SharePoint...'
$ct = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
$uploadUri = "https://graph.microsoft.com/v1.0/drives/$($drive.id)/root:/$SpFolderPath/${fileName}:/content"
$upload = Invoke-MgGraphRequest -Method PUT -Uri $uploadUri -InputFilePath $localPath -ContentType $ct -OutputType PSObject
Write-Ok "Uploaded: $($upload.webUrl)"

Disconnect-MgGraph | Out-Null
Write-Ok 'Done.'