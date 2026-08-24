<#
.SYNOPSIS
    Exports all Conditional Access policies from Entra ID with GUIDs resolved
    to display names (users, groups, roles, apps, named locations).

.DESCRIPTION
    Produces:
      1. CAPolicies_Raw_<timestamp>.json      - full raw policy objects (lossless)
      2. CAPolicies_Resolved_<timestamp>.json - policies with GUIDs resolved to names
      3. CAPolicies_Summary_<timestamp>.csv   - one-row-per-policy flat summary
      4. NamedLocations_<timestamp>.json      - all named locations (incl. IP ranges)

.PARAMETER TenantId
    Directory (tenant) ID to connect to. Omit to use the home tenant of the
    signed-in account.

.PARAMETER OutputFolder
    Where the four output files are written. Default: <script folder>\Exports.

.EXAMPLE
    .\Export-ConditionalAccessPolicy.ps1 -TenantId '<tenant-id>' -OutputFolder .\Exports

.NOTES
    When to use  : An auditor asks for the Conditional Access policies in writing, or after an incident you need to diff today's CA against three months ago.
    Why it exists: A GUID in a CA policy can be a user, group, role, service principal, appId or named location and the policy body does not say which. The resolver tries five endpoints, then role templates (CA stores role TEMPLATE ids), then servicePrincipals(appId=). Raw JSON, resolved JSON and a flat CSV for diffing are all kept.
    Required Graph scopes (read-only):
      Policy.Read.All, Directory.Read.All, Application.Read.All
    READ-ONLY. No directory writes.
#>

[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$OutputFolder = (Join-Path $PSScriptRoot 'Exports')
)

#region Setup
$ErrorActionPreference = 'Stop'
$timestamp = Get-Date -Format 'yyyyMMdd_HHmm'

if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

$requiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Identity.SignIns')
foreach ($m in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        Write-Host "Installing module $m..." -ForegroundColor Yellow
        Install-Module $m -Scope CurrentUser -Force
    }
}

Write-Host "Connecting to Microsoft Graph (read-only scopes)..." -ForegroundColor Cyan
$connectParams = @{ Scopes = @('Policy.Read.All', 'Directory.Read.All', 'Application.Read.All') }
if ($TenantId) { $connectParams.TenantId = $TenantId }
Connect-MgGraph @connectParams -NoWelcome
#endregion

#region Fetch policies and named locations
# $top is a page size, not a limit: Graph is free to return fewer items and a nextLink,
# and a request that ignores the nextLink then silently exports a partial policy set.
# A CA export that is quietly missing policies is worse than no export, so follow it.
function Get-GraphCollection {
    param([Parameter(Mandatory)][string]$Uri)
    $items = New-Object System.Collections.Generic.List[object]
    $next = $Uri
    $pages = 0
    while ($next) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop
        if ($resp.value) { $items.AddRange(@($resp.value)) }
        $next = $resp.'@odata.nextLink'
        $pages++
    }
    if ($pages -gt 1) { Write-Host "  -> $pages pages followed" -ForegroundColor DarkGray }
    return $items
}

Write-Host "Fetching Conditional Access policies..." -ForegroundColor Cyan
$policies = Get-GraphCollection -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies?$top=999'

Write-Host "  -> $($policies.Count) policies found" -ForegroundColor Green

Write-Host "Fetching named locations..." -ForegroundColor Cyan
$namedLocations = Get-GraphCollection -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations?$top=999'
$locationMap = @{}
foreach ($loc in $namedLocations) { $locationMap[$loc.id] = $loc.displayName }
#endregion

#region GUID resolution helpers (cached lookups)
$guidCache = @{}

function Resolve-DirectoryObject {
    param([string]$Id)

    if ([string]::IsNullOrWhiteSpace($Id)) { return $Id }
    # Pass through well-known keywords (All, None, GuestsOrExternalUsers, etc.)
    if ($Id -notmatch '^[0-9a-fA-F\-]{36}$') { return $Id }
    if ($guidCache.ContainsKey($Id)) { return $guidCache[$Id] }

    $resolved = $null
    foreach ($endpoint in @('users', 'groups', 'directoryRoles', 'servicePrincipals', 'directoryObjects')) {
        try {
            $obj = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/$endpoint/$Id" -ErrorAction Stop
            $type = ($obj.'@odata.context' -split '#')[-1] -replace '/\$entity', ''
            $resolved = "$($obj.displayName) [$type]"
            break
        } catch { continue }
    }
    # Role template IDs (CA stores role *template* IDs, not activated role IDs)
    if (-not $resolved) {
        try {
            $obj = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/directoryRoleTemplates/$Id" -ErrorAction Stop
            $resolved = "$($obj.displayName) [roleTemplate]"
        } catch { }
    }
    # App IDs used in targetResources are appIds, not object IDs
    if (-not $resolved) {
        try {
            $sp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals(appId='$Id')" -ErrorAction Stop
            $resolved = "$($sp.displayName) [app]"
        } catch { }
    }
    if (-not $resolved) { $resolved = "$Id [UNRESOLVED]" }

    $guidCache[$Id] = $resolved
    return $resolved
}

function Resolve-IdList {
    param($Ids)
    if ($null -eq $Ids) { return $null }
    return @($Ids | ForEach-Object { Resolve-DirectoryObject $_ })
}

function Resolve-LocationList {
    param($Ids)
    if ($null -eq $Ids) { return $null }
    return @($Ids | ForEach-Object {
        if ($locationMap.ContainsKey($_)) { "$($locationMap[$_])" } else { $_ }
    })
}
#endregion

#region Build resolved policy objects
Write-Host "Resolving GUIDs to display names (this can take a minute)..." -ForegroundColor Cyan
$resolvedPolicies = foreach ($p in $policies) {

    $c = $p.conditions

    [PSCustomObject]@{
        DisplayName       = $p.displayName
        Id                = $p.id
        State             = $p.state            # enabled / disabled / enabledForReportingButNotEnforced
        CreatedDateTime   = $p.createdDateTime
        ModifiedDateTime  = $p.modifiedDateTime

        Users = [PSCustomObject]@{
            IncludeUsers  = Resolve-IdList $c.users.includeUsers
            ExcludeUsers  = Resolve-IdList $c.users.excludeUsers
            IncludeGroups = Resolve-IdList $c.users.includeGroups
            ExcludeGroups = Resolve-IdList $c.users.excludeGroups
            IncludeRoles  = Resolve-IdList $c.users.includeRoles
            ExcludeRoles  = Resolve-IdList $c.users.excludeRoles
            IncludeGuestsOrExternalUsers = $c.users.includeGuestsOrExternalUsers
            ExcludeGuestsOrExternalUsers = $c.users.excludeGuestsOrExternalUsers
        }

        TargetResources = [PSCustomObject]@{
            IncludeApplications = Resolve-IdList $c.applications.includeApplications
            ExcludeApplications = Resolve-IdList $c.applications.excludeApplications
            IncludeUserActions  = $c.applications.includeUserActions
            ApplicationFilter   = $c.applications.applicationFilter
        }

        Conditions = [PSCustomObject]@{
            UserRiskLevels        = $c.userRiskLevels
            SignInRiskLevels      = $c.signInRiskLevels
            InsiderRiskLevels     = $c.insiderRiskLevels
            ClientAppTypes        = $c.clientAppTypes
            Platforms_Include     = $c.platforms.includePlatforms
            Platforms_Exclude     = $c.platforms.excludePlatforms
            Locations_Include     = Resolve-LocationList $c.locations.includeLocations
            Locations_Exclude     = Resolve-LocationList $c.locations.excludeLocations
            DeviceFilter          = $c.devices.deviceFilter
            AuthenticationFlows   = $c.authenticationFlows
        }

        GrantControls = [PSCustomObject]@{
            Operator                    = $p.grantControls.operator
            BuiltInControls             = $p.grantControls.builtInControls
            AuthenticationStrength      = $p.grantControls.authenticationStrength.displayName
            CustomAuthenticationFactors = $p.grantControls.customAuthenticationFactors
            TermsOfUse                  = $p.grantControls.termsOfUse
        }

        SessionControls = [PSCustomObject]@{
            SignInFrequency          = $p.sessionControls.signInFrequency
            PersistentBrowser        = $p.sessionControls.persistentBrowser
            CloudAppSecurity         = $p.sessionControls.cloudAppSecurity
            ApplicationEnforcedRestrictions = $p.sessionControls.applicationEnforcedRestrictions
            ContinuousAccessEvaluation = $p.sessionControls.continuousAccessEvaluation
            DisableResilienceDefaults  = $p.sessionControls.disableResilienceDefaults
        }
    }
}
#endregion

#region Build flat CSV summary
function Join-Safe { param($x) if ($null -eq $x) { '' } else { ($x | Where-Object { $_ }) -join '; ' } }

$summary = foreach ($rp in $resolvedPolicies) {
    [PSCustomObject]@{
        DisplayName        = $rp.DisplayName
        State              = $rp.State
        IncludeUsers       = Join-Safe $rp.Users.IncludeUsers
        ExcludeUsers       = Join-Safe $rp.Users.ExcludeUsers
        IncludeGroups      = Join-Safe $rp.Users.IncludeGroups
        ExcludeGroups      = Join-Safe $rp.Users.ExcludeGroups
        IncludeRoles       = Join-Safe $rp.Users.IncludeRoles
        ExcludeRoles       = Join-Safe $rp.Users.ExcludeRoles
        IncludeApps        = Join-Safe $rp.TargetResources.IncludeApplications
        ExcludeApps        = Join-Safe $rp.TargetResources.ExcludeApplications
        UserActions        = Join-Safe $rp.TargetResources.IncludeUserActions
        ClientAppTypes     = Join-Safe $rp.Conditions.ClientAppTypes
        LocationsInclude   = Join-Safe $rp.Conditions.Locations_Include
        LocationsExclude   = Join-Safe $rp.Conditions.Locations_Exclude
        UserRisk           = Join-Safe $rp.Conditions.UserRiskLevels
        SignInRisk         = Join-Safe $rp.Conditions.SignInRiskLevels
        GrantOperator      = $rp.GrantControls.Operator
        GrantControls      = Join-Safe $rp.GrantControls.BuiltInControls
        AuthStrength       = $rp.GrantControls.AuthenticationStrength
        SignInFrequency    = if ($rp.SessionControls.SignInFrequency.isEnabled) { "$($rp.SessionControls.SignInFrequency.value) $($rp.SessionControls.SignInFrequency.type)" } else { '' }
        PersistentBrowser  = $rp.SessionControls.PersistentBrowser.mode
        Modified           = $rp.ModifiedDateTime
    }
}
#endregion

#region Write outputs
$rawPath      = Join-Path $OutputFolder "CAPolicies_Raw_$timestamp.json"
$resolvedPath = Join-Path $OutputFolder "CAPolicies_Resolved_$timestamp.json"
$csvPath      = Join-Path $OutputFolder "CAPolicies_Summary_$timestamp.csv"
$locPath      = Join-Path $OutputFolder "NamedLocations_$timestamp.json"

$policies         | ConvertTo-Json -Depth 20 | Out-File $rawPath      -Encoding utf8
$resolvedPolicies | ConvertTo-Json -Depth 10 | Out-File $resolvedPath -Encoding utf8
$summary          | Export-Csv $csvPath -NoTypeInformation -Encoding utf8
$namedLocations   | ConvertTo-Json -Depth 10 | Out-File $locPath      -Encoding utf8

Write-Host ""
Write-Host "Export complete:" -ForegroundColor Green
Write-Host "  Raw JSON      : $rawPath"
Write-Host "  Resolved JSON : $resolvedPath"
Write-Host "  Summary CSV   : $csvPath"
Write-Host "  Named locations: $locPath"
Write-Host ""
Write-Host "Policy states:" -ForegroundColor Cyan
$policies | Group-Object state | ForEach-Object { Write-Host ("  {0,-45} {1}" -f $_.Name, $_.Count) }

Disconnect-MgGraph | Out-Null
#endregion
