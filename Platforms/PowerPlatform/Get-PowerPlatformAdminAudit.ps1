<#
.SYNOPSIS
    Audits Power Platform environments for users assigned the System Administrator
    or Environment Maker role, across the entire tenant.

.DESCRIPTION
    Requires the Microsoft.PowerApps.Administration.PowerShell module.
    Must be run by a user with Power Platform (Global/Dynamics 365) Service Admin
    or Power Platform Administrator rights.

.PARAMETER DiagnosticEnvironmentName
    Optional. One or more EnvironmentName (GUID) values to deep-dive after the main
    audit. Useful when the main scan returns 0 results tenant-wide, to confirm whether
    that's genuinely empty (implicit admin / group-based access, no explicit records)
    or a permissions/API issue on the signed-in account.

.NOTES
    Install module (if not already present):
        Install-Module -Name Microsoft.PowerApps.Administration.PowerShell -Force -AllowClobber

.EXAMPLE
    .\Get-PowerPlatformAdminAudit.ps1

.EXAMPLE
    .\Get-PowerPlatformAdminAudit.ps1 -DiagnosticEnvironmentName "3f2b1a...-env-guid","a91c...-env-guid"
#>

param(
    [Parameter(Mandatory = $false)]
    [string[]]$DiagnosticEnvironmentName
)

# ---- Connect ----
Import-Module Microsoft.PowerApps.Administration.PowerShell -ErrorAction Stop

Write-Host "Connecting to Power Platform..." -ForegroundColor Cyan
Add-PowerAppsAccount

# ---- Config ----
# Matched by substring (case-insensitive) since exact RoleType string values can vary
# slightly by module version (e.g. "SysAdmin" vs "System Administrator").
$RolePatterns = @("Admin", "Maker")
$Results = @()
$AllRoleTypesSeen = @{}

# ---- Get all environments ----
Write-Host "Retrieving environments..." -ForegroundColor Cyan
$Environments = Get-AdminPowerAppEnvironment

Write-Host "Found $($Environments.Count) environment(s). Scanning role assignments..." -ForegroundColor Cyan

foreach ($Env in $Environments) {
    $EnvName = $Env.DisplayName
    $EnvId = $Env.EnvironmentName

    try {
        # NOTE: Get-AdminPowerAppRoleAssignment is for APP-level roles (requires -AppName).
        # For ENVIRONMENT-level roles (System Administrator / Environment Maker / System Customizer)
        # we need Get-AdminPowerAppEnvironmentRoleAssignment instead.
        $RoleAssignments = Get-AdminPowerAppEnvironmentRoleAssignment -EnvironmentName $EnvId -ErrorAction Stop
    } catch {
        Write-Warning "Could not retrieve role assignments for '$EnvName': $($_.Exception.Message)"
        continue
    }

    foreach ($Role in $RoleAssignments) {
        if (-not $AllRoleTypesSeen.ContainsKey($Role.RoleType)) {
            $AllRoleTypesSeen[$Role.RoleType] = 0
        }
        $AllRoleTypesSeen[$Role.RoleType]++

        $IsMatch = $false
        foreach ($Pattern in $RolePatterns) {
            if ($Role.RoleType -match $Pattern) { $IsMatch = $true; break }
        }
        if ($IsMatch) {
            $Results += [PSCustomObject]@{
                Environment    = $EnvName
                EnvironmentId  = $EnvId
                RoleType       = $Role.RoleType
                PrincipalType  = $Role.PrincipalType
                PrincipalName  = $Role.PrincipalDisplayName
                PrincipalEmail = $Role.PrincipalEmail
            }
        }
    }
}

# ---- Output ----
if ($Results.Count -eq 0) {
    Write-Host "No System Administrator / Environment Maker assignments found." -ForegroundColor Yellow
} else {
    Write-Host "`n=== Power Platform Admin/Maker Role Assignments ===" -ForegroundColor Green
    $Results | Sort-Object Environment, RoleType, PrincipalName |
    Format-Table Environment, RoleType, PrincipalName, PrincipalEmail, PrincipalType -AutoSize
}

Write-Host "`nTotal assignments found: $($Results.Count)" -ForegroundColor Cyan

if ($AllRoleTypesSeen.Count -gt 0) {
    Write-Host "`n(Diagnostic) All distinct RoleType values encountered across all environments:" -ForegroundColor DarkGray
    $AllRoleTypesSeen.GetEnumerator() | Sort-Object Name | ForEach-Object {
        Write-Host "  - $($_.Key): $($_.Value) assignment(s)" -ForegroundColor DarkGray
    }
}

# ---- Deep-dive diagnostic (self-contained) ----
# If the main scan came back empty, this either means access is granted implicitly
# (tenant admin roles, or environment security-group restriction) rather than via
# explicit role assignment records, or the signed-in account lacks rights to see them.
# Passing -DiagnosticEnvironmentName runs a raw dump against specific environment(s)
# to tell the two apart.

if ($DiagnosticEnvironmentName -and $DiagnosticEnvironmentName.Count -gt 0) {
    Write-Host "`n=== Deep-dive: raw role assignment dump per requested environment ===" -ForegroundColor Magenta
    foreach ($EnvId in $DiagnosticEnvironmentName) {
        $MatchedEnv = $Environments | Where-Object { $_.EnvironmentName -eq $EnvId }
        $Label = if ($MatchedEnv) { $MatchedEnv.DisplayName } else { "(name not found in tenant list)" }

        Write-Host "`n--- $EnvId  [$Label] ---" -ForegroundColor Magenta
        try {
            $Raw = Get-AdminPowerAppEnvironmentRoleAssignment -EnvironmentName $EnvId -ErrorAction Stop
            if ($Raw.Count -eq 0) {
                Write-Host "  No explicit role assignment records returned." -ForegroundColor Yellow
                Write-Host "  This is expected if access here relies on: (a) implicit tenant-admin rights," -ForegroundColor DarkGray
                Write-Host "  or (b) an environment security-group restriction rather than per-user roles." -ForegroundColor DarkGray
            } else {
                $Raw | Format-List *
            }
        } catch {
            Write-Warning "  Failed to query '$EnvId': $($_.Exception.Message)"
        }
    }
} elseif ($Results.Count -eq 0) {
    Write-Host "`n(Tip) To confirm whether this empty result is expected, re-run with:" -ForegroundColor DarkGray
    Write-Host "  .\Get-PowerPlatformAdminAudit.ps1 -DiagnosticEnvironmentName `"<EnvironmentName-GUID>`"" -ForegroundColor DarkGray
    Write-Host "`nHere are the environments found, for copy/paste:" -ForegroundColor DarkGray
    $Environments | Select-Object DisplayName, EnvironmentName |
    Sort-Object DisplayName | Format-Table -AutoSize
}