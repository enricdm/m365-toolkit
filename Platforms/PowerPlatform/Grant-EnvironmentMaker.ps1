<#
.SYNOPSIS
    Grants a user the Environment Maker role on a specific Power Platform environment
    (least-privilege — scoped access, not tenant/environment admin).

.DESCRIPTION
    Use this for requests like "I need to import my solution into the team environment"
    where the requester only needs to build/import in that one environment, not manage it.

    NOTE: If the target environment has a Dataverse database,
    Environment Maker alone lets the user access Power Apps/Flow authoring surfaces, but
    solution IMPORT into a Dataverse-backed environment also requires a Dataverse security
    role with import rights (e.g. "System Customizer"). That role is assigned separately,
    via Power Platform Admin Center > Environment > Settings > Users + permissions >
    Security roles (a one-time UI action — see notes at bottom of this script).

.PARAMETER EnvironmentId
    The target environment's EnvironmentName (GUID), e.g. from the audit script output.

.PARAMETER UserPrincipalName
    The UPN/email of the user to grant access to, e.g. user@contoso.com

.PARAMETER Execute
    Required to actually assign the role. Without it the script prints what it would
    do and exits without connecting or changing anything.

.NOTES
    When to use  : Requests of the form 'I need to import my solution into the team environment'.
    Why it exists: Grants Environment Maker on one environment rather than environment or tenant admin, checks -Execute before connecting to anything, and inspects the result explicitly because Set-AdminPowerAppEnvironmentRoleAssignment can return an error object as normal output instead of throwing. The manual Dataverse follow-up is printed rather than assumed.
    Requires modules:
        Microsoft.PowerApps.Administration.PowerShell
        Microsoft.Graph.Users   (for resolving UPN -> Azure AD Object ID)

    Install if missing:
        Install-Module Microsoft.PowerApps.Administration.PowerShell -Force -AllowClobber
        Install-Module Microsoft.Graph.Users -Force -AllowClobber

.EXAMPLE
    # Dry run - shows what would happen, connects to nothing
    .\Grant-EnvironmentMaker.ps1 -EnvironmentId '<environment-guid>' -UserPrincipalName 'user@contoso.com'

.EXAMPLE
    # Apply the role assignment
    .\Grant-EnvironmentMaker.ps1 -EnvironmentId '<environment-guid>' -UserPrincipalName 'user@contoso.com' -Execute
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$EnvironmentId,

    [Parameter(Mandatory = $true)]
    [string]$UserPrincipalName,

    [switch]$Execute
)

if (-not $Execute) {
    Write-Host "DRY-RUN. Would grant 'Environment Maker' on environment '$EnvironmentId' to '$UserPrincipalName'." -ForegroundColor Yellow
    Write-Host "DRY-RUN. Nothing was changed. Re-run with -Execute to apply." -ForegroundColor Yellow
    return
}

Import-Module Microsoft.PowerApps.Administration.PowerShell -ErrorAction Stop
Import-Module Microsoft.Graph.Users -ErrorAction Stop

# ---- Connect ----
Write-Host "Connecting to Power Platform..." -ForegroundColor Cyan
Add-PowerAppsAccount

Write-Host "Connecting to Microsoft Graph (to resolve user object ID)..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.Read.All" -NoWelcome

# ---- Resolve user ----
Write-Host "Resolving '$UserPrincipalName'..." -ForegroundColor Cyan
# -UserId resolves the UPN to exactly one object; -Filter returns a collection, and a
# multi-match would silently make $ObjectId an array.
try {
    $MgUser = Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
} catch {
    Write-Error "User '$UserPrincipalName' not found in Entra ID: $($_.Exception.Message). Aborting."
    return
}

if (-not $MgUser) {
    Write-Error "User '$UserPrincipalName' not found in Entra ID. Aborting."
    return
}

$ObjectId = $MgUser.Id
Write-Host "Resolved to Object ID: $ObjectId ($($MgUser.DisplayName))" -ForegroundColor Green

# ---- Confirm environment ----
$TargetEnv = Get-AdminPowerAppEnvironment -EnvironmentName $EnvironmentId -ErrorAction Stop
if (-not $TargetEnv) {
    Write-Error "Environment '$EnvironmentId' not found. Aborting."
    return
}
Write-Host "Target environment: $($TargetEnv.DisplayName) [$EnvironmentId]" -ForegroundColor Green

# ---- Assign Environment Maker role (least privilege) ----
Write-Host "`nAssigning 'Environment Maker' role..." -ForegroundColor Cyan
try {
    $AssignResult = Set-AdminPowerAppEnvironmentRoleAssignment `
        -EnvironmentName $EnvironmentId `
        -RoleName EnvironmentMaker `
        -PrincipalType User `
        -PrincipalObjectId $ObjectId `
        -ErrorAction Stop

    # This cmdlet can return an error object (e.g. 403) as normal output rather than
    # throwing a terminating exception, so check explicitly instead of assuming success.
    if ($AssignResult -and ($AssignResult.Code -or $AssignResult.Error)) {
        Write-Error "Assignment failed (API returned an error object):"
        $AssignResult | Format-List *
        return
    }

    Write-Host "Success: '$($MgUser.DisplayName)' granted Environment Maker on '$($TargetEnv.DisplayName)'." -ForegroundColor Green
} catch {
    Write-Error "Failed to assign role: $($_.Exception.Message)"
    return
}

# ---- Verify ----
Write-Host "`nVerifying assignment..." -ForegroundColor Cyan
Get-AdminPowerAppEnvironmentRoleAssignment -EnvironmentName $EnvironmentId |
Where-Object { $_.PrincipalObjectId -eq $ObjectId } |
Format-Table PrincipalDisplayName, RoleType, PrincipalEmail -AutoSize

Write-Host "`n--- NEXT STEP (manual, Dataverse-backed environments only) ---" -ForegroundColor Yellow
Write-Host "This environment has a Dataverse database. To complete solution IMPORT rights:" -ForegroundColor Yellow
Write-Host "  1. Power Platform Admin Center > Environments > $($TargetEnv.DisplayName)" -ForegroundColor Yellow
Write-Host "  2. Settings > Users + permissions > Security roles" -ForegroundColor Yellow
Write-Host "  3. Assign '$UserPrincipalName' the 'System Customizer' security role" -ForegroundColor Yellow
Write-Host "     (grants solution import privileges without full System Administrator access)" -ForegroundColor Yellow