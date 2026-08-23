<#
.SYNOPSIS
    Migrates an app registration's GitHub OIDC federated credentials from
    subject-mode (one branch each) to flexible federated identity credentials
    (FFIC), one per repository, covering every branch.

.DESCRIPTION
    Subject-mode federated credentials pin an exact 'sub' claim, so a repo needs
    one credential per branch and the 20-credential limit is hit quickly. FFIC
    uses a claims-matching expression instead, so a single credential covers
    'refs/heads/*' for the whole repository.

    The script backs up the current credentials to JSON, prints them, asks for an
    explicit 'yes', then deletes ALL existing federated credentials on the app and
    recreates one FFIC per repository.

.PARAMETER TenantId
    Directory (tenant) ID to connect to.

.PARAMETER AppId
    Application (client) ID of the app registration whose federated credentials
    are replaced.

.PARAMETER GitHubOrg
    GitHub organisation (or user) that owns the repositories.

.PARAMETER Repository
    Repository names to create credentials for. One FFIC is created per entry.

.PARAMETER BackupPath
    Folder for the pre-change JSON backup. Default: <script folder>\Exports.

.PARAMETER NameTrimPattern
    Regex removed from the repository name when building the credential's display
    name, so 'azure-myproduct-infrastructure' becomes 'myproduct'. Set to '^$' to
    keep the repository name verbatim.

.EXAMPLE
    .\Update-FederatedCredential.ps1 -TenantId '<tenant-id>' -AppId '<client-id>' `
        -GitHubOrg 'contoso' -Repository 'platform-infra','platform-shared'

.NOTES
    Scopes: Application.ReadWrite.All
    DESTRUCTIVE: deletes every existing federated credential on the target app.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$TenantId,
    [Parameter(Mandatory)][string]$AppId,
    [Parameter(Mandatory)][string]$GitHubOrg,
    [Parameter(Mandatory)][string[]]$Repository,
    [string]$BackupPath = (Join-Path $PSScriptRoot 'Exports'),
    [string]$NameTrimPattern = '^azure-|-infrastructure$'
)

# --- Connect ---
Connect-MgGraph -Scopes "Application.ReadWrite.All" -TenantId $TenantId

$appId = $AppId
$githubOrg = $GitHubOrg
$repos = $Repository

# --- Get the Application object ---
$appResponse = Invoke-MgGraphRequest -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/applications?`$filter=appId eq '$appId'"
$appObjectId = $appResponse.value[0].id

if (-not $appObjectId) {
    Write-Error "App Registration not found"
    return
}
Write-Host "Found App Registration. Object ID: $appObjectId" -ForegroundColor Green

# --- BACKUP ---
$backupDir = $BackupPath
if (-not (Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir | Out-Null }
$backupPath = "$backupDir\federatedCreds-backup-$(Get-Date -Format 'yyyyMMdd-HHmmss').json"
$current = Invoke-MgGraphRequest -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/applications/$appObjectId/federatedIdentityCredentials"
$current.value | ConvertTo-Json -Depth 10 | Out-File $backupPath
Write-Host "Backup saved to: $backupPath" -ForegroundColor Cyan

Write-Host "`nCurrent federated credentials:" -ForegroundColor Yellow
$current.value | Select-Object name, subject, @{N='expression';E={$_.claimsMatchingExpression.value}} | Format-Table -AutoSize

# --- Confirmation prompt ---
$confirm = Read-Host "`nProceed to DELETE all current credentials and recreate as FFIC? (yes/no)"
if ($confirm -ne "yes") {
    Write-Host "Aborted." -ForegroundColor Red
    return
}

# --- DELETE all existing federated credentials ---
Write-Host "`nDeleting existing credentials..." -ForegroundColor Yellow
foreach ($cred in $current.value) {
    try {
        Invoke-MgGraphRequest -Method DELETE `
            -Uri "https://graph.microsoft.com/v1.0/applications/$appObjectId/federatedIdentityCredentials/$($cred.id)"
        Write-Host "  Deleted: $($cred.name)" -ForegroundColor Gray
    }
    catch {
        Write-Error "Failed to delete $($cred.name): $_"
    }
}

# --- CREATE new FFIC credentials (one per repo, all branches) ---
Write-Host "`nCreating new FFIC credentials..." -ForegroundColor Yellow
foreach ($repo in $repos) {
    $shortName = $repo -replace $NameTrimPattern, ""
    $credName = "github-$githubOrg-$shortName-all-branches"
    $expression = "claims['sub'] matches 'repo:$githubOrg/${repo}:ref:refs/heads/*'"

    $body = @{
        name        = $credName
        issuer      = "https://token.actions.githubusercontent.com"
        subject     = $null
        description = "GitHub OIDC - $repo - all branches (FFIC)"
        audiences   = @("api://AzureADTokenExchange")
        claimsMatchingExpression = @{
            value           = $expression
            languageVersion = 1
        }
    }

    try {
        $result = Invoke-MgGraphRequest -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/applications/$appObjectId/federatedIdentityCredentials" `
            -Body ($body | ConvertTo-Json -Depth 10) `
            -ContentType "application/json"

        Write-Host "  Created: $credName" -ForegroundColor Green
        Write-Host "    Expression: $expression" -ForegroundColor DarkGray
    }
    catch {
        Write-Error "Failed to create $credName : $_"
    }
}

# --- Verify final state ---
Write-Host "`nFinal state:" -ForegroundColor Cyan
$final = Invoke-MgGraphRequest -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/applications/$appObjectId/federatedIdentityCredentials"
$final.value | Select-Object name, @{N='expression';E={$_.claimsMatchingExpression.value}} | Format-Table -AutoSize

Write-Host "Done. $($final.value.Count) federated credentials configured." -ForegroundColor Green
