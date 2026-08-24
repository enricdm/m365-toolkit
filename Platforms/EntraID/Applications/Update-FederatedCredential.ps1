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

.PARAMETER Execute
    Required to change anything. Without it the script prints the credentials it
    would create and exits without connecting to the tenant.

.EXAMPLE
    .\Update-FederatedCredential.ps1 -TenantId '<tenant-id>' -AppId '<client-id>' `
        -GitHubOrg 'contoso' -Repository 'platform-infra','platform-shared'

    Dry run. Shows the credential names that would be created, connects to nothing.

.EXAMPLE
    .\Update-FederatedCredential.ps1 -TenantId '<tenant-id>' -AppId '<client-id>' `
        -GitHubOrg 'contoso' -Repository 'platform-infra','platform-shared' -Execute

    Real run. Backs up the current credentials, verifies the backup is readable and
    complete, prompts for confirmation, then replaces them.

.NOTES
    When to use  : A GitHub Actions pipeline fails with 'no matching federated identity record' on a new branch and the app registration is already full of federated credentials.
    Why it exists: Subject-mode credentials pin an exact sub claim, so a repo needs one per branch and the 20-credential ceiling arrives fast. FFIC uses a claims-matching expression instead, so one credential covers every branch of a repository.
    Scopes: Application.ReadWrite.All
    DESTRUCTIVE with -Execute: deletes every existing federated credential on the
    target app. The pre-change backup is written, read back and checked for
    completeness first; if any of that fails, nothing is deleted.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$TenantId,
    [Parameter(Mandatory)][string]$AppId,
    [Parameter(Mandatory)][string]$GitHubOrg,
    [Parameter(Mandatory)][string[]]$Repository,
    [string]$BackupPath = (Join-Path $PSScriptRoot 'Exports'),
    [string]$NameTrimPattern = '^azure-|-infrastructure$',

    # Nothing is deleted or created without this. The check happens before the first
    # write AND before authenticating, so a run without it cannot touch the tenant.
    [switch]$Execute
)

$ErrorActionPreference = 'Stop'

if (-not $Execute) {
    Write-Host ""
    Write-Host "DRY RUN. Nothing will be deleted or created." -ForegroundColor Yellow
    Write-Host "This script DELETES every federated identity credential on app $AppId and" -ForegroundColor Yellow
    Write-Host "recreates one FFIC per repository. Re-run with -Execute to apply:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Repositories that would get an all-branches credential:" -ForegroundColor Cyan
    foreach ($r in $Repository) {
        $short = $r -replace $NameTrimPattern, ""
        Write-Host ("    {0}  ->  github-{1}-{2}-all-branches" -f $r, $GitHubOrg, $short)
    }
    Write-Host ""
    Write-Host "  Run with -Execute to see the current credentials, take a backup and replace them." -ForegroundColor Cyan
    return
}

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
# The backup is the only way back from the deletion below, so it is verified rather than
# assumed. Writing it and carrying on regardless is how a backup step becomes decoration:
# if the directory cannot be created or the file does not land, nothing is deleted.
$backupDir = $BackupPath
if (-not (Test-Path $backupDir)) {
    try { New-Item -ItemType Directory -Path $backupDir -ErrorAction Stop | Out-Null }
    catch { throw "Cannot create the backup folder '$backupDir': $($_.Exception.Message). Nothing has been deleted." }
}
$backupPath = "$backupDir\federatedCreds-backup-$(Get-Date -Format 'yyyyMMdd-HHmmss').json"
$current = Invoke-MgGraphRequest -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/applications/$appObjectId/federatedIdentityCredentials"

$existing = @($current.value)
try {
    ,$existing | ConvertTo-Json -Depth 10 | Out-File -LiteralPath $backupPath -Encoding utf8 -ErrorAction Stop
} catch {
    throw "Could not write the backup to '$backupPath': $($_.Exception.Message). Nothing has been deleted."
}

# Read it back. An Out-File that returned without throwing is not proof of a usable file.
if (-not (Test-Path -LiteralPath $backupPath)) {
    throw "The backup file '$backupPath' does not exist after writing it. Nothing has been deleted."
}
$verify = $null
try { $verify = @(Get-Content -LiteralPath $backupPath -Raw -ErrorAction Stop | ConvertFrom-Json) }
catch { throw "The backup at '$backupPath' is not readable JSON: $($_.Exception.Message). Nothing has been deleted." }
if ($verify.Count -ne $existing.Count) {
    throw "The backup holds $($verify.Count) credential(s) but the app has $($existing.Count). Nothing has been deleted."
}
Write-Host "Backup saved and verified ($($existing.Count) credential(s)): $backupPath" -ForegroundColor Cyan

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
