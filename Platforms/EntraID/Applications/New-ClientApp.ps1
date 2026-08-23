<#
.SYNOPSIS
    Bulk-creates client app registrations for certificate-based machine-to-machine
    authentication, one per supplied name.

.DESCRIPTION
    For each requested application:
      - Creates the App Registration (single tenant, no redirect URI, no credentials)
      - Creates the corresponding Service Principal (required later for app role assignment)
      - Assigns the supplied owner to both objects
      - Stamps governance metadata in the Notes field
    Existing applications with the same displayName are skipped, so the script is
    safe to re-run. Outputs a summary table with the Application (client) IDs.

    No credentials are created: the owner is expected to upload their own
    certificate afterwards.

.PARAMETER Name
    One or more application display names to create.

.PARAMETER OwnerUpn
    UPN of the account to set as owner of every created application and service
    principal.

.PARAMETER Notes
    Free-text governance metadata stamped into the application's Notes field.

.EXAMPLE
    .\New-ClientApp.ps1 -Name 'APP-Client-01','APP-Client-02' -OwnerUpn 'admin@contoso.com'

.NOTES
    Requires Microsoft.Graph.Applications + Microsoft.Graph.Users.
    Scopes: Application.ReadWrite.All, User.Read.All
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string[]]$Name,

    [Parameter(Mandatory)]
    [string]$OwnerUpn,

    [string]$Notes = 'Client auth (cert-based M2M). Owner self-configures certificates.'
)

$AppNames = $Name

# ---------------- Connect ----------------
Connect-MgGraph -Scopes "Application.ReadWrite.All", "User.Read.All" -NoWelcome

$Owner = Get-MgUser -UserId $OwnerUpn -ErrorAction Stop
Write-Host "Owner resolved: $($Owner.DisplayName) [$($Owner.Id)]" -ForegroundColor Cyan

# ---------------- Create ----------------
$Results = foreach ($Name in $AppNames) {

    # Idempotency guard
    $Existing = Get-MgApplication -Filter "displayName eq '$Name'" -ErrorAction SilentlyContinue
    if ($Existing) {
        Write-Warning "$Name already exists (AppId $($Existing.AppId)) - skipping."
        continue
    }

    Write-Host "Creating $Name..." -ForegroundColor Yellow

    $App = New-MgApplication -DisplayName $Name `
        -SignInAudience "AzureADMyOrg" `
        -Notes $Notes

    # Service principal (needed later for app role assignment against the API app)
    $Sp = New-MgServicePrincipal -AppId $App.AppId

    # Owner assignment
    $OwnerRef = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($Owner.Id)" }
    New-MgApplicationOwnerByRef     -ApplicationId    $App.Id -BodyParameter $OwnerRef
    New-MgServicePrincipalOwnerByRef -ServicePrincipalId $Sp.Id -BodyParameter $OwnerRef

    [PSCustomObject]@{
        DisplayName = $Name
        AppClientId = $App.AppId
        AppObjectId = $App.Id
        SpObjectId  = $Sp.Id
    }
}

# ---------------- Summary ----------------
Write-Host "`nDone. Summary:" -ForegroundColor Green
$Results | Format-Table -AutoSize

# Copy-paste block: name -> client id
Write-Host "`nName / client id:" -ForegroundColor Cyan
$Results | ForEach-Object { Write-Host ("  {0} - {1}" -f $_.DisplayName, $_.AppClientId) }
