#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Applications, Microsoft.Graph.Identity.SignIns, Microsoft.Graph.Users, Microsoft.Graph.Groups

<#
.SYNOPSIS
    Exports all Entra ID App Registrations with hygiene-relevant data to CSV.

.DESCRIPTION
    Connects to Microsoft Graph using interactive (delegated) authentication and exports:
      - App Registration details (name, App ID, object ID, creation date, publisher domain, audience)
      - Owners (display name + UPN)
      - Service Principal presence, status, verified publisher, Microsoft-disabled status
      - Service Principal last sign-in activity (signInActivity — persisted beyond log window)
      - Sign-in activity from audit logs (last 30 days, delegated + non-interactive + SP)
      - Secrets and certificates with expiration dates + credential lifetime flags
      - Federated credentials (OIDC / workload identity) with issuer + subject
      - API permissions (application + delegated, admin consent status)
      - Pre-authorized applications (implicit trust relationships)
      - App roles defined by this app
      - Directory role memberships of the SP (privileged roles)
      - Group memberships of the SP
      - Identifier URIs, redirect URIs, homepage, logout URL
      - Hygiene flag columns for quick Excel filtering

.NOTES
    When to use  : Someone asks which app registrations can be deleted, or an expired secret has just broken an integration and you need to know how many more are queued up behind it.
    Why it exists: No portal view joins credentials, owners, API permissions and activity in one place. It also chunks the sign-in log query day by day, because Graph's server-side cursor times out mid-pagination on large tenants, and it tells you how many chunks came back empty.
    Required Graph permissions (delegated, granted to your user):
      Application.Read.All
      Directory.Read.All
      AuditLog.Read.All
      User.Read.All
      RoleManagement.Read.Directory
      GroupMember.Read.All

    Install modules if needed:
      Install-Module Microsoft.Graph -Scope CurrentUser

    Output: AppRegistrations_<timestamp>.csv in the current directory.
#>

[CmdletBinding()]
param(
    [string]$TenantId = "",          # Leave empty to use the home tenant of your account
    [int]$SignInLogDays = 30,        # Days to look back in sign-in logs (max 30 without P2, up to 90 with P2)
    [int]$LongLifetimeDays = 730,    # Flag credentials with lifetime > this many days (default 2 years)
    [int]$ExpiryWarningDays = 60,    # Report credentials expiring within this many days
    [string]$OutputPath = "."        # Folder for the CSV output
)

#region ── Connect ────────────────────────────────────────────────────────────

$connectParams = @{
    Scopes = @(
        "Application.Read.All",
        "Directory.Read.All",
        "AuditLog.Read.All",
        "User.Read.All",
        "RoleManagement.Read.Directory",
        "GroupMember.Read.All"
    )
}
if ($TenantId) { $connectParams.TenantId = $TenantId }

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph @connectParams -NoWelcome

#endregion

#region ── Load reference data ───────────────────────────────────────────────

Write-Host "Loading service principals (for permission + role resolution)..." -ForegroundColor Cyan

$allServicePrincipals = Get-MgServicePrincipal -All -Property `
    Id, AppId, DisplayName, AppRoles, AccountEnabled,
    SignInActivity, VerifiedPublisher, DisabledByMicrosoftStatus

$spByAppId = @{}
foreach ($sp in $allServicePrincipals) {
    $spByAppId[$sp.AppId] = $sp
}

# All app roles across all SPs (for permission name resolution)
$allAppRoles = $allServicePrincipals | ForEach-Object { $_.AppRoles } | Where-Object { $_ }

# Load directory roles to map principalId -> role name
Write-Host "Loading directory role assignments..." -ForegroundColor Cyan
$dirRoleAssignments = @{}
# Flag_HasDirectoryRole is derived from this table, so a role whose members fail to load
# makes privileged service principals look unprivileged. Count the failures and say so,
# rather than letting the flag read "No" for a reason that has nothing to do with the app.
$roleReadFailures = [System.Collections.Generic.List[string]]::new()
try {
    $activeRoles = Get-MgDirectoryRole -All -ErrorAction Stop
    foreach ($role in $activeRoles) {
        try {
            $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All -ErrorAction Stop
        } catch {
            $roleReadFailures.Add("$($role.DisplayName): $($_.Exception.Message)")
            continue
        }
        foreach ($m in $members) {
            if (-not $dirRoleAssignments.ContainsKey($m.Id)) {
                $dirRoleAssignments[$m.Id] = @()
            }
            $dirRoleAssignments[$m.Id] += $role.DisplayName
        }
    }
} catch {
    Write-Warning "Could not load directory roles at all: $_. Flag_HasDirectoryRole will be False on every row for that reason, not because no app holds a role."
}
if ($roleReadFailures.Count -gt 0) {
    Write-Warning "$($roleReadFailures.Count) directory role(s) could not be enumerated. Any service principal holding only those roles will show Flag_HasDirectoryRole = False incorrectly:"
    $roleReadFailures | ForEach-Object { Write-Warning "    $_" }
}

#endregion

#region ── Sign-in activity from audit logs ───────────────────────────────────

Write-Host "Querying sign-in logs (last $SignInLogDays days)..." -ForegroundColor Cyan

$since = (Get-Date).AddDays(-$SignInLogDays).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

# Sign-in logs — paginate manually via Invoke-MgGraphRequest because Get-MgAuditLogSignIn -All
# has known issues with skip-token expiration on large tenants.
#
# Additionally, querying 30 days in one go frequently fails mid-pagination with BadRequest
# on large tenants (Microsoft appears to time out server-side cursors after extended runs).
# We split the query into daily chunks to keep each query short, and retry on transient failures.
$signInLogs = @{}

function Get-SignInPage {
    param(
        [string]$Uri,
        [hashtable]$Store,
        [string]$Label,
        [int]$MaxRetries = 2
    )
    $pageCount = 0
    $entryCount = 0
    $currentUri = $Uri
    $retryCount = 0
    while ($currentUri) {
        try {
            $response = Invoke-MgGraphRequest -Method GET -Uri $currentUri -ErrorAction Stop
            foreach ($entry in $response.value) {
                $appId = $entry.appId
                if (-not $appId) { continue }
                $date = [datetime]$entry.createdDateTime
                if (-not $Store.ContainsKey($appId) -or $Store[$appId] -lt $date) {
                    $Store[$appId] = $date
                }
                $entryCount++
            }
            $currentUri = $response.'@odata.nextLink'
            $pageCount++
            $retryCount = 0  # Reset retry counter on success
        }
        catch {
            $retryCount++
            if ($retryCount -le $MaxRetries) {
                Write-Host "    [$Label] Page $pageCount failed, retrying ($retryCount/$MaxRetries)..." -ForegroundColor DarkYellow
                Start-Sleep -Seconds (2 * $retryCount)
                continue
            }
            Write-Warning "    [$Label] Gave up at page $pageCount after $MaxRetries retries: $($_.Exception.Message)"
            break
        }
    }
    return @{ Pages = $pageCount; Entries = $entryCount }
}

function Invoke-ChunkedSignInQuery {
    param(
        [string]$BaseUrl,
        [string]$FilterPrefix,   # empty or " and signInEventTypes/any(t: t eq 'servicePrincipal')"
        [int]$DaysBack,
        [hashtable]$Store,
        [string]$Label
    )
    $now = (Get-Date).ToUniversalTime()
    $totalPages = 0
    $totalEntries = 0
    $failedChunks = 0

    # Iterate backward day by day from today
    for ($i = 0; $i -lt $DaysBack; $i++) {
        $chunkEnd = $now.AddDays(-$i)
        $chunkStart = $now.AddDays(-$i - 1)
        $startStr = $chunkStart.ToString("yyyy-MM-ddTHH:mm:ssZ")
        $endStr   = $chunkEnd.ToString("yyyy-MM-ddTHH:mm:ssZ")

        $filter = "createdDateTime ge $startStr and createdDateTime lt $endStr$FilterPrefix"
        $encFilter = [System.Web.HttpUtility]::UrlEncode($filter)
        $chunkUri = "$BaseUrl`?`$filter=$encFilter&`$select=appId,createdDateTime&`$top=1000"

        $before = $Store.Count
        $result = Get-SignInPage -Uri $chunkUri -Store $Store -Label "$Label day-$($i+1)" -MaxRetries 2
        $totalPages += $result.Pages
        $totalEntries += $result.Entries

        $gained = $Store.Count - $before
        if ($result.Pages -eq 0 -and $result.Entries -eq 0) { $failedChunks++ }

        # Progress ping every 5 days
        if (($i + 1) % 5 -eq 0) {
            Write-Host "    [$Label] Processed $($i+1)/$DaysBack days | $totalPages pages | $totalEntries entries | $($Store.Count) distinct apps" -ForegroundColor DarkGray
        }
    }

    Write-Host "  [$Label] Complete: $totalPages pages, $totalEntries entries across $DaysBack days" -ForegroundColor Gray
    if ($failedChunks -gt 0) {
        Write-Warning "  [$Label] $failedChunks of $DaysBack daily chunks returned no data (may have failed or genuinely had no activity)"
    }
}

# Ensure URL encoding helper is available
Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue

# User sign-ins (interactive + non-interactive) on v1.0 — daily chunks
Invoke-ChunkedSignInQuery `
    -BaseUrl "https://graph.microsoft.com/v1.0/auditLogs/signIns" `
    -FilterPrefix "" `
    -DaysBack $SignInLogDays `
    -Store $signInLogs `
    -Label "User sign-ins"

# Service principal sign-ins (app-only auth) on /beta — daily chunks
Write-Host "Querying service principal sign-in logs..." -ForegroundColor Cyan
Invoke-ChunkedSignInQuery `
    -BaseUrl "https://graph.microsoft.com/beta/auditLogs/signIns" `
    -FilterPrefix " and signInEventTypes/any(t: t eq 'servicePrincipal')" `
    -DaysBack $SignInLogDays `
    -Store $signInLogs `
    -Label "SP sign-ins"

Write-Host "  Combined sign-in coverage: $($signInLogs.Count) distinct apps with activity in last $SignInLogDays days." -ForegroundColor Green

#endregion

#region ── Enumerate App Registrations ───────────────────────────────────────

Write-Host "Loading app registrations..." -ForegroundColor Cyan

$apps = Get-MgApplication -All -Property `
    Id, AppId, DisplayName, CreatedDateTime,
    SignInAudience, Tags, Notes, Description,
    PasswordCredentials, KeyCredentials,
    RequiredResourceAccess, PublisherDomain,
    Web, Spa, PublicClient,
    IdentifierUris, AppRoles, Api,
    IsFallbackPublicClient, DisabledByMicrosoftStatus

Write-Host "  Found $($apps.Count) app registrations. Processing..." -ForegroundColor Gray

$export = [System.Collections.Generic.List[PSCustomObject]]::new()

$i = 0
foreach ($app in $apps) {
    $i++
    Write-Progress -Activity "Processing app registrations" -Status "$i / $($apps.Count): $($app.DisplayName)" -PercentComplete (($i / $apps.Count) * 100)

    # ── Owners ──────────────────────────────────────────────────────────────
    $ownerNames = try {
        (Get-MgApplicationOwner -ApplicationId $app.Id -All |
            ForEach-Object {
                $name = if ($_.AdditionalProperties.ContainsKey("displayName")) { $_.AdditionalProperties["displayName"] } else { $_.Id }
                $upn  = if ($_.AdditionalProperties.ContainsKey("userPrincipalName")) { $_.AdditionalProperties["userPrincipalName"] } else { "" }
                if ($upn) { "$name <$upn>" } else { $name }
            }) -join " | "
    } catch { "Error reading owners" }

    # ── Service Principal ────────────────────────────────────────────────────
    $sp = $spByAppId[$app.AppId]
    $spExists  = if ($sp) { "Yes" } else { "No" }
    $spEnabled = if ($sp) { if ($sp.AccountEnabled) { "Yes" } else { "No" } } else { "N/A" }

    # Verified publisher
    $verifiedPublisher = if ($sp -and $sp.VerifiedPublisher -and $sp.VerifiedPublisher.DisplayName) {
        $sp.VerifiedPublisher.DisplayName
    } else { "Not verified" }

    # Disabled by Microsoft
    $disabledByMs = if ($app.DisabledByMicrosoftStatus) { $app.DisabledByMicrosoftStatus } else { "No" }

    # SP last sign-in (from signInActivity — persists beyond log rolling window)
    $spLastSignIn = "N/A"
    if ($sp -and $sp.SignInActivity) {
        if ($sp.SignInActivity.LastSignInDateTime) {
            $spLastSignIn = $sp.SignInActivity.LastSignInDateTime.ToString("yyyy-MM-dd HH:mm:ss")
        }
    }

    # ── Secrets (PasswordCredentials) ───────────────────────────────────────
    $secrets = $app.PasswordCredentials
    $secretDetails = "None"
    $longLifetimeSecrets = 0
    if ($secrets) {
        $detailParts = @()
        foreach ($s in $secrets) {
            $daysLeft = [math]::Round((($s.EndDateTime) - (Get-Date)).TotalDays)
            $lifetime = [math]::Round((($s.EndDateTime) - ($s.StartDateTime)).TotalDays)
            $status = if ($daysLeft -lt 0) { "EXPIRED" } elseif ($daysLeft -le 30) { "EXPIRING_SOON" } else { "OK" }
            if ($lifetime -gt $LongLifetimeDays) { $longLifetimeSecrets++ }
            $name = if ($s.DisplayName) { $s.DisplayName } else { "(unnamed)" }
            $detailParts += "$name | Expires: $($s.EndDateTime.ToString('yyyy-MM-dd')) [$status, $daysLeft days] | Lifetime: $lifetime days"
        }
        $secretDetails = $detailParts -join " || "
    }

    $secretCount  = if ($secrets) { $secrets.Count } else { 0 }
    $expiredCount = if ($secrets) { @($secrets | Where-Object { $_.EndDateTime -lt (Get-Date) }).Count } else { 0 }
    $expiringSoon = if ($secrets) { @($secrets | Where-Object { $_.EndDateTime -gt (Get-Date) -and $_.EndDateTime -lt (Get-Date).AddDays($ExpiryWarningDays) }).Count } else { 0 }
    $nextSecretExpiry = if ($secrets) {
        $upcoming = @($secrets | Where-Object { $_.EndDateTime -gt (Get-Date) } | Sort-Object EndDateTime)
        if ($upcoming.Count -gt 0) { $upcoming[0].EndDateTime.ToString("yyyy-MM-dd") } else { "" }
    } else { "" }

    # ── Certificates (KeyCredentials) ────────────────────────────────────────
    $certs = $app.KeyCredentials
    $certDetails = "None"
    $certExpiredCount = 0
    $certExpiringSoon = 0
    $nextCertExpiry = ""
    if ($certs) {
        $detailParts = @()
        foreach ($c in $certs) {
            $daysLeft = [math]::Round((($c.EndDateTime) - (Get-Date)).TotalDays)
            $status = if ($daysLeft -lt 0) { "EXPIRED" } elseif ($daysLeft -le $ExpiryWarningDays) { "EXPIRING_SOON" } else { "OK" }
            if ($daysLeft -lt 0) { $certExpiredCount++ }
            elseif ($daysLeft -le $ExpiryWarningDays) { $certExpiringSoon++ }
            $name = if ($c.DisplayName) { $c.DisplayName } else { "(unnamed)" }
            $detailParts += "$name | Expires: $($c.EndDateTime.ToString('yyyy-MM-dd')) [$status, $daysLeft days] | Type: $($c.Type)"
        }
        $certDetails = $detailParts -join " || "
        $upcomingCerts = @($certs | Where-Object { $_.EndDateTime -gt (Get-Date) } | Sort-Object EndDateTime)
        if ($upcomingCerts.Count -gt 0) { $nextCertExpiry = $upcomingCerts[0].EndDateTime.ToString("yyyy-MM-dd") }
    }
    $certCount = if ($certs) { $certs.Count } else { 0 }

    # ── Federated Credentials (OIDC / workload identity) ────────────────────
    $fedCredDetails = "None"
    $fedCredCount = 0
    try {
        $fedCreds = Get-MgApplicationFederatedIdentityCredential -ApplicationId $app.Id -ErrorAction Stop
        if ($fedCreds) {
            $fedCredCount = $fedCreds.Count
            $detailParts = @()
            foreach ($f in $fedCreds) {
                $audiences = if ($f.Audiences) { $f.Audiences -join "," } else { "(none)" }
                $detailParts += "$($f.Name) | Issuer: $($f.Issuer) | Subject: $($f.Subject) | Audiences: $audiences"
            }
            $fedCredDetails = $detailParts -join " || "
        }
    } catch { $fedCredDetails = "Error reading federated credentials" }

    # ── Permissions ──────────────────────────────────────────────────────────
    $permissionDetail = ""
    $adminConsent = "N/A"
    $appPermCount = 0
    $delPermCount = 0

    if ($sp) {
        try {
            $appRoleAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -All
            $oauth2Grants = Get-MgServicePrincipalOauth2PermissionGrant -ServicePrincipalId $sp.Id -All

            $permList = @()

            foreach ($assignment in $appRoleAssignments) {
                $resourceSp   = $allServicePrincipals | Where-Object Id -eq $assignment.ResourceId
                $resourceName = if ($resourceSp) { $resourceSp.DisplayName } else { $assignment.ResourceId }
                $roleId       = $assignment.AppRoleId
                $roleName     = ($allAppRoles | Where-Object Id -eq $roleId | Select-Object -First 1).Value
                if (-not $roleName) { $roleName = $roleId }
                $permList += "[APP] $resourceName / $roleName"
                $appPermCount++
            }

            foreach ($grant in $oauth2Grants) {
                $resourceSp   = $allServicePrincipals | Where-Object Id -eq $grant.ResourceId
                $resourceName = if ($resourceSp) { $resourceSp.DisplayName } else { $grant.ResourceId }
                foreach ($scope in ($grant.Scope -split " " | Where-Object { $_ })) {
                    $permList += "[DEL] $resourceName / $scope"
                    $delPermCount++
                }
            }

            $permissionDetail = $permList -join " | "

            $adminConsent = if ($oauth2Grants | Where-Object { $_.ConsentType -eq "AllPrincipals" }) {
                "Yes"
            } elseif ($appRoleAssignments.Count -gt 0) {
                "Yes (App perms require admin consent)"
            } else {
                "No"
            }
        }
        catch { $permissionDetail = "Error reading permissions"; $adminConsent = "Error" }
    } else {
        $permList = @()
        foreach ($rra in $app.RequiredResourceAccess) {
            $resourceSp   = $allServicePrincipals | Where-Object AppId -eq $rra.ResourceAppId
            $resourceName = if ($resourceSp) { $resourceSp.DisplayName } else { $rra.ResourceAppId }
            foreach ($ra in $rra.ResourceAccess) {
                $type = if ($ra.Type -eq "Role") { "[APP]"; $appPermCount++ } else { "[DEL]"; $delPermCount++ }
                $permList += "$type $resourceName / $($ra.Id)"
            }
        }
        $permissionDetail = if ($permList) { $permList -join " | " } else { "None declared" }
        $adminConsent = "No SP (uninstalled)"
    }

    # ── Pre-authorized applications ──────────────────────────────────────────
    $preAuthApps = "None"
    $preAuthCount = 0
    if ($app.Api -and $app.Api.PreAuthorizedApplications) {
        $preAuthCount = $app.Api.PreAuthorizedApplications.Count
        $detailParts = @()
        foreach ($pa in $app.Api.PreAuthorizedApplications) {
            $resolvedSp = $allServicePrincipals | Where-Object AppId -eq $pa.AppId | Select-Object -First 1
            $resolvedName = if ($resolvedSp) { $resolvedSp.DisplayName } else { $pa.AppId }
            $scopeCount = if ($pa.DelegatedPermissionIds) { $pa.DelegatedPermissionIds.Count } else { 0 }
            $detailParts += "$resolvedName ($scopeCount scopes)"
        }
        $preAuthApps = $detailParts -join " | "
    }

    # ── App roles defined by this app ────────────────────────────────────────
    $definedRoles = "None"
    $definedRoleCount = 0
    if ($app.AppRoles) {
        $definedRoleCount = $app.AppRoles.Count
        $definedRoles = ($app.AppRoles | ForEach-Object { "$($_.Value) ($(if($_.IsEnabled){'enabled'}else{'disabled'}))" }) -join " | "
    }

    # ── Directory role memberships ───────────────────────────────────────────
    $spDirectoryRoles = "None"
    if ($sp -and $dirRoleAssignments.ContainsKey($sp.Id)) {
        $spDirectoryRoles = $dirRoleAssignments[$sp.Id] -join " | "
    }

    # ── Group memberships of the SP ──────────────────────────────────────────
    $spGroups = "None"
    $spGroupCount = 0
    if ($sp) {
        try {
            $memberOf = Get-MgServicePrincipalMemberOf -ServicePrincipalId $sp.Id -All -ErrorAction Stop
            $groupNames = $memberOf |
                Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.group' } |
                ForEach-Object { $_.AdditionalProperties['displayName'] }
            if ($groupNames) {
                $spGroupCount = $groupNames.Count
                $spGroups = $groupNames -join " | "
            }
        } catch { $spGroups = "Error" }
    }

    # ── Last used (sign-in log) ───────────────────────────────────────────────
    $lastUsedFromLog = if ($signInLogs.ContainsKey($app.AppId)) {
        $signInLogs[$app.AppId].ToString("yyyy-MM-dd HH:mm:ss")
    } else {
        "No activity in last $SignInLogDays days"
    }

    # ── URIs ──────────────────────────────────────────────────────────────────
    $redirectUris = @()
    if ($app.Web -and $app.Web.RedirectUris) { $redirectUris += $app.Web.RedirectUris }
    if ($app.Spa -and $app.Spa.RedirectUris) { $redirectUris += $app.Spa.RedirectUris }
    if ($app.PublicClient -and $app.PublicClient.RedirectUris) { $redirectUris += $app.PublicClient.RedirectUris }
    $redirectUriStr = if ($redirectUris) { $redirectUris -join " | " } else { "None" }

    $homepage = if ($app.Web -and $app.Web.HomePageUrl) { $app.Web.HomePageUrl } else { "" }
    $logoutUrl = if ($app.Web -and $app.Web.LogoutUrl) { $app.Web.LogoutUrl } else { "" }
    $identifierUris = if ($app.IdentifierUris) { $app.IdentifierUris -join " | " } else { "None" }

    # ── Assemble record ───────────────────────────────────────────────────────
    $export.Add([PSCustomObject]@{
        # Identity
        DisplayName                 = $app.DisplayName
        AppId                       = $app.AppId
        ObjectId                    = $app.Id
        CreatedDateTime             = if ($app.CreatedDateTime) { $app.CreatedDateTime.ToString("yyyy-MM-dd") } else { "" }
        PublisherDomain             = $app.PublisherDomain
        SignInAudience              = $app.SignInAudience
        Description                 = $app.Description
        DisabledByMicrosoft         = $disabledByMs

        # Ownership
        Owners                      = $ownerNames

        # Service Principal
        HasServicePrincipal         = $spExists
        ServicePrincipalEnabled     = $spEnabled
        VerifiedPublisher           = $verifiedPublisher
        SP_LastSignInActivity       = $spLastSignIn
        LastUsed_SignInLog          = $lastUsedFromLog

        # Secrets
        SecretCount                 = $secretCount
        SecretsExpired              = $expiredCount
        SecretsExpiringSoon         = $expiringSoon
        SecretsLongLifetime         = $longLifetimeSecrets
        NextSecretExpiry            = $nextSecretExpiry
        SecretDetails               = $secretDetails

        # Certificates
        CertificateCount            = $certCount
        CertificatesExpired         = $certExpiredCount
        CertificatesExpiringSoon    = $certExpiringSoon
        NextCertificateExpiry       = $nextCertExpiry
        CertificateDetails          = $certDetails

        # Federated credentials
        FederatedCredentialCount    = $fedCredCount
        FederatedCredentialDetails  = $fedCredDetails

        # Permissions
        AppPermissionCount          = $appPermCount
        DelegatedPermissionCount    = $delPermCount
        AdminConsentGranted         = $adminConsent
        Permissions                 = $permissionDetail

        # Trust relationships
        PreAuthorizedAppCount       = $preAuthCount
        PreAuthorizedApps           = $preAuthApps

        # App roles defined
        DefinedAppRoleCount         = $definedRoleCount
        DefinedAppRoles             = $definedRoles

        # Privileged access
        SP_DirectoryRoles           = $spDirectoryRoles
        SP_GroupMembershipCount     = $spGroupCount
        SP_GroupMemberships         = $spGroups

        # App configuration
        IdentifierUris              = $identifierUris
        RedirectURIs                = $redirectUriStr
        Homepage                    = $homepage
        LogoutUrl                   = $logoutUrl
        IsFallbackPublicClient      = $app.IsFallbackPublicClient
        Notes                       = $app.Notes
        Tags                        = ($app.Tags -join ", ")

        # Hygiene flags (quick-filter columns)
        Flag_NoOwner                = [string]::IsNullOrWhiteSpace($ownerNames)
        Flag_HasExpiredSecret       = ($expiredCount -gt 0)
        Flag_HasExpiringSecret      = ($expiringSoon -gt 0)
        Flag_HasExpiredCert         = ($certExpiredCount -gt 0)
        Flag_HasExpiringCert        = ($certExpiringSoon -gt 0)
        Flag_LongLifetimeSecret     = ($longLifetimeSecrets -gt 0)
        Flag_NoActivity             = (-not $signInLogs.ContainsKey($app.AppId))
        Flag_NoServicePrincipal     = ($spExists -eq "No")
        Flag_SPDisabled             = ($spEnabled -eq "No")
        Flag_HasDirectoryRole       = ($spDirectoryRoles -ne "None")
        Flag_MultiTenant            = ($app.SignInAudience -ne "AzureADMyOrg")
        Flag_PersonalMSAccount      = ($app.SignInAudience -eq "AzureADandPersonalMicrosoftAccount")
        Flag_HasAppPermissions      = ($appPermCount -gt 0)
        Flag_UnverifiedPublisher    = ($verifiedPublisher -eq "Not verified")
        Flag_DisabledByMicrosoft    = ($disabledByMs -and $disabledByMs -ne "No" -and $disabledByMs -ne "None")
        Flag_HasFederatedCreds      = ($fedCredCount -gt 0)
        Flag_HasPreAuthorizedApps   = ($preAuthCount -gt 0)
    })
}

Write-Progress -Activity "Processing app registrations" -Completed

#endregion

#region ── Export ─────────────────────────────────────────────────────────────

$timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFile = Join-Path $OutputPath "AppRegistrations_$timestamp.csv"

$export | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8

Write-Host "`nExport complete." -ForegroundColor Green
Write-Host "File: $outputFile" -ForegroundColor Yellow
Write-Host "Total apps exported: $($export.Count)"
Write-Host ""
Write-Host "── Hygiene summary ─────────────────────────────────────────────"
Write-Host "  No owner assigned:          $(($export | Where-Object Flag_NoOwner).Count)"
Write-Host "  Expired secrets:            $(($export | Where-Object Flag_HasExpiredSecret).Count)"
Write-Host "  Secrets expiring ($ExpiryWarningDays d):    $(($export | Where-Object Flag_HasExpiringSecret).Count)"
Write-Host "  Expired certificates:       $(($export | Where-Object Flag_HasExpiredCert).Count)"
Write-Host "  Certificates expiring ($ExpiryWarningDays d): $(($export | Where-Object Flag_HasExpiringCert).Count)"
Write-Host "  Long-lifetime secrets:      $(($export | Where-Object Flag_LongLifetimeSecret).Count)"
Write-Host "  No activity:                $(($export | Where-Object Flag_NoActivity).Count)"
Write-Host "  No service principal:       $(($export | Where-Object Flag_NoServicePrincipal).Count)"
Write-Host "  SP disabled:                $(($export | Where-Object Flag_SPDisabled).Count)"
Write-Host "  Has directory roles:        $(($export | Where-Object Flag_HasDirectoryRole).Count)"
Write-Host "  Has federated credentials:  $(($export | Where-Object Flag_HasFederatedCreds).Count)"
Write-Host "  Has pre-authorized apps:    $(($export | Where-Object Flag_HasPreAuthorizedApps).Count)"
Write-Host "  Multi-tenant:               $(($export | Where-Object Flag_MultiTenant).Count)"
Write-Host "  Personal MS account audience: $(($export | Where-Object Flag_PersonalMSAccount).Count)"
Write-Host "  App-level permissions:      $(($export | Where-Object Flag_HasAppPermissions).Count)"
Write-Host "  Disabled by Microsoft:      $(($export | Where-Object Flag_DisabledByMicrosoft).Count)"
Write-Host "────────────────────────────────────────────────────────────────"

Disconnect-MgGraph | Out-Null

#endregion
