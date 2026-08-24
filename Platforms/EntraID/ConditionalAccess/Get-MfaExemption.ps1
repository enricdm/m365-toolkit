<#
.SYNOPSIS
    Reports the users who are genuinely exempt from MFA, expanding exception groups
    to their full transitive membership (direct members + nested-group members).

.DESCRIPTION
    Answers "who is actually exempt from MFA?" - not "who appears in an exclusion
    list somewhere?", which is a much larger and much less useful set.

    A user is counted as EXEMPT when they are a member of the master exception
    group (transitively, including nested groups), or are directly listed in
    ExcludeUsers on an MFA-enforcing policy.

    WHY THE DEFAULT NARROWS. Being excluded from a Conditional Access policy is not
    the same as being exempt from MFA. The largest exclusion lists in a real tenant
    belong to policies whose grant control is BLOCK, and excluding somebody from a
    policy that blocks is the opposite of exempting them from MFA. An earlier
    version of this script reported every exclusion it found: on the tenant it was
    written against that produced 17,682 rows where the real answer was 192, a
    factor of 92, and the delivered report had to be filtered by hand afterwards.
    The number this script prints is the number somebody quotes to security, so it
    defaults to the one that means what its name says.

    Use -AllExclusions for the full exclusion dump. That is a legitimate question -
    "where does this user appear in any exclusion list" - but it is a different one,
    and it is labelled as such rather than being the default.

    SCOPING NOTE: by default ALL enabled MFA-enforcing policies are evaluated. In a
    tenant with many app/country-scoped policies this widens the input set - being
    excluded from e.g. a Salesforce-only MFA policy does not mean a user is exempt
    from the org-wide MFA requirement. Use -PolicyName / -PolicyId to scope to the
    baseline policy. The per-group and per-policy diagnostics printed during the run
    show which groups contribute what.

    READ-ONLY. No directory writes.

.PARAMETER PolicyName
    Wildcard filter on policy DisplayName (e.g. '*external*'). Restricts evaluation
    to matching MFA-enforcing policies.

.PARAMETER PolicyId
    One or more policy IDs to restrict evaluation to.

.PARAMETER TenantId
    Directory (tenant) ID to connect to.

.PARAMETER ClientId
.PARAMETER CertThumbprint
    App-only certificate auth. Supply both, or omit both for interactive
    delegated auth.

.PARAMETER MasterExceptionGroup
    Display name of the master exception group always expanded transitively.
    Rename to match your own tenant's exception group.

.PARAMETER AllExclusions
    Report every exclusion found, not only true exemptions. Includes exclusions from
    policies that BLOCK, which are not MFA exemptions. Use when the question is
    "where does this user appear in any exclusion list", not "who is exempt".

.PARAMETER IncludeReportOnly
    Also evaluate policies in 'enabledForReportingButNotEnforced' state.

.PARAMETER CompareCsv
    Optional path to a CSV with a UserPrincipalName column already triaged. Excluded
    users not present are flagged InReport = No (the set the report is missing).

.EXAMPLE
    .\Get-MfaExemption.ps1 -TenantId '<tenant-id>'
    True exemptions only, across all enabled MFA-enforcing policies.
    The run reports how many further exclusion rows were not counted, and why.

.EXAMPLE
    .\Get-MfaExemption.ps1 -TenantId '<tenant-id>' -PolicyName '*Require MFA*' `
        -MasterExceptionGroup 'CA-Exception-MFA'
    Scoped to the baseline policy. This is the number to quote to security.

.NOTES
    When to use  : Security asks how many people are exempt from MFA and the number the portal shows does not match reality.
    Why it exists: Two things the portal hides. Excluded groups are expanded transitively, including nested groups, so the count is the real population and not the names listed on the policy. And exclusion is separated from exemption: an exclusion from a BLOCK policy is not an MFA exemption, and counting it as one overstated a real tenant by 92x.
    Graph permissions: Policy.Read.All, Group.Read.All, GroupMember.Read.All,
                       User.Read.All, AuditLog.Read.All, Directory.Read.All
    Modules: Microsoft.Graph.Authentication, .Identity.SignIns, .Groups, .Users, .Reports
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertThumbprint,
    [string]$PolicyName,
    [string[]]$PolicyId,
    [string]$MasterExceptionGroup = 'CA-Exception-MFA',
    [switch]$IncludeReportOnly,
    [switch]$AllExclusions,
    [string]$CompareCsv,
    [string]$ExportDir = (Join-Path $PSScriptRoot 'Exports')
)

function Write-Step($m) { Write-Host "`n=> $m" -ForegroundColor Cyan }
function Write-OK  ($m) { Write-Host "   [OK]  $m" -ForegroundColor Green }
function Write-Warn($m) { Write-Host "   [!]   $m" -ForegroundColor Yellow }
function Die       ($m) { Write-Host "   [X]   $m" -ForegroundColor Red; exit 1 }

# ---- connect --------------------------------------------------------------
Write-Step "Connecting to Microsoft Graph (tenant $TenantId)"
try {
    if ($ClientId -and $CertThumbprint) {
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId `
            -CertificateThumbprint $CertThumbprint -NoWelcome -ErrorAction Stop
        Write-OK "App-only certificate auth"
    } else {
        Connect-MgGraph -TenantId $TenantId -NoWelcome -ErrorAction Stop -Scopes `
            'Policy.Read.All', 'Group.Read.All', 'GroupMember.Read.All',
        'User.Read.All', 'AuditLog.Read.All', 'Directory.Read.All'
        Write-Warn "Interactive delegated auth (no ClientId/CertThumbprint supplied)"
    }
} catch { Die "Graph connect failed: $($_.Exception.Message)" }

# ---- gather MFA-enforcing policies ---------------------------------------
Write-Step "Reading Conditional Access policies"
$wantStates = @('enabled'); if ($IncludeReportOnly) { $wantStates += 'enabledForReportingButNotEnforced' }
$allPolicies = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop
$mfaPolicies = $allPolicies | Where-Object {
    $_.State -in $wantStates -and (
        ($_.GrantControls.BuiltInControls -contains 'mfa') -or ($null -ne $_.GrantControls.AuthenticationStrength) )
}
if ($PolicyName) { $mfaPolicies = $mfaPolicies | Where-Object { $_.DisplayName -like $PolicyName } }
if ($PolicyId) { $mfaPolicies = $mfaPolicies | Where-Object { $_.Id -in $PolicyId } }
Write-OK ("{0} policies total, {1} MFA-enforcing in scope" -f $allPolicies.Count, $mfaPolicies.Count)
if (-not $mfaPolicies) { Die "No MFA-enforcing policies matched the filters." }
if (-not $PolicyName -and -not $PolicyId -and $mfaPolicies.Count -gt 5) {
    Write-Warn "Evaluating $($mfaPolicies.Count) policies - counts will include app/scoped exclusions. Use -PolicyName/-PolicyId to scope to the baseline."
}

# ---- helpers --------------------------------------------------------------
$groupNameCache = @{}
function Get-GroupName($id) {
    if ($groupNameCache.ContainsKey($id)) { return $groupNameCache[$id] }
    $n = try { (Get-MgGroup -GroupId $id -Property DisplayName -ErrorAction Stop).DisplayName } catch { $id }
    $groupNameCache[$id] = $n; return $n
}
$groupUserCache = @{}
# Groups whose membership could not be read. Tracked because a failed expansion here
# silently SHRINKS the exclusion list: the users behind that group simply never appear,
# and the report then understates how many people are exempt from MFA. A short list is
# the dangerous failure mode for this particular report, so it is reported loudly at the
# end rather than swallowed.
$unresolvedGroups = [System.Collections.Generic.List[string]]::new()
function Get-TransitiveUserIds($groupId) {
    if ($groupUserCache.ContainsKey($groupId)) { return $groupUserCache[$groupId] }
    try {
        $ids = Get-MgGroupTransitiveMember -GroupId $groupId -All -ErrorAction Stop |
        Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.user' } |
        Select-Object -ExpandProperty Id
    } catch {
        $name = Get-GroupName $groupId
        $unresolvedGroups.Add("$name ($groupId): $($_.Exception.Message)")
        Write-Warn "Could not expand group '$name' - its members are MISSING from this report, which makes the exclusion count too LOW: $($_.Exception.Message)"
        $ids = @()
    }
    $groupUserCache[$groupId] = $ids; return $ids
}

$excluded = @{}
$groupStats = New-Object System.Collections.Generic.List[object]
function Add-Excluded($userId, $source, $policy) {
    if (-not $excluded.ContainsKey($userId)) {
        $excluded[$userId] = [ordered]@{ Sources = [System.Collections.Generic.HashSet[string]]::new()
            Policies                             = [System.Collections.Generic.HashSet[string]]::new() 
        } 
    }
    [void]$excluded[$userId].Sources.Add($source)
    [void]$excluded[$userId].Policies.Add($policy)
}

# ---- expand exclusions ----------------------------------------------------
Write-Step "Expanding exclusions (direct users, groups, nested groups)"
foreach ($p in $mfaPolicies) {
    $u = $p.Conditions.Users
    $direct = @($u.ExcludeUsers | Where-Object { $_ })
    foreach ($uid in $direct) { Add-Excluded $uid 'Direct user exclusion' $p.DisplayName }
    if ($direct.Count) { $groupStats.Add([pscustomobject]@{ Policy = $p.DisplayName; Group = '(direct users)'; Users = $direct.Count }) }
    foreach ($gid in @($u.ExcludeGroups | Where-Object { $_ })) {
        $gname = Get-GroupName $gid; $members = Get-TransitiveUserIds $gid
        foreach ($mid in $members) { Add-Excluded $mid "Group: $gname" $p.DisplayName }
        $groupStats.Add([pscustomobject]@{ Policy = $p.DisplayName; Group = $gname; Users = @($members).Count })
        Write-Host ("      {0,-45} via '{1}' -> {2} users" -f $p.DisplayName.Substring(0, [Math]::Min(45, $p.DisplayName.Length)), $gname, @($members).Count)
    }
    if (@($u.ExcludeRoles | Where-Object { $_ }).Count) {
        Write-Warn ("Policy '{0}' excludes directory role(s) - role members not expanded (users/groups scope only)" -f $p.DisplayName) 
    }
}

Write-Step "Expanding master exception group '$MasterExceptionGroup' (transitive)"
$me = Get-MgGroup -Filter "displayName eq '$MasterExceptionGroup'" -Property Id, DisplayName -ErrorAction SilentlyContinue
if ($me) {
    $m = Get-TransitiveUserIds $me.Id
    foreach ($mid in $m) { Add-Excluded $mid "Nested via $MasterExceptionGroup" '(master exception group)' }
    Write-OK ("{0} transitive user members" -f @($m).Count)
} else { Write-Warn "Group '$MasterExceptionGroup' not found - skipping explicit expansion" }
Write-OK ("{0} distinct excluded users" -f $excluded.Count)

# ---- BULK user resolution (replaces per-user Get-MgUser) ------------------
Write-Step "Bulk-loading directory users (single paged pull)"
$userMap = @{}
Get-MgUser -All -Property Id, UserPrincipalName, DisplayName, AccountEnabled, UserType, Mail -ErrorAction Stop |
ForEach-Object { $userMap[$_.Id] = $_ }
Write-OK ("{0} users cached" -f $userMap.Count)

# ---- MFA registration report ---------------------------------------------
Write-Step "Loading authentication method registration details"
$reg = @{}
try {
    Get-MgReportAuthenticationMethodUserRegistrationDetail -All -ErrorAction Stop | ForEach-Object { $reg[$_.Id] = $_ }
    Write-OK ("{0} registration records" -f $reg.Count) 
} catch { Write-Warn "Could not load registration report: $($_.Exception.Message)" }

# ---- optional compare set -------------------------------------------------
$known = $null
if ($CompareCsv -and (Test-Path $CompareCsv)) {
    $known = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    Import-Csv $CompareCsv | ForEach-Object {
        $v = $_.UserPrincipalName; if (-not $v) { $v = ($_.PSObject.Properties | Select-Object -First 1).Value }
        if ($v) { [void]$known.Add($v.Trim()) } }
    Write-OK ("$($known.Count) known UPNs from $CompareCsv")
}

# ---- build output (local lookups, no per-user calls) ----------------------
Write-Step "Building output"
$rows = foreach ($uid in $excluded.Keys) {
    $usr = $userMap[$uid]; $rd = $reg[$uid]
    $upn = if ($usr) { $usr.UserPrincipalName } else { $uid }
    [pscustomobject]@{
        UserPrincipalName = $upn
        DisplayName       = if ($usr) { $usr.DisplayName }else { '(unresolved)' }
        AccountEnabled    = if ($usr) { $usr.AccountEnabled }else { $null }
        UserType          = if ($usr) { $usr.UserType }else { $null }
        MfaRegistered     = if ($rd) { $rd.IsMfaRegistered }else { 'unknown' }
        IsMfaCapable      = if ($rd) { $rd.IsMfaCapable }else { 'unknown' }
        MethodsRegistered = if ($rd) { ($rd.MethodsRegistered -join '; ') }else { '' }
        ExcludedVia       = ($excluded[$uid].Sources -join ' | ')
        Policies          = (($excluded[$uid].Policies | Where-Object { $_ -ne '(master exception group)' }) -join ' | ')
        InReport          = if ($known -ne $null) { if ($known.Contains($upn)) { 'Yes' }else { 'No' } } else { 'n/a' }
        ObjectId          = $uid
    }
}
# ---- narrow to true exemptions -------------------------------------------
# Being excluded from a policy is not the same as being exempt from MFA. The
# largest exclusion lists in a real tenant belong to policies whose grant control
# is BLOCK - excluding someone from a policy that blocks is the opposite of
# exempting them from MFA. Reporting those as "MFA exemptions" overstates the
# number enormously; on the tenant this was written against, by 92x.
#
# A true exemption is membership of the master exception group, or a direct user
# exclusion on an MFA-enforcing policy. Everything else is reported only with
# -AllExclusions, which is a different question and is labelled as one.
$allRows = $rows
if (-not $AllExclusions) {
    $rows = @($rows | Where-Object {
            $_.ExcludedVia -match [regex]::Escape($MasterExceptionGroup) -or
            $_.ExcludedVia -match 'Direct user exclusion'
        })
    $dropped = $allRows.Count - $rows.Count
    if ($dropped -gt 0) {
        Write-OK ("{0} exemptions. {1} further exclusion rows were NOT counted - they are exclusions from other policies, including block policies, which are not MFA exemptions. Use -AllExclusions to see them." -f $rows.Count, $dropped)
    }
    if ($rows.Count -eq 0 -and $allRows.Count -gt 0) {
        Write-Warn "No row matched the exemption criteria, but $($allRows.Count) exclusion rows exist. Check -MasterExceptionGroup names your real exception group, or use -AllExclusions."
    }
}

$rows = $rows | Sort-Object `
@{E = { if ($_.AccountEnabled -eq $true) { 0 }else { 1 } } }, `
@{E = { if ($_.MfaRegistered -eq $false) { 0 }elseif ($_.MfaRegistered -eq 'unknown') { 1 }else { 2 } } }, `
    UserPrincipalName

# ---- export + diagnostics -------------------------------------------------
if (-not (Test-Path $ExportDir)) { New-Item -ItemType Directory -Path $ExportDir | Out-Null }
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$csv = Join-Path $ExportDir "CA-MFA-Exemptions_$stamp.csv"
$rows | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8

$diag = Join-Path $ExportDir "CA-MFA-ExemptionSources_$stamp.csv"
$groupStats | Sort-Object Users -Descending | Export-Csv -Path $diag -NoTypeInformation -Encoding UTF8

Write-Step "Done"
Write-OK "Exported $($rows.Count) excluded users -> $csv"
Write-OK "Exclusion-source breakdown -> $diag"

# A group that could not be expanded means users are MISSING from the export, so the
# total below is a floor rather than a count. Said here, in red, next to the number.
if ($unresolvedGroups.Count -gt 0) {
    Write-Warn ("{0} group(s) could not be expanded. The {1} figure above is a LOWER BOUND - the members of these groups are exempt but are NOT in the export:" -f $unresolvedGroups.Count, $rows.Count)
    $unresolvedGroups | ForEach-Object { Write-Warn "    $_" }
    Write-Warn "    Do not quote this total as the number of MFA-exempt users until these resolve."
}
Write-Host "`n   Top exclusion sources (which group/policy drives the count):" -ForegroundColor Cyan
$groupStats | Sort-Object Users -Descending | Select-Object -First 12 | Format-Table -AutoSize

$vuln = @($rows | Where-Object { $_.AccountEnabled -eq $true -and $_.MfaRegistered -ne $true })
Write-Warn ("{0} enabled with NO MFA registered (or unknown) - self-enrolment risk" -f $vuln.Count)
if ($known -ne $null) {
    $gap = @($rows | Where-Object { $_.InReport -eq 'No' })
    Write-Warn ("{0} excluded users NOT in {1}" -f $gap.Count, (Split-Path $CompareCsv -Leaf)) 
}