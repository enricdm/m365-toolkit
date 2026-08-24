<#
.SYNOPSIS
    Enumerates every user excluded from MFA enforcement across Conditional Access
    policies, expanding excluded GROUPS to their full transitive membership
    (direct members + nested-group members).

.DESCRIPTION
    A user is treated as excluded from MFA when, on any in-scope MFA-enforcing CA
    policy, they are:
      (a) directly listed in ExcludeUsers, OR
      (b) a member of a group in ExcludeGroups, OR
      (c) a member of a group nested inside such a group (transitive).

    SCOPING NOTE: by default ALL enabled MFA-enforcing policies are evaluated. In a
    tenant with many app/country-scoped policies this over-counts - being excluded
    from e.g. a Salesforce-only MFA policy does not mean a user is exempt from the
    org-wide MFA requirement. Use -PolicyName / -PolicyId to scope to the baseline
    policy when the goal is to measure true MFA exemptions. The per-group and
    per-policy diagnostics printed during the run show which groups inflate the count.

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

.PARAMETER IncludeReportOnly
    Also evaluate policies in 'enabledForReportingButNotEnforced' state.

.PARAMETER CompareCsv
    Optional path to a CSV with a UserPrincipalName column already triaged. Excluded
    users not present are flagged InReport = No (the set the report is missing).

.EXAMPLE
    .\Get-MfaExclusion.ps1 -TenantId '<tenant-id>'
    All enabled MFA-enforcing policies. Over-counts in a tenant with many scoped
    policies — read the diagnostics table before quoting the total.

.EXAMPLE
    .\Get-MfaExclusion.ps1 -TenantId '<tenant-id>' -PolicyName '*Require MFA*' `
        -MasterExceptionGroup 'CA-Exception-MFA'
    Scoped to the baseline policy — this is the number that means "exempt from MFA".

.NOTES
    When to use  : Security asks how many people are exempt from MFA and the number the portal shows does not match reality.
    Why it exists: Excluded groups are expanded transitively, including nested groups, which is exactly what the policy view hides. The per-group diagnostics show which group is inflating the count, and the run warns that evaluating every MFA policy over-counts.
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
$rows = $rows | Sort-Object `
@{E = { if ($_.AccountEnabled -eq $true) { 0 }else { 1 } } }, `
@{E = { if ($_.MfaRegistered -eq $false) { 0 }elseif ($_.MfaRegistered -eq 'unknown') { 1 }else { 2 } } }, `
    UserPrincipalName

# ---- export + diagnostics -------------------------------------------------
if (-not (Test-Path $ExportDir)) { New-Item -ItemType Directory -Path $ExportDir | Out-Null }
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$csv = Join-Path $ExportDir "CA-MFA-Exclusions_$stamp.csv"
$rows | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8

$diag = Join-Path $ExportDir "CA-MFA-ExclusionSources_$stamp.csv"
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