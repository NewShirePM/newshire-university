<#
.SYNOPSIS
  Rewrites the role targeting stored on Learning Paths and Courses so it uses
  canonical job titles. Run this after employee-lifecycle\scripts\provision-job-roles.ps1
  -RemapTitles, and after audit-role-targeting.ps1 has shown you what's stale.

.DESCRIPTION
  Remapping PEOPLE does not fix a PATH still aimed at a retired title. A path
  targeting "Owner/Operator" now matches nobody, so it silently assigns to
  no one - the same failure mode, one level up.

  This applies $RoleRemap below to:
    LearningPaths.Roles        (comma-separated, or a multi-value column)
    TrainingCourses.CourseRoles

  Behaviour:
    - A role mapping to $null is DROPPED from the targeting.
    - Duplicates after remapping are collapsed (two old titles can map to one
      canonical role).
    - Original order is preserved; unchanged items are not written at all.
    - "All" is passed through untouched.
    - Anything non-canonical that ISN'T in the remap table is left alone and
      reported at the end, so nothing changes silently.

  The column shape is detected per item: if SharePoint returns an array
  (multi-value choice) it writes an array, otherwise a comma-joined string -
  matching what the app itself writes.

.PARAMETER WhatIf
  Dry run. Prints the full before/after role list for every item that would
  change, and writes nothing. Run this first.

.EXAMPLE
  .\fix-role-targeting.ps1 -WhatIf
  .\fix-role-targeting.ps1
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$SiteId   = "newshirepmcom.sharepoint.com,f5d74a99-8b23-477c-aed0-c1682efa5de1,0aa4ab90-0ddc-4a81-a803-06d4f2e1a7d8",
    [string]$TenantId = "5e932bcd-838c-4dae-b838-f6d22d7c6b8a",
    [string]$ClientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
)
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ============================================================================
#  ROLE REMAP for course/path TARGETING  (confirmed by Brandy, 2026-08-28)
#  Mirrors $TitleRemap in provision-job-roles.ps1, plus the two titles that
#  only ever existed inside NewShire University and were never real job titles.
#  $null = drop the role from the targeting entirely.
#
#  NOT remapped: 'Leasing Agent'. It was invented by this app and nobody holds
#  it yet, but it is a role NewShire is hiring into, so it was ADDED to the
#  canonical list instead. The paths already targeting it are correct as-is.
# ============================================================================
$RoleRemap = @{
    'Owner/Operator'                   = 'Director of Operations'
    'Assistant Property Manager'       = 'Delinquency and Collections Manager'
    'Part-time Leasing Assist - 1099'  = '1099 Leasing'
    'Maintenance Tech - Fusion Pointe' = 'Maintenance Technician'
    'Area Director'                    = $null   # drop - never a real role here
    'Service Manager'                  = $null   # drop - superseded by Maintenance Supervisor
}

function Write-Step($m) { Write-Host "==> $m" -ForegroundColor Cyan }
function Write-OK   ($m) { Write-Host "    [OK]    $m" -ForegroundColor Green }
function Write-Skip ($m) { Write-Host "    [SKIP]  $m" -ForegroundColor DarkYellow }
function Write-New  ($m) { Write-Host "    [APPLY] $m" -ForegroundColor Yellow }
function Write-Warn ($m) { Write-Host "    [WARN]  $m" -ForegroundColor Magenta }

Write-Step "Requesting device code"
$scope  = "https://graph.microsoft.com/Sites.ReadWrite.All offline_access"
$dcResp = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/devicecode" `
    -Body @{ client_id = $ClientId; scope = $scope } -ContentType 'application/x-www-form-urlencoded'

Write-Host ""
Write-Host "  +==============================================================+" -ForegroundColor Yellow
Write-Host "  |  ACTION REQUIRED - sign in to authorise the change           |" -ForegroundColor Yellow
Write-Host "  |                                                              |" -ForegroundColor Yellow
Write-Host "  |  1. Open:   https://login.microsoft.com/device               |" -ForegroundColor Yellow
Write-Host "  |  2. Enter:  $($dcResp.user_code.PadRight(50))|" -ForegroundColor Yellow
Write-Host "  |  3. Sign in as bturner@newshirepm.com                        |" -ForegroundColor Yellow
Write-Host "  +==============================================================+" -ForegroundColor Yellow
Write-Host ""

$expiresAt    = (Get-Date).AddSeconds([int]$dcResp.expires_in - 5)
$pollInterval = [int]$dcResp.interval; if ($pollInterval -lt 5) { $pollInterval = 5 }
$token = $null
while ((Get-Date) -lt $expiresAt) {
    Start-Sleep -Seconds $pollInterval
    try {
        $tokenResp = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
            -Body @{ grant_type = 'urn:ietf:params:oauth:grant-type:device_code'; client_id = $ClientId; device_code = $dcResp.device_code } `
            -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
        $token = $tokenResp.access_token; break
    } catch {
        $err = $null; try { $err = ($_.ErrorDetails.Message | ConvertFrom-Json) } catch {}
        if ($err -and $err.error -eq 'authorization_pending') { Write-Host '.' -NoNewline -ForegroundColor DarkGray; continue }
        if ($err -and $err.error -eq 'slow_down')              { $pollInterval += 5; continue }
        Write-Host ""
        $msg = $_.Exception.Message
        if ($err -and $err.error_description) { $msg = $err.error_description }
        Write-Error $msg; exit 1
    }
}
Write-Host ""
if (-not $token) { Write-Error "Authentication did not complete within the device-code lifetime."; exit 1 }
Write-OK "Authenticated"
$headers = @{ Authorization = "Bearer $token" }

function Invoke-Graph {
    param([string]$Method,[string]$Path,[object]$Body)
    $uri = "https://graph.microsoft.com/v1.0$Path"
    if ($Body) { return Invoke-RestMethod -Method $Method -Uri $uri -Headers $headers -Body ($Body | ConvertTo-Json -Depth 20 -Compress) -ContentType 'application/json' }
    return Invoke-RestMethod -Method $Method -Uri $uri -Headers $headers
}
function Get-AllGraph {
    param([string]$Path)
    $out = @()
    $next = "https://graph.microsoft.com/v1.0$Path"
    while ($next) {
        $page = Invoke-RestMethod -Method GET -Uri $next -Headers $headers
        $out += $page.value
        $next = $page.'@odata.nextLink'
    }
    return $out
}
function Split-Roles($raw) {
    if ($null -eq $raw) { return @() }
    $parts = @()
    if ($raw -is [array]) { foreach ($r in $raw) { $parts += ("$r" -split ',') } }
    else { $parts = ("$raw" -split ',') }
    return @($parts | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

Write-Step "Reading lists"
$lists = (Invoke-Graph GET "/sites/$SiteId/lists?`$select=id,name,displayName&`$top=200").value
function Find-List { param($Name) return ($lists | Where-Object { $_.name -eq $Name -or $_.displayName -eq $Name } | Select-Object -First 1) }

$lJobRoles = Find-List 'NS_JobRoles'
if (-not $lJobRoles) { Write-Error "NS_JobRoles not found. Run employee-lifecycle\scripts\provision-job-roles.ps1 first."; exit 1 }
$canonical = @{}
foreach ($it in (Get-AllGraph "/sites/$SiteId/lists/$($lJobRoles.id)/items?expand=fields&`$top=200")) {
    if ($it.fields.RoleActive -eq $false) { continue }
    if ($it.fields.Title) { $canonical[$it.fields.Title.Trim().ToLower()] = $it.fields.Title.Trim() }
}
Write-OK "$($canonical.Count) canonical role(s)"

$remapCI = @{}
foreach ($k in $RoleRemap.Keys) { $remapCI[$k.Trim().ToLower()] = $RoleRemap[$k] }

$changed = 0; $clean = 0; $leftAlone = @{}

foreach ($spec in @(
    @{ Name = 'LearningPaths';   Field = 'Roles';       Label = 'Path'   },
    @{ Name = 'TrainingCourses'; Field = 'CourseRoles'; Label = 'Course' }
)) {
    $list = Find-List $spec.Name
    if (-not $list) { Write-Warn "$($spec.Name) not found - skipped"; continue }

    Write-Host ""
    Write-Step "$($spec.Label)s ($($spec.Name).$($spec.Field))"

    foreach ($it in (Get-AllGraph "/sites/$SiteId/lists/$($list.id)/items?expand=fields&`$top=500")) {
        $raw = $it.fields.($spec.Field)
        $before = Split-Roles $raw
        if ($before.Count -eq 0) { continue }   # empty = targets everyone

        $after = @()
        foreach ($r in $before) {
            if ($r -eq 'All') { $after += $r; continue }
            $key = $r.ToLower()
            if ($remapCI.ContainsKey($key)) {
                $to = $remapCI[$key]
                if ($to) { $after += $to }      # $null => dropped
            }
            elseif ($canonical.ContainsKey($key)) { $after += $canonical[$key] }
            else {
                # Not canonical and not in the remap table: leave it, and say so.
                $after += $r
                if (-not $leftAlone.ContainsKey($r)) { $leftAlone[$r] = 0 }
                $leftAlone[$r]++
            }
        }
        # Collapse duplicates created by two old titles mapping to one role,
        # preserving first-seen order.
        $seen = @{}; $dedup = @()
        foreach ($r in $after) { if (-not $seen.ContainsKey($r.ToLower())) { $seen[$r.ToLower()] = $true; $dedup += $r } }
        $after = $dedup

        if (($before -join '|') -eq ($after -join '|')) { $clean++; continue }

        $title = $it.fields.Title
        Write-Host "    $($spec.Label) '$title'" -ForegroundColor White
        Write-Host "        before: $($before -join ', ')" -ForegroundColor DarkGray
        Write-Host "        after : $($after  -join ', ')" -ForegroundColor Cyan

        if ($after.Count -eq 0) {
            Write-Warn "Every target role was dropped. Skipping - an empty Roles value means"
            Write-Warn "'applies to everyone', which is almost certainly not what you want here."
            Write-Warn "Fix this one by hand in NewShire University."
            continue
        }

        if ($PSCmdlet.ShouldProcess("$($spec.Name) item #$($it.id) ('$title')", "$($spec.Field): $($before -join ', ') -> $($after -join ', ')")) {
            # Match the stored shape: array in, array out; string in, string out.
            $value = if ($raw -is [array]) { $after } else { ($after -join ',') }
            Invoke-Graph PATCH "/sites/$SiteId/lists/$($list.id)/items/$($it.id)/fields" @{ $spec.Field = $value } | Out-Null
            Write-New "updated"
        }
        $changed++
    }
}

Write-Host ""
Write-OK "$changed item(s) rewritten, $clean already clean"
if ($leftAlone.Count) {
    Write-Warn "$($leftAlone.Count) role(s) are neither canonical nor in the remap table - left untouched:"
    foreach ($k in ($leftAlone.Keys | Sort-Object)) { Write-Host "             '$k'  x$($leftAlone[$k])" -ForegroundColor Magenta }
    Write-Host "           Either add them to NS_JobRoles or to `$RoleRemap above." -ForegroundColor Gray
}

Write-Host ""
Write-Step "Done."
Write-Host "  Re-run .\audit-role-targeting.ps1 to confirm section 1 comes back clean." -ForegroundColor Gray
