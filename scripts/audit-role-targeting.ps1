<#
.SYNOPSIS
  READ-ONLY. Reports every place NewShire University's role targeting no longer
  lines up with the canonical job-role list (NS_JobRoles).

.DESCRIPTION
  Courses (CourseRoles) and Learning Paths (Roles) target people by EXACT
  STRING MATCH against Employees.JobTitle. A target role that matches nobody
  fails silently - the course simply never appears for anyone, with no error.

  This script surfaces three kinds of mismatch:

    1. Targeting a role that isn't canonical
       e.g. a path aimed at "Area Director" when the list says
       "Regional/Portfolio Manager". Nobody is assigned.

    2. An employee whose JobTitle isn't canonical
       They match no course or path targeting at all.

    3. A canonical role nobody currently holds
       Not a fault - just worth knowing before you build a path around it.

  Writes nothing. Fix what it finds in the apps: course/path roles in
  NewShire University (the pickers now offer only canonical roles), and job
  titles in Employee Lifecycle.

.EXAMPLE
  .\audit-role-targeting.ps1
#>
[CmdletBinding()]
param(
    [string]$SiteId   = "newshirepmcom.sharepoint.com,f5d74a99-8b23-477c-aed0-c1682efa5de1,0aa4ab90-0ddc-4a81-a803-06d4f2e1a7d8",
    [string]$TenantId = "5e932bcd-838c-4dae-b838-f6d22d7c6b8a",
    [string]$ClientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
)
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Write-Step($m) { Write-Host "==> $m" -ForegroundColor Cyan }
function Write-OK   ($m) { Write-Host "    [OK]    $m" -ForegroundColor Green }
function Write-Bad  ($m) { Write-Host "    [MISS]  $m" -ForegroundColor Red }
function Write-Warn ($m) { Write-Host "    [NOTE]  $m" -ForegroundColor Magenta }

Write-Step "Requesting device code (read-only)"
$scope  = "https://graph.microsoft.com/Sites.Read.All offline_access"
$dcResp = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/devicecode" `
    -Body @{ client_id = $ClientId; scope = $scope } -ContentType 'application/x-www-form-urlencoded'

Write-Host ""
Write-Host "  +==============================================================+" -ForegroundColor Yellow
Write-Host "  |  ACTION REQUIRED - sign in (read-only)                       |" -ForegroundColor Yellow
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
function Find-List {
    param($Lists, [string]$Name)
    return ($Lists | Where-Object { $_.name -eq $Name -or $_.displayName -eq $Name } | Select-Object -First 1)
}
# Both Courses and Paths store roles as a comma-separated string; Paths can also
# come back as a multi-value array. Flatten either shape.
function Split-Roles($raw) {
    if ($null -eq $raw) { return @() }
    $parts = @()
    if ($raw -is [array]) { foreach ($r in $raw) { $parts += ("$r" -split ',') } }
    else { $parts = ("$raw" -split ',') }
    return $parts | ForEach-Object { $_.Trim() } | Where-Object { $_ }
}

Write-Step "Reading lists"
$lists = (Invoke-RestMethod -Method GET -Headers $headers `
    -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/lists?`$select=id,name,displayName&`$top=200").value

$lJobRoles = Find-List $lists 'NS_JobRoles'
$lEmployees = Find-List $lists 'Employees'
$lCourses  = Find-List $lists 'TrainingCourses'
$lPaths    = Find-List $lists 'LearningPaths'

if (-not $lJobRoles) { Write-Error "NS_JobRoles not found. Run employee-lifecycle\scripts\provision-job-roles.ps1 first."; exit 1 }
if (-not $lEmployees) { Write-Error "Employees list not found."; exit 1 }

$canonical = @{}
foreach ($it in (Get-AllGraph "/sites/$SiteId/lists/$($lJobRoles.id)/items?expand=fields&`$top=200")) {
    if ($it.fields.RoleActive -eq $false) { continue }
    if ($it.fields.Title) { $canonical[$it.fields.Title.Trim().ToLower()] = $it.fields.Title.Trim() }
}
Write-OK "$($canonical.Count) canonical role(s)"

$employees = Get-AllGraph "/sites/$SiteId/lists/$($lEmployees.id)/items?expand=fields&`$top=500"
$heldCI = @{}
$badTitles = @{}
foreach ($e in $employees) {
    if ($e.fields.EmployeeActive -eq $false) { continue }
    $t = ("$($e.fields.JobTitle)").Trim()
    if (-not $t) { continue }
    $heldCI[$t.ToLower()] = $true
    if (-not $canonical.ContainsKey($t.ToLower())) {
        if (-not $badTitles.ContainsKey($t)) { $badTitles[$t] = 0 }
        $badTitles[$t]++
    }
}
Write-OK "$($employees.Count) employee row(s)"

# ── 1. Targeting that matches no canonical role ─────────────────────────────
Write-Host ""
Write-Step "1. Course / path targeting that matches no canonical role"
$offenders = 0
foreach ($pair in @(
    @{ List = $lCourses; Field = 'CourseRoles'; Label = 'Course' },
    @{ List = $lPaths;   Field = 'Roles';       Label = 'Path'   }
)) {
    if (-not $pair.List) { Write-Warn "$($pair.Label) list not found - skipped"; continue }
    foreach ($it in (Get-AllGraph "/sites/$SiteId/lists/$($pair.List.id)/items?expand=fields&`$top=500")) {
        $roles = Split-Roles $it.fields.($pair.Field)
        if (-not $roles) { continue }   # empty = targets everyone, which is fine
        $bad = @($roles | Where-Object { $_ -ne 'All' -and -not $canonical.ContainsKey($_.ToLower()) })
        if ($bad.Count) {
            Write-Bad "$($pair.Label) '$($it.fields.Title)' targets: $($bad -join ', ')"
            $offenders++
        }
    }
}
if ($offenders -eq 0) { Write-OK "None - all targeting uses canonical roles" }

# ── 2. Employees whose title isn't canonical ────────────────────────────────
Write-Host ""
Write-Step "2. Active employees whose JobTitle isn't canonical"
if ($badTitles.Count -eq 0) { Write-OK "None" }
else {
    foreach ($k in ($badTitles.Keys | Sort-Object)) { Write-Bad "'$k'  x$($badTitles[$k])  - matches no targeting at all" }
    Write-Host "           Fix in Employee Lifecycle (Job Title is now a picker), or add a" -ForegroundColor Gray
    Write-Host "           remap entry in employee-lifecycle\scripts\provision-job-roles.ps1." -ForegroundColor Gray
}

# ── 3. Canonical roles nobody holds ─────────────────────────────────────────
Write-Host ""
Write-Step "3. Canonical roles nobody currently holds"
$unheld = @($canonical.Values | Where-Object { -not $heldCI.ContainsKey($_.ToLower()) } | Sort-Object)
if ($unheld.Count -eq 0) { Write-OK "None - every canonical role is held by someone" }
else { foreach ($r in $unheld) { Write-Warn "$r" }; Write-Host "           Fine if these are roles you plan to hire into." -ForegroundColor Gray }

Write-Host ""
Write-Step "Done (nothing was modified)."
