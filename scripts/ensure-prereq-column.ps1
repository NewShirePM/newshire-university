<#
.SYNOPSIS
  One-time setup for NewShire University course prerequisites.

  Adds one column:
    TrainingCourses:  PrerequisiteCourseCode (Text)

  How it works in the app:
    - PrerequisiteCourseCode holds a COURSE CODE (e.g. "FHC 101"), not an item id.
      Item ids differ between the live site and any rebuild, and they are meaningless
      to a human reading the SharePoint list. A code survives re-provisioning.
    - A learner who has never PASSED the prerequisite sees the target course locked:
      greyed on the Training Library card, not clickable, and blocked on deep link.
    - The gate is a CURRENT certification, not merely a past pass. If the prerequisite
      carries a recert interval and the learner has let it lapse, the dependent course
      re-locks until they recertify. A prerequisite with no recert interval stays
      satisfied forever once passed. A cert inside its 30-day "expiring" window still
      counts - it has not lapsed yet.
    - A prerequisite naming a course that does not exist yet locks nobody. That is a
      data error, not a gate, so it fails open.

  Set the value from the app: Admin > Courses > edit a course > Prerequisite.
  Leave it as None for most courses. Current intended use is FHC 301 -> FHC 101.

  Safe to re-run — it skips anything already present.

.NOTES
  Requires PnP.PowerShell:  Install-Module PnP.PowerShell -Scope CurrentUser
  Run with:  pwsh ./scripts/ensure-prereq-column.ps1
#>

param(
  [string]$SiteUrl     = "https://newshirepmcom.sharepoint.com/sites/NewShirePM",
  [string]$ClientId    = "32e75ffa-747a-4cf0-8209-6a19150c4547",
  [string]$CoursesList = "TrainingCourses"
)

$ErrorActionPreference = "Stop"

Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

function Add-FieldIfMissing {
  param([string]$List, [string]$InternalName, [string]$Xml)
  $existingField = Get-PnPField -List $List | Where-Object InternalName -eq $InternalName
  if ($existingField) {
    Write-Host ("  [OK]    {0}.{1} already exists" -f $List, $InternalName) -ForegroundColor Green
  } else {
    Add-PnPFieldFromXml -List $List -FieldXml $Xml | Out-Null
    Write-Host ("  [ADDED] {0}.{1}" -f $List, $InternalName) -ForegroundColor Cyan
  }
}

$lists = (Get-PnPList).Title
if ($lists -notcontains $CoursesList) {
  Write-Host "`n'$CoursesList' does not exist on this site. Re-run with the correct -SiteUrl." -ForegroundColor Red
  Disconnect-PnPOnline
  exit 1
}

Write-Host "`nEnsuring prerequisite column ..." -ForegroundColor Cyan

Add-FieldIfMissing -List $CoursesList -InternalName "PrerequisiteCourseCode" `
  -Xml "<Field Type='Text' DisplayName='PrerequisiteCourseCode' Name='PrerequisiteCourseCode' StaticName='PrerequisiteCourseCode' MaxLength='32' />"

# ── Integrity report ─────────────────────────────────────────────────────────
# A prerequisite pointing at a code that does not exist fails open in the app, so it
# will not lock anyone out — but it also will not do anything, and nobody will notice.
# Report it here instead.
Write-Host "`nChecking existing prerequisite values ..." -ForegroundColor Cyan

$courses = Get-PnPListItem -List $CoursesList -PageSize 500
$codes = @{}
foreach ($c in $courses) {
  $code = $c.FieldValues["CourseCode"]
  if ($code) { $codes[$code.Trim()] = $c.FieldValues["Title"] }
}

$issues = 0
foreach ($c in $courses) {
  $pre = $c.FieldValues["PrerequisiteCourseCode"]
  if (-not $pre) { continue }
  $pre = $pre.Trim()
  $title = $c.FieldValues["Title"]
  $own = $c.FieldValues["CourseCode"]
  if ($pre -eq $own) {
    Write-Host ("  [ERROR] '{0}' lists itself as its own prerequisite" -f $title) -ForegroundColor Red
    $issues++
  } elseif (-not $codes.ContainsKey($pre)) {
    Write-Host ("  [WARN]  '{0}' requires '{1}', which is not a course code on this site" -f $title, $pre) -ForegroundColor Yellow
    $issues++
  } else {
    Write-Host ("  [OK]    '{0}' requires '{1}'" -f $title, $pre) -ForegroundColor Green
  }
}

# Cycle check. Two courses each requiring the other locks both permanently.
foreach ($c in $courses) {
  $pre = $c.FieldValues["PrerequisiteCourseCode"]
  $own = $c.FieldValues["CourseCode"]
  if (-not $pre -or -not $own) { continue }
  $other = $courses | Where-Object { $_.FieldValues["CourseCode"] -eq $pre.Trim() } | Select-Object -First 1
  if ($other -and $other.FieldValues["PrerequisiteCourseCode"] -eq $own.Trim()) {
    Write-Host ("  [ERROR] Cycle: '{0}' and '{1}' require each other. Both are permanently locked." -f $own, $pre) -ForegroundColor Red
    $issues++
  }
}

if ($issues -eq 0) { Write-Host "  No prerequisite issues found." -ForegroundColor Green }

Write-Host "`nDone. Set prerequisites in the app: Admin > Courses > edit > Prerequisite." -ForegroundColor Cyan
Disconnect-PnPOnline
