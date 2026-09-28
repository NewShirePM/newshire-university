<#
.SYNOPSIS
  One-time setup for NewShire University pre-hire courses.

  Adds one column:
    TrainingCourses:  PrehireAvailable (Yes/No, default No)

  How it works in the app:
    - A pre-hire is anyone in Employees whose StartDate is still in the future.
      Employee Lifecycle creates that row when the onboarding journey is started,
      so start the journey (and create their M365 account) as soon as the offer is
      accepted.
    - Before StartDate they see ONLY courses with PrehireAvailable = Yes, all
      optional. No Training Library, no due dates, no reminder emails, and they are
      left out of Team Compliance (shown as PRE-HIRE, not counted in the rate).
    - Anything they pass counts toward their real learning paths on Day 1.
    - On StartDate they flip to the normal view automatically. Their new-hire
      assignment email goes out that day too.

  Set the flag from the app: Manage > Courses > edit a course > Pre-hire Access.

  Safe to re-run — it skips anything already present.

.NOTES
  Requires PnP.PowerShell:  Install-Module PnP.PowerShell -Scope CurrentUser
  Run with:  pwsh ./scripts/ensure-prehire-column.ps1
#>

param(
  [string]$SiteUrl     = "https://newshirepmcom.sharepoint.com/sites/NewShirePM",
  [string]$ClientId    = "7f310acf-12b1-4ba9-a113-c027614268b9",
  [string]$CoursesList = "TrainingCourses"
)

$ErrorActionPreference = "Stop"

Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

$lists = (Get-PnPList).Title
if ($lists -notcontains $CoursesList) {
  Write-Host "`n'$CoursesList' does not exist on this site. Re-run with the correct -SiteUrl." -ForegroundColor Red
  Disconnect-PnPOnline
  exit 1
}

Write-Host "`nEnsuring pre-hire column ..." -ForegroundColor Cyan

$existingField = Get-PnPField -List $CoursesList | Where-Object InternalName -eq "PrehireAvailable"
if ($existingField) {
  Write-Host "  [OK]    $CoursesList.PrehireAvailable already exists" -ForegroundColor Green
} else {
  Add-PnPFieldFromXml -List $CoursesList -FieldXml "<Field Type='Boolean' DisplayName='PrehireAvailable' Name='PrehireAvailable' StaticName='PrehireAvailable'><Default>0</Default></Field>" | Out-Null
  Write-Host "  [ADDED] $CoursesList.PrehireAvailable" -ForegroundColor Cyan
}

Write-Host "`nDone. Flag courses in the app: Manage > Courses > edit > Pre-hire Access." -ForegroundColor Cyan
Disconnect-PnPOnline
