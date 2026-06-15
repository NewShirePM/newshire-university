<#
.SYNOPSIS
  One-time setup for NewShire University written-lesson support.

  1. Signs in to the NewShire PM SharePoint site (interactive / browser).
  2. Reports which TrainingCourses / LearningPaths / TrainingLessons / TrainingQuizzes
     lists currently exist (answers "do the lists exist yet?").
  3. Adds the `LessonBody` column to TrainingLessons if it is missing, so written
     lesson content (HTML) saves from the app and the Import Course tool.

  Safe to re-run — it skips anything already present.

.NOTES
  Requires the PnP.PowerShell module:  Install-Module PnP.PowerShell -Scope CurrentUser
  Run with:  pwsh ./scripts/ensure-lessonbody-column.ps1
#>

param(
  [string]$SiteUrl  = "https://vanrockre.sharepoint.com/sites/NewshirePM",
  # Entra app registration used for PnP interactive sign-in (must allow public client flows
  # and have delegated SharePoint permissions). Override if you use a dedicated PnP app.
  [string]$ClientId = "32e75ffa-747a-4cf0-8209-6a19150c4547",
  [string]$LessonsList = "TrainingLessons"
)

$ErrorActionPreference = "Stop"

Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

# ---- 1. Inventory the training lists -------------------------------------------------
$expected = @("TrainingCourses","LearningPaths","TrainingLessons","TrainingQuizzes",
              "TrainingCompletions","TrainingEnrollments","TrainingAssignments")
Write-Host "`nTraining list inventory:" -ForegroundColor Cyan
$existing = (Get-PnPList).Title
foreach ($name in $expected) {
  if ($existing -contains $name) { Write-Host ("  [OK]      {0}" -f $name) -ForegroundColor Green }
  else                           { Write-Host ("  [MISSING] {0}" -f $name) -ForegroundColor Yellow }
}

# ---- 2. Ensure the LessonBody column exists ------------------------------------------
if ($existing -notcontains $LessonsList) {
  Write-Host "`n'$LessonsList' does not exist on this site. Cannot add LessonBody column." -ForegroundColor Red
  Write-Host "If the training lists live elsewhere, re-run with -SiteUrl <correct site>." -ForegroundColor Red
  Disconnect-PnPOnline
  exit 1
}

$field = Get-PnPField -List $LessonsList | Where-Object InternalName -eq "LessonBody"
if ($field) {
  Write-Host "`n'LessonBody' already exists on '$LessonsList' — nothing to do." -ForegroundColor Green
} else {
  Write-Host "`nAdding 'LessonBody' (multi-line plain text) to '$LessonsList' ..." -ForegroundColor Cyan
  # Plain multi-line text (RichText FALSE) so the app's HTML is stored verbatim and
  # not re-escaped/altered by SharePoint's rich-text editor.
  $xml = "<Field Type='Note' DisplayName='LessonBody' Name='LessonBody' " +
         "StaticName='LessonBody' RichText='FALSE' NumLines='20' " +
         "UnlimitedLengthInDocumentLibrary='TRUE' />"
  Add-PnPFieldFromXml -List $LessonsList -FieldXml $xml | Out-Null
  Write-Host "Done. Written lessons can now be saved." -ForegroundColor Green
}

Disconnect-PnPOnline
Write-Host "`nComplete." -ForegroundColor Cyan
