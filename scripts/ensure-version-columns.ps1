<#
.SYNOPSIS
  One-time setup for NewShire University course content versioning.

  Adds the columns that let a course be "updated with re-training required":
    TrainingCourses:      Version (Number), ReqVersion (Number), VersionNote (text), VersionDate (DateTime)
    TrainingCompletions:  CompletedVersion (Number)

  How it works in the app:
    - Version      = the course's current content version (bumped on every published update).
    - ReqVersion   = the minimum completed version that still counts. Bumped to the new Version
                     only when an update "requires re-training" — that marks every older completion
                     stale, so the course shows as required/overdue again for those staff.
    - CompletedVersion = the Version stamped on a completion when the quiz is passed.

  Safe to re-run — it skips anything already present.

.NOTES
  Requires PnP.PowerShell:  Install-Module PnP.PowerShell -Scope CurrentUser
  Run with:  pwsh ./scripts/ensure-version-columns.ps1
#>

param(
  [string]$SiteUrl  = "https://newshirepmcom.sharepoint.com/sites/NewShirePM",
  # "NewShire Migration Tool" app in the current post-carve-out tenant. The previous
  # default, 32e75ffa-747a-4cf0-8209-6a19150c4547, lived in Vanrock's old tenant and
  # now fails sign-in with AADSTS700016.
  [string]$ClientId = "7f310acf-12b1-4ba9-a113-c027614268b9",
  [string]$CoursesList     = "TrainingCourses",
  [string]$CompletionsList = "TrainingCompletions"
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
foreach ($l in @($CoursesList, $CompletionsList)) {
  if ($lists -notcontains $l) {
    Write-Host "`n'$l' does not exist on this site. Re-run with the correct -SiteUrl." -ForegroundColor Red
    Disconnect-PnPOnline
    exit 1
  }
}

Write-Host "`nEnsuring versioning columns ..." -ForegroundColor Cyan

# TrainingCourses
Add-FieldIfMissing -List $CoursesList -InternalName "Version" `
  -Xml "<Field Type='Number' DisplayName='Version' Name='Version' StaticName='Version' Decimals='0'><Default>1</Default></Field>"
Add-FieldIfMissing -List $CoursesList -InternalName "ReqVersion" `
  -Xml "<Field Type='Number' DisplayName='ReqVersion' Name='ReqVersion' StaticName='ReqVersion' Decimals='0'><Default>1</Default></Field>"
Add-FieldIfMissing -List $CoursesList -InternalName "VersionNote" `
  -Xml "<Field Type='Note' DisplayName='VersionNote' Name='VersionNote' StaticName='VersionNote' RichText='FALSE' NumLines='3' />"
Add-FieldIfMissing -List $CoursesList -InternalName "VersionDate" `
  -Xml "<Field Type='DateTime' DisplayName='VersionDate' Name='VersionDate' StaticName='VersionDate' Format='DateTime' />"

# TrainingCompletions
Add-FieldIfMissing -List $CompletionsList -InternalName "CompletedVersion" `
  -Xml "<Field Type='Number' DisplayName='CompletedVersion' Name='CompletedVersion' StaticName='CompletedVersion' Decimals='0'><Default>1</Default></Field>"

Disconnect-PnPOnline
Write-Host "`nComplete. Existing courses/completions default to version 1, so nothing is marked stale until you publish a 'require re-training' update." -ForegroundColor Cyan
