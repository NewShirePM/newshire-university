<#
.SYNOPSIS
  Provisions the AppFolio Academy external courses (APF 151-160) in NewShire University,
  plus the columns and the TrainingAcknowledgments list they depend on.

  Author: Brandy Turner, NewShire Property Management
  Source: AppFolio_Academy_Catalog_Map.xlsx (Include = Y rows, plus the two "New" standalone
          courses). Every URL below was verified to load on training.appfolio.com on 9/24/2026.

.DESCRIPTION
  These are How-to-Use-AppFolio courses. They do NOT replace NewShire SOP (OPS) or policy
  courses. There is no quiz. A learner opens each lesson on AppFolio Academy in a new tab,
  then signs a per-lesson acknowledgment in the app. Signing is disabled until the learner
  has opened the lesson. When every lesson is acknowledged the app writes a normal
  TrainingCompletions record, so paths, compliance reporting and the Monday report all work
  unchanged.

  What this script does (safe to re-run — everything is skip-if-present):
    1. TrainingCourses  + CourseType (Text), ExternalProvider (Text)
    2. TrainingLessons  + ExternalURL (Text)
    3. Creates TrainingAcknowledgments (the signed attestations), with versioning on and
       item-level write security (people can only edit items they created).
    4. Creates each course as CourseStatus = "Coming Soon" and CourseType = "External".
       A course whose CourseCode already exists is SKIPPED, lessons included.
    5. Attaches courses to learning paths per $PathMap below. A missing path is a warning,
       not an error. Use -SkipPaths to attach nothing and do it in the app instead.

  After it runs: review the courses in Admin > Courses, then switch each to Active. Going
  Active sends the usual go-live emails.

  ROLES: CourseRoles uses the canonical NS_JobRoles titles, not the catalog abbreviations.
  "Assistant Property Manager" is not a canonical title, so APM targeting maps to nothing
  and is omitted. "1099 Leasing" is deliberately left off every course: required training
  for 1099 contractors is a worker-classification risk. They can still take any course
  voluntarily from the Training Library.

.EXAMPLE
  pwsh ./scripts/provision-appfolio-academy-courses.ps1
  pwsh ./scripts/provision-appfolio-academy-courses.ps1 -SkipPaths
  pwsh ./scripts/provision-appfolio-academy-courses.ps1 -WhatIf     # report only, writes nothing

.NOTES
  Requires PnP.PowerShell:  Install-Module PnP.PowerShell -Scope CurrentUser
#>

[CmdletBinding()]
param(
  [string]$SiteUrl  = "https://newshirepmcom.sharepoint.com/sites/NewShirePM",
  # "NewShire Migration Tool" app in the current tenant (same as ensure-prereq-column.ps1).
  [string]$ClientId = "7f310acf-12b1-4ba9-a113-c027614268b9",
  [switch]$SkipPaths,
  [switch]$WhatIf
)

$ErrorActionPreference = "Stop"
$Provider = "AppFolio Academy"

# ── Learning path attachment ─────────────────────────────────────────────────
# Path name (exact Title in LearningPaths) -> course codes. Course-level role targeting
# still decides who inside the path sees each course. APF 351, 353 and 160 are electives
# and attach to nothing. Edit freely before running.
$PathMap = [ordered]@{
  "New Hire Onboarding"    = @("APF 151")
  "Leasing Professional"   = @("APF 154", "APF 155", "APF 156", "APF 157", "APF 159")
  "Maintenance Technician" = @("APF 158")
  "Maintenance Supervisor" = @("APF 157", "APF 158")
  "Property Manager"       = @("APF 152", "APF 154", "APF 155", "APF 156", "APF 157", "APF 158", "APF 159", "APF 255")
}

$Courses = @(
  @{ Code='APF 151'; Title='Getting Started in AppFolio'; Sort=1051; RecertDays=0; DurationMin=32
     Roles='Leasing Agent, Maintenance Technician, Maintenance Supervisor, Property Manager, Regional/Portfolio Manager, Operations Manager, Director of Operations, Delinquency and Collections Manager, Executive Assistant, Virtual Assistant'
     Description='AppFolio Academy''s Getting Started path: navigating AppFolio Property Manager, general and personal settings, user roles and permissions, and AppFolio help resources.'
     Lessons=@(
      @{ Order=1; Title='Navigating AppFolio Property Manager'; Min=6; Url='https://training.appfolio.com/path/getting-started-in-appfolio/navigating-appfolio-property-manager-1' },
      @{ Order=2; Title='General Settings'; Min=6; Url='https://training.appfolio.com/path/getting-started-in-appfolio/general-settings' },
      @{ Order=3; Title='Configuring Personal Settings'; Min=6; Url='https://training.appfolio.com/path/getting-started-in-appfolio/configuring-personal-settings' },
      @{ Order=4; Title='User Roles & Permissions'; Min=5; Url='https://training.appfolio.com/path/getting-started-in-appfolio/user-roles-permissions' },
      @{ Order=5; Title='AppFolio Help Resources'; Min=6; Url='https://training.appfolio.com/path/getting-started-in-appfolio/appfolio-help-resources-1' },
      @{ Order=6; Title='Knowledge Check: Getting Started in AppFolio'; Min=3; Url='https://training.appfolio.com/path/getting-started-in-appfolio/knowledge-check-getting-started-in-appfolio' }
     ) },
  @{ Code='APF 152'; Title='Accounting: Receivables'; Sort=1052; RecertDays=0; DurationMin=30
     Roles='Property Manager, Regional/Portfolio Manager, Operations Manager, Delinquency and Collections Manager, Virtual Assistant'
     Description='AppFolio Academy''s Accounting: Receivables path: resident charges, receipts, bank deposits, resident credits, and NSFs.'
     Lessons=@(
      @{ Order=1; Title='Resident Charges'; Min=5; Url='https://training.appfolio.com/path/accounting-receivables/resident-charges' },
      @{ Order=2; Title='Receipts'; Min=8; Url='https://training.appfolio.com/path/accounting-receivables/receipts-course' },
      @{ Order=3; Title='Bank Deposits'; Min=5; Url='https://training.appfolio.com/path/accounting-receivables/bank-deposits' },
      @{ Order=4; Title='Resident Credits'; Min=5; Url='https://training.appfolio.com/path/accounting-receivables/resident-credits' },
      @{ Order=5; Title='Process NSF''s'; Min=4; Url='https://training.appfolio.com/path/accounting-receivables/process-nsfs' },
      @{ Order=6; Title='Knowledge Check: Accounting: Receivables'; Min=3; Url='https://training.appfolio.com/path/accounting-receivables/knowledge-check-accounting-receivables' }
     ) },
  @{ Code='APF 154'; Title='Rental Applications & Leases'; Sort=1054; RecertDays=0; DurationMin=28
     Roles='Leasing Agent, Property Manager, Regional/Portfolio Manager, Operations Manager, Virtual Assistant'
     Description='AppFolio Academy''s Rental Applications & Leases path: configuring, receiving, and converting rental applications, and online leases. AppFolio mechanics only; NewShire screening criteria are covered in NewShire policy and SOP courses.'
     Lessons=@(
      @{ Order=1; Title='Configuring Online Rental Applications'; Min=7; Url='https://training.appfolio.com/path/rental-applications-leases/configuring-rental-applications' },
      @{ Order=2; Title='Receiving Rental Applications'; Min=5; Url='https://training.appfolio.com/path/rental-applications-leases/receiving-rental-applications' },
      @{ Order=3; Title='Create and Configure Online Leases'; Min=5; Url='https://training.appfolio.com/path/rental-applications-leases/create-and-configure-online-leases' },
      @{ Order=4; Title='Converting Rental Applications'; Min=8; Url='https://training.appfolio.com/path/rental-applications-leases/converting-rental-applications' },
      @{ Order=5; Title='Knowledge Check: Rental Applications & Leases'; Min=3; Url='https://training.appfolio.com/path/rental-applications-leases/knowledge-check-rental-applications-leases' }
     ) },
  @{ Code='APF 155'; Title='Marketing and Lead Management'; Sort=1055; RecertDays=0; DurationMin=22
     Roles='Leasing Agent, Property Manager, Regional/Portfolio Manager, Operations Manager, Virtual Assistant'
     Description='AppFolio Academy''s Marketing and Lead Management path: marketing vacancies, guest cards, prospect communication, and showings.'
     Lessons=@(
      @{ Order=1; Title='Marketing Vacancies'; Min=6; Url='https://training.appfolio.com/path/marketing-and-lead-management/marketing-vacancies-1' },
      @{ Order=2; Title='Guest Card Submission'; Min=3; Url='https://training.appfolio.com/path/marketing-and-lead-management/guest-card-submission' },
      @{ Order=3; Title='Guest Card Communication and Activities'; Min=6; Url='https://training.appfolio.com/path/marketing-and-lead-management/guest-card-communication-and-activities' },
      @{ Order=4; Title='Showings'; Min=4; Url='https://training.appfolio.com/path/marketing-and-lead-management/showings-course' },
      @{ Order=5; Title='Knowledge Check: Marketing and Lead Management'; Min=3; Url='https://training.appfolio.com/path/marketing-and-lead-management/knowledge-check-marketing-and-lead-management' }
     ) },
  @{ Code='APF 156'; Title='Resident Lifecycle'; Sort=1056; RecertDays=0; DurationMin=65
     Roles='Leasing Agent, Property Manager, Regional/Portfolio Manager, Operations Manager, Virtual Assistant'
     Description='AppFolio Academy''s Resident Lifecycle path: move-ins, the resident online portal, and move-outs. Move-out policy and South Carolina deposit rules are covered in OPS 301.'
     Lessons=@(
      @{ Order=1; Title='Resident Move Ins'; Min=14; Url='https://training.appfolio.com/path/resident-lifecycle/resident-move-ins' },
      @{ Order=2; Title='Resident Online Portal'; Min=30; Url='https://training.appfolio.com/path/resident-lifecycle/resident-online-portal' },
      @{ Order=3; Title='Resident Move Outs'; Min=18; Url='https://training.appfolio.com/path/resident-lifecycle/resident-move-outs' },
      @{ Order=4; Title='Knowledge Check: Resident Lifecycle'; Min=3; Url='https://training.appfolio.com/path/resident-lifecycle/knowledge-check-resident-lifecycle' }
     ) },
  @{ Code='APF 157'; Title='Communications'; Sort=1057; RecertDays=0; DurationMin=15
     Roles='Leasing Agent, Maintenance Supervisor, Property Manager, Regional/Portfolio Manager, Operations Manager, Delinquency and Collections Manager, Virtual Assistant'
     Description='AppFolio Academy''s Communications path: texting, emailing, and letters in AppFolio.'
     Lessons=@(
      @{ Order=1; Title='Texting in AppFolio'; Min=4; Url='https://training.appfolio.com/path/communications/texting-in-appfolio' },
      @{ Order=2; Title='Emailing in AppFolio'; Min=4; Url='https://training.appfolio.com/path/communications/emailing-in-appfolio' },
      @{ Order=3; Title='Letters in AppFolio'; Min=4; Url='https://training.appfolio.com/path/communications/letters-in-appfolio' },
      @{ Order=4; Title='Knowledge Check: Communications'; Min=3; Url='https://training.appfolio.com/path/communications/knowledge-check-communications' }
     ) },
  @{ Code='APF 158'; Title='Maintenance Basics'; Sort=1058; RecertDays=0; DurationMin=37
     Roles='Maintenance Technician, Maintenance Supervisor, Property Manager, Regional/Portfolio Manager, Operations Manager, Virtual Assistant'
     Description='AppFolio Academy''s Maintenance Basics path: service requests and work orders, estimates, billing and closing out work orders, the maintenance tech mobile view, the vendor portal, and inspections.'
     Lessons=@(
      @{ Order=1; Title='Service Requests and Work Orders'; Min=7; Url='https://training.appfolio.com/path/maintenance-basics/service-requests-and-work-orders-1' },
      @{ Order=2; Title='Managing Work Orders'; Min=4; Url='https://training.appfolio.com/path/maintenance-basics/managing-work-orders' },
      @{ Order=3; Title='Work Order Estimates'; Min=5; Url='https://training.appfolio.com/path/maintenance-basics/work-order-estimates' },
      @{ Order=4; Title='Billing & Closing Out a Work Order'; Min=7; Url='https://training.appfolio.com/path/maintenance-basics/billing-closing-out-a-work-order' },
      @{ Order=5; Title='Maintenance Tech Mobile View'; Min=3; Url='https://training.appfolio.com/path/maintenance-basics/maintenance-tech-mobile-view' },
      @{ Order=6; Title='Vendor Portal'; Min=4; Url='https://training.appfolio.com/path/maintenance-basics/vendor-portal' },
      @{ Order=7; Title='Inspections'; Min=4; Url='https://training.appfolio.com/path/maintenance-basics/inspections-course-1' },
      @{ Order=8; Title='Knowledge Check: Maintenance Basics'; Min=3; Url='https://training.appfolio.com/path/maintenance-basics/knowledge-check-maintenance-basics' }
     ) },
  @{ Code='APF 159'; Title='Screening Applicants'; Sort=1059; RecertDays=0; DurationMin=11
     Roles='Leasing Agent, Property Manager, Regional/Portfolio Manager, Operations Manager, Virtual Assistant'
     Description='AppFolio Academy course on using AppFolio''s residential screening service. AppFolio mechanics only; NewShire screening criteria are covered in NewShire policy.'
     Lessons=@(
      @{ Order=1; Title='Screening Applicants'; Min=11; Url='https://training.appfolio.com/screening-applicants' }
     ) },
  @{ Code='APF 160'; Title='Marketing Your Business to Attract Leads & Owners'; Sort=1060; RecertDays=0; DurationMin=0
     Roles='Property Manager, Regional/Portfolio Manager, Operations Manager, Director of Operations'
     Description='AppFolio Academy recorded session (March 2022) on marketing a property management business to attract more leads and owners.'
     Lessons=@(
      @{ Order=1; Title='How to Effectively Market Your Business to Attract More Leads & Owners'; Min=$null; Url='https://training.appfolio.com/how-to-effectively-market-your-business-to-attract-more-leads-owners' }
     ) },
  @{ Code='APF 255'; Title='Affordable Housing: Housing Choice Voucher'; Sort=1155; RecertDays=0; DurationMin=0
     Roles='Property Manager, Regional/Portfolio Manager, Operations Manager, Delinquency and Collections Manager, Virtual Assistant'
     Description='AppFolio Academy lessons on managing subsidy programs and entering subsidized rent receipts in bulk.'
     Lessons=@(
      @{ Order=1; Title='Manage Subsidy Programs'; Min=$null; Url='https://training.appfolio.com/manage-subsidy-programs' },
      @{ Order=2; Title='Enter Subsidized Rent Receipts in Bulk'; Min=$null; Url='https://training.appfolio.com/enter-subsidized-rent-receipts-in-bulk' }
     ) },
  @{ Code='APF 351'; Title='Accounting Transactions Certification'; Sort=1251; RecertDays=0; DurationMin=0
     Roles='Property Manager, Regional/Portfolio Manager, Operations Manager'
     Description='AppFolio Academy''s Accounting Transactions Certification: accounting tools, receivables, payables, management fees, and paying owners, ending in the certification exam.'
     Lessons=@(
      @{ Order=1; Title='Academy Certifications Overview: Monthly Transactions'; Min=$null; Url='https://training.appfolio.com/path/accounting-transactions-certification/transactions-certification-overview' },
      @{ Order=2; Title='Accounting Tools'; Min=$null; Url='https://training.appfolio.com/path/accounting-transactions-certification/accounting-tools-course' },
      @{ Order=3; Title='Accounts Receivable Entry'; Min=$null; Url='https://training.appfolio.com/path/accounting-transactions-certification/accounts-receivable-entry' },
      @{ Order=4; Title='Accounts Receivable Tasks'; Min=$null; Url='https://training.appfolio.com/path/accounting-transactions-certification/accounts-receivable-tasks' },
      @{ Order=5; Title='Accounts Payable'; Min=$null; Url='https://training.appfolio.com/path/accounting-transactions-certification/accounts-payable' },
      @{ Order=6; Title='Management Fees & Additional Fees'; Min=$null; Url='https://training.appfolio.com/path/accounting-transactions-certification/management-fees-additional-fees' },
      @{ Order=7; Title='Paying Owners'; Min=$null; Url='https://training.appfolio.com/path/accounting-transactions-certification/paying-owners' },
      @{ Order=8; Title='Transactions Certification Exam'; Min=$null; Url='https://training.appfolio.com/path/accounting-transactions-certification/transactions-certification-exam' }
     ) },
  @{ Code='APF 353'; Title='Lead to Lease Certification'; Sort=1253; RecertDays=0; DurationMin=0
     Roles='Leasing Agent, Property Manager'
     Description='AppFolio Academy''s Lead to Lease Certification: marketing units, leads and showings, rental applications, leases and renewals, and the move-in flow, ending in the final exam.'
     Lessons=@(
      @{ Order=1; Title='Lead to Lease Certification Overview'; Min=$null; Url='https://training.appfolio.com/path/leasing-lead-to-lease/lead-to-lease-certification-overview' },
      @{ Order=2; Title='Lead to Lease: Marketing your Units'; Min=$null; Url='https://training.appfolio.com/path/leasing-lead-to-lease/lead-to-lease-marketing-your-units' },
      @{ Order=3; Title='Lead to Lease: Leads & Scheduling Showings'; Min=$null; Url='https://training.appfolio.com/path/leasing-lead-to-lease/lead-to-lease-leads-scheduling-showings' },
      @{ Order=4; Title='Lead to Lease: Rental Applications'; Min=$null; Url='https://training.appfolio.com/path/leasing-lead-to-lease/lead-to-lease-rental-applications' },
      @{ Order=5; Title='Lead to Lease: Leases, Addenda, & Renewals'; Min=$null; Url='https://training.appfolio.com/path/leasing-lead-to-lease/lead-to-lease-leases-addenda-renewals' },
      @{ Order=6; Title='Lead to Lease: The Move In Flow'; Min=$null; Url='https://training.appfolio.com/path/leasing-lead-to-lease/lead-to-lease-the-move-in-flow' },
      @{ Order=7; Title='Lead to Lease Certification Final Exam'; Min=$null; Url='https://training.appfolio.com/path/leasing-lead-to-lease/lead-to-lease-certification-final-exam' }
     ) }
)

# ── Sanity checks before touching SharePoint ────────────────────────────────
$codes = $Courses | ForEach-Object { $_.Code }
if (($codes | Select-Object -Unique).Count -ne $codes.Count) { throw "Duplicate course code in `$Courses." }
foreach ($c in $Courses) {
  foreach ($l in $c.Lessons) {
    if ($l.Url -notmatch '^https://training\.appfolio\.com/') { throw "$($c.Code) lesson '$($l.Title)' has a non-AppFolio URL: $($l.Url)" }
  }
}
foreach ($p in $PathMap.Keys) { foreach ($code in $PathMap[$p]) { if ($codes -notcontains $code) { throw "PathMap '$p' names $code, which is not in `$Courses." } } }
Write-Host ("Loaded {0} courses / {1} lessons." -f $Courses.Count, ($Courses | ForEach-Object { $_.Lessons.Count } | Measure-Object -Sum).Sum) -ForegroundColor Cyan
if ($WhatIf) { Write-Host "-WhatIf: nothing will be written." -ForegroundColor Yellow }

Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

$lists = (Get-PnPList).Title
foreach ($required in "TrainingCourses", "TrainingLessons", "LearningPaths") {
  if ($lists -notcontains $required) { Write-Host "'$required' does not exist on this site. Re-run with the correct -SiteUrl." -ForegroundColor Red; exit 1 }
}

function Add-FieldIfMissing {
  param([string]$List, [string]$InternalName, [string]$Xml)
  $existing = Get-PnPField -List $List | Where-Object InternalName -eq $InternalName
  if ($existing) { Write-Host ("  [OK]    {0}.{1}" -f $List, $InternalName) -ForegroundColor Green; return }
  if ($WhatIf) { Write-Host ("  [WOULD ADD] {0}.{1}" -f $List, $InternalName) -ForegroundColor Yellow; return }
  Add-PnPFieldFromXml -List $List -FieldXml $Xml | Out-Null
  Write-Host ("  [ADDED] {0}.{1}" -f $List, $InternalName) -ForegroundColor Cyan
}
function TextField([string]$n, [int]$max = 255) { "<Field Type='Text' DisplayName='$n' Name='$n' StaticName='$n' MaxLength='$max' />" }

# ── 1-2. Columns ─────────────────────────────────────────────────────────────
Write-Host "`nColumns ..." -ForegroundColor Cyan
Add-FieldIfMissing -List "TrainingCourses" -InternalName "CourseType"       -Xml (TextField "CourseType" 32)
Add-FieldIfMissing -List "TrainingCourses" -InternalName "ExternalProvider" -Xml (TextField "ExternalProvider" 64)
Add-FieldIfMissing -List "TrainingLessons" -InternalName "ExternalURL"      -Xml (TextField "ExternalURL" 255)

# ── 3. TrainingAcknowledgments ───────────────────────────────────────────────
$AckList = "TrainingAcknowledgments"
Write-Host "`n$AckList ..." -ForegroundColor Cyan
if ($lists -notcontains $AckList) {
  if ($WhatIf) { Write-Host "  [WOULD CREATE] $AckList" -ForegroundColor Yellow }
  else { New-PnPList -Title $AckList -Template GenericList -OnQuickLaunch:$false | Out-Null; Write-Host "  [CREATED] $AckList" -ForegroundColor Cyan }
} else { Write-Host "  [OK]    $AckList exists" -ForegroundColor Green }
if (-not $WhatIf -or $lists -contains $AckList) {
  Add-FieldIfMissing -List $AckList -InternalName "AckEmployeeEmail" -Xml (TextField "AckEmployeeEmail")
  Add-FieldIfMissing -List $AckList -InternalName "AckEmployeeName"  -Xml (TextField "AckEmployeeName")
  Add-FieldIfMissing -List $AckList -InternalName "AckCourseID"      -Xml "<Field Type='Number' DisplayName='AckCourseID' Name='AckCourseID' StaticName='AckCourseID' Decimals='0' />"
  Add-FieldIfMissing -List $AckList -InternalName "AckCourseCode"    -Xml (TextField "AckCourseCode" 32)
  Add-FieldIfMissing -List $AckList -InternalName "AckCourseTitle"   -Xml (TextField "AckCourseTitle")
  Add-FieldIfMissing -List $AckList -InternalName "AckLessonID"      -Xml "<Field Type='Number' DisplayName='AckLessonID' Name='AckLessonID' StaticName='AckLessonID' Decimals='0' />"
  Add-FieldIfMissing -List $AckList -InternalName "AckLessonTitle"   -Xml (TextField "AckLessonTitle")
  Add-FieldIfMissing -List $AckList -InternalName "AckExternalURL"   -Xml (TextField "AckExternalURL")
  Add-FieldIfMissing -List $AckList -InternalName "AckVersion"       -Xml (TextField "AckVersion" 32)
  Add-FieldIfMissing -List $AckList -InternalName "AckText"          -Xml "<Field Type='Note' DisplayName='AckText' Name='AckText' StaticName='AckText' NumLines='8' RichText='FALSE' />"
  Add-FieldIfMissing -List $AckList -InternalName "LaunchedAt"       -Xml "<Field Type='DateTime' DisplayName='LaunchedAt' Name='LaunchedAt' StaticName='LaunchedAt' Format='DateTime' />"
  Add-FieldIfMissing -List $AckList -InternalName "AcknowledgedAt"   -Xml "<Field Type='DateTime' DisplayName='AcknowledgedAt' Name='AcknowledgedAt' StaticName='AcknowledgedAt' Format='DateTime' />"
  if (-not $WhatIf) {
    # Evidence hygiene: keep every version of every record, and let people edit only what
    # they created (WriteSecurity 2). Reading stays open (ReadSecurity 1) so managers and
    # the app can see everyone's acknowledgments.
    Set-PnPList -Identity $AckList -EnableVersioning $true -MajorVersions 500 | Out-Null
    $l = Get-PnPList -Identity $AckList
    $l.ReadSecurity = 1; $l.WriteSecurity = 2; $l.Update(); Invoke-PnPQuery
    Write-Host "  [OK]    versioning on, item-level write security set" -ForegroundColor Green
  }
}

# ── Lesson lookup column: resolve the real internal name instead of guessing ──
$lookup = Get-PnPField -List "TrainingLessons" | Where-Object { $_.TypeAsString -eq "Lookup" -and $_.InternalName -like "CourseID*" } | Select-Object -First 1
if (-not $lookup) { $lookup = Get-PnPField -List "TrainingLessons" | Where-Object { $_.TypeAsString -eq "Lookup" } | Select-Object -First 1 }
if (-not $lookup) { throw "No lookup column to TrainingCourses found on TrainingLessons." }
$LessonCourseField = $lookup.InternalName
Write-Host "`nTrainingLessons course lookup column: $LessonCourseField" -ForegroundColor Cyan

# ── 4. Courses + lessons ─────────────────────────────────────────────────────
Write-Host "`nCourses ..." -ForegroundColor Cyan
$existingCourses = Get-PnPListItem -List "TrainingCourses" -PageSize 500
$courseIdByCode = @{}
foreach ($e in $existingCourses) { $cc = $e.FieldValues["CourseCode"]; if ($cc) { $courseIdByCode[$cc.Trim()] = $e.Id } }

$created = 0; $skipped = 0
foreach ($c in $Courses) {
  if ($courseIdByCode.ContainsKey($c.Code)) {
    Write-Host ("  [SKIP]  {0} already exists (item {1}) — not touched" -f $c.Code, $courseIdByCode[$c.Code]) -ForegroundColor DarkYellow
    $skipped++; continue
  }
  if ($WhatIf) { Write-Host ("  [WOULD CREATE] {0} — {1} ({2} lessons)" -f $c.Code, $c.Title, $c.Lessons.Count) -ForegroundColor Yellow; continue }
  $item = Add-PnPListItem -List "TrainingCourses" -Values @{
    "Title"             = $c.Title
    "CourseCode"        = $c.Code
    "CourseDescription" = $c.Description
    "Category"          = "Systems"
    "DurationMin"       = $c.DurationMin
    "RecertDays"        = $c.RecertDays
    "SortOrder"         = $c.Sort
    "CourseActive"      = $true
    "CourseStatus"      = "Coming Soon"
    "CourseRoles"       = $c.Roles
    "CourseType"        = "External"
    "ExternalProvider"  = $Provider
  }
  $courseIdByCode[$c.Code] = $item.Id
  foreach ($l in $c.Lessons) {
    $v = @{
      "Title"           = $l.Title
      $LessonCourseField = $item.Id
      "LessonSortOrder" = $l.Order
      "ExternalURL"     = $l.Url
    }
    if ($null -ne $l.Min) { $v["LessonDurationMin"] = $l.Min }
    Add-PnPListItem -List "TrainingLessons" -Values $v | Out-Null
  }
  Write-Host ("  [ADDED] {0} — {1} (item {2}, {3} lessons)" -f $c.Code, $c.Title, $item.Id, $c.Lessons.Count) -ForegroundColor Cyan
  $created++
}

# ── 5. Learning paths ────────────────────────────────────────────────────────
if ($SkipPaths) {
  Write-Host "`n-SkipPaths: no path changes. Attach courses in the app (Admin > Paths)." -ForegroundColor Yellow
} else {
  Write-Host "`nLearning paths ..." -ForegroundColor Cyan
  foreach ($pathName in $PathMap.Keys) {
    $path = Get-PnPListItem -List "LearningPaths" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$pathName</Value></Eq></Where></Query></View>"
    if (-not $path) { Write-Host "  [WARN]  Path '$pathName' not found — skipped. Attach in the app or fix `$PathMap." -ForegroundColor Red; continue }
    $ids = @(); if ($path.FieldValues["CourseIDs"]) { $ids = $path.FieldValues["CourseIDs"].Split(",") | ForEach-Object { $_.Trim() } | Where-Object { $_ } }
    $added = @()
    foreach ($code in $PathMap[$pathName]) {
      if (-not $courseIdByCode.ContainsKey($code)) { continue }   # only happens under -WhatIf
      $id = [string]$courseIdByCode[$code]
      if ($ids -notcontains $id) { $ids += $id; $added += $code }
    }
    if ($added.Count -eq 0) { Write-Host "  [OK]    $pathName — nothing new to attach" -ForegroundColor Green; continue }
    if ($WhatIf) { Write-Host "  [WOULD ATTACH] $pathName <- $($added -join ', ')" -ForegroundColor Yellow; continue }
    Set-PnPListItem -List "LearningPaths" -Identity $path.Id -Values @{ "CourseIDs" = ($ids -join ",") } | Out-Null
    Write-Host "  [ATTACHED] $pathName <- $($added -join ', ')" -ForegroundColor Cyan
  }
}

Write-Host ("`nDone. Created {0}, skipped {1}. New courses are 'Coming Soon' — review, then set Active in the app." -f $created, $skipped) -ForegroundColor Cyan
Disconnect-PnPOnline
