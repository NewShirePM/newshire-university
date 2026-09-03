<#
.SYNOPSIS
  Provisions the final FHC 101 course quiz (40 questions) into SharePoint.

  Replaces every existing TrainingQuizzes item tied to FHC 101 with the finalized
  set below, so the SharePoint list matches FHC-101-Quiz-Final-40.md exactly. That
  file is the source of truth for wording -- this script is just the delivery
  mechanism. If you ever hand-edit a question, edit the markdown first, then
  re-run this.

  Site: https://newshirepmcom.sharepoint.com/sites/NewShirePM
    (NOT vanrockre.sharepoint.com -- that was the pre-migration tenant name and
    is stale in the older prerequisite-column script, which has been patched.)

  What it does, in order:
    1. Connects interactively (your own login, your own MFA).
    2. Looks up the FHC 101 course record in TrainingCourses by CourseCode, to
       get the item ID the quiz questions link to.
    3. Pre-flight: confirms TrainingQuizzes actually has every field this script
       writes, and stops before changing anything if one is missing. This runs
       under -DryRun too, so a dry run really does exercise the field names.
    4. Reports how many TrainingQuizzes items currently exist for that course.
    5. Unless -DryRun is passed, recycles those existing items and adds the 40
       below in their place, in order. Recycled items are recoverable from the
       site recycle bin, so a failed run mid-flight is undoable.
    6. Verifies the final count is exactly 40 and reports any mismatch.

  Run once with -DryRun first if you want to see the before-state without
  changing anything:
    pwsh ./provision-fhc101-quiz.ps1 -DryRun

  Then for real:
    pwsh ./provision-fhc101-quiz.ps1

.NOTES
  Requires PnP.PowerShell:  Install-Module PnP.PowerShell -Scope CurrentUser

  A note on history: deleting and re-adding quiz items changes their SharePoint
  item IDs. Any employee's PAST quiz attempt already has its own snapshot of
  what was asked (question text, options, and their answer) saved on the
  completion record itself -- that snapshot doesn't reference these items live,
  so old completion history is unaffected either way. Nothing about past scores
  or pass/fail status changes.
#>

param(
  [string]$SiteUrl     = "https://newshirepmcom.sharepoint.com/sites/NewShirePM",
  # "NewShire Migration Tool" app, registered in the CURRENT post-carve-out tenant
  # (5e932bcd-838c-4dae-b838-f6d22d7c6b8a). The previous default here was
  # 32e75ffa-747a-4cf0-8209-6a19150c4547, an app registration that lived in
  # Vanrock's old tenant -- it now fails with AADSTS700016. The July 2026
  # migration repointed the front-end SPA config but missed these PnP scripts.
  [string]$ClientId    = "7f310acf-12b1-4ba9-a113-c027614268b9",
  [string]$CoursesList = "TrainingCourses",
  [string]$QuizList    = "TrainingQuizzes",
  [string]$CourseCode  = "FHC 101",
  [switch]$DryRun
)

$ErrorActionPreference = "Stop"

# -- The 40 questions --------------------------------------------------------
# Order here is display order (QuizSortOrder). Grouped by module in comments
# only -- the list itself is flat, matching how the app reads it.

$Questions = @(

# -- Module 1 -- Fair Housing Act Foundations --------------------------------
@{ Title="The Fair Housing Act was originally signed into law in which year?"
   OptionA="1964"; OptionB="1968"; OptionC="1988"; OptionD="1974"; CorrectAnswer="B" }

@{ Title="How many federal protected classes does the Fair Housing Act establish?"
   OptionA="Five"; OptionB="Six"; OptionC="Seven"; OptionD="Nine"; CorrectAnswer="C" }

@{ Title="The 1988 amendment to the Fair Housing Act added which two protected classes?"
   OptionA="Sex and religion"; OptionB="Familial status and disability"
   OptionC="Race and color"; OptionD="National origin and religion"; CorrectAnswer="B" }

@{ Title="What is the difference between disparate treatment and disparate impact?"
   OptionA="Disparate treatment requires proof of a bad motive; disparate impact never does"
   OptionB="Disparate treatment is one person treated differently because of a protected characteristic - no bad motive required; disparate impact is one neutral rule, applied evenly, that lands unevenly across a protected class"
   OptionC="Disparate treatment applies to housing; disparate impact applies to employment"
   OptionD="There is no legal difference between the two"; CorrectAnswer="B" }

@{ Title="Two questions come up often enough that staff should have the answer memorized. What is the correct response to 'What kinds of people live here?'"
   OptionA="It's a great mix of families and young professionals"
   OptionB="We rent to anyone who qualifies"
   OptionC="I'd have to check who's currently renting"
   OptionD="This one leans more toward retirees"; CorrectAnswer="B" }

@{ Title="A small owner-occupied building of four units or fewer can, in narrow circumstances, qualify for a Fair Housing Act exemption. How many units does it apply to?"
   OptionA="Two"; OptionB="Four"; OptionC="Six"; OptionD="Eight"; CorrectAnswer="B" }

@{ Title="Do any Fair Housing Act exemptions apply to NewShire?"
   OptionA="Yes - the small-building exemption applies to our single-family homes"
   OptionB="Yes - a religious organization exemption applies"
   OptionC="No - the exemption collapses the moment an agent is involved, and NewShire is always the agent"
   OptionD="It depends on the specific property and owner"; CorrectAnswer="C" }

@{ Title="A property owner instructs you to tell families a unit is already taken because they're worried about children on the stairs. What should you do?"
   OptionA="Follow the instruction since it's the owner's building"
   OptionB="Decline it, explain why in one sentence, and get the Property Manager on it the same day"
   OptionC="Comply, but document that the owner gave the instruction"
   OptionD="Tell the owner to make the call themselves"; CorrectAnswer="B" }

@{ Title="What is South Carolina's state fair housing enforcement agency?"
   OptionA="SC Department of Consumer Affairs"; OptionB="SC Human Affairs Commission"
   OptionC="SC Real Estate Commission"; OptionD="SC Housing Authority"; CorrectAnswer="B" }

@{ Title="Which of the following is NOT one of the specific practices the Fair Housing Act prohibits?"
   OptionA="Advertising that states a preference or limitation"
   OptionB="Offering different terms, conditions, or privileges to different residents"
   OptionC="Requiring the same application documents from every applicant, applied consistently"
   OptionD="Falsely representing that a unit is unavailable"; CorrectAnswer="C" }

# -- Module 2 -- Protected Classes in Depth ----------------------------------
@{ Title="Which of the following is a standalone fair housing violation, distinct from race?"
   OptionA="Discrimination based on color"; OptionB="Discrimination based on income level"
   OptionC="Discrimination based on criminal history"; OptionD="Discrimination based on credit score"; CorrectAnswer="A" }

@{ Title="National origin protections extend to which of the following?"
   OptionA="Only country of birth as listed on identification documents"
   OptionB="Birthplace, ancestry, culture, language, and accent"
   OptionC="Only persons who are not U.S. citizens"
   OptionD="Country of birth only if the person is a legal permanent resident"; CorrectAnswer="B" }

@{ Title="Suggesting a prospect 'might be more comfortable' at a different property based on who lives in the neighborhood is an example of:"
   OptionA="Reasonable customer service"; OptionB="Disparate impact"
   OptionC="Steering"; OptionD="Blockbusting"; CorrectAnswer="C" }

@{ Title="Fair housing protections on the basis of religion cover:"
   OptionA="Only religions recognized by the IRS as tax-exempt"
   OptionB="Belief, practice, and the absence of belief"
   OptionC="Only religions practiced by a majority of local residents"
   OptionD="Religious beliefs but not religious practices like dress or dietary needs"; CorrectAnswer="B" }

@{ Title="The Supreme Court's 2020 decision in Bostock v. Clayton County held that:"
   OptionA="The Fair Housing Act explicitly lists gender identity as a protected class"
   OptionB="Discrimination based on sexual orientation or gender identity is inherently discrimination because of sex"
   OptionC="States may choose whether to apply sex-based protections to LGBTQ individuals"
   OptionD="Federal housing protections for LGBTQ individuals were eliminated"; CorrectAnswer="B" }

@{ Title="HUD's Keating guidance on a two-person-per-bedroom occupancy standard is best described as:"
   OptionA="A strict federal rule that cannot be exceeded under any circumstances"
   OptionB="A safe harbor that protects landlords from all familial status claims"
   OptionC="A presumptively reasonable starting point that can be rebutted by specific circumstances"
   OptionD="A guideline that applies only to public housing and Section 8 properties"; CorrectAnswer="C" }

@{ Title="Which factor does HUD NOT consider when evaluating whether an occupancy standard is reasonable?"
   OptionA="Size of bedrooms and overall unit square footage"
   OptionB="The race or national origin of the applicant family"
   OptionC="Age of children in the household"
   OptionD="Configuration of the unit, including dens or bonus rooms"; CorrectAnswer="B" }

@{ Title="Under disability protections, being 'regarded as' having a disability means:"
   OptionA="The person must provide documentation of the perceived condition"
   OptionB="If you treat someone as though they have a disability - even if they do not - that is discrimination"
   OptionC="The person believes they have a disability but has not been diagnosed"
   OptionD="A landlord may ask if someone appears to have a disability"; CorrectAnswer="B" }

@{ Title="Which protected class is the only one that creates affirmative obligations for housing providers?"
   OptionA="Familial status"; OptionB="National origin"; OptionC="Religion"; OptionD="Disability"; CorrectAnswer="D" }

@{ Title="A property manager tells a family with children, 'This is a quiet community - mostly professionals.' This statement is an example of:"
   OptionA="Accurate marketing of the property"
   OptionB="Steering or discouraging a household based on familial status"
   OptionC="A reasonable description of the tenant demographic"
   OptionD="An occupancy standard disclosure"; CorrectAnswer="B" }

# -- Module 3 -- Reasonable Accommodations and Modifications ----------------
@{ Title="A reasonable accommodation is a change to:"
   OptionA="The physical structure of a unit or common area"
   OptionB="A rule, policy, practice, or service"
   OptionC="The tenant's lease terms only"
   OptionD="Federal or state fair housing regulations"; CorrectAnswer="B" }

@{ Title="Who pays for a reasonable modification, and does that ever change?"
   OptionA="The resident always pays, regardless of property type"
   OptionB="The resident pays on a conventional property; NewShire pays on a Section 504 property - and NewShire does not currently manage any Section 504 properties"
   OptionC="NewShire always pays for modifications"
   OptionD="The cost is split evenly between resident and landlord"; CorrectAnswer="B" }

@{ Title="Which of the following is true about a reasonable accommodation request?"
   OptionA="The resident must use the words 'reasonable accommodation' for it to count"
   OptionB="It must be submitted on NewShire's official form before staff are required to act on it"
   OptionC="It does not have to be in writing, does not require any specific phrase, and may be made by someone other than the resident on their behalf"
   OptionD="It can only be made during the application process"; CorrectAnswer="C" }

@{ Title="When requesting verification for a non-obvious disability, what may the Property Manager ask for - and what may they never ask for?"
   OptionA="May ask for a specific diagnosis; may never ask for a doctor's note"
   OptionB="May ask for confirmation that a disability exists and that it's connected to what was requested; may never ask for a diagnosis, medical records, or a specific type of provider"
   OptionC="May ask for full medical records; may never ask for a third party's opinion"
   OptionD="May ask for nothing at all in any circumstance"; CorrectAnswer="B" }

@{ Title="What are the three valid grounds for denying a reasonable accommodation or modification request?"
   OptionA="Cost, inconvenience, and resident history"
   OptionB="Undue financial and administrative burden, fundamental alteration of operations, and direct threat to others or property"
   OptionC="Property age, lease term remaining, and owner preference"
   OptionD="Any of NewShire's standard operating policies"; CorrectAnswer="B" }

@{ Title="If a specific accommodation request is validly denied on one of the three grounds, what does NewShire still owe the resident?"
   OptionA="Nothing further - the conversation ends with the denial"
   OptionB="An interactive process - going back and offering an alternative that meets the same need"
   OptionC="A written apology"
   OptionD="A refund of the application fee"; CorrectAnswer="B" }

@{ Title="An accommodation request sits unanswered in an inbox for five weeks. Under NewShire's standard, this is best understood as:"
   OptionA="Not a problem, since nothing was said no to"
   OptionB="A refusal - an unreasonable delay is treated as a denial, and the clock started the day the resident asked"
   OptionC="Acceptable as long as the resident doesn't follow up"
   OptionD="Only a problem once the resident files a formal complaint"; CorrectAnswer="B" }

@{ Title="On May 22, 2026, HUD rescinded its 2020 guidance on assistance animals. What actually changed, and what didn't?"
   OptionA="The Fair Housing Act itself was amended to remove assistance animal protections"
   OptionB="HUD's enforcement office will no longer pursue providers for denying untrained emotional support animals - but the statute, the courts, and a resident's right to sue privately are unchanged"
   OptionC="Assistance animals are no longer a protected accommodation under any circumstance"
   OptionD="The rescission only applies to properties in South Carolina"; CorrectAnswer="B" }

@{ Title="Regardless of whether an animal is a trained service animal or an emotional support animal, how should a request involving one be handled at the counter?"
   OptionA="Charge the standard pet fee, since policy doesn't distinguish between the two"
   OptionB="Never process it as a pet - no fee, no approval or denial at the counter - and escalate it to the Property Manager the same day"
   OptionC="Approve it automatically if the resident says it's a service animal"
   OptionD="Deny it if the animal hasn't been through professional training"; CorrectAnswer="B" }

@{ Title="A resident presents an online-purchased ESA certificate you don't recognize. What should you do?"
   OptionA="Tell the resident it's not sufficient because it isn't from a physician"
   OptionB="Accept it as final and process the accommodation yourself"
   OptionC="Note what was submitted, don't comment on whether it's sufficient, and get it to the Property Manager the same day"
   OptionD="Deny the request and inform the resident their only option is to appeal"; CorrectAnswer="C" }

# -- Module 4 -- Enforcement and Consequences --------------------------------
@{ Title="A fair housing complaint against NewShire can reach us three different ways. Which of the following correctly matches each door to its filing window?"
   OptionA="HUD - 90 days; SCHAC - 1 year; Federal court - 6 months"
   OptionB="HUD - 1 year, no cost or attorney required; SCHAC - 180 days; Federal court - 2 years, no agency approval needed"
   OptionC="HUD - 30 days; SCHAC - 30 days; Federal court - 1 year"
   OptionD="All three require an attorney and a filing fee to start"; CorrectAnswer="B" }

@{ Title="During a HUD investigation, in addition to the complainant's own file, HUD will typically also request:"
   OptionA="Nothing else - only the complainant's file is reviewed"
   OptionB="The files of everyone else who was treated the same way, for comparison"
   OptionC="A character reference from the Property Manager"
   OptionD="Financial records for the entire company"; CorrectAnswer="B" }

@{ Title="What is the maximum civil penalty for a first fair housing violation, and how does it increase for repeat violations?"
   OptionA="`$26,262 for a first violation; `$65,653 within 5 years; `$131,308 within 7 years or more"
   OptionB="A flat `$50,000 regardless of how many violations"
   OptionC="`$10,000 for a first violation, doubling with each repeat"
   OptionD="There is no civil penalty - only actual damages apply"; CorrectAnswer="A" }

@{ Title="In a federal court fair housing case, are actual damages capped?"
   OptionA="Yes - capped at `$50,000"; OptionB="Yes - capped at the amount of the civil penalty"
   OptionC="No - actual damages, including emotional distress, are uncapped"
   OptionD="Actual damages are not available in fair housing cases"; CorrectAnswer="C" }

@{ Title="Under HUD's liability rule (24 CFR 100.7), which of the following is true?"
   OptionA="Only an individual who personally commits the act can be held liable - the company cannot"
   OptionB="The company can be liable for its own conduct, for failing to correct an employee, or for failing to correct a third party - and an employee who carries out a discriminatory instruction can also be named personally, regardless of who gave the order"
   OptionC="Liability requires proof the company intended to discriminate"
   OptionD="An employee who was following a supervisor's instruction cannot be held individually liable"; CorrectAnswer="B" }

@{ Title="Fair housing testers contact a property to see how it responds, with no intention of renting. Under Havens Realty Corp. v. Coleman (1982), do they have legal standing to sue?"
   OptionA="No - a plaintiff must genuinely intend to rent to have standing"
   OptionB="Yes - the injury is being given false information about available housing, and no intention to rent is required"
   OptionC="Only if they are employed by a government agency"
   OptionD="Only if they can prove they lost money as a result"; CorrectAnswer="B" }

@{ Title="In paired fair housing testing, two testers are matched on everything relevant except one protected characteristic - and staff are never expected to know which one it is. What actually protects you from a testing complaint?"
   OptionA="Trying to identify which callers might be testers and giving them extra care"
   OptionB="Treating every inquiry the same way, logged consistently, regardless of who's asking or how you're feeling that day"
   OptionC="Only providing information in writing so there's no record of verbal conversations"
   OptionD="Limiting the information you share with everyone to reduce risk"; CorrectAnswer="B" }

@{ Title="If you learn that a fair housing complaint has been filed, what should you do in the first hour?"
   OptionA="Contact the complainant directly to try to resolve it"
   OptionB="Notify the Property Manager immediately and preserve all records - emails, notes, texts, call logs, applications"
   OptionC="Wait for official instructions before telling anyone"
   OptionD="Discuss it with coworkers to compare notes"; CorrectAnswer="B" }

@{ Title="After a complaint is filed, changing how you treat the complainant - either more favorably or more strictly than before - may be interpreted as:"
   OptionA="A good-faith effort to resolve the situation"
   OptionB="Retaliation, which is independently unlawful under the Act"
   OptionC="Evidence that NewShire is taking the complaint seriously"
   OptionD="Required corrective action under the Fair Housing Act"; CorrectAnswer="B" }

@{ Title="If HUD finds reasonable cause that discrimination occurred, what happens next?"
   OptionA="HUD immediately issues the maximum penalty"
   OptionB="Either side may elect to move the case to federal court instead of a HUD administrative hearing, and that choice affects what can be awarded"
   OptionC="The case is automatically dismissed unless the complainant appeals"
   OptionD="NewShire has no further options and must accept HUD's administrative ruling"; CorrectAnswer="B" }

)

if ($Questions.Count -ne 40) {
  Write-Host "Expected 40 questions, found $($Questions.Count). Stopping before touching SharePoint." -ForegroundColor Red
  exit 1
}

Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

# -- Find the FHC 101 course record -------------------------------------------
$courses = Get-PnPListItem -List $CoursesList -PageSize 500
$course = $courses | Where-Object { ($_.FieldValues["CourseCode"] -as [string]).Trim() -eq $CourseCode } | Select-Object -First 1

if (-not $course) {
  Write-Host "No course found in '$CoursesList' with CourseCode '$CourseCode'. Nothing to do." -ForegroundColor Red
  Disconnect-PnPOnline
  exit 1
}
$courseId = $course.Id
Write-Host "Found '$CourseCode' as course item $courseId ('$($course.FieldValues['Title'])')." -ForegroundColor Green

# -- Pre-flight: confirm the quiz list has the fields we are about to write -----
# This runs before anything is deleted, and before the -DryRun exit, on purpose.
# A wrong internal name is the one failure mode that would otherwise destroy the
# existing quiz and then fail to add its replacement, and a dry run that skipped
# this check would tell us nothing about it.
$required = @("Title","OptionA","OptionB","OptionC","OptionD","CorrectAnswer","QuizSortOrder","QuizCourseID")
$quizFields = Get-PnPField -List $QuizList
$present    = $quizFields | ForEach-Object { $_.InternalName }
$missing    = $required | Where-Object { $present -notcontains $_ }

if ($missing) {
  Write-Host "'$QuizList' is missing expected field(s): $($missing -join ', ')" -ForegroundColor Red
  Write-Host "Nothing has been changed. Internal names actually on the list:" -ForegroundColor Yellow
  $quizFields | Where-Object { -not $_.Hidden } | ForEach-Object {
    Write-Host "  $($_.InternalName)  [$($_.TypeAsString)]" -ForegroundColor Gray
  }
  Disconnect-PnPOnline
  exit 1
}

$courseField = $quizFields | Where-Object { $_.InternalName -eq "QuizCourseID" }
Write-Host "Pre-flight OK: all $($required.Count) expected fields exist. QuizCourseID is type '$($courseField.TypeAsString)'." -ForegroundColor Green
if ($courseField.TypeAsString -notin @("Lookup","Number","Integer","Text","Counter")) {
  Write-Host "  Heads up: writing a bare course id ($courseId) into a '$($courseField.TypeAsString)' field may not behave as expected." -ForegroundColor Yellow
}

# -- Find existing quiz items for this course ----------------------------------
$existing = Get-PnPListItem -List $QuizList -PageSize 500 | Where-Object {
  $lookupId = $_.FieldValues["QuizCourseID"]
  if ($lookupId -is [Microsoft.SharePoint.Client.FieldLookupValue]) { $lookupId = $lookupId.LookupId }
  $lookupId -eq $courseId
}
Write-Host "Found $($existing.Count) existing '$QuizList' items for $CourseCode." -ForegroundColor Cyan

if ($DryRun) {
  Write-Host "`n-DryRun set. Would recycle $($existing.Count) existing item(s) and add $($Questions.Count) new item(s). No changes made." -ForegroundColor Yellow
  Disconnect-PnPOnline
  exit 0
}

# -- Delete existing, add new ---------------------------------------------------
if ($existing.Count -gt 0) {
  Write-Host "Recycling $($existing.Count) existing item(s) ..." -ForegroundColor Cyan
  foreach ($item in $existing) {
    Remove-PnPListItem -List $QuizList -Identity $item.Id -Recycle -Force
  }
}

Write-Host "Adding $($Questions.Count) question(s) ..." -ForegroundColor Cyan
$order = 1
$added = 0
foreach ($q in $Questions) {
  $values = @{
    Title          = $q.Title
    OptionA        = $q.OptionA
    OptionB        = $q.OptionB
    OptionC        = $q.OptionC
    OptionD        = $q.OptionD
    CorrectAnswer  = $q.CorrectAnswer
    QuizSortOrder  = $order
    QuizCourseID   = $courseId
  }
  try {
    Add-PnPListItem -List $QuizList -Values $values | Out-Null
    $added++
  } catch {
    Write-Host "  [ERROR] Question $order failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  If this mentions 'QuizCourseID', run: Get-PnPField -List '$QuizList' | Select InternalName,TypeAsString" -ForegroundColor Yellow
    Write-Host "  and confirm the lookup field's internal name matches what this script assumes." -ForegroundColor Yellow
  }
  $order++
}

# -- Verify ---------------------------------------------------------------------
$final = Get-PnPListItem -List $QuizList -PageSize 500 | Where-Object {
  $lookupId = $_.FieldValues["QuizCourseID"]
  if ($lookupId -is [Microsoft.SharePoint.Client.FieldLookupValue]) { $lookupId = $lookupId.LookupId }
  $lookupId -eq $courseId
}

Write-Host "`nAdded $added of $($Questions.Count) questions." -ForegroundColor Cyan
if ($final.Count -eq 40) {
  Write-Host "Verified: $CourseCode now has exactly 40 quiz items in '$QuizList'." -ForegroundColor Green
} else {
  Write-Host "WARNING: $CourseCode has $($final.Count) quiz items, expected 40. Review before this goes live." -ForegroundColor Red
}

Disconnect-PnPOnline
