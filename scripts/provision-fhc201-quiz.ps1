<#
.SYNOPSIS
  Provisions the FHC 201 course quiz (40 questions) into SharePoint.

  FHC 201 -- Fair Housing in Daily Operations. Four modules, ten questions each:
  M1 Marketing & Leasing, M2 During Tenancy, M3 Maintenance Operations,
  M4 Scenario-Based Decision Making.

  Replaces every existing TrainingQuizzes item tied to FHC 201 with the set
  below, so re-running is idempotent rather than additive. (The first draft of
  this script was add-only, which silently doubled the quiz on a second run.)

  Site: https://newshirepmcom.sharepoint.com/sites/NewShirePM

  What it does, in order:
    1. Connects interactively (your own login, your own MFA).
    2. Looks up the FHC 201 course record in TrainingCourses by CourseCode, to
       get the item ID the quiz questions link to.
    3. Pre-flight: confirms TrainingQuizzes actually has every field this script
       writes, and stops before changing anything if one is missing. This runs
       under -DryRun too, so a dry run really does exercise the field names.
    4. Reports how many TrainingQuizzes items currently exist for that course.
    5. Unless -DryRun is passed, recycles those existing items and adds the 40
       below in their place. Recycled items are recoverable from the site
       recycle bin, so a failed run mid-flight is undoable.
    6. Verifies the final count is exactly 40 and reports any mismatch.

  Run once with -DryRun first to see the before-state without changing anything:
    pwsh -File ./provision-fhc201-quiz.ps1 -DryRun

  Then for real:
    pwsh -File ./provision-fhc201-quiz.ps1

.NOTES
  Requires PnP.PowerShell, which is PowerShell 7+ only (CompatiblePSEditions:
  Core). It will not load under Windows PowerShell 5.1.

  Course-lookup field: QuizCourseID, with no LookupId suffix. Confirmed live
  against this list on 2026-09-03 -- the pre-flight below reports its type and
  it came back 'Lookup'. QuizCourseIDLookupId is the Graph write shortcut and is
  correct for the app's own Graph calls; PnP takes the target item id under the
  base internal name. The app's reader accepts either.

  Correct-answer letters are a deliberate 10 A / 10 B / 10 C / 10 D balance, so
  the pattern of correct letters isn't itself a study shortcut.

  Deleting and re-adding quiz items changes their SharePoint item IDs. Past quiz
  attempts keep their own snapshot of what was asked on the completion record,
  so old completion history and scores are unaffected.
#>

param(
  [string]$SiteUrl     = "https://newshirepmcom.sharepoint.com/sites/NewShirePM",
  # "NewShire Migration Tool" app in the current post-carve-out tenant. The first
  # draft of this script used 32e75ffa-747a-4cf0-8209-6a19150c4547 against
  # vanrockre.sharepoint.com/sites/NewshirePM -- both belong to Vanrock's old
  # tenant and now fail sign-in with AADSTS700016.
  [string]$ClientId    = "7f310acf-12b1-4ba9-a113-c027614268b9",
  [string]$CoursesList = "TrainingCourses",
  [string]$QuizList    = "TrainingQuizzes",
  [string]$CourseCode  = "FHC 201",
  [switch]$DryRun
)

$ErrorActionPreference = "Stop"

# -- The 40 questions --------------------------------------------------------
# Grouped by module in comments only -- the list itself is flat, matching how the
# app reads it. QuizSortOrder is baked into each entry (1-40) rather than
# computed from position, so the module grouping stays explicit.

$questions = @(

    # ── Module 1 — Fair Housing in Marketing & Leasing ─────────────────────────
    @{ Title="What federal regulation makes it unlawful to publish an ad indicating a preference, limitation, or discrimination based on a protected class?"
       OptionA="24 CFR 100.75"; OptionB="24 CFR 100.7"; OptionC="24 CFR 100.600"; OptionD="24 CFR 100.204"
       CorrectAnswer="A"; QuizSortOrder=1 }

    @{ Title="Under the 'ordinary reader' standard from Ragin v. New York Times, what determines whether an ad violates the advertising rule?"
       OptionA="Whether the ad was ever actually published"
       OptionB="What an ordinary reader, seeing the ad cold, would understand it to say"
       OptionC="Whether the ad used any words from HUD's old prohibited-word list"
       OptionD="What the person who wrote the ad meant to say"
       CorrectAnswer="B"; QuizSortOrder=2 }

    @{ Title="Which of these listing phrases is safe to publish because it describes the unit rather than a person?"
       OptionA="'Perfect for a single professional'"; OptionB="'Quiet, mature community'"
       OptionC="'One bedroom, six hundred square feet'"; OptionD="'Walking distance to First Baptist'"
       CorrectAnswer="C"; QuizSortOrder=3 }

    @{ Title="In the showings section, what is the module's working rule for deciding what to show a prospect?"
       OptionA="Show only the units the prospect specifically asks about"
       OptionB="Show the full building only to prospects who mention moving with family"
       OptionC="Let the leasing agent's judgment decide based on what the prospect seems to need"
       OptionD="If you would not skip it for one prospect, do not skip it for any prospect"
       CorrectAnswer="D"; QuizSortOrder=4 }

    @{ Title="In the first version of Scenario One (Priya), why was her showing a violation even though she never refused to show the prospect anything?"
       OptionA="She quoted a different rent for the accessible unit"
       OptionB="She decided, based on the prospect's mention of stairs, what units and information the prospect was allowed to see"
       OptionC="She required the prospect to fill out a separate application"
       OptionD="She refused to answer the prospect's questions"
       CorrectAnswer="B"; QuizSortOrder=5 }

    @{ Title="In the corrected version of that scenario, what did Priya do differently?"
       OptionA="She presented the same full menu of floors, the elevator schedule, and the accessible unit to the prospect, the same as she would to anyone, and let the prospect choose"
       OptionB="She asked the prospect to prove her disability before continuing the tour"
       OptionC="She skipped the elevator schedule to save time"
       OptionD="She sent the prospect directly to the Property Manager instead of continuing the tour"
       CorrectAnswer="A"; QuizSortOrder=6 }

    @{ Title="Per the screening section, what should happen when a genuine exception to written screening criteria is made for one applicant?"
       OptionA="Nothing — exceptions do not need to be documented"
       OptionB="The exception should only ever be granted once per property"
       OptionC="The applicant should be asked to sign a waiver instead"
       OptionD="Document the exception with the actual business reason, and apply that same exception the next time the same situation comes up"
       CorrectAnswer="D"; QuizSortOrder=7 }

    @{ Title="Why does the module say screening paperwork matters, beyond just having a consistent policy?"
       OptionA="Screening notes are not something HUD ever asks for"
       OptionB="Paperwork is only reviewed if the applicant requests it"
       OptionC="When HUD investigates a complaint, it pulls the files of everyone else who applied around the same time to check they were held to the same standard"
       OptionD="SharePoint requires a completed file before an application can be approved"
       CorrectAnswer="C"; QuizSortOrder=8 }

    @{ Title="What changed about sharing neighborhood data (like crime statistics or school ratings) per HUD's April 2026 guidance?"
       OptionA="Sharing it does not by itself violate the steering prohibition, as long as it's delivered consistently and without discriminatory intent"
       OptionB="Sharing this data with prospects is now completely prohibited"
       OptionC="Agents are now required to bring it up unprompted with every prospect"
       OptionD="It replaced the 'ordinary reader' standard for advertising"
       CorrectAnswer="A"; QuizSortOrder=9 }

    @{ Title="In Scenario Two, why was Reese's recommendation to a prospect wearing a hijab a violation, even though he was trying to be welcoming?"
       OptionA="He quoted her a different rent than other prospects"
       OptionB="He refused to show her any units"
       OptionC="He recommended a different property based on an assumption about her background, something nobody asked him to do"
       OptionD="He asked her directly about her religion"
       CorrectAnswer="C"; QuizSortOrder=10 }

    # ── Module 2 — Fair Housing During Tenancy ──────────────────────────────────
    @{ Title="According to the enforcement section, what makes enforcing a lease term against one resident more strictly than another a fair housing violation?"
       OptionA="The lease itself must specifically mention a protected class"
       OptionB="Whether the response would have been different for a different resident, even if the lease never mentions a protected class"
       OptionC="Only formal written notices count as enforcement"
       OptionD="It's only a violation if the resident files a written complaint"
       CorrectAnswer="B"; QuizSortOrder=11 }

    @{ Title="In the first version of Scenario One (Holly, the guest policy), what made her handling of the second resident's violation a problem?"
       OptionA="She never gave anyone a warning for the guest policy"
       OptionB="She charged a late fee that wasn't in the lease"
       OptionC="She refused to enforce the guest policy at all"
       OptionD="A different resident's identical first-time violation got only a verbal reminder with no notice in the file, but this resident's identical violation got a formal notice — same margin, same policy, different response"
       CorrectAnswer="D"; QuizSortOrder=12 }

    @{ Title="What did Holly do differently in the corrected version of that scenario?"
       OptionA="She checked the file for how the last comparable violation was actually handled before deciding the response"
       OptionB="She waived the guest policy entirely"
       OptionC="She asked her Property Manager to make the call instead"
       OptionD="She gave every resident a formal notice regardless of history"
       CorrectAnswer="A"; QuizSortOrder=13 }

    @{ Title="Per the module, what determines whether a maintenance request has been triaged properly?"
       OptionA="Requests should always be handled in the exact order they're received, with no exceptions"
       OptionB="Maintenance staff should ask the resident's household size before scheduling"
       OptionC="Response time, priority, and thoroughness are terms and conditions of the tenancy — they cannot depend on who is asking"
       OptionD="Emergency requests are the only ones that matter for fair housing purposes"
       CorrectAnswer="C"; QuizSortOrder=14 }

    @{ Title="A resident says, 'Can someone put a bar by the tub? My balance is not what it used to be.' How should this be handled?"
       OptionA="As a routine repair ticket, scheduled like any other maintenance request"
       OptionB="It should be denied unless the resident provides medical documentation on the spot"
       OptionC="It should be ignored unless the resident specifically says the word 'disability'"
       OptionD="As a modification request — the resident linked a physical limitation to a requested change, so it routes through the accommodation process, not just a work order"
       CorrectAnswer="D"; QuizSortOrder=15 }

    @{ Title="Under 24 CFR 100.600, when does NewShire have a duty to act on harassment between two residents based on a protected class?"
       OptionA="Only if the harassment happens on NewShire property during business hours"
       OptionB="Once NewShire knows, or should know, that a resident is being harassed by another resident because of a protected class"
       OptionC="Only if both residents file a formal complaint"
       OptionD="NewShire has no duty to act on harassment between residents — that is a police matter"
       CorrectAnswer="B"; QuizSortOrder=16 }

    @{ Title="In the first version of Scenario Two (Andre, at the front desk), what was the problem with telling the resident 'that's really between you two'?"
       OptionA="He described repeated unwanted comments and following in the parking lot from a neighbor, which is harassment based on sex — treating it as a personal matter instead of logging and escalating it is the violation"
       OptionB="Andre should have called the police immediately instead"
       OptionC="He should have offered the resident a lease termination on the spot"
       OptionD="He was not allowed to speak to the resident about the complaint at all"
       CorrectAnswer="A"; QuizSortOrder=17 }

    @{ Title="What is the correct front-desk response to a harassment complaint, per the corrected version of that scenario?"
       OptionA="Investigate the claim yourself before deciding whether to report it"
       OptionB="Tell the resident to work it out directly with their neighbor first"
       OptionC="Wait until a second resident files a similar complaint before taking action"
       OptionD="Hear it, log it exactly as the resident described it, and get it to the Property Manager the same day"
       CorrectAnswer="D"; QuizSortOrder=18 }

    @{ Title="Which of the following is true about a noise complaint versus a harassment complaint, per this module?"
       OptionA="They are always the same thing and should be handled identically"
       OptionB="Neither one is ever NewShire's responsibility"
       OptionC="A noise complaint is between residents; a slur or unwanted conduct tied to who someone is is between that resident and NewShire"
       OptionD="Only a Property Manager is allowed to receive either type of complaint"
       CorrectAnswer="C"; QuizSortOrder=19 }

    @{ Title="What earlier course and module first taught the 'hear it, log it, get it to the Property Manager' process for a modification or accommodation request?"
       OptionA="FHC 101, Module 3"; OptionB="FHC 201, Module 1"
       OptionC="WPS 101"; OptionD="FHC 101, Module 1"
       CorrectAnswer="A"; QuizSortOrder=20 }

    # ── Module 3 — Fair Housing in Maintenance Operations ──────────────────────
    @{ Title="Per the NFHA's 2025 Fair Housing Trends Report cited in this module, what percentage of all tracked fair housing complaints were about disability?"
       OptionA="54.59%"; OptionB="25%"; OptionC="7.13%"; OptionD="15.58%"
       CorrectAnswer="A"; QuizSortOrder=21 }

    @{ Title="What are the three honest factors allowed to move a maintenance request up or down the triage queue?"
       OptionA="How long the resident has lived there, household size, and how the request was submitted"
       OptionB="Whether the resident has ever filed a complaint before"
       OptionC="Safety severity, health-hazard risk, and how long the ticket has been open"
       OptionD="The technician's personal judgment about urgency"
       CorrectAnswer="C"; QuizSortOrder=22 }

    @{ Title="In the first version of Scenario One (Nate), why was moving the leaking supply line ticket down the list a violation?"
       OptionA="The line wasn't actually leaking"
       OptionB="A leaking supply line is a water-damage risk, which put it near the top of the list by the three honest factors — he moved it down anyway based on an assumption about a household with children"
       OptionC="He didn't have the parts on hand"
       OptionD="The resident never submitted a work order"
       CorrectAnswer="B"; QuizSortOrder=23 }

    @{ Title="In this module, what is the test for whether a triage decision was made properly?"
       OptionA="Whether the reason could be said out loud to the resident's face — if the honest reason would embarrass you, it was not the real reason"
       OptionB="Whether the resident was satisfied with the outcome"
       OptionC="Whether the ticket was closed within 24 hours"
       OptionD="Whether a supervisor approved the order in advance"
       CorrectAnswer="A"; QuizSortOrder=24 }

    @{ Title="What rule from FHC 101, Module 2 does this module explicitly carry over for recognizing a request at the door?"
       OptionA="Always ask the resident directly if they have a disability"
       OptionB="Only log a request if it's submitted in writing"
       OptionC="Write down what a person asked for — never write down what you think they asked"
       OptionD="Refer every unusual observation to the Regional Manager"
       CorrectAnswer="C"; QuizSortOrder=25 }

    @{ Title="In the first version of Scenario Two (Carmen), what made writing 'Resident appears disabled — recommend ADA modification' on the work order a violation?"
       OptionA="The work order form doesn't have room for notes"
       OptionB="She should have called the resident before starting the repair"
       OptionC="ADA modifications are never NewShire's responsibility"
       OptionD="She diagnosed a disability herself, based only on what she observed, and put that diagnosis in writing without the resident ever asking for anything"
       CorrectAnswer="D"; QuizSortOrder=26 }

    @{ Title="Under South Carolina law, does a maintenance technician still need to announce their intent to enter before going in on a tenant-requested repair, even though the work order itself serves as the tenant's notice?"
       OptionA="No — submitting a work order removes any further notice obligation"
       OptionB="Only if the resident is home at the time"
       OptionC="Yes — the work order removes the 24-hour wait that applies to other entries, but the technician still must announce intent to enter, such as a knock, call, or text"
       OptionD="Only for emergency repairs"
       CorrectAnswer="C"; QuizSortOrder=27 }

    @{ Title="When does the 24-hour notice floor under South Carolina law actually apply, per this module?"
       OptionA="To every entry, including tenant-requested repairs"
       OptionB="To entries the resident did not request — inspections, showings, anything landlord-initiated"
       OptionC="It never applies in South Carolina"
       OptionD="Only to entries scheduled after 8:00 p.m."
       CorrectAnswer="B"; QuizSortOrder=28 }

    @{ Title="If a resident has asked to be called or texted before anyone enters — whether on the ticket itself or as an accommodation on file — what should happen?"
       OptionA="It's a nice-to-have, but the standard notice rules still control"
       OptionB="Only a Property Manager can honor that kind of request"
       OptionC="It only applies to emergency repairs"
       OptionD="That request controls; skipping it because a ticket was already filed is not a defense, and if it's tied to a disability it's a fair housing issue"
       CorrectAnswer="D"; QuizSortOrder=29 }

    @{ Title="What are the three rules for maintenance documentation taught in this module?"
       OptionA="Photograph every repair, get a signature, and file it within 30 days"
       OptionB="Log only tickets over $500, write the outcome only, and always note the resident's household composition"
       OptionC="Log every ticket the same way, write the triage reason not just the outcome, and never write a characteristic that was not asked about"
       OptionD="Log the technician's name, the date, and nothing else"
       CorrectAnswer="C"; QuizSortOrder=30 }

    # ── Module 4 — Scenario-Based Decision Making ───────────────────────────────
    @{ Title="Per this module, what is true about the rules it covers?"
       OptionA="They only apply to Property Managers and above"
       OptionB="Every rule already exists somewhere in FHC 101 or the first three modules of this course — what changes here is the pace"
       OptionC="They are new law not covered anywhere else in the curriculum"
       OptionD="They replace the rules taught in FHC 101"
       CorrectAnswer="B"; QuizSortOrder=31 }

    @{ Title="A leasing agent tells a prospect who mentioned her kids, 'Honestly, this building leans more toward young professionals.' What is this?"
       OptionA="Not a violation, since it wasn't said with any intent to discriminate"
       OptionB="Acceptable because it describes the building's general atmosphere"
       OptionC="A violation only if the prospect complains about it"
       OptionD="A violation — an unsolicited comment about who fits, tied to familial status, said out loud during a showing"
       CorrectAnswer="D"; QuizSortOrder=32 }

    @{ Title="A leasing agent tells an applicant, 'I can't waive the income requirement — it's the same for every applicant.' What is this?"
       OptionA="Not a violation — one standard, applied the same way to everyone who asks, even when the answer disappoints someone"
       OptionB="A violation, because the applicant was denied"
       OptionC="A violation because income requirements are never allowed"
       OptionD="Only acceptable if the applicant is told in writing"
       CorrectAnswer="A"; QuizSortOrder=33 }

    @{ Title="A maintenance supervisor tells a coworker, 'Her English isn't great, so I always send someone else for her tickets.' What is this?"
       OptionA="Not a violation, since it's about efficiency, not discrimination"
       OptionB="A violation — an assignment decision built on an assumption about national origin, not on anything the resident has ever actually asked for"
       OptionC="Only a violation if the resident finds out about the comment"
       OptionD="Acceptable as long as the same technician is always sent"
       CorrectAnswer="B"; QuizSortOrder=34 }

    @{ Title="A Regional Manager approves one routine lease renewal within a day, but lets a comparable renewal for a resident with a documented disability sit untouched for three weeks with no explanation. Is this a fair housing violation?"
       OptionA="No, because nobody said anything discriminatory"
       OptionB="Yes — silence and delay are decisions too, and an unexplained timeline disparity can be evidence of a violation even without a single spoken word"
       OptionC="No, because renewal timing is entirely at the manager's discretion"
       OptionD="Only if the resident specifically asks why it's taking so long"
       CorrectAnswer="B"; QuizSortOrder=35 }

    @{ Title="Which of these is one of the five red-flag phrases this module identifies as a reason to stop mid-sentence?"
       OptionA="'Let me check the file and get back to you.'"
       OptionB="'I'll have that scheduled by end of day.'"
       OptionC="'Thank you for letting me know.'"
       OptionD="'That's really between you two.'"
       CorrectAnswer="D"; QuizSortOrder=36 }

    @{ Title="According to the 'red-flag behaviors' section, what can be a fair housing violation even without a single discriminatory word ever being spoken?"
       OptionA="Delay, unequal effort, and silence — decisions that show up in the record whether anyone meant them or not"
       OptionB="Nothing — a violation always requires something said out loud"
       OptionC="Only actions taken by a Property Manager or above"
       OptionD="Only actions that are reported by a third party"
       CorrectAnswer="A"; QuizSortOrder=37 }

    @{ Title="What are the three rules for a defensible record taught in this module's documentation section?"
       OptionA="Log only interactions that resulted in a complaint"
       OptionB="Log the same information for every interaction, write the reason not just the outcome, and never write a characteristic nobody asked about"
       OptionC="Write the outcome only, and skip the reason to save time"
       OptionD="Only document interactions involving a protected class"
       CorrectAnswer="B"; QuizSortOrder=38 }

    @{ Title="Per the escalation section, what should you do about something that feels off, even if you're not fully sure why?"
       OptionA="Wait until you're certain before saying anything"
       OptionB="Only escalate if a resident files a formal written complaint first"
       OptionC="Say something to whoever you would normally report to, the same day — you don't have to be right, you have to say it"
       OptionD="Handle it yourself first, then report it only if it turns out to be serious"
       CorrectAnswer="C"; QuizSortOrder=39 }

    @{ Title="If a fair housing concern involves the person you would normally report to, what does this module say to do?"
       OptionA="Let it go, since there's no one else to tell"
       OptionB="Only HR is authorized to receive that kind of complaint"
       OptionC="Wait for your next scheduled review to bring it up"
       OptionD="Go to their supervisor instead — the chain does not stop just because the person in it is the problem"
       CorrectAnswer="D"; QuizSortOrder=40 }
)

if ($questions.Count -ne 40) {
  Write-Host "Expected 40 questions, found $($questions.Count). Stopping before touching SharePoint." -ForegroundColor Red
  exit 1
}

# QuizSortOrder is hand-written per question here, so verify it really is 1-40
# with no duplicates or gaps before anything is written.
$orders = @($questions | ForEach-Object { $_.QuizSortOrder } | Sort-Object -Unique)
if ($orders.Count -ne 40 -or $orders[0] -ne 1 -or $orders[-1] -ne 40) {
  Write-Host "QuizSortOrder is not 40 distinct values from 1 to 40. Stopping." -ForegroundColor Red
  exit 1
}

Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

# -- Find the FHC 201 course record -------------------------------------------
$courses = Get-PnPListItem -List $CoursesList -PageSize 500
$course = $courses | Where-Object { ($_.FieldValues["CourseCode"] -as [string]).Trim() -eq $CourseCode } | Select-Object -First 1

if (-not $course) {
  Write-Host "No course found in '$CoursesList' with CourseCode '$CourseCode'. Create the course record before running this script." -ForegroundColor Red
  Disconnect-PnPOnline
  exit 1
}
$courseId = $course.Id
Write-Host "Found '$CourseCode' as course item $courseId ('$($course.FieldValues['Title'])')." -ForegroundColor Green

# -- Pre-flight: confirm the quiz list has the fields we are about to write -----
# This runs before anything is deleted, and before the -DryRun exit, on purpose.
# A wrong internal name is the one failure mode that would otherwise destroy the
# existing quiz and then fail to add its replacement.
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
  Write-Host "`n-DryRun set. Would recycle $($existing.Count) existing item(s) and add $($questions.Count) new item(s). No changes made." -ForegroundColor Yellow
  Disconnect-PnPOnline
  exit 0
}

# -- Recycle existing, add new ---------------------------------------------------
if ($existing.Count -gt 0) {
  Write-Host "Recycling $($existing.Count) existing item(s) ..." -ForegroundColor Cyan
  # -Recycle emits a RecycleBinItemId per item. Collect them rather than letting
  # bare GUIDs scroll past -- the recycle bin itself is the recovery UI, so a
  # count is all that's useful here.
  $recycled = @()
  foreach ($item in $existing) {
    $recycled += Remove-PnPListItem -List $QuizList -Identity $item.Id -Recycle -Force
  }
  Write-Host "  $($recycled.Count) item(s) moved to the site recycle bin (recoverable there if this run goes wrong)." -ForegroundColor Gray
}

Write-Host "Adding $($questions.Count) question(s) ..." -ForegroundColor Cyan
$added = 0
foreach ($q in $questions | Sort-Object { $_.QuizSortOrder }) {
  $values = @{
    Title          = $q.Title
    OptionA        = $q.OptionA
    OptionB        = $q.OptionB
    OptionC        = $q.OptionC
    OptionD        = $q.OptionD
    CorrectAnswer  = $q.CorrectAnswer
    QuizSortOrder  = $q.QuizSortOrder
    QuizCourseID   = $courseId
  }
  try {
    Add-PnPListItem -List $QuizList -Values $values | Out-Null
    $added++
  } catch {
    Write-Host "  [ERROR] Question $($q.QuizSortOrder) failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  If this mentions 'QuizCourseID', run: Get-PnPField -List '$QuizList' | Select InternalName,TypeAsString" -ForegroundColor Yellow
    Write-Host "  and confirm the lookup field's internal name matches what this script assumes." -ForegroundColor Yellow
  }
}

# -- Verify ---------------------------------------------------------------------
$final = Get-PnPListItem -List $QuizList -PageSize 500 | Where-Object {
  $lookupId = $_.FieldValues["QuizCourseID"]
  if ($lookupId -is [Microsoft.SharePoint.Client.FieldLookupValue]) { $lookupId = $lookupId.LookupId }
  $lookupId -eq $courseId
}

Write-Host "`nAdded $added of $($questions.Count) questions." -ForegroundColor Cyan
if ($final.Count -eq 40) {
  Write-Host "Verified: $CourseCode now has exactly 40 quiz items in '$QuizList'." -ForegroundColor Green
} else {
  Write-Host "WARNING: $CourseCode has $($final.Count) quiz items, expected 40. Review before this goes live." -ForegroundColor Red
}

Disconnect-PnPOnline
