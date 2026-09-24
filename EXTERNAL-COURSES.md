# External Courses (AppFolio Academy)

Author: Brandy Turner, NewShire Property Management

External courses are How-to-Use-AppFolio training hosted on AppFolio Academy
(training.appfolio.com). They are added to NewShire SOP and policy courses and don't replace them.

## How a learner completes one
1. Open the course in NewShire University and click a lesson.
2. **Open in AppFolio Academy** opens the lesson in a new tab. The app records the first time it's opened.
3. Complete the lesson on AppFolio Academy.
4. Back in the app, tick the checkbox and **Sign Acknowledgment**. Signing stays disabled until the lesson has been opened.
5. When every lesson is signed, the app writes a normal `TrainingCompletions` record (Score 100,
   `Answers` = `{"type":"acknowledgment",...}`). Paths, compliance and reports treat it like any
   other completion and show it as "Acknowledged".

There is no quiz. Admins using "View as" cannot sign for someone else.

## What is stored (TrainingAcknowledgments, one row per lesson)
Employee email and name, course ID, code and title, lesson ID and title, the AppFolio URL, the
acknowledgment wording version (`ACK_VERSION`), **the full rendered text they agreed to**
(`AckText`), `LaunchedAt` (first open) and `AcknowledgedAt`. Versioning is on, and people can
only edit rows they created.

If you change the wording in `ackText()`, bump `ACK_VERSION`. Records that are already signed
keep their original text.

## Setup
`pwsh ./scripts/provision-appfolio-academy-courses.ps1` (add `-WhatIf` first to preview it, or
`-SkipPaths` to leave learning paths alone). It adds the columns and the list, and creates APF 151–160
as Coming Soon. Switch each course to Active in Admin > Courses when it's ready.

To make any other course external, set **Course Type** to External in the course form. Each
lesson then gets an **External Course URL** field.
