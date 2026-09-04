# newshire-university

NewShire University — internal training & compliance LMS (React/Babel CDN + MSAL + SharePoint Lists).

- **App:** [`index.html`](index.html) loads [`training-lms.jsx`](training-lms.jsx).
- **Building courses from SOPs:** see [`COURSE-PACKAGE.md`](COURSE-PACKAGE.md) for the SOP → course pipeline.
  Generated course packages live in [`course-packages/`](course-packages/); one-time setup scripts in [`scripts/`](scripts/).

## Automated emails

Certification reminders, new-assignment emails and the Monday manager compliance
report run unattended on a schedule — GitHub Actions + app-only Graph, no Power
Automate, $0. See [`scripts/SETUP-NOTIFICATIONS.md`](scripts/SETUP-NOTIFICATIONS.md).

They used to fire only when an admin signed in, which silently lost reminders:
the cert tiers match an **exact** day count (30 / 14 / 0 / −7), so a day with no
login missed that reminder permanently.

- [`.github/workflows/notifications.yml`](.github/workflows/notifications.yml) —
  daily at 12:00 UTC with a 14:00 UTC backup; `workflow_dispatch` offers a
  preview-only dry run.
- [`scripts/send-notifications.mjs`](scripts/send-notifications.mjs) — the runner.
  Contains a **port** of the pure helpers in `training-lms.jsx`; change due-date,
  cert-expiry, versioning or role-matching rules in both places.

Event-driven emails (quiz results, enrolment, manual assignment, course go-live,
the manual compliance-reminder button) still send from the app.

## Job roles

Course and learning-path targeting matches **exactly** against
`Employees.JobTitle`. Both this app and Employee Lifecycle read their role
picker from the shared **`NS_JobRoles`** list, so the two can't drift apart — a
target role that matches no employee assigns the course to nobody, silently.

- The list is owned and seeded by
  [`employee-lifecycle/scripts/provision-job-roles.ps1`](../employee-lifecycle/scripts/provision-job-roles.ps1);
  full write-up in [`docs/job-roles.md`](../employee-lifecycle/docs/job-roles.md).
- [`scripts/audit-role-targeting.ps1`](scripts/audit-role-targeting.ps1) — read-only.
  Lists courses and paths aimed at a role nobody holds, employees whose title
  isn't canonical, and canonical roles nobody holds.
- [`scripts/fix-role-targeting.ps1`](scripts/fix-role-targeting.ps1) — rewrites
  stale course/path targeting onto canonical roles. Supports `-WhatIf`; refuses
  to leave an item with empty targeting (empty means "everyone").
- App access level reads `Employees.UniversityRole` (set in Employee Lifecycle →
  App Permissions), falling back to the legacy `AccessLevel` column.
