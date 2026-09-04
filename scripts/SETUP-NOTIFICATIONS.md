# Scheduled notifications — GitHub Actions setup

Makes NewShire University's training emails send on a schedule instead of only
when an admin happens to sign in. No Power Automate, no Azure, $0 on
GitHub-hosted runners.

Until the three secrets below exist, the scheduled run is a harmless no-op — it
logs `Secrets not set` and exits cleanly.

## What it replaces

| Email | Was | Now |
|---|---|---|
| Cert expiring 30 / 14 / 0 days | Admin login | Daily, scheduled |
| Cert expired 7+ days (to admins) | Admin login | Daily, scheduled |
| New training assigned | Admin login | Daily, scheduled |
| Monday manager compliance report | **Power Automate** | Mondays, scheduled |

The login trigger wasn't just inconvenient — the certificate tiers match an
**exact** day count (30 / 14 / 0 / −7 days to expiry). If nobody signed in on
the day someone hit 30 days out, that reminder was missed *permanently*, because
the tier never matches again.

Quiz results, enrolment confirmations, manual course assignment, go-live
announcements and the manual compliance-reminder button all still send from the
app — they're triggered by a person doing something, so they don't need a
schedule.

## 1. Create the sender mailbox

App-only auth has no "signed-in user", so mail goes out as a named mailbox
rather than as whoever was logged in. (Previously staff received these from
*you*, personally.)

M365 admin center → **Teams & groups** → **Shared mailboxes** → **Add**:
- Name: `NewShire University`
- Email: `university@newshirepm.com`

Shared mailboxes are free and need no licence. Replies land there rather than in
your inbox — worth adding yourself as a member so you can see them.

## 2. Create the app registration

> The browser app uses a *delegated* SPA registration. This job runs with nobody
> signed in, so it needs **application** permissions plus a secret.
>
> Use a **new** registration rather than the PM Hub one. `Mail.Send` as an
> application permission is broad, and there's no reason the PM Hub task
> generator should be able to send mail.

1. Entra admin center → **App registrations** → **New registration**
   - Name: `NewShire University Notifier`
   - Single tenant → Register
2. **API permissions** → Add → Microsoft Graph → **Application permissions**:
   - `Sites.ReadWrite.All` — read the lists, write the dedup log
   - `Mail.Send` — send the emails

   Then **Grant admin consent**. Both must show a green check, and both must say
   *Application*, not *Delegated*.
3. **Certificates & secrets** → **New client secret** → copy the **Value**
   immediately. It is shown once and never again.
4. **Overview** → copy the **Application (client) ID** and **Directory (tenant) ID**.

### Recommended: restrict `Mail.Send` to the one mailbox

By default `Mail.Send` lets this app send as **any** mailbox in the tenant. One
command scopes it to `university@` only:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser   # first time only
Connect-ExchangeOnline
New-ApplicationAccessPolicy `
  -AppId <Application (client) ID from step 2.4> `
  -PolicyScopeGroupId university@newshirepm.com `
  -AccessRight RestrictAccess `
  -Description "NewShire University notifier - university mailbox only"
```

Verify with:
`Test-ApplicationAccessPolicy -Identity university@newshirepm.com -AppId <client id>`

Policy changes can take up to ~30 minutes to apply.

## 3. Add the repo secrets

GitHub → this repo → **Settings** → **Secrets and variables** → **Actions** →
**New repository secret**:

| Secret | Value |
|---|---|
| `NSU_TENANT_ID` | Directory (tenant) ID |
| `NSU_CLIENT_ID` | Application (client) ID |
| `NSU_CLIENT_SECRET` | the secret **Value** from step 2.3 |

Optional: `NSU_SENDER` (defaults to `university@newshirepm.com`) and
`NSU_SITE_ID` (defaults to the NewShirePM site).

> Secrets are encrypted and never appear in logs, even though this repo is
> public. The script only ever logs counts and SharePoint item ids — never
> employee names or email addresses — because Actions logs on a public repo are
> world-readable.

## 4. Test it — dry run first

Actions tab → **NewShire University notifications** → **Run workflow**. Leave
**"Preview only"** ticked. It does everything except send, and prints what it
*would* have sent:

```
== Certification reminders ==
  [dry-run] would send: Certification expires in 30 days: FH1 — Fair Housing
```

When that looks right, run it again with the box **unticked**.

### The first real run sends no assignment emails

`runAssignmentScan` seeds a baseline the first time: it records everyone's
current assignments and sends nothing, so nobody gets a blast for training they
were already assigned weeks ago. You'll see:

```
  baseline seeded (N existing assignments recorded, 0 emails sent)
```

Only genuinely new assignments notify from the second run onwards.

## 5. Turn off the Power Automate flow

Once a *scheduled* (not just manual) Monday run has worked, switch off the old
Power Automate Monday manager report so the two don't both send. Both running is
not silently duplicated — the `NotificationLog` dedup key is per-day — but the
flow is the thing you're trying to be rid of.

## Schedule

`.github/workflows/notifications.yml` runs at **12:00 UTC** with a **14:00 UTC**
backup — 8 AM and 10 AM EDT (7/9 AM EST). GitHub cron is UTC and ignores
daylight saving, and scheduled jobs are best-effort: they routinely run hours
late under load. Those times are late enough in the ET day that even a long
delay stays on the same Eastern date, which matters because the cert tiers are
exact-day. The second run is free — every job dedups, so it never double-sends.

The Monday report checks the weekday itself (in Eastern time), so it needs no
separate workflow.

## Maintenance

**The client secret expires.** Entra caps them at 24 months. When it lapses the
workflow fails loudly — token request errors, the run exits non-zero, and GitHub
emails you about the failed run. Fix by creating a new secret on the same app
registration and updating `NSU_CLIENT_SECRET`. Worth a calendar reminder a month
out.

**The logic is duplicated.** `scripts/send-notifications.mjs` contains a port of
the pure helpers from `training-lms.jsx` — `getCertStatus`, `getCourseDueDate`,
`getPathDueStatus`, `courseMatchesRole`, `isTrainingExempt` and friends. If you
change how due dates, cert expiry, content versioning or role matching work in
the app, change them in both places. This is the one real cost of running the
logic outside the browser.
