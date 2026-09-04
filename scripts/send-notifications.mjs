// ============================================================
// NewShire University — scheduled notification runner
// ============================================================
// Runs unattended in GitHub Actions. Replaces the in-app scans, which only
// fired when an admin happened to sign in — and because the certification
// tiers match an EXACT day count (30 / 14 / 0 / -7 days to expiry), a day
// with no admin login meant that reminder was missed permanently. Also
// replaces the Power Automate Monday manager report.
//
// App-only Microsoft Graph auth (client credentials). Mirrors the logic that
// used to live in training-lms.jsx; the helper block below is a faithful port
// of the pure functions there. If you change scoring, cert-expiry, due-date or
// role-matching rules in the app, change them here too.
//
// Env (GitHub Actions secrets):
//   NSU_TENANT_ID      Directory (tenant) id
//   NSU_CLIENT_ID      App registration (application) client id
//   NSU_CLIENT_SECRET  Client secret for that app reg
// Optional:
//   NSU_SITE_ID          Graph site id (defaults to the NewShirePM site)
//   NSU_SENDER           Mailbox the emails are sent from
//   NSU_DRY_RUN          "true" = do everything except send/log. Safe preview.
//
// IMPORTANT: this repo is PUBLIC and Actions logs are world-readable.
// Never log employee names or email addresses — counts only.
// ============================================================

const TENANT = process.env.NSU_TENANT_ID;
const CLIENT_ID = process.env.NSU_CLIENT_ID;
const CLIENT_SECRET = process.env.NSU_CLIENT_SECRET;
const SITE_ID = process.env.NSU_SITE_ID ||
  "newshirepmcom.sharepoint.com,f5d74a99-8b23-477c-aed0-c1682efa5de1,0aa4ab90-0ddc-4a81-a803-06d4f2e1a7d8";
const SENDER = process.env.NSU_SENDER || "university@newshirepm.com";
const DRY_RUN = String(process.env.NSU_DRY_RUN || "").toLowerCase() === "true";

if (!TENANT || !CLIENT_ID || !CLIENT_SECRET) {
  // No-op rather than fail, so scheduled runs don't send daily failure
  // notifications before the secrets are configured.
  console.log("Secrets not set (NSU_TENANT_ID / NSU_CLIENT_ID / NSU_CLIENT_SECRET). Skipping — see scripts/SETUP-NOTIFICATIONS.md.");
  process.exit(0);
}

const GRAPH = "https://graph.microsoft.com/v1.0";
const SITE = `${GRAPH}/sites/${SITE_ID}`;
const L = {
  users: "Employees",
  courses: "TrainingCourses",
  paths: "LearningPaths",
  completions: "TrainingCompletions",
  notifications: "NotificationLog",
  config: "AppConfig",
};

// Dates are evaluated in Eastern time (South Carolina), DST-aware — otherwise a
// job running at 04:00 UTC would compute "tomorrow" for half the year and the
// exact-day cert tiers would fire on the wrong day.
const ET_TZ = "America/New_York";
const TODAY = new Date().toLocaleDateString("en-CA", { timeZone: ET_TZ }); // YYYY-MM-DD
const ET_WEEKDAY = new Date().toLocaleDateString("en-US", { timeZone: ET_TZ, weekday: "long" });

let EMAILS_PAUSED = false;

// ---------- auth ----------
async function getToken() {
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });
  const r = await fetch(`https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });
  if (!r.ok) throw new Error(`token ${r.status}: ${await r.text()}`);
  return (await r.json()).access_token;
}

// ---------- graph ----------
async function gAll(token, listName, query = "") {
  let url = `${SITE}/lists/${encodeURIComponent(listName)}/items?expand=fields&$top=999${query}`;
  const out = [];
  while (url) {
    const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!r.ok) throw new Error(`GET ${listName} ${r.status}: ${await r.text().catch(() => "")}`);
    const d = await r.json();
    out.push(...(d.value || []));
    url = d["@odata.nextLink"] || null;
  }
  return out;
}
async function gCreate(token, listName, fields) {
  const r = await fetch(`${SITE}/lists/${encodeURIComponent(listName)}/items`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ fields }),
  });
  if (!r.ok) throw new Error(`POST ${listName} ${r.status}: ${await r.text().catch(() => "")}`);
  return r.json();
}

// ---------- email ----------
const EMAIL_COLORS = { teal: "#1C3740", gold: "#CDA04B", lightBg: "#F7F8F7", white: "#FFFFFF", border: "#E8EAEA", gray: "#28434C" };

function emailTemplate(bodyHtml, subject) {
  return `<!DOCTYPE html><html><head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;background:${EMAIL_COLORS.lightBg};font-family:'Segoe UI',Arial,sans-serif">
<table width="100%" cellpadding="0" cellspacing="0" style="background:${EMAIL_COLORS.lightBg};padding:24px 0">
<tr><td align="center">
<table width="600" cellpadding="0" cellspacing="0" style="background:${EMAIL_COLORS.white};border-radius:6px;border:1px solid ${EMAIL_COLORS.border};overflow:hidden">
  <tr><td style="background:${EMAIL_COLORS.teal};padding:16px 24px;border-bottom:2px solid ${EMAIL_COLORS.gold}">
    <div style="color:#FFF;font-size:15px;font-weight:700;letter-spacing:0.05em">NEWSHIRE UNIVERSITY</div>
    <div style="color:${EMAIL_COLORS.gold};font-size:11px;letter-spacing:0.08em;text-transform:uppercase">Training &amp; Compliance</div>
  </td></tr>
  <tr><td style="padding:24px;color:${EMAIL_COLORS.gray};font-size:14px;line-height:1.6">
    ${bodyHtml}
  </td></tr>
  <tr><td style="padding:14px 24px;background:${EMAIL_COLORS.lightBg};border-top:1px solid ${EMAIL_COLORS.border};font-size:11px;color:#7A8585">
    Automated message from NewShire University. Please do not reply to this email.
  </td></tr>
</table>
</td></tr></table></body></html>`;
}

let sentCount = 0;
// App-only auth has no "me", so mail goes out as an explicit mailbox rather
// than as whoever was signed in. Pass the RAW body — this wraps it exactly
// once (the in-app version had callers double-wrapping the template).
async function sendEmail(token, to, subject, bodyHtml) {
  if (EMAILS_PAUSED) { console.log(`  [paused] suppressed: ${subject}`); return; }
  if (DRY_RUN) { console.log(`  [dry-run] would send: ${subject}`); sentCount++; return; }
  const r = await fetch(`${GRAPH}/users/${encodeURIComponent(SENDER)}/sendMail`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({
      message: {
        subject: `[NewShire University] ${subject}`,
        body: { contentType: "HTML", content: emailTemplate(bodyHtml, subject) },
        toRecipients: (Array.isArray(to) ? to : [to]).map(a => ({ emailAddress: { address: a } })),
      },
      saveToSentItems: true,
    }),
  });
  if (!r.ok && r.status !== 202) throw new Error(`sendMail ${r.status}: ${await r.text().catch(() => "")}`);
  sentCount++;
}

// ============================================================
// PORTED HELPERS — keep in step with training-lms.jsx
// ============================================================
const splitList = raw => {
  if (raw == null) return [];
  const parts = Array.isArray(raw) ? raw.flatMap(r => String(r).split(",")) : String(raw).split(",");
  return parts.map(s => s.trim()).filter(Boolean);
};
const daysBetween = (d1, d2) => Math.round((new Date(d2) - new Date(d1)) / 86400000);
const courseFmt = c => (c ? (c.code ? `${c.code} — ${c.name}` : c.name) : "");
const isTrainingExempt = emp => (emp.role || "").toLowerCase().includes("owner");
const courseMatchesRole = (course, role) => !course.roles || course.roles.length === 0 || course.roles.includes(role);
const getEmployeePaths = (emp, paths) => paths.filter(p => p.roles.includes("All") || p.roles.includes(emp.role));
const isVersionStale = (comp, course) =>
  !!comp && comp.status === "passed" && !!course && (comp.completedVersion || 1) < (course.reqVersion || 1);

function getCertStatus(completion, course) {
  if (!completion || completion.status !== "passed") return "incomplete";
  if (isVersionStale(completion, course)) return "expired";
  if (!course.recertDays && !completion.certExpires) return "current";
  const expiry = completion.certExpires;
  if (!expiry) return "current";
  const daysLeft = daysBetween(TODAY, expiry);
  if (daysLeft < 0) return "expired";
  if (daysLeft <= 30) return "expiring";
  return "current";
}

function getCourseDueDate(course, path, employee) {
  if (!path.dueDays || !employee.hireDate) return null;
  const hire = new Date(employee.hireDate);
  const courseBaseline = course.activatedDate ? new Date(course.activatedDate)
    : course.createdDate ? new Date(course.createdDate) : hire;
  const baseline = hire > courseBaseline ? hire : courseBaseline;
  return new Date(baseline.getTime() + path.dueDays * 86400000).toISOString().split("T")[0];
}

function getPathDueDate(path, employee, courses) {
  if (!path.dueDays || !employee.hireDate) return null;
  const pathCourses = (path.courseIds || []).map(id => courses.find(c => c.id === id)).filter(Boolean);
  if (pathCourses.length === 0) return null;
  let latest = null;
  for (const course of pathCourses) {
    const d = getCourseDueDate(course, path, employee);
    if (d && (!latest || d > latest)) latest = d;
  }
  return latest;
}

function getPathProgress(path, employeeId, completions, courses, employeeRole) {
  const applicableIds = path.courseIds.filter(cid => {
    const course = courses.find(c => c.id === cid);
    return course && course.status === "Active" && courseMatchesRole(course, employeeRole);
  });
  const total = applicableIds.length;
  if (total === 0) return { total: 0, completed: 0, pct: 100 };
  const completed = applicableIds.filter(cid => {
    const comp = completions.filter(c => c.employeeId === employeeId && c.courseId === cid && c.status === "passed");
    if (comp.length === 0) return false;
    const course = courses.find(c => c.id === cid);
    const latest = comp.sort((a, b) => b.completedDate.localeCompare(a.completedDate))[0];
    return getCertStatus(latest, course) !== "expired";
  }).length;
  return { total, completed, pct: Math.round((completed / total) * 100) };
}

function getPathDueStatus(path, employee, completions, courses) {
  if (!path.dueDays || !employee.hireDate) return { dueDate: null, status: null };
  if (isTrainingExempt(employee)) return { dueDate: null, status: null };
  const applicableCourseIds = (path.courseIds || []).filter(cid => {
    const course = courses.find(c => c.id === cid);
    return course && course.status === "Active" && courseMatchesRole(course, employee.role);
  });
  if (applicableCourseIds.length === 0) return { dueDate: null, status: null };

  const progress = getPathProgress(path, employee.id, completions, courses, employee.role);
  const pathDueDate = getPathDueDate(path, employee, courses);
  if (!pathDueDate) return { dueDate: null, status: null };
  if (progress.pct >= 100) return { dueDate: pathDueDate, status: "complete" };

  const stillOpen = cid => {
    const comp = completions.filter(c => c.employeeId === employee.id && c.courseId === cid && c.status === "passed");
    if (comp.length === 0) return true;
    const course = courses.find(c => c.id === cid);
    const latest = comp.sort((a, b) => b.completedDate.localeCompare(a.completedDate))[0];
    return getCertStatus(latest, course) === "expired";
  };

  let earliestOverdue = null;
  for (const cid of applicableCourseIds) {
    if (!stillOpen(cid)) continue;
    const due = getCourseDueDate(courses.find(c => c.id === cid), path, employee);
    if (due && due < TODAY && (!earliestOverdue || due < earliestOverdue)) earliestOverdue = due;
  }
  if (earliestOverdue) return { dueDate: earliestOverdue, status: "overdue" };

  for (const cid of applicableCourseIds) {
    if (!stillOpen(cid)) continue;
    const due = getCourseDueDate(courses.find(c => c.id === cid), path, employee);
    if (due && daysBetween(TODAY, due) <= 7) return { dueDate: due, status: "due-soon" };
  }
  return { dueDate: pathDueDate, status: null };
}

// ============================================================
// LOAD
// ============================================================
async function loadAll(token) {
  const [usersRaw, coursesRaw, pathsRaw, compsRaw, notifRaw, configRaw] = await Promise.all([
    gAll(token, L.users), gAll(token, L.courses), gAll(token, L.paths),
    gAll(token, L.completions), gAll(token, L.notifications), gAll(token, L.config).catch(() => []),
  ]);

  for (const item of configRaw) {
    const f = item.fields || {};
    if (f.Title === "EmailsPaused") EMAILS_PAUSED = f.Value === "true";
  }

  const employees = usersRaw.map(i => {
    const f = i.fields || {};
    return {
      id: String(i.id),
      name: f.Title || "",
      email: (f.Email || "").toLowerCase(),
      role: f.JobTitle || "",
      // Matches the app: ELC's App Permissions matrix owns UniversityRole;
      // AccessLevel is the legacy fallback.
      appRole: f.UniversityRole || f.AccessLevel || "Employee",
      reportsTo: (f.ManagerEmail || "").toLowerCase(),
      hireDate: f.StartDate ? String(f.StartDate).split("T")[0] : null,
      active: f.EmployeeActive !== false,
    };
  });
  const byEmail = {};
  for (const e of employees) byEmail[e.email] = e;

  const courses = coursesRaw.map(i => {
    const f = i.fields || {};
    return {
      id: String(i.id),
      name: f.Title || "",
      code: f.CourseCode || "",
      recertDays: f.RecertDays || null,
      status: f.CourseStatus || (f.CourseActive === false ? "Archived" : "Active"),
      roles: splitList(f.CourseRoles),
      createdDate: i.createdDateTime ? i.createdDateTime.split("T")[0] : null,
      activatedDate: f.ActivatedDate ? String(f.ActivatedDate).split("T")[0] : null,
      reqVersion: f.ReqVersion || 1,
    };
  });

  const paths = pathsRaw.filter(i => i.fields?.PathActive !== false).map(i => {
    const f = i.fields || {};
    return {
      id: String(i.id),
      name: f.Title || "",
      roles: splitList(f.Roles),
      courseIds: splitList(f.CourseIDs),
      required: f.Required !== false,
      dueDays: f.DueDays || null,
    };
  });

  const completions = compsRaw.map(i => {
    const f = i.fields || {};
    const email = (f.EmployeeEmail || "").toLowerCase();
    const emp = byEmail[email];
    return {
      employeeId: emp ? emp.id : email,
      courseId: String(f.CompCourseIDLookupId || f.CompCourseID || ""),
      completedDate: (f.CompletedDate || "").split("T")[0],
      status: (f.CompStatus || "").toLowerCase(),
      certExpires: f.CertExpires ? String(f.CertExpires).split("T")[0] : null,
      completedVersion: f.CompletedVersion || 1,
    };
  });

  const notifKeys = new Set(notifRaw.map(i => i.fields?.NotificationKey || "").filter(Boolean));

  return { employees, courses, paths, completions, notifKeys };
}

const logged = new Set();
async function logNotification(token, key) {
  if (DRY_RUN) { logged.add(key); return; }
  try {
    await gCreate(token, L.notifications, { Title: key, NotificationKey: key, SentDate: new Date().toISOString() });
    logged.add(key);
  } catch (e) { console.error("  failed to log notification:", e.message); }
}

// ============================================================
// 1. CERT EXPIRATION
// ============================================================
async function runCertScan(token, d, adminEmails) {
  let sent = 0;
  for (const emp of d.employees.filter(e => e.active && e.email)) {
    for (const course of d.courses) {
      if (!course.recertDays) continue;
      const passing = d.completions.filter(c => c.employeeId === emp.id && c.courseId === course.id && c.status === "passed");
      if (passing.length === 0) continue;
      const latest = passing.sort((a, b) => b.completedDate.localeCompare(a.completedDate))[0];
      if (!latest.certExpires) continue;

      const daysLeft = daysBetween(TODAY, latest.certExpires);
      let tier = null;
      if (daysLeft === 30) tier = "30day";
      else if (daysLeft === 14) tier = "14day";
      else if (daysLeft === 0) tier = "today";
      else if (daysLeft === -7) tier = "7past";
      if (!tier) continue;

      const key = `cert_${emp.id}_${course.id}_${tier}_${TODAY}`;
      if (d.notifKeys.has(key) || logged.has(key)) continue;

      try {
        if (tier === "7past") {
          if (adminEmails.length) {
            await sendEmail(token, adminEmails,
              `OVERDUE: ${emp.name} — ${courseFmt(course)} certification expired`,
              `<p style="color:#C44B3B;font-weight:600">Certification has been expired for 7+ days.</p>
               <p><strong>${emp.name}</strong> (${emp.role}) — <strong>${courseFmt(course)}</strong></p>
               <p>Expired: <strong>${latest.certExpires}</strong></p>
               <p style="font-size:13px;color:#7A8585">Please follow up directly to ensure recertification is completed.</p>`);
          }
        } else {
          const urgency = tier === "today" ? "expires today" : tier === "14day" ? "expires in 14 days" : "expires in 30 days";
          await sendEmail(token, emp.email,
            `Certification ${urgency}: ${courseFmt(course)}`,
            `<p>Hi ${emp.name.split(" ")[0]},</p>
             <p>Your certification for <strong>${courseFmt(course)}</strong> <strong>${urgency}</strong> (${latest.certExpires}).</p>
             <p>Log in to NewShire University to recertify by retaking the course quiz.</p>`);
          if (adminEmails.length) {
            await sendEmail(token, adminEmails,
              `Cert ${urgency}: ${emp.name} — ${courseFmt(course)}`,
              `<p><strong>${emp.name}</strong> (${emp.role}) — <strong>${courseFmt(course)}</strong> certification ${urgency}.</p>
               <p>Expiration date: <strong>${latest.certExpires}</strong></p>`);
          }
        }
        await logNotification(token, key);
        sent++;
      } catch (e) {
        // Log and continue — one bad address must not stop the whole run.
        console.error(`  cert email failed (employee #${emp.id}, course #${course.id}):`, e.message);
      }
    }
  }
  return sent;
}

// ============================================================
// 2. NEW ASSIGNMENTS
// ============================================================
async function runAssignmentScan(token, d) {
  const BASELINE = "assigned_baseline_v1";
  const baselineDone = d.notifKeys.has(BASELINE);

  const assignedFor = (emp) => {
    const reqPaths = getEmployeePaths(emp, d.paths).filter(p => p.required);
    const ids = [...new Set(reqPaths.flatMap(p => p.courseIds))];
    return ids
      .map(cid => d.courses.find(c => c.id === cid))
      .filter(c => c && c.status === "Active" && courseMatchesRole(c, emp.role))
      .map(c => {
        const path = reqPaths.find(p => p.dueDays && p.courseIds.includes(c.id));
        return { course: c, dueDate: path ? getCourseDueDate(c, path, emp) : null };
      });
  };

  const activeEmps = d.employees.filter(e => e.active && !isTrainingExempt(e) && e.email);

  // First run records current assignments WITHOUT emailing, so nobody gets a
  // blast for training they were already assigned.
  if (!baselineDone) {
    let seeded = 0;
    for (const emp of activeEmps) {
      for (const { course } of assignedFor(emp)) { await logNotification(token, `assigned_${emp.id}_${course.id}`); seeded++; }
    }
    await logNotification(token, BASELINE);
    console.log(`  baseline seeded (${seeded} existing assignments recorded, 0 emails sent)`);
    return 0;
  }

  let sent = 0;
  for (const emp of activeEmps) {
    const newly = assignedFor(emp).filter(({ course }) =>
      !d.notifKeys.has(`assigned_${emp.id}_${course.id}`) && !logged.has(`assigned_${emp.id}_${course.id}`));
    if (newly.length === 0) continue;
    const rows = newly.map(({ course, dueDate }) =>
      `<li><strong>${courseFmt(course)}</strong>${dueDate ? ` &mdash; due ${new Date(dueDate).toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" })}` : ""}</li>`).join("");
    const subject = newly.length > 1 ? `New training assigned (${newly.length} courses)` : `New training assigned: ${courseFmt(newly[0].course)}`;
    try {
      await sendEmail(token, emp.email, subject,
        `<p>Hi ${emp.name.split(" ")[0]},</p>
         <p>The following training ${newly.length > 1 ? "courses have" : "course has"} been assigned to you in NewShire University:</p>
         <ul>${rows}</ul>
         <p>Log in to NewShire University to get started.</p>`);
      for (const { course } of newly) await logNotification(token, `assigned_${emp.id}_${course.id}`);
      sent++;
    } catch (e) {
      console.error(`  assignment email failed (employee #${emp.id}):`, e.message);
    }
  }
  return sent;
}

// ============================================================
// 3. MONDAY MANAGER REPORT  (replaces the Power Automate flow)
// ============================================================
async function runMondayManagerReport(token, d) {
  if (ET_WEEKDAY !== "Monday") return 0;
  const weekKey = `monday_report_${TODAY}`;
  if (d.notifKeys.has(weekKey) || logged.has(weekKey)) { console.log("  already sent today"); return 0; }

  const managers = d.employees.filter(mgr =>
    mgr.active && mgr.email &&
    d.employees.some(e => e.active && e.id !== mgr.id && (e.reportsTo === mgr.id || e.reportsTo === mgr.email)));

  let sent = 0;
  for (const mgr of managers) {
    const reports = d.employees.filter(e => e.active && e.id !== mgr.id && (e.reportsTo === mgr.id || e.reportsTo === mgr.email));
    const issues = [];

    for (const emp of reports) {
      if (isTrainingExempt(emp)) continue;
      for (const path of getEmployeePaths(emp, d.paths).filter(p => p.required)) {
        const { dueDate, status } = getPathDueStatus(path, emp, d.completions, d.courses);
        if (status === "overdue") issues.push({ emp, type: "overdue", detail: `${path.name} was due ${dueDate}` });
        else if (status === "due-soon") issues.push({ emp, type: "due-soon", detail: `${path.name} due ${dueDate}` });

        for (const cid of path.courseIds) {
          const course = d.courses.find(c => c.id === cid);
          if (!course || !course.recertDays) continue;
          if (!courseMatchesRole(course, emp.role)) continue;
          const passing = d.completions.filter(c => c.employeeId === emp.id && c.courseId === cid && c.status === "passed");
          if (passing.length === 0) continue;
          const latest = passing.sort((a, b) => b.completedDate.localeCompare(a.completedDate))[0];
          if (!latest.certExpires) continue;
          const daysLeft = daysBetween(TODAY, latest.certExpires);
          if (daysLeft < 0) issues.push({ emp, type: "expired", detail: `${courseFmt(course)} cert expired ${latest.certExpires}` });
          else if (daysLeft <= 30) issues.push({ emp, type: "expiring", detail: `${courseFmt(course)} cert expires ${latest.certExpires} (${daysLeft}d)` });
        }
      }
    }

    if (issues.length === 0) continue; // no news is good news

    const colorMap = { overdue: "#C44B3B", expired: "#C44B3B", "due-soon": "#D4960A", expiring: "#D4960A" };
    const labelMap = { overdue: "OVERDUE", expired: "EXPIRED", "due-soon": "DUE SOON", expiring: "EXPIRING" };
    const rows = issues.map(i =>
      `<tr>
        <td style="padding:8px 12px;border-bottom:1px solid #E8EAEA;font-size:13px">${i.emp.name}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #E8EAEA;font-size:13px"><span style="display:inline-block;padding:2px 8px;border-radius:9999px;font-size:11px;font-weight:600;color:${colorMap[i.type]};background:${colorMap[i.type]}15">${labelMap[i.type]}</span></td>
        <td style="padding:8px 12px;border-bottom:1px solid #E8EAEA;font-size:13px">${i.detail}</td>
      </tr>`).join("");

    try {
      await sendEmail(token, mgr.email,
        `Weekly Compliance Report — ${reports.length} Direct Report${reports.length > 1 ? "s" : ""}`,
        `<p>Hi ${mgr.name.split(" ")[0]},</p>
         <p>Here is this week's compliance summary for your ${reports.length} direct report${reports.length > 1 ? "s" : ""}.</p>
         <table style="width:100%;border-collapse:collapse;margin:16px 0">
           <tr style="background:#EDF4F7">
             <th style="padding:8px 12px;text-align:left;font-size:12px;color:#28434C;border-bottom:2px solid #D6E7EC">Employee</th>
             <th style="padding:8px 12px;text-align:left;font-size:12px;color:#28434C;border-bottom:2px solid #D6E7EC">Status</th>
             <th style="padding:8px 12px;text-align:left;font-size:12px;color:#28434C;border-bottom:2px solid #D6E7EC">Detail</th>
           </tr>
           ${rows}
         </table>
         <p style="font-size:13px;color:#7A8585">${issues.length} item${issues.length > 1 ? "s" : ""} requiring attention this week.</p>`);
      sent++;
    } catch (e) {
      console.error(`  manager report failed (manager #${mgr.id}):`, e.message);
    }
  }

  if (sent > 0) await logNotification(token, weekKey);
  return sent;
}

// ============================================================
// MAIN
// ============================================================
(async () => {
  console.log(`NewShire University notifications — ${TODAY} (${ET_WEEKDAY}, ${ET_TZ})${DRY_RUN ? " [DRY RUN]" : ""}`);
  const token = await getToken();
  const d = await loadAll(token);
  console.log(`Loaded ${d.employees.length} employees, ${d.courses.length} courses, ${d.paths.length} paths, ${d.completions.length} completions.`);

  if (EMAILS_PAUSED) console.log("EmailsPaused is set in AppConfig — sends will be suppressed.");

  const adminEmails = d.employees
    .filter(e => e.active && e.email && String(e.appRole).trim().toLowerCase() === "admin")
    .map(e => e.email);
  console.log(`${adminEmails.length} admin recipient(s).`);

  let failed = 0;
  const results = {};
  for (const [label, fn] of [
    ["Certification reminders", () => runCertScan(token, d, adminEmails)],
    ["New assignments", () => runAssignmentScan(token, d)],
    ["Monday manager report", () => runMondayManagerReport(token, d)],
  ]) {
    console.log(`\n== ${label} ==`);
    try { results[label] = await fn(); console.log(`  ${results[label]} sent`); }
    catch (e) { failed++; console.error(`  FAILED: ${e.message}`); }
  }

  console.log(`\nDone. ${sentCount} email(s) ${DRY_RUN ? "would have been " : ""}sent across ${Object.keys(results).length} job(s).`);
  // Fail the workflow run if a whole job threw, so a broken scan is visible in
  // the Actions tab rather than silently producing nothing every night.
  if (failed > 0) process.exit(1);
})().catch(e => { console.error("Fatal:", e.message); process.exit(1); });
