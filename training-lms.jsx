const { useState, useEffect, useCallback, useRef, useMemo, createContext, useContext } = React;

// Favicon — gold box with teal "U" for University
(()=>{const l=document.querySelector("link[rel='icon']")||document.createElement("link");l.rel="icon";l.href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'><rect width='32' height='32' rx='6' fill='%23CDA04B'/><text x='16' y='23' font-family='Georgia,serif' font-size='22' font-weight='bold' fill='%2328434C' text-anchor='middle'>U</text></svg>";document.head.appendChild(l)})();

// ============================================================
// CONFIG — Update per-app
// ============================================================
const CONFIG = {
  clientId: "32e75ffa-747a-4cf0-8209-6a19150c4547",
  tenantId: "33575d04-ca7b-4396-8011-9eaea4030b46",
  siteId: "vanrockre.sharepoint.com,a02c1cd8-9f1f-4827-8286-7b6b7ce74232,01202419-6625-4499-b0d5-8ceb1cffdba3",
  lists: {
    users:        "Employees",
    courses:      "TrainingCourses",
    paths:        "LearningPaths",
    lessons:      "TrainingLessons",
    quizzes:      "TrainingQuizzes",
    completions:  "TrainingCompletions",
    enrollments:  "TrainingEnrollments",
    assignments:  "TrainingAssignments",
    config:       "AppConfig",
    notifications: "NotificationLog",
  },
  appName: "NEWSHIRE UNIVERSITY",
  // Set false to run in demo mode with hardcoded data (no SharePoint connection)
  isConfigured: true,
  passingScore: 80,
  adminEmail: "bturner@newshirepm.com", // Fallback admin for notifications
};

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const SITE_URL = `${GRAPH_BASE}/sites/${CONFIG.siteId}`;

// Global email pause flag — checked by sendEmail before every send
// DEFAULT: false (active) for production. Toggle via Settings tab.
let EMAIL_PAUSED = false;

// ============================================================
// REACT CONTEXT — All components pull data from here
// ============================================================
const DataContext = createContext(null);
function useData() { return useContext(DataContext); }

// ============================================================
// MSAL AUTHENTICATION
// ============================================================
const MSAL_CONFIG = {
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: window.location.origin,
  },
  cache: { cacheLocation: "sessionStorage" },
};
const SCOPES = ["Sites.ReadWrite.All", "User.Read", "Mail.Send"];

let _msalInstance = null;
const MSAL_CDN = "https://unpkg.com/@azure/msal-browser@2.38.3/lib/msal-browser.min.js";

async function loadMsalScript() {
  // If loaded via index.html <script> tag, it's already available
  if (window.msal) return;
  // Check if script tag exists but hasn't finished loading
  const existing = document.querySelector(`script[src="${MSAL_CDN}"]`);
  if (existing) {
    return new Promise((resolve, reject) => {
      if (window.msal) { resolve(); return; }
      existing.addEventListener("load", () => window.msal ? resolve() : reject(new Error("MSAL loaded but not on window")));
      existing.addEventListener("error", () => reject(new Error("Failed to load MSAL library from CDN")));
      // Timeout fallback — script may have already loaded before listener attached
      setTimeout(() => window.msal ? resolve() : null, 500);
    });
  }
  // Dynamic injection fallback
  return new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = MSAL_CDN;
    s.onload = resolve;
    s.onerror = () => reject(new Error("Failed to load MSAL library from CDN"));
    document.head.appendChild(s);
  });
}

async function getMsal() {
  if (_msalInstance) return _msalInstance;
  await loadMsalScript();
  const msal = window.msal;
  if (!msal) throw new Error("MSAL library not available after CDN load.");
  _msalInstance = new msal.PublicClientApplication(MSAL_CONFIG);
  await _msalInstance.initialize();
  return _msalInstance;
}

async function msalLogin() {
  const instance = await getMsal();
  const accounts = instance.getAllAccounts();
  if (accounts.length > 0) return accounts[0];
  try {
    const response = await instance.loginPopup({ scopes: SCOPES });
    return response.account;
  } catch (err) {
    if (err.errorCode === "user_cancelled") return null;
    throw err;
  }
}

async function msalGetToken(account) {
  const instance = await getMsal();
  try {
    const response = await instance.acquireTokenSilent({ scopes: SCOPES, account });
    return response.accessToken;
  } catch {
    const response = await instance.acquireTokenPopup({ scopes: SCOPES, account });
    return response.accessToken;
  }
}

// ============================================================
// SHAREPOINT GRAPH API SERVICE
// ============================================================
function listUrl(listName) {
  return `${SITE_URL}/lists/${listName}/items`;
}

async function spGet(token, listName, opts = {}) {
  const { filter, expand = "fields", top = 200 } = opts;
  let all = [];
  let url = `${listUrl(listName)}?expand=${expand}&$top=${top}`;
  if (filter) url += `&$filter=${encodeURIComponent(filter)}`;
  while (url) {
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) throw new Error(`GET ${listName} failed (${res.status})`);
    const data = await res.json();
    all = all.concat(data.value || []);
    url = data["@odata.nextLink"] || null;
  }
  return all;
}

async function spCreate(token, listName, fields) {
  const res = await fetch(listUrl(listName), {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ fields }),
  });
  if (!res.ok) throw new Error(`POST ${listName} failed (${res.status})`);
  return res.json();
}

async function spDelete(token, listName, itemId) {
  const res = await fetch(`${listUrl(listName)}/${itemId}`, {
    method: "DELETE",
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok && res.status !== 204) throw new Error(`DELETE ${listName}/${itemId} failed (${res.status})`);
}

async function spUpdate(token, listName, itemId, fields) {
  const res = await fetch(`${listUrl(listName)}/${itemId}/fields`, {
    method: "PATCH",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify(fields),
  });
  if (!res.ok) throw new Error(`PATCH ${listName}/${itemId} failed (${res.status})`);
  return res.json();
}

// ============================================================
// DATA NORMALIZATION — SharePoint fields → app shape
// ============================================================
function normalizeEmployees(items) {
  return items.map(item => {
    const f = item.fields;
    return {
      id: String(item.id),
      name: f.Title || "",
      email: (f.Email || "").toLowerCase(),
      role: f.JobTitle || "",
      appRole: f.AccessLevel || "Employee",
      reportsTo: (f.ManagerEmail || "").toLowerCase(),
      hireDate: f.StartDate ? f.StartDate.split("T")[0] : null,
      active: f.EmployeeActive !== false,
    };
  });
}

function normalizeCourses(items) {
  return items
    .map(item => {
      const f = item.fields;
      let status = f.CourseStatus || (f.CourseActive === false ? "Archived" : "Active");
      const rolesRaw = f.CourseRoles || "";
      const roles = rolesRaw ? rolesRaw.split(",").map(s => s.trim()).filter(Boolean) : [];
      return {
        id: String(item.id),
        name: f.Title || "",
        code: f.CourseCode || "",
        description: f.CourseDescription || "",
        category: f.Category || "Onboarding",
        durationMin: f.DurationMin || 0,
        recertDays: f.RecertDays || null,
        passingScore: f.PassingScore || CONFIG.passingScore,
        sortOrder: f.SortOrder || 999,
        status: status,
        roles: roles, // empty = all roles, populated = only these roles
        createdDate: item.createdDateTime ? item.createdDateTime.split("T")[0] : null,
      };
    })
    .sort((a, b) => a.sortOrder - b.sortOrder);
}

function normalizePaths(items) {
  return items
    .filter(item => item.fields.PathActive !== false)
    .map(item => {
      const f = item.fields;
      const courseIdStr = f.CourseIDs || "";
      const courseIds = courseIdStr.split(",").map(s => s.trim()).filter(Boolean);
      let roles = [];
      const rolesRaw = f.Roles;
      if (Array.isArray(rolesRaw)) {
        // If SP returned an array, flatten any comma-separated entries within
        roles = rolesRaw.flatMap(r => r.split(",").map(s => s.trim())).filter(Boolean);
      } else if (typeof rolesRaw === "string" && rolesRaw.trim()) {
        roles = rolesRaw.split(",").map(s => s.trim()).filter(Boolean);
      }
      return {
        id: String(item.id),
        name: f.Title || "",
        description: f.PathDescription || "",
        roles,
        courseIds,
        required: f.Required !== false,
        recertDays: f.RecertDays || null,
        dueDays: f.DueDays || null,
      };
    });
}

function normalizeLessons(items) {
  return items.map(item => {
    const f = item.fields;
    const courseId = String(f.CourseIDLookupId || f.CourseID || "");
    return {
      id: String(item.id),
      courseId,
      title: f.Title || "",
      order: f.LessonSortOrder || 1,
      durationMin: f.LessonDurationMin || 0,
      videoUrl: f.VideoURL || null,
      documentUrl: f.DocumentURL || null,
      documentTitle: f.DocumentTitle || null,
      // Supplements: stored as JSON array in SupplementURL, or legacy single URL
      supplements: (() => {
        const raw = f.SupplementURL || "";
        if (!raw) return [];
        try {
          const parsed = JSON.parse(raw);
          if (Array.isArray(parsed)) return parsed;
        } catch {}
        // Legacy single URL format — wrap in array
        return [{ title: f.SupplementTitle || "Supplemental Material", url: raw }];
      })(),
    };
  }).sort((a, b) => a.order - b.order);
}

function normalizeQuizzes(items) {
  const map = {};
  const sorted = [...items].sort((a, b) => (a.fields.QuizSortOrder || 0) - (b.fields.QuizSortOrder || 0));
  for (const item of sorted) {
    const f = item.fields;
    const courseId = String(f.QuizCourseIDLookupId || f.QuizCourseID || "");
    if (!courseId) continue;
    if (!map[courseId]) map[courseId] = { questions: [] };
    map[courseId].questions.push({
      id: String(item.id),
      question: f.Title || "",
      options: { A: f.OptionA || "", B: f.OptionB || "", C: f.OptionC || "", D: f.OptionD || "" },
      correct: f.CorrectAnswer || "A",
    });
  }
  return map;
}

function normalizeCompletions(items, employeesByEmail) {
  return items.map(item => {
    const f = item.fields;
    const email = (f.EmployeeEmail || "").toLowerCase();
    const emp = employeesByEmail[email];
    return {
      id: String(item.id),
      employeeId: emp ? emp.id : email,
      courseId: String(f.CompCourseIDLookupId || f.CompCourseID || ""),
      completedDate: (f.CompletedDate || "").split("T")[0],
      score: f.Score || 0,
      status: (f.CompStatus || "").toLowerCase(),
      certExpires: f.CertExpires ? f.CertExpires.split("T")[0] : null,
    };
  });
}

function normalizeEnrollments(items, employeesByEmail) {
  return items.map(item => {
    const f = item.fields;
    const email = (f.EnrollEmployeeEmail || "").toLowerCase();
    const emp = employeesByEmail[email];
    return {
      spItemId: item.id,
      employeeId: emp ? emp.id : email,
      courseId: String(f.EnrollCourseIDLookupId || f.EnrollCourseID || ""),
      enrolledDate: (f.EnrolledDate || "").split("T")[0],
    };
  });
}

function normalizeAssignments(items, employeesByEmail) {
  return items.map(item => {
    const f = item.fields;
    const empEmail = (f.AssignEmployeeEmail || "").toLowerCase();
    const assignerEmail = (f.AssignedByEmail || "").toLowerCase();
    const emp = employeesByEmail[empEmail];
    const assigner = employeesByEmail[assignerEmail];
    return {
      id: String(item.id),
      employeeId: emp ? emp.id : empEmail,
      courseId: String(f.AssignCourseIDLookupId || f.AssignCourseID || ""),
      assignedBy: assigner ? assigner.name : assignerEmail,
      assignedById: assigner ? assigner.id : assignerEmail,
      assignedDate: (f.AssignedDate || "").split("T")[0],
      dueDate: f.AssignDueDate ? f.AssignDueDate.split("T")[0] : null,
      notes: f.AssignNotes || "",
      status: f.AssignStatus || "Assigned", // Assigned, Completed, Dismissed
    };
  });
}

// ============================================================
// DATA LOADER — parallel fetch of all 8 lists
// ============================================================
async function loadAllData(token) {
  const L = CONFIG.lists;
  const [usersRaw, coursesRaw, pathsRaw, lessonsRaw, quizzesRaw, completionsRaw, enrollmentsRaw, assignmentsRaw, configRaw] =
    await Promise.all([
      spGet(token, L.users), spGet(token, L.courses), spGet(token, L.paths),
      spGet(token, L.lessons), spGet(token, L.quizzes), spGet(token, L.completions),
      spGet(token, L.enrollments), spGet(token, L.assignments).catch(() => []), spGet(token, L.config),
    ]);
  const employees = normalizeEmployees(usersRaw);
  const employeesByEmail = {};
  for (const e of employees) employeesByEmail[e.email] = e;
  const courses = normalizeCourses(coursesRaw);
  const paths = normalizePaths(pathsRaw);
  const lessons = normalizeLessons(lessonsRaw);
  const quizzes = normalizeQuizzes(quizzesRaw);
  const completions = normalizeCompletions(completionsRaw, employeesByEmail);
  const enrollments = normalizeEnrollments(enrollmentsRaw, employeesByEmail);
  const assignments = normalizeAssignments(assignmentsRaw, employeesByEmail);
  // Apply config overrides
  for (const item of configRaw) {
    const f = item.fields;
    if (f.Title === "DefaultPassingScore" && f.Value) CONFIG.passingScore = parseInt(f.Value, 10) || 80;
    if (f.Title === "EmailsPaused") EMAIL_PAUSED = f.Value === "true";
  }
  return { employees, employeesByEmail, courses, paths, lessons, quizzes, completions, enrollments, assignments };
}

// ============================================================
// WRITE OPERATIONS — POST/DELETE to SharePoint
// ============================================================
async function submitQuizToSP(token, employee, course, score, passed, answersJson) {
  const certExpires = passed && course.recertDays
    ? new Date(Date.now() + course.recertDays * 86400000).toISOString() : null;
  const fields = {
    Title: `${employee.name} - ${course.name} - ${new Date().toISOString().split("T")[0]}`,
    EmployeeEmail: employee.email,
    CompCourseIDLookupId: parseInt(course.id, 10),
    CompletedDate: new Date().toISOString(),
    Score: score,
    CompStatus: passed ? "Passed" : "Failed",
    Answers: answersJson,
  };
  if (certExpires) fields.CertExpires = certExpires;
  const result = await spCreate(token, CONFIG.lists.completions, fields);
  return {
    id: String(result.id),
    employeeId: employee.id,
    courseId: course.id,
    completedDate: new Date().toISOString().split("T")[0],
    score,
    status: passed ? "passed" : "failed",
    certExpires: certExpires ? certExpires.split("T")[0] : null,
  };
}

async function createEnrollmentSP(token, employee, courseId) {
  const fields = {
    Title: `${employee.name} - Course ${courseId}`,
    EnrollEmployeeEmail: employee.email,
    EnrollCourseIDLookupId: parseInt(courseId, 10),
    EnrolledDate: new Date().toISOString(),
  };
  const result = await spCreate(token, CONFIG.lists.enrollments, fields);
  return {
    spItemId: result.id,
    employeeId: employee.id,
    courseId,
    enrolledDate: new Date().toISOString().split("T")[0],
  };
}

async function deleteEnrollmentSP(token, spItemId) {
  await spDelete(token, CONFIG.lists.enrollments, spItemId);
}

// ============================================================
// EMAIL SERVICE — Graph API Mail.Send with branded HTML
// ============================================================
const EMAIL_COLORS = { teal: "#1C3740", gold: "#CDA04B", gray: "#3E4A4A", lightBg: "#F7F8F7", white: "#FFFFFF", border: "#D0D8DC" };

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
  <tr><td style="padding:16px 24px;border-top:1px solid ${EMAIL_COLORS.border};font-size:11px;color:#7A8585;text-align:center">
    NewShire Property Management &mdash; This is an automated message from NewShire University
  </td></tr>
</table>
</td></tr></table></body></html>`;
}

async function sendEmail(token, to, subject, bodyHtml) {
  if (EMAIL_PAUSED) { console.log(`[EMAIL PAUSED] Suppressed: "${subject}" to ${Array.isArray(to) ? to.join(", ") : to}`); return; }
  const toRecipients = (Array.isArray(to) ? to : [to]).map(email => ({ emailAddress: { address: email } }));
  const res = await fetch(`${GRAPH_BASE}/me/sendMail`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({
      message: {
        subject: `[NewShire University] ${subject}`,
        body: { contentType: "HTML", content: emailTemplate(bodyHtml, subject) },
        toRecipients,
      },
      saveToSentItems: false,
    }),
  });
  if (!res.ok && res.status !== 202) {
    console.error(`sendMail failed (${res.status}):`, await res.text().catch(() => ""));
  }
}

// ============================================================
// NOTIFICATION LOG — dedup prevents repeat sends
// ============================================================
async function getNotificationsSentToday(token) {
  const today = new Date().toISOString().split("T")[0];
  try {
    const items = await spGet(token, CONFIG.lists.notifications, {
      filter: `fields/SentDate ge '${today}T00:00:00Z'`,
    });
    return items.map(i => i.fields.NotificationKey || "");
  } catch { return []; }
}

async function logNotification(token, key) {
  try {
    await spCreate(token, CONFIG.lists.notifications, {
      Title: key,
      NotificationKey: key,
      SentDate: new Date().toISOString(),
    });
  } catch (e) { console.error("Failed to log notification:", e); }
}

// ============================================================
// DUE DATE UTILITIES
// ============================================================
// Per-course due date: max(hireDate, courseCreatedDate) + path.dueDays
// A course that went active after the employee was hired gets its own deadline
function getCourseDueDate(course, path, employee) {
  if (!path.dueDays || !employee.hireDate) return null;
  const hire = new Date(employee.hireDate);
  const courseCreated = course.createdDate ? new Date(course.createdDate) : hire;
  const baseline = hire > courseCreated ? hire : courseCreated;
  const due = new Date(baseline.getTime() + path.dueDays * 86400000);
  return due.toISOString().split("T")[0];
}

// Path-level due date: the latest individual course due date in the path
// This gives a realistic "when should all courses be done" answer
function getPathDueDate(path, employee, courses) {
  if (!path.dueDays || !employee.hireDate) return null;
  const pathCourses = (path.courseIds || []).map(id => courses.find(c => c.id === id)).filter(Boolean);
  if (pathCourses.length === 0) return null;
  let latestDue = null;
  for (const course of pathCourses) {
    const d = getCourseDueDate(course, path, employee);
    if (d && (!latestDue || d > latestDue)) latestDue = d;
  }
  return latestDue;
}

function getPathDueStatus(path, employee, completions, courses, learningPaths) {
  const today = new Date().toISOString().split("T")[0];
  if (!path.dueDays || !employee.hireDate) return { dueDate: null, status: null };

  // Exempt roles have no deadlines
  if (isTrainingExempt(employee)) return { dueDate: null, status: null };

  // Only count Active courses that match this employee's role
  const applicableCourseIds = (path.courseIds || []).filter(cid => {
    const course = courses.find(c => c.id === cid);
    return course && course.status === "Active" && courseMatchesRole(course, employee.role);
  });

  // No active courses in this path yet — nothing to be overdue on
  if (applicableCourseIds.length === 0) return { dueDate: null, status: null };

  const progress = getPathProgress(path.id, employee.id, completions, courses, learningPaths, employee.role);
  const pathDueDate = getPathDueDate(path, employee, courses);
  if (!pathDueDate) return { dueDate: null, status: null };

  // Complete — no deadline concerns
  if (progress.pct >= 100) return { dueDate: pathDueDate, status: "complete" };

  // Check if any INCOMPLETE Active courses are actually past their individual due dates
  let hasOverdueCourse = false;
  let earliestOverdueCourseDate = null;

  for (const cid of applicableCourseIds) {
    const comp = completions.filter(c => c.employeeId === employee.id && c.courseId === cid && c.status === "passed");
    const course = courses.find(c => c.id === cid);
    if (comp.length > 0) {
      const latest = comp.sort((a, b) => b.completedDate.localeCompare(a.completedDate))[0];
      if (getCertStatus(latest, course) !== "expired") continue;
    }
    const courseDue = getCourseDueDate(course, path, employee);
    if (courseDue && courseDue < today) {
      hasOverdueCourse = true;
      if (!earliestOverdueCourseDate || courseDue < earliestOverdueCourseDate) earliestOverdueCourseDate = courseDue;
    }
  }

  if (hasOverdueCourse) return { dueDate: earliestOverdueCourseDate, status: "overdue" };

  // Not overdue — check if due soon (any incomplete course due within 7 days)
  for (const cid of applicableCourseIds) {
    const comp = completions.filter(c => c.employeeId === employee.id && c.courseId === cid && c.status === "passed");
    const course = courses.find(c => c.id === cid);
    if (comp.length > 0) {
      const latest = comp.sort((a, b) => b.completedDate.localeCompare(a.completedDate))[0];
      if (getCertStatus(latest, course) !== "expired") continue;
    }
    const courseDue = getCourseDueDate(course, path, employee);
    if (courseDue) {
      const daysLeft = Math.round((new Date(courseDue) - new Date(today)) / 86400000);
      if (daysLeft <= 7) return { dueDate: courseDue, status: "due-soon" };
    }
  }

  return { dueDate: pathDueDate, status: "on-track" };
}

// ============================================================
// EVENT-DRIVEN EMAILS — fire from app handlers
// ============================================================
async function sendQuizResultEmail(token, employee, course, score, passed, adminEmails) {
  if (passed) {
    // Email employee: congratulations + cert info
    const certLine = course.recertDays
      ? `<p>Your certification is valid for <strong>${course.recertDays} days</strong>. You will receive a reminder before it expires.</p>`
      : "";
    await sendEmail(token, employee.email,
      `Course Passed: ${course.name}`,
      `<p>Hi ${employee.name.split(" ")[0]},</p>
       <p>Congratulations! You passed <strong>${course.name}</strong> with a score of <strong>${score}%</strong>.</p>
       ${certLine}
       <p style="margin-top:16px;font-size:13px;color:#7A8585">Keep up the great work.</p>`
    );
  } else {
    // Email admin(s): failure notification
    for (const adminEmail of adminEmails) {
      await sendEmail(token, adminEmail,
        `Quiz Failed: ${employee.name} — ${course.name}`,
        `<p><strong>${employee.name}</strong> (${employee.role}) did not pass the quiz for <strong>${course.name}</strong>.</p>
         <table style="border-collapse:collapse;margin:12px 0">
           <tr><td style="padding:6px 16px 6px 0;color:#7A8585;font-size:13px">Score</td><td style="padding:6px 0;font-weight:600;color:#C44B3B">${score}%</td></tr>
           <tr><td style="padding:6px 16px 6px 0;color:#7A8585;font-size:13px">Required</td><td style="padding:6px 0;font-weight:600">${CONFIG.passingScore}%</td></tr>
         </table>
         <p style="font-size:13px;color:#7A8585">The employee can retake the quiz after reviewing the course material.</p>`
      );
    }
  }
}

async function sendEnrollmentEmail(token, employee, course) {
  await sendEmail(token, employee.email,
    `Enrolled: ${course.name}`,
    `<p>Hi ${employee.name.split(" ")[0]},</p>
     <p>You have been enrolled in <strong>${course.name}</strong>.</p>
     <p style="font-size:13px;color:#7A8585">${course.category} &middot; ${course.durationMin} minutes</p>
     <p>Log in to NewShire University to begin the course.</p>`
  );
}

// ============================================================
// CERT EXPIRATION SCANNER — runs on admin login
// ============================================================
async function runCertExpirationScan(token, employees, completions, courses, adminEmails) {
  const today = new Date().toISOString().split("T")[0];
  const sentKeys = await getNotificationsSentToday(token);
  let sent = 0;

  for (const emp of employees.filter(e => e.active)) {
    for (const course of courses) {
      if (!course.recertDays) continue;
      const passing = completions.filter(c => c.employeeId === emp.id && c.courseId === course.id && c.status === "passed");
      if (passing.length === 0) continue;
      const latest = passing.sort((a, b) => b.completedDate.localeCompare(a.completedDate))[0];
      if (!latest.certExpires) continue;

      const daysLeft = Math.round((new Date(latest.certExpires) - new Date(today)) / 86400000);
      let tier = null;
      if (daysLeft === 30) tier = "30day";
      else if (daysLeft === 14) tier = "14day";
      else if (daysLeft === 0) tier = "today";
      else if (daysLeft === -7) tier = "7past";
      if (!tier) continue;

      const key = `cert_${emp.id}_${course.id}_${tier}_${today}`;
      if (sentKeys.includes(key)) continue;

      if (tier === "7past") {
        // Escalation: admin only, high priority
        for (const ae of adminEmails) {
          await sendEmail(token, ae,
            `OVERDUE: ${emp.name} — ${course.name} certification expired`,
            `<p style="color:#C44B3B;font-weight:600">Certification has been expired for 7+ days.</p>
             <p><strong>${emp.name}</strong> (${emp.role}) — <strong>${course.name}</strong></p>
             <p>Expired: <strong>${latest.certExpires}</strong></p>
             <p style="font-size:13px;color:#7A8585">Please follow up directly to ensure recertification is completed.</p>`
          );
        }
      } else {
        const urgency = tier === "today" ? "expires today" : tier === "14day" ? "expires in 14 days" : "expires in 30 days";
        // Email employee
        await sendEmail(token, emp.email,
          `Certification ${urgency}: ${course.name}`,
          `<p>Hi ${emp.name.split(" ")[0]},</p>
           <p>Your certification for <strong>${course.name}</strong> <strong>${urgency}</strong> (${latest.certExpires}).</p>
           <p>Log in to NewShire University to recertify by retaking the course quiz.</p>`
        );
        // Email admin
        for (const ae of adminEmails) {
          await sendEmail(token, ae,
            `Cert ${urgency}: ${emp.name} — ${course.name}`,
            `<p><strong>${emp.name}</strong> (${emp.role}) — <strong>${course.name}</strong> certification ${urgency}.</p>
             <p>Expiration date: <strong>${latest.certExpires}</strong></p>`
          );
        }
      }
      await logNotification(token, key);
      sent++;
    }
  }
  return sent;
}

// ============================================================
// MONDAY MANAGER REPORT — runs on Monday admin login
// ============================================================
async function runMondayManagerReport(token, employees, completions, courses, learningPaths) {
  const today = new Date();
  if (today.getDay() !== 1) return 0; // Monday only
  const todayStr = today.toISOString().split("T")[0];
  const sentKeys = await getNotificationsSentToday(token);
  const weekKey = `monday_report_${todayStr}`;
  if (sentKeys.includes(weekKey)) return 0;

  // Find all managers (anyone with direct reports)
  const managers = employees.filter(mgr => {
    if (!mgr.active) return false;
    return employees.some(e => e.active && e.id !== mgr.id && (e.reportsTo === mgr.id || e.reportsTo === mgr.email));
  });

  let sent = 0;
  for (const mgr of managers) {
    const reports = employees.filter(e =>
      e.active && e.id !== mgr.id && (e.reportsTo === mgr.id || e.reportsTo === mgr.email)
    );
    if (reports.length === 0) continue;

    const issues = [];
    for (const emp of reports) {
      const paths = getEmployeePaths(emp, learningPaths);
      for (const path of paths) {
        if (!path.required) continue;
        // Check due date
        const { dueDate, status: dueStatus } = getPathDueStatus(path, emp, completions, courses, learningPaths);
        if (dueStatus === "overdue") {
          issues.push({ emp, type: "overdue", detail: `${path.name} was due ${dueDate}` });
        } else if (dueStatus === "due-soon") {
          issues.push({ emp, type: "due-soon", detail: `${path.name} due ${dueDate}` });
        }
        // Check cert expirations
        for (const cid of path.courseIds) {
          const course = courses.find(c => c.id === cid);
          if (!course || !course.recertDays) continue;
          if (!courseMatchesRole(course, emp.role)) continue;
          const passing = completions.filter(c => c.employeeId === emp.id && c.courseId === cid && c.status === "passed");
          if (passing.length === 0) continue;
          const latest = passing.sort((a, b) => b.completedDate.localeCompare(a.completedDate))[0];
          if (!latest.certExpires) continue;
          const daysLeft = Math.round((new Date(latest.certExpires) - new Date(todayStr)) / 86400000);
          if (daysLeft < 0) {
            issues.push({ emp, type: "expired", detail: `${course.name} cert expired ${latest.certExpires}` });
          } else if (daysLeft <= 30) {
            issues.push({ emp, type: "expiring", detail: `${course.name} cert expires ${latest.certExpires} (${daysLeft}d)` });
          }
        }
      }
    }

    if (issues.length === 0) continue; // No news is good news — skip clean managers

    const colorMap = { overdue: "#C44B3B", expired: "#C44B3B", "due-soon": "#D4960A", expiring: "#D4960A" };
    const labelMap = { overdue: "OVERDUE", expired: "EXPIRED", "due-soon": "DUE SOON", expiring: "EXPIRING" };
    const rows = issues.map(i =>
      `<tr>
        <td style="padding:8px 12px;border-bottom:1px solid #E8EAEA;font-size:13px">${i.emp.name}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #E8EAEA;font-size:13px">
          <span style="display:inline-block;padding:2px 8px;border-radius:9999px;font-size:11px;font-weight:600;color:${colorMap[i.type]};background:${colorMap[i.type]}15">${labelMap[i.type]}</span>
        </td>
        <td style="padding:8px 12px;border-bottom:1px solid #E8EAEA;font-size:13px">${i.detail}</td>
      </tr>`
    ).join("");

    await sendEmail(token, mgr.email,
      `Weekly Compliance Report — ${reports.length} Direct Reports`,
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
       <p style="font-size:13px;color:#7A8585">${issues.length} item${issues.length > 1 ? "s" : ""} requiring attention this week.</p>`
    );
    sent++;
  }

  if (sent > 0) await logNotification(token, weekKey);
  return sent;
}

// ============================================================
// BRAND PALETTE — NewShire Light Theme
// ============================================================
const C = {
  headerBg: "#1C3740", headerHover: "#213F4A",
  teal700: "#28434C", teal600: "#2F5260", teal500: "#3A6577", teal400: "#4A7E91",
  teal100: "#D6E7EC", teal50: "#EDF4F7",
  gold700: "#9E7B2F", gold600: "#B8922E", gold500: "#CDA04B", gold400: "#D4AF61",
  gold100: "#F8F0DB", gold50: "#FFFBF0",
  white: "#FFFFFF", pageBg: "#F7F8F7",
  gray100: "#E8EAEA", gray200: "#D0D8DC", gray300: "#A8B0B0", gray400: "#7A8585",
  gray600: "#3E4A4A", dark: "#1A2A30",
  success: "#2D8A5A", successBg: "rgba(45,138,90,0.08)", successBdr: "rgba(45,138,90,0.25)",
  error: "#C44B3B", errorBg: "rgba(196,75,59,0.06)", errorBdr: "rgba(196,75,59,0.25)",
  warning: "#D4960A", warningBg: "rgba(212,150,10,0.08)", warningBdr: "rgba(212,150,10,0.25)",
  info: "#4A78B0", infoBg: "rgba(74,120,176,0.08)", infoBdr: "rgba(74,120,176,0.25)",
};

const font = "'Source Sans 3', 'Segoe UI', -apple-system, sans-serif";
const mono = "'Source Code Pro', Consolas, monospace";

// ============================================================
// STYLES
// ============================================================
const S = {
  page: { fontFamily: font, background: C.pageBg, minHeight: "100vh", color: C.teal700 },
  header: { background: C.headerBg, borderBottom: `2px solid ${C.gold500}`, padding: "0 20px", display: "flex", alignItems: "center", justifyContent: "space-between", height: 56 },
  headerTitle: { color: "#FFFFFF", fontSize: 16, fontWeight: 700, letterSpacing: "0.05em" },
  headerSubtitle: { color: C.gold500, fontSize: 11, letterSpacing: "0.08em", textTransform: "uppercase" },
  headerUser: { color: C.teal100, fontSize: 13, display: "flex", alignItems: "center", gap: 10 },
  tabBar: { background: C.white, borderBottom: `1px solid ${C.gray200}`, display: "flex", gap: 0, padding: "0 20px", overflowX: "auto" },
  tab: (active) => ({ padding: "12px 20px", fontSize: 14, fontWeight: active ? 600 : 400, color: active ? C.teal700 : C.gray400, borderBottom: active ? `2px solid ${C.gold500}` : "2px solid transparent", cursor: "pointer", whiteSpace: "nowrap", background: "none", border: "none", fontFamily: font }),
  content: { maxWidth: 1200, margin: "0 auto", padding: "24px 20px" },
  card: { background: C.white, border: `1px solid ${C.gray200}`, borderRadius: 6, boxShadow: "0 1px 3px rgba(28,55,64,0.06)", padding: 20, marginBottom: 16 },
  cardTitle: { fontSize: 18, fontWeight: 600, color: C.teal700, paddingBottom: 12, borderBottom: `1px solid ${C.gray100}`, marginBottom: 16 },
  label: { display: "block", fontSize: 14, fontWeight: 500, color: C.teal700, marginBottom: 4 },
  input: { width: "100%", padding: "10px 12px", fontSize: 15, fontFamily: font, color: C.teal700, background: C.white, border: `1px solid ${C.gray200}`, borderRadius: 4, outline: "none", boxSizing: "border-box" },
  select: { width: "100%", padding: "10px 12px", fontSize: 15, fontFamily: font, color: C.teal700, background: C.white, border: `1px solid ${C.gray200}`, borderRadius: 4, cursor: "pointer", boxSizing: "border-box" },
  btnPrimary: { display: "inline-flex", alignItems: "center", gap: 8, padding: "10px 20px", fontSize: 14, fontWeight: 600, fontFamily: font, color: "#FFFFFF", background: C.headerBg, border: "none", borderRadius: 4, cursor: "pointer" },
  btnSecondary: { display: "inline-flex", alignItems: "center", gap: 8, padding: "10px 20px", fontSize: 14, fontWeight: 600, fontFamily: font, color: C.teal700, background: C.white, border: `1px solid ${C.teal100}`, borderRadius: 4, cursor: "pointer" },
  btnDanger: { display: "inline-flex", alignItems: "center", gap: 8, padding: "10px 20px", fontSize: 14, fontWeight: 600, fontFamily: font, color: "#FFFFFF", background: C.error, border: "none", borderRadius: 4, cursor: "pointer" },
  btnSmall: { padding: "6px 14px", fontSize: 13 },
  th: { textAlign: "left", padding: "10px 12px", fontSize: 13, fontWeight: 500, color: C.teal700, background: C.teal50, borderBottom: `2px solid ${C.teal100}` },
  td: { padding: "10px 12px", fontSize: 14, color: C.gray600, borderBottom: `1px solid ${C.gray100}` },
  badge: (type) => {
    const map = { success: { color: C.success, bg: C.successBg }, error: { color: C.error, bg: C.errorBg }, warning: { color: C.warning, bg: C.warningBg }, info: { color: C.info, bg: C.infoBg }, neutral: { color: C.gray400, bg: C.gray100 } };
    const m = map[type] || map.neutral;
    return { display: "inline-flex", alignItems: "center", padding: "2px 10px", fontSize: 12, fontWeight: 600, borderRadius: 9999, textTransform: "uppercase", letterSpacing: "0.03em", color: m.color, background: m.bg };
  },
  kpiCard: { background: C.white, border: `1px solid ${C.gray200}`, borderRadius: 6, padding: 16, textAlign: "center", flex: 1, minWidth: 140 },
  kpiLabel: { fontSize: 12, fontWeight: 500, color: C.gray400, textTransform: "uppercase", letterSpacing: "0.05em" },
  kpiValue: { fontSize: 30, fontWeight: 700, color: C.teal700, fontFamily: mono },
  row: { display: "flex", gap: 16, flexWrap: "wrap" },
  progressBar: (pct, color = C.success) => ({
    height: 8, borderRadius: 4, background: C.gray100, position: "relative", overflow: "hidden",
    _fill: { position: "absolute", top: 0, left: 0, height: "100%", width: `${Math.min(100, pct)}%`, background: color, borderRadius: 4, transition: "width 0.4s ease" },
  }),
};

// ============================================================
// DEMO DATA
// ============================================================
// ============================================================
// LEARNING PATHS (8 total)
// ============================================================
const DEMO_LEARNING_PATHS = [
  { id: "lp1", name: "New Hire Onboarding", description: "Required for all new employees within 30 days of hire. Covers NewShire orientation, industry basics, Fair Housing, harassment prevention, safety, communication, and technology.", roles: ["All"], courseIds: ["c1","c2","c3","c4","c5","c6","c7","c8"], required: true, dueDays: 30 },
  { id: "lp2", name: "Fair Housing Annual Recertification", description: "Annual recertification required for all staff per HUD compliance.", roles: ["All"], courseIds: ["c3","c4"], required: true, recertDays: 365 },
  { id: "lp3", name: "Leasing Certification", description: "Required for leasing agents and property managers. Covers leasing workflows, screening, and resident onboarding.", roles: ["Leasing Agent", "Property Manager"], courseIds: ["c9","c10"], required: true, dueDays: 60 },
  { id: "lp4", name: "Maintenance Technician Track", description: "Required for maintenance staff. Covers fundamentals, preventive maintenance, and customer service.", roles: ["Maintenance Tech", "Maintenance Supervisor"], courseIds: ["c11","c12","c13"], required: true, dueDays: 60 },
  { id: "lp5", name: "Safety and Environmental Compliance", description: "Required for maintenance staff and property managers. Covers mold, lead paint, and environmental hazards.", roles: ["Maintenance Tech", "Maintenance Supervisor", "Property Manager"], courseIds: ["c14","c15","c16"], required: true, dueDays: 90 },
  { id: "lp6", name: "Inspections and Quality Control", description: "Required for maintenance supervisors and property managers. Covers inspection standards and risk management.", roles: ["Maintenance Supervisor", "Property Manager"], courseIds: ["c17","c18"], required: true, dueDays: 90 },
  { id: "lp7", name: "Supervisory and Financial Track", description: "Required for supervisors and managers. Covers financials, budgets, and performance metrics.", roles: ["Maintenance Supervisor", "Property Manager", "Owner/Operator"], courseIds: ["c19","c20"], required: true, dueDays: 90 },
  { id: "lp8", name: "AMI and Income-Restricted Housing", description: "Income verification using the NewShire AMI Calculator, HUD compliance, and annual recertification workflows.", roles: ["Property Manager", "Leasing Agent"], courseIds: ["c21"], required: true, recertDays: 365, dueDays: 60 },
];

// ============================================================
// COURSES (21 total)
// ============================================================
const DEMO_COURSES = [
  // ── Core / Onboarding (c1-c8: All employees) ──
  { id: "c1", name: "Welcome to NewShire", description: "Introduces NewShire's mission, values, expectations, and compliance first culture.", category: "Onboarding", durationMin: 20, lessonIds: ["l1","l2"] },
  { id: "c2", name: "Property Management 101", description: "Explains how the property management industry works, key roles, and basic terminology.", category: "Onboarding", durationMin: 25, lessonIds: ["l3","l4"] },
  { id: "c3", name: "Fair Housing Law and Compliance", description: "Covers federal Fair Housing laws, protected classes, and prohibited practices.", category: "Compliance", durationMin: 45, recertDays: 365, lessonIds: ["l5","l6","l7"] },
  { id: "c4", name: "Fair Housing in Daily Operations", description: "Applies Fair Housing principles to real world leasing, maintenance, and resident interactions.", category: "Compliance", durationMin: 35, recertDays: 365, lessonIds: ["l8","l9","l10"] },
  { id: "c5", name: "Harassment Prevention and Hostile Work Environment", description: "Defines harassment, reinforces zero tolerance for hostile behavior from anyone, and explains reporting obligations.", category: "Compliance", durationMin: 30, recertDays: 365, lessonIds: ["l11","l12","l13"] },
  { id: "c6", name: "Personal Safety and Situational Awareness", description: "Teaches employees how to recognize unsafe situations, set boundaries, and prioritize personal safety.", category: "Compliance", durationMin: 25, lessonIds: ["l14","l15"] },
  { id: "c7", name: "Communication Etiquette and Professional Conduct", description: "Sets standards for professional communication, documentation, and respectful interactions.", category: "Operations", durationMin: 25, lessonIds: ["l16","l17"] },
  { id: "c8", name: "Technology Basics for NewShire Operations", description: "Introduces core systems, data accuracy expectations, and cybersecurity awareness.", category: "Systems", durationMin: 30, lessonIds: ["l18","l19"] },
  // ── Leasing Track (c9-c10) ──
  { id: "c9", name: "Leasing and Resident Onboarding", description: "Required for leasing and office staff. Covers consistent leasing, screening, and move in processes.", category: "Leasing", durationMin: 40, lessonIds: ["l20","l21"] },
  { id: "c10", name: "Leasing Fundamentals for New Professionals", description: "Builds foundational leasing skills for employees new to the industry.", category: "Leasing", durationMin: 35, lessonIds: ["l22","l23"] },
  // ── Maintenance Track (c11-c13) ──
  { id: "c11", name: "Maintenance Fundamentals for Multifamily and Single Family Housing", description: "Required for maintenance staff. Covers roles, work orders, safety, and compliance basics.", category: "Maintenance", durationMin: 40, lessonIds: ["l24","l25","l26"] },
  { id: "c12", name: "Preventive Maintenance and Inspections", description: "Required for maintenance staff. Focuses on inspections, preventive care, and documentation.", category: "Maintenance", durationMin: 35, lessonIds: ["l27","l28"] },
  { id: "c13", name: "Maintenance Customer Service and Communication", description: "Required for maintenance staff. Sets expectations for resident interactions and professionalism.", category: "Maintenance", durationMin: 25, lessonIds: ["l29","l30"] },
  // ── Safety & Environmental (c14-c16) ──
  { id: "c14", name: "Mold and Mildew Awareness and Response", description: "Required for maintenance and managers. Covers moisture issues, response protocols, and documentation.", category: "Compliance", durationMin: 30, lessonIds: ["l31","l32"] },
  { id: "c15", name: "Lead Based Paint Compliance", description: "Required where applicable. Covers disclosure and handling requirements for pre 1978 properties.", category: "Compliance", durationMin: 25, lessonIds: ["l33","l34"] },
  { id: "c16", name: "Health and Environmental Hazards", description: "Covers asbestos awareness, pest control coordination, and environmental safety basics.", category: "Compliance", durationMin: 30, lessonIds: ["l35","l36","l37"] },
  // ── Supervisory / Management (c17-c20) ──
  { id: "c17", name: "Property Inspections and Quality Control", description: "Required for maintenance and managers. Establishes inspection standards and documentation requirements.", category: "Operations", durationMin: 35, lessonIds: ["l38","l39"] },
  { id: "c18", name: "Risk Management and Site Safety", description: "Required for supervisors and managers. Covers hazard mitigation and incident reporting.", category: "Operations", durationMin: 30, lessonIds: ["l40","l41"] },
  { id: "c19", name: "Financial Foundations for Property Management", description: "Required for supervisors and managers. Introduces budgets, expenses, and financial responsibility.", category: "Operations", durationMin: 35, lessonIds: ["l42","l43"] },
  { id: "c20", name: "Understanding Property Performance Metrics", description: "Required for supervisors and managers. Explains occupancy, delinquency, and performance indicators.", category: "Operations", durationMin: 30, lessonIds: ["l44","l45"] },
  // ── AMI (c21) ──
  { id: "c21", name: "AMI and Income-Restricted Housing Compliance", description: "Covers Area Median Income calculations using the NewShire AMI Calculator, income verification, household composition, documentation requirements, and annual recertification workflows.", category: "Compliance", durationMin: 50, recertDays: 365, lessonIds: ["l46","l47","l48"] },
];

// ============================================================
// LESSONS (49 total — 2-3 per course)
// ============================================================
const DEMO_LESSONS = [
  // C1: Welcome to NewShire
  { id: "l1", courseId: "c1", order: 1, title: "Our Mission, Values, and Culture", type: "video", videoUrl: "", durationMin: 10, hasDocument: true, docTitle: "NewShire Employee Handbook.pdf" },
  { id: "l2", courseId: "c1", order: 2, title: "What Compliance First Means at NewShire", type: "video", videoUrl: "", durationMin: 10, hasDocument: false },
  // C2: Property Management 101
  { id: "l3", courseId: "c2", order: 1, title: "How Property Management Works", type: "video", videoUrl: "", durationMin: 12, hasDocument: false },
  { id: "l4", courseId: "c2", order: 2, title: "Key Roles and Industry Terminology", type: "video", videoUrl: "", durationMin: 13, hasDocument: false },
  // C3: Fair Housing Law and Compliance
  { id: "l5", courseId: "c3", order: 1, title: "Protected Classes Under Federal Law", type: "video", videoUrl: "", durationMin: 15, hasDocument: true, docTitle: "FHA Protected Classes Quick Reference.pdf" },
  { id: "l6", courseId: "c3", order: 2, title: "Prohibited Practices and Advertising Compliance", type: "video", videoUrl: "", durationMin: 15, hasDocument: false },
  { id: "l7", courseId: "c3", order: 3, title: "Reasonable Accommodations and Modifications", type: "video", videoUrl: "", durationMin: 15, hasDocument: true, docTitle: "Reasonable Accommodation Request Form.pdf" },
  // C4: Fair Housing in Daily Operations
  { id: "l8", courseId: "c4", order: 1, title: "Fair Housing During Showings and Leasing", type: "video", videoUrl: "", durationMin: 12, hasDocument: false },
  { id: "l9", courseId: "c4", order: 2, title: "Fair Housing in Maintenance and Resident Interactions", type: "video", videoUrl: "", durationMin: 12, hasDocument: false },
  { id: "l10", courseId: "c4", order: 3, title: "Recognizing and Reporting Violations", type: "video", videoUrl: "", durationMin: 11, hasDocument: false },
  // C5: Harassment Prevention
  { id: "l11", courseId: "c5", order: 1, title: "Defining Harassment and Hostile Work Environment", type: "video", videoUrl: "", durationMin: 10, hasDocument: false },
  { id: "l12", courseId: "c5", order: 2, title: "Reporting Obligations and Bystander Responsibility", type: "video", videoUrl: "", durationMin: 10, hasDocument: false },
  { id: "l13", courseId: "c5", order: 3, title: "Zero Tolerance from Residents, Vendors, and Coworkers", type: "video", videoUrl: "", durationMin: 10, hasDocument: false },
  // C6: Personal Safety
  { id: "l14", courseId: "c6", order: 1, title: "Recognizing Unsafe Situations in the Field", type: "video", videoUrl: "", durationMin: 12, hasDocument: false },
  { id: "l15", courseId: "c6", order: 2, title: "Setting Boundaries and De-escalation", type: "video", videoUrl: "", durationMin: 13, hasDocument: false },
  // C7: Communication Etiquette
  { id: "l16", courseId: "c7", order: 1, title: "Professional Communication Standards", type: "video", videoUrl: "", durationMin: 12, hasDocument: true, docTitle: "Communication Templates.pdf" },
  { id: "l17", courseId: "c7", order: 2, title: "Documentation and Respectful Interactions", type: "video", videoUrl: "", durationMin: 13, hasDocument: false },
  // C8: Technology Basics
  { id: "l18", courseId: "c8", order: 1, title: "Core Systems and Data Accuracy", type: "video", videoUrl: "", durationMin: 15, hasDocument: false },
  { id: "l19", courseId: "c8", order: 2, title: "Cybersecurity Awareness for Property Management", type: "video", videoUrl: "", durationMin: 15, hasDocument: false },
  // C9: Leasing and Resident Onboarding
  { id: "l20", courseId: "c9", order: 1, title: "Consistent Leasing and Screening Processes", type: "video", videoUrl: "", durationMin: 20, hasDocument: false },
  { id: "l21", courseId: "c9", order: 2, title: "Move-In Procedures and Resident Onboarding", type: "video", videoUrl: "", durationMin: 20, hasDocument: true, docTitle: "Move-In Checklist.pdf" },
  // C10: Leasing Fundamentals
  { id: "l22", courseId: "c10", order: 1, title: "Lead Response and Prospect Follow-Up", type: "video", videoUrl: "", durationMin: 17, hasDocument: true, docTitle: "Follow-Up Email Templates.pdf" },
  { id: "l23", courseId: "c10", order: 2, title: "Application Processing and Lease Execution", type: "video", videoUrl: "", durationMin: 18, hasDocument: false },
  // C11: Maintenance Fundamentals
  { id: "l24", courseId: "c11", order: 1, title: "Maintenance Roles and Responsibilities", type: "video", videoUrl: "", durationMin: 12, hasDocument: false },
  { id: "l25", courseId: "c11", order: 2, title: "Work Order Lifecycle and Safety Basics", type: "video", videoUrl: "", durationMin: 15, hasDocument: false },
  { id: "l26", courseId: "c11", order: 3, title: "Compliance and Documentation Requirements", type: "video", videoUrl: "", durationMin: 13, hasDocument: false },
  // C12: Preventive Maintenance
  { id: "l27", courseId: "c12", order: 1, title: "Inspection Types and Schedules", type: "video", videoUrl: "", durationMin: 17, hasDocument: false },
  { id: "l28", courseId: "c12", order: 2, title: "Preventive Care Programs and Documentation", type: "video", videoUrl: "", durationMin: 18, hasDocument: true, docTitle: "Preventive Maintenance Checklist.pdf" },
  // C13: Maintenance Customer Service
  { id: "l29", courseId: "c13", order: 1, title: "Resident Interactions During Service Calls", type: "video", videoUrl: "", durationMin: 12, hasDocument: false },
  { id: "l30", courseId: "c13", order: 2, title: "Professionalism and Communication Standards for Techs", type: "video", videoUrl: "", durationMin: 13, hasDocument: false },
  // C14: Mold and Mildew
  { id: "l31", courseId: "c14", order: 1, title: "Moisture Sources and Mold Prevention", type: "video", videoUrl: "", durationMin: 15, hasDocument: false },
  { id: "l32", courseId: "c14", order: 2, title: "Response Protocols and Documentation", type: "video", videoUrl: "", durationMin: 15, hasDocument: true, docTitle: "Mold Response Protocol.pdf" },
  // C15: Lead Based Paint
  { id: "l33", courseId: "c15", order: 1, title: "EPA Disclosure Requirements for Pre-1978 Properties", type: "video", videoUrl: "", durationMin: 12, hasDocument: true, docTitle: "Lead Paint Disclosure Form.pdf" },
  { id: "l34", courseId: "c15", order: 2, title: "Handling and Renovation Requirements", type: "video", videoUrl: "", durationMin: 13, hasDocument: false },
  // C16: Health and Environmental Hazards
  { id: "l35", courseId: "c16", order: 1, title: "Asbestos Awareness and Identification", type: "video", videoUrl: "", durationMin: 10, hasDocument: false },
  { id: "l36", courseId: "c16", order: 2, title: "Pest Control Coordination and Safety", type: "video", videoUrl: "", durationMin: 10, hasDocument: false },
  { id: "l37", courseId: "c16", order: 3, title: "Environmental Compliance Basics", type: "video", videoUrl: "", durationMin: 10, hasDocument: false },
  // C17: Property Inspections and QC
  { id: "l38", courseId: "c17", order: 1, title: "Inspection Standards and Checklists", type: "video", videoUrl: "", durationMin: 17, hasDocument: true, docTitle: "Property Inspection Checklist.pdf" },
  { id: "l39", courseId: "c17", order: 2, title: "Quality Control Documentation and Follow-Up", type: "video", videoUrl: "", durationMin: 18, hasDocument: false },
  // C18: Risk Management
  { id: "l40", courseId: "c18", order: 1, title: "Hazard Identification and Mitigation", type: "video", videoUrl: "", durationMin: 15, hasDocument: false },
  { id: "l41", courseId: "c18", order: 2, title: "Incident Reporting and Documentation", type: "video", videoUrl: "", durationMin: 15, hasDocument: true, docTitle: "Incident Report Template.pdf" },
  // C19: Financial Foundations
  { id: "l42", courseId: "c19", order: 1, title: "Budgets, Expenses, and Financial Responsibility", type: "video", videoUrl: "", durationMin: 17, hasDocument: false },
  { id: "l43", courseId: "c19", order: 2, title: "Owner Reporting and Financial Documentation", type: "video", videoUrl: "", durationMin: 18, hasDocument: false },
  // C20: Performance Metrics
  { id: "l44", courseId: "c20", order: 1, title: "Occupancy and Delinquency Tracking", type: "video", videoUrl: "", durationMin: 15, hasDocument: false },
  { id: "l45", courseId: "c20", order: 2, title: "Performance Indicators and Benchmarking", type: "video", videoUrl: "", durationMin: 15, hasDocument: false },
  // C21: AMI and Income-Restricted Housing
  { id: "l46", courseId: "c21", order: 1, title: "Area Median Income Basics and the NewShire AMI Calculator", type: "video", videoUrl: "", durationMin: 15, hasDocument: false },
  { id: "l47", courseId: "c21", order: 2, title: "Income Verification and Documentation Requirements", type: "video", videoUrl: "", durationMin: 15, hasDocument: true, docTitle: "Income Verification Checklist.pdf" },
  { id: "l48", courseId: "c21", order: 3, title: "Annual Recertification Workflows", type: "video", videoUrl: "", durationMin: 20, hasDocument: false },
];

// ============================================================
// QUIZZES (5 questions per course = 105 total)
// ============================================================
const DEMO_QUIZZES = {
  c1: { courseId: "c1", passingScore: 80, questions: [
    { id: "c1q1", text: "What is the foundation of NewShire's operational culture?", options: ["Revenue growth at all costs", "Compliance first", "Speed over accuracy", "Owner preference above all else"], correct: 1 },
    { id: "c1q2", text: "NewShire Property Management operates as a:", options: ["Franchise model open to anyone", "Referral-only property management company", "Real estate brokerage", "Government housing authority"], correct: 1 },
    { id: "c1q3", text: "Which markets does NewShire primarily serve?", options: ["Atlanta and Savannah", "Greenville-Spartanburg and Charlotte", "Columbia and Charleston", "Raleigh and Durham"], correct: 1 },
    { id: "c1q4", text: "What types of properties does NewShire manage?", options: ["Only luxury apartments", "Single-family rentals, multifamily communities, and income-restricted AMI housing", "Commercial office buildings only", "Only Section 8 housing"], correct: 1 },
    { id: "c1q5", text: "When uncertain about a process, your first priority should be:", options: ["Figure it out yourself", "Ask the tenant what they think", "Check established SOPs and escalate to your supervisor", "Skip it and move on"], correct: 2 },
  ]},
  c2: { courseId: "c2", passingScore: 80, questions: [
    { id: "c2q1", text: "The primary role of a property management company is to:", options: ["Sell real estate", "Manage properties on behalf of owners, protecting their investment and serving residents", "Provide mortgage lending", "Only collect rent"], correct: 1 },
    { id: "c2q2", text: "Which of the following is NOT a core function of property management?", options: ["Leasing and tenant screening", "Maintenance coordination", "Property appraisal for sale", "Rent collection and financial reporting"], correct: 2 },
    { id: "c2q3", text: "NOI stands for:", options: ["Net Occupancy Index", "Net Operating Income", "National Owner Insurance", "Non-Operating Investment"], correct: 1 },
    { id: "c2q4", text: "A 'turn' in property management refers to:", options: ["A lease renewal", "The process of preparing a unit for a new tenant after move-out", "A rent increase", "A maintenance emergency"], correct: 1 },
    { id: "c2q5", text: "Who is the property manager's primary client?", options: ["The tenant", "The property owner", "The city government", "The insurance company"], correct: 1 },
  ]},
  c3: { courseId: "c3", passingScore: 80, questions: [
    { id: "c3q1", text: "Which of the following is NOT a federally protected class under the Fair Housing Act?", options: ["Race", "Religion", "Marital Status", "National Origin"], correct: 2 },
    { id: "c3q2", text: "A tenant with a disability requests a support animal in a no-pets property. The correct response is:", options: ["Deny because the lease prohibits pets", "Approve as a reasonable accommodation without a pet deposit", "Charge a pet deposit", "Require veterinary certification"], correct: 1 },
    { id: "c3q3", text: "Which advertising phrase could violate fair housing law?", options: ["Spacious 2BR near park", "Perfect for young professionals", "Updated kitchen and bath", "On-site laundry available"], correct: 1 },
    { id: "c3q4", text: "How many federally protected classes exist under the Fair Housing Act?", options: ["Five", "Seven", "Nine", "Twelve"], correct: 1 },
    { id: "c3q5", text: "A prospect asks about the racial demographics of the neighborhood. You should:", options: ["Provide Census data", "Suggest they research it themselves", "Decline and redirect to available unit features", "Only share if the prospect is a minority"], correct: 2 },
  ]},
  c4: { courseId: "c4", passingScore: 80, questions: [
    { id: "c4q1", text: "When showing a unit to a family with children, which statement is a fair housing violation?", options: ["The unit has two bedrooms and one bathroom", "This unit is on the third floor, which might not be ideal for your kids", "The lease term is 12 months", "Utilities are tenant responsibility"], correct: 1 },
    { id: "c4q2", text: "A maintenance tech notices a resident using a wheelchair has trouble with the front door. The correct response is:", options: ["Ignore it — accessibility is not maintenance's job", "Report the issue to the property manager as a potential reasonable modification need", "Tell the resident to file a complaint with HUD", "Fix it only if the tenant pays for it"], correct: 1 },
    { id: "c4q3", text: "Treating some residents differently in maintenance response times based on their background is:", options: ["Acceptable if the repairs are minor", "A fair housing violation regardless of intent", "Only a problem if someone complains", "Allowed if it is based on lease terms"], correct: 1 },
    { id: "c4q4", text: "If a coworker makes a discriminatory comment about an applicant, you should:", options: ["Laugh it off", "Report it to your supervisor immediately", "Wait and see if it happens again", "Only report if the applicant hears it"], correct: 1 },
    { id: "c4q5", text: "Consistent application of leasing criteria means:", options: ["Using different standards for different prospects", "Applying the same screening standards to every applicant regardless of protected class", "Making exceptions for referrals", "Lowering standards to fill vacancies"], correct: 1 },
  ]},
  c5: { courseId: "c5", passingScore: 80, questions: [
    { id: "c5q1", text: "NewShire's harassment policy applies to behavior from:", options: ["Only coworkers", "Anyone — coworkers, residents, vendors, and visitors", "Only supervisors", "Only during work hours"], correct: 1 },
    { id: "c5q2", text: "A resident repeatedly makes inappropriate comments to a leasing agent. The correct action is:", options: ["Ignore it because the resident is a customer", "Report it to management and document each incident", "Confront the resident alone", "Avoid the resident and say nothing"], correct: 1 },
    { id: "c5q3", text: "A hostile work environment is created when:", options: ["A coworker disagrees with you in a meeting", "Unwelcome conduct based on protected characteristics is severe or pervasive enough to alter work conditions", "You receive a negative performance review", "Your manager assigns you extra tasks"], correct: 1 },
    { id: "c5q4", text: "Retaliation against someone who reports harassment is:", options: ["Acceptable if the report was unfounded", "Strictly prohibited and itself a violation", "Only prohibited if the reporter is a supervisor", "Not a real concern"], correct: 1 },
    { id: "c5q5", text: "If you witness harassment but are not the target, your obligation is to:", options: ["Mind your own business", "Report it through the established process", "Wait for the victim to report it", "Confront the harasser directly"], correct: 1 },
  ]},
  c6: { courseId: "c6", passingScore: 80, questions: [
    { id: "c6q1", text: "Before entering a vacant unit alone, you should:", options: ["Just go in — it is your job", "Notify someone of your location and expected return time", "Only worry if it is nighttime", "Bring a weapon"], correct: 1 },
    { id: "c6q2", text: "A prospect becomes aggressive during a showing. Your first priority is:", options: ["Calm them down at all costs", "Remove yourself from the situation and call for help", "Complete the showing to avoid a complaint", "Argue your position"], correct: 1 },
    { id: "c6q3", text: "Setting professional boundaries with residents means:", options: ["Being rude to maintain distance", "Keeping interactions professional and not sharing personal information", "Refusing to help with any requests", "Only communicating via email"], correct: 1 },
    { id: "c6q4", text: "When conducting property inspections, which practice increases safety?", options: ["Going alone without telling anyone", "Working in pairs when possible and keeping your phone accessible", "Inspecting only during off-hours", "Skipping vacant units"], correct: 1 },
    { id: "c6q5", text: "If you feel unsafe at any point during a work task, the correct response is:", options: ["Push through — safety concerns are overblown", "Leave the situation, report it, and do not return without support", "Finish the task first then report", "Only leave if someone else is with you"], correct: 1 },
  ]},
  c7: { courseId: "c7", passingScore: 80, questions: [
    { id: "c7q1", text: "All written tenant communications should be:", options: ["Casual and friendly with emojis", "Professional, factual, and documented in the property management system", "As brief as possible with no details", "Sent only when the tenant initiates contact"], correct: 1 },
    { id: "c7q2", text: "When a tenant sends a hostile email, your first step is:", options: ["Reply with the same energy", "Acknowledge the concern professionally, provide a timeline, and document the interaction", "Ignore it", "Forward it to the owner immediately"], correct: 1 },
    { id: "c7q3", text: "Which communication channel is preferred for documented tenant interactions?", options: ["Text message", "Phone call", "Email through the property management platform", "In-person only"], correct: 2 },
    { id: "c7q4", text: "Including personal opinions in a tenant notice is:", options: ["Fine if you are being honest", "Unprofessional and potentially a liability — stick to facts, dates, and required actions", "Required for transparency", "Only a problem in legal proceedings"], correct: 1 },
    { id: "c7q5", text: "Documentation of every tenant interaction is important because:", options: ["It creates busy work", "It protects the company legally and maintains audit trails", "Tenants expect it", "It is only needed for problem tenants"], correct: 1 },
  ]},
  c8: { courseId: "c8", passingScore: 80, questions: [
    { id: "c8q1", text: "Data accuracy in the property management platform is critical because:", options: ["It looks nice", "Financial reports, owner statements, and compliance audits all depend on accurate data", "Only the accountant uses it", "It does not really matter"], correct: 1 },
    { id: "c8q2", text: "Which is the safest password practice?", options: ["Using the same password for everything", "Using unique, complex passwords and a password manager", "Writing passwords on sticky notes", "Sharing passwords with your team for convenience"], correct: 1 },
    { id: "c8q3", text: "If you receive a suspicious email requesting login credentials, you should:", options: ["Click the link and check", "Do not click anything, report it to your supervisor, and delete the email", "Forward it to the whole team", "Reply and ask if it is real"], correct: 1 },
    { id: "c8q4", text: "When entering data into any system, which practice prevents errors?", options: ["Enter quickly to save time", "Double-check entries before saving, especially financial amounts and dates", "Let someone else verify later", "Only review if it seems wrong"], correct: 1 },
    { id: "c8q5", text: "Sharing tenant personal information outside of authorized systems is:", options: ["Fine if you trust the person", "A data privacy violation that could expose the company to liability", "Only a problem if the tenant finds out", "Acceptable over text for quick responses"], correct: 1 },
  ]},
  c9: { courseId: "c9", passingScore: 80, questions: [
    { id: "c9q1", text: "The leasing process at NewShire must be consistent because:", options: ["It is faster that way", "Inconsistency creates fair housing risk and operational errors", "Tenants prefer it", "It is only important during audits"], correct: 1 },
    { id: "c9q2", text: "Before approving any rental application, what must be verified?", options: ["Just the credit score", "Income, credit, rental history, and background per established screening criteria", "Only employment", "Social media profiles"], correct: 1 },
    { id: "c9q3", text: "A move-in inspection should be completed:", options: ["Within the first month", "Before or on the day the tenant takes possession", "After the first rent payment", "Only if the tenant requests it"], correct: 1 },
    { id: "c9q4", text: "What is the purpose of documenting pre-existing conditions at move-in?", options: ["To make the file look complete", "To protect both the tenant and the company during deposit accounting at move-out", "It is not required", "To charge the tenant later"], correct: 1 },
    { id: "c9q5", text: "The resident welcome package should include:", options: ["Just the keys", "Lease copy, move-in inspection, emergency contacts, maintenance request procedures, and community rules", "Only the rules", "Nothing — just hand them the keys"], correct: 1 },
  ]},
  c10: { courseId: "c10", passingScore: 80, questions: [
    { id: "c10q1", text: "The recommended response time for a new prospect inquiry is:", options: ["Within 24 hours", "Within 1 business hour", "By end of week", "When you have time"], correct: 1 },
    { id: "c10q2", text: "How many follow-up touches should a prospect receive after initial contact?", options: ["One and done", "At least 3 to 5 across different channels", "Follow up only if they contact you again", "10 calls per day"], correct: 1 },
    { id: "c10q3", text: "During a property showing, which statement is a fair housing violation?", options: ["The unit has two bedrooms", "This neighborhood has great schools for your kids", "Suggesting a unit based on the prospect's ethnicity", "Discussing pet policies"], correct: 2 },
    { id: "c10q4", text: "If a prospect does not meet screening criteria, the correct action is:", options: ["Deny and document the specific criteria not met", "Make an exception because they seem nice", "Ask them to apply again later", "Ignore the application"], correct: 0 },
    { id: "c10q5", text: "A lease should be reviewed with the tenant:", options: ["Only if they ask questions", "Clause by clause to ensure understanding before signing", "Just have them sign and send them a copy", "Only the first and last page"], correct: 1 },
  ]},
  c11: { courseId: "c11", passingScore: 80, questions: [
    { id: "c11q1", text: "The first step when receiving a maintenance work order is:", options: ["Drive to the property immediately", "Review the request details, classify priority, and determine if in-house or vendor dispatch is appropriate", "Call the tenant and ask if it can wait", "Forward it to the supervisor"], correct: 1 },
    { id: "c11q2", text: "Which of the following is classified as an emergency work order?", options: ["Leaky faucet", "No heat in winter, active flooding, or gas leak", "Squeaky door hinge", "Burned out light bulb"], correct: 1 },
    { id: "c11q3", text: "Before entering a tenant-occupied unit for non-emergency maintenance, you must:", options: ["Just go in with your key", "Provide proper notice as required by state law and coordinate with the tenant", "Only knock once", "Enter when the tenant is not home to avoid disruption"], correct: 1 },
    { id: "c11q4", text: "Completing a work order in the system after finishing a repair is:", options: ["Optional if you told the tenant", "Required — it closes the loop, documents the work, and maintains accurate records", "Only needed for expensive repairs", "The office staff's job, not maintenance"], correct: 1 },
    { id: "c11q5", text: "Personal protective equipment (PPE) should be used:", options: ["Only when the supervisor is watching", "Whenever the task requires it, per safety guidelines, no exceptions", "Only for major projects", "Never — it slows you down"], correct: 1 },
  ]},
  c12: { courseId: "c12", passingScore: 80, questions: [
    { id: "c12q1", text: "The purpose of preventive maintenance is:", options: ["To create extra work", "To identify and address issues before they become costly repairs or safety hazards", "Only to satisfy insurance requirements", "To keep techs busy during slow periods"], correct: 1 },
    { id: "c12q2", text: "HVAC filters in rental units should be inspected or replaced:", options: ["Only when the tenant complains", "On a regular schedule, typically every 30 to 90 days depending on the system", "Once a year", "Never — that is the tenant's responsibility"], correct: 1 },
    { id: "c12q3", text: "Documenting the condition during an inspection requires:", options: ["A verbal description to the office", "Photos, written notes, and completion of the inspection checklist in the system", "Just a text to the property manager", "Nothing if everything looks fine"], correct: 1 },
    { id: "c12q4", text: "A preventive maintenance schedule should be:", options: ["Created once and never updated", "Reviewed and adjusted seasonally based on property needs and historical data", "Left to each tech to manage individually", "Only for commercial properties"], correct: 1 },
    { id: "c12q5", text: "Which item is commonly included in a seasonal inspection checklist?", options: ["Tenant credit check", "Smoke detector testing, gutter clearance, and water heater inspection", "Lease renewal status", "Rent payment history"], correct: 1 },
  ]},
  c13: { courseId: "c13", passingScore: 80, questions: [
    { id: "c13q1", text: "When entering a tenant's unit for a scheduled repair, the tech should:", options: ["Go straight to work without speaking to the tenant", "Greet the resident, explain what work will be done, and provide an estimated timeframe", "Ask the tenant to leave while work is done", "Only communicate if there is a problem"], correct: 1 },
    { id: "c13q2", text: "A tenant is upset about a repair delay. The tech should:", options: ["Argue that it is not their fault", "Acknowledge the frustration, provide a realistic update, and remain professional", "Ignore the complaint and focus on the repair", "Tell them to call the office"], correct: 1 },
    { id: "c13q3", text: "After completing a repair, the tech should:", options: ["Leave without saying anything", "Show the resident the completed work, explain any follow-up needed, and close the work order", "Only update the system if it was a major repair", "Text the property manager and leave"], correct: 1 },
    { id: "c13q4", text: "Professionalism during a service call includes:", options: ["Playing music loudly while working", "Wearing appropriate attire, respecting the tenant's space, and cleaning up after the work", "Talking on the phone during the repair", "Using the tenant's bathroom without asking"], correct: 1 },
    { id: "c13q5", text: "If a tenant asks about an unrelated repair during your visit, you should:", options: ["Ignore it", "Note the request and advise them to submit a work order so it is documented and tracked", "Fix it immediately without documenting", "Tell them it is not your problem"], correct: 1 },
  ]},
  c14: { courseId: "c14", passingScore: 80, questions: [
    { id: "c14q1", text: "The primary cause of mold growth in residential units is:", options: ["Old paint", "Uncontrolled moisture from leaks, condensation, or poor ventilation", "Dirty carpets", "Hot weather"], correct: 1 },
    { id: "c14q2", text: "When a tenant reports visible mold, the first step is:", options: ["Tell them to clean it with bleach", "Document the report, inspect the area, identify the moisture source, and escalate per protocol", "Ignore it if it is a small area", "Schedule a repair for next month"], correct: 1 },
    { id: "c14q3", text: "Which areas are most susceptible to mold growth?", options: ["Living rooms only", "Bathrooms, kitchens, laundry areas, basements, and around windows with condensation", "Only exterior walls", "Garages"], correct: 1 },
    { id: "c14q4", text: "Documentation for a mold complaint should include:", options: ["Just a note in the file", "Photos, date of report, inspection findings, moisture source, remediation steps, and tenant communication", "Only if the tenant threatens legal action", "Nothing if you fix it quickly"], correct: 1 },
    { id: "c14q5", text: "Maintenance staff should NOT attempt to remediate mold:", options: ["Ever", "When the affected area exceeds 10 square feet or involves HVAC systems — professional remediation is required", "Only if it is black mold", "Unless they have gloves"], correct: 1 },
  ]},
  c15: { courseId: "c15", passingScore: 80, questions: [
    { id: "c15q1", text: "Lead based paint disclosure is required for properties built before:", options: ["1990", "1978", "1985", "2000"], correct: 1 },
    { id: "c15q2", text: "The EPA pamphlet 'Protect Your Family From Lead in Your Home' must be provided to:", options: ["Only families with children", "All tenants and buyers of pre-1978 housing before signing a lease or purchase agreement", "Only if lead is confirmed present", "Only in certain states"], correct: 1 },
    { id: "c15q3", text: "Renovation work that disturbs lead paint in a pre-1978 unit requires:", options: ["No special precautions", "Compliance with EPA's RRP Rule, including certified renovators and lead-safe work practices", "Only a dust mask", "Just opening the windows"], correct: 1 },
    { id: "c15q4", text: "Failure to provide lead paint disclosure can result in:", options: ["A verbal warning", "Penalties up to $19,507 per violation, treble damages, and tenant right to void the lease", "No consequences", "A small fine only"], correct: 1 },
    { id: "c15q5", text: "The signed lead paint disclosure form should be kept:", options: ["For one year", "For at least three years as required by federal law", "Only during the lease term", "It does not need to be kept"], correct: 1 },
  ]},
  c16: { courseId: "c16", passingScore: 80, questions: [
    { id: "c16q1", text: "If suspected asbestos-containing material is found during a renovation, you should:", options: ["Remove it yourself", "Stop work immediately, do not disturb it, and notify management for professional assessment", "Cover it with new material", "It is only a concern in commercial buildings"], correct: 1 },
    { id: "c16q2", text: "Pest control in rental properties should be:", options: ["The tenant's sole responsibility", "Coordinated by management using licensed providers, with proper notice to tenants", "Done only when tenants complain", "Handled by maintenance techs with store-bought products"], correct: 1 },
    { id: "c16q3", text: "Which is a common environmental hazard in older multifamily properties?", options: ["New construction dust", "Asbestos in floor tiles, pipe insulation, or popcorn ceilings", "Solar panel glare", "None — older buildings are safe"], correct: 1 },
    { id: "c16q4", text: "Radon testing in rental properties is recommended:", options: ["Never", "Especially for ground-floor and basement units, as radon is a leading cause of lung cancer", "Only in commercial buildings", "Only if the tenant requests it"], correct: 1 },
    { id: "c16q5", text: "An Integrated Pest Management (IPM) approach focuses on:", options: ["Spraying chemicals on a regular schedule", "Prevention, sealing entry points, eliminating food and water sources, and targeted treatment as needed", "Only responding after infestations", "Making tenants responsible for all pest control"], correct: 1 },
  ]},
  c17: { courseId: "c17", passingScore: 80, questions: [
    { id: "c17q1", text: "The purpose of routine property inspections is:", options: ["To spy on tenants", "To identify maintenance needs, safety hazards, and lease violations before they escalate", "Only to satisfy insurance requirements", "To justify rent increases"], correct: 1 },
    { id: "c17q2", text: "How much notice is required before entering an occupied unit for a non-emergency inspection in South Carolina?", options: ["No notice required", "24 hours written notice is industry standard and recommended best practice", "48 hours", "One week"], correct: 1 },
    { id: "c17q3", text: "A quality control inspection after a turn should verify:", options: ["Just that the unit is clean", "All repairs completed, appliances functioning, cleaning standards met, and the unit is rent-ready", "Only paint and carpet", "Nothing — just list it"], correct: 1 },
    { id: "c17q4", text: "Inspection findings must be documented with:", options: ["A verbal summary to the office", "Photos, written notes, action items, and completion tracking in the management system", "Nothing if everything is fine", "A text to the maintenance team"], correct: 1 },
    { id: "c17q5", text: "When an inspection reveals a lease violation, the correct process is:", options: ["Ignore minor violations", "Document it, issue appropriate notice per the lease and state law, and follow the enforcement timeline", "Evict immediately", "Call the tenant and yell at them"], correct: 1 },
  ]},
  c18: { courseId: "c18", passingScore: 80, questions: [
    { id: "c18q1", text: "A slip-and-fall hazard is reported in a common area. The first action is:", options: ["Wait for the next maintenance round", "Address it immediately — mark the area, mitigate the hazard, and document the response", "Only fix it if someone falls", "Tell residents to be careful"], correct: 1 },
    { id: "c18q2", text: "An incident report should be completed:", options: ["Only for serious injuries", "For any incident involving injury, property damage, or potential liability, no matter how minor", "Only if the tenant demands it", "Once a month in a batch"], correct: 1 },
    { id: "c18q3", text: "Which is an example of proactive risk management?", options: ["Fixing things only when they break", "Regular safety audits, preventive maintenance, adequate lighting, and clear signage", "Waiting for insurance to tell you what to fix", "Only managing risks after a lawsuit"], correct: 1 },
    { id: "c18q4", text: "Tree limbs overhanging a parking area are:", options: ["An aesthetic issue only", "A liability risk that should be trimmed and documented as part of site safety", "Only a concern during storms", "The utility company's problem"], correct: 1 },
    { id: "c18q5", text: "After a severe weather event, the property should be:", options: ["Left alone until tenants report problems", "Inspected promptly for damage, hazards documented, and emergency repairs prioritized", "Only checked if insurance requires it", "Inspected next week during regular rounds"], correct: 1 },
  ]},
  c19: { courseId: "c19", passingScore: 80, questions: [
    { id: "c19q1", text: "Net Operating Income (NOI) is calculated as:", options: ["Total revenue minus mortgage payments", "Gross rental income minus operating expenses, not including debt service", "Rent collected minus taxes", "Revenue minus all expenses including capital improvements"], correct: 1 },
    { id: "c19q2", text: "A budget variance report shows:", options: ["Only revenue", "The difference between budgeted amounts and actual performance, flagging areas that need attention", "Only expenses", "Year-over-year rent changes"], correct: 1 },
    { id: "c19q3", text: "Owner distributions should be:", options: ["Sent whenever the owner asks", "Calculated after all operating expenses, reserves, and contractual obligations are accounted for", "The full rent amount collected", "Sent before paying vendors"], correct: 1 },
    { id: "c19q4", text: "Maintaining a reserve fund for each property protects against:", options: ["Nothing — it is wasted money", "Unexpected repairs, vacancies, and cash flow disruptions", "Only natural disasters", "Tax obligations"], correct: 1 },
    { id: "c19q5", text: "When reviewing a property's financial performance, delinquency rate refers to:", options: ["The number of vacant units", "The percentage of total rent owed that is past due", "The owner's tax liability", "Maintenance costs per unit"], correct: 1 },
  ]},
  c20: { courseId: "c20", passingScore: 80, questions: [
    { id: "c20q1", text: "Occupancy rate is calculated as:", options: ["Vacant units divided by total units", "Occupied units divided by total units, expressed as a percentage", "Lease applications divided by showings", "Revenue divided by expenses"], correct: 1 },
    { id: "c20q2", text: "A property with 95% occupancy and 8% delinquency has:", options: ["No problems", "A collections issue — high occupancy means little if units are occupied by non-paying tenants", "Perfect performance", "Too many vacant units"], correct: 1 },
    { id: "c20q3", text: "Average days on market measures:", options: ["How long a property has been managed", "The average number of days a vacant unit takes to lease from listing to signed lease", "Time between maintenance requests", "Tenant satisfaction ratings"], correct: 1 },
    { id: "c20q4", text: "Lease renewal rate is an important metric because:", options: ["It does not matter", "Higher renewals reduce turnover costs and vacancy loss, directly protecting NOI", "It only affects marketing", "Owners do not care about it"], correct: 1 },
    { id: "c20q5", text: "Cost per turn includes:", options: ["Only paint and carpet", "All expenses to make a unit rent-ready: cleaning, repairs, materials, labor, and lost rent during vacancy", "Just the cleaning fee", "Only vendor invoices"], correct: 1 },
  ]},
  c21: { courseId: "c21", passingScore: 80, questions: [
    { id: "c21q1", text: "AMI stands for:", options: ["Annual Maximum Income", "Area Median Income", "Average Monthly Income", "Assessed Market Index"], correct: 1 },
    { id: "c21q2", text: "The NewShire AMI Calculator determines eligibility based on:", options: ["Only the applicant's salary", "Total household income, household size, the applicable AMI percentage tier, and current HUD income limits", "Credit score and income combined", "Rent amount only"], correct: 1 },
    { id: "c21q3", text: "Household income for AMI qualification includes income from:", options: ["Only the lease signer", "All household members age 18 and older", "Only employed members", "Only the highest earner"], correct: 1 },
    { id: "c21q4", text: "Annual income recertification for AMI units must be completed:", options: ["Every 6 months", "Annually, before the certification expiration date, with all required documentation", "Only at initial move-in", "Every 2 years"], correct: 1 },
    { id: "c21q5", text: "Acceptable income verification documents include:", options: ["A verbal statement from the applicant", "Pay stubs, tax returns, employer verification letters, benefit award letters, and bank statements", "Only a single bank statement", "A credit report"], correct: 1 },
  ]},
};

// ============================================================
// EMPLOYEES (16 total — real NewShire org chart)
// ============================================================
const DEMO_EMPLOYEES = [
  // ── Owner/Operator ──
  { id: "e1", name: "John White", email: "Jwhite@vanrockre.com", role: "Owner/Operator", appRole: "Employee", reportsTo: null, hireDate: "2025-05-31", active: true },
  // ── Regional Manager → reports to John ──
  { id: "e2", name: "Cara Munson", email: "cmunson@newshirepm.com", role: "Property Manager", appRole: "Employee", reportsTo: "e1", hireDate: "2025-10-29", active: true },
  // ── Brandy (Admin) → reports to Cara ──
  { id: "e3", name: "Brandy Turner", email: "bturner@newshirepm.com", role: "Property Manager", appRole: "Admin", reportsTo: "e2", hireDate: "2025-11-11", active: true },
  // ── Property Managers → report to Cara ──
  { id: "e4", name: "Amanda Bradshaw", email: "abradshaw@newshirepm.com", role: "Property Manager", appRole: "Employee", reportsTo: "e2", hireDate: "2025-10-10", active: true },
  { id: "e5", name: "Lisa Roberts", email: "lroberts@newshirepm.com", role: "Property Manager", appRole: "Employee", reportsTo: "e2", hireDate: "2026-01-26", active: true },
  { id: "e6", name: "Ziba Rhemtulla", email: "zrhemtulla@newshirepm.com", role: "Property Manager", appRole: "Employee", reportsTo: "e2", hireDate: "2025-07-28", active: true },
  { id: "e7", name: "Monique Burgos", email: "Mburgos@newshirepm.com", role: "Property Manager", appRole: "Employee", reportsTo: "e2", hireDate: "2026-02-25", active: true },
  // ── Leasing Agent → reports to Cara ──
  { id: "e8", name: "Leslie Byrd", email: "lbyrd@newshirepm.com", role: "Leasing Agent", appRole: "Employee", reportsTo: "e2", hireDate: "2024-01-09", active: true },
  // ── Service Manager → reports to Cara ──
  { id: "e9", name: "Gerald Harblin", email: "gharblin@newshirepm.com", role: "Maintenance Supervisor", appRole: "Employee", reportsTo: "e2", hireDate: "2025-09-29", active: true },
  // ── Service Techs → report to Gerald ──
  { id: "e10", name: "Alex Said", email: "asaid@newshirepm.com", role: "Maintenance Tech", appRole: "Employee", reportsTo: "e9", hireDate: "2025-09-22", active: true },
  { id: "e11", name: "Chuck Sloan", email: "csloan@newshirepm.com", role: "Maintenance Tech", appRole: "Employee", reportsTo: "e9", hireDate: "2025-11-12", active: true },
  { id: "e12", name: "Rooky Duncan", email: "rduncan@newshirepm.com", role: "Maintenance Tech", appRole: "Employee", reportsTo: "e9", hireDate: "2026-01-29", active: true },
  // ── Virtual Assistants → report to Brandy ──
  { id: "e13", name: "Aljon Yabut", email: "ayabut@newshirepm.com", role: "Virtual Assistant", appRole: "Employee", reportsTo: "e3", hireDate: "2026-02-17", active: true },
  { id: "e14", name: "Moreblessings Mancama", email: "MMancama@newshirepm.com", role: "Virtual Assistant", appRole: "Employee", reportsTo: "e3", hireDate: "2026-01-20", active: true },
  { id: "e15", name: "Robyn Friedman", email: "RFriedman@newshirepm.com", role: "Virtual Assistant", appRole: "Employee", reportsTo: "e3", hireDate: "2026-01-20", active: true },
  { id: "e16", name: "Sarah Chen", email: "schen@newshirepm.com", role: "Virtual Assistant", appRole: "Employee", reportsTo: "e3", hireDate: "2026-02-10", active: true },
];

// ── Org Tree Utilities ──
// Get all direct and indirect reports for a given employee (recursive downward walk).
// ReportsTo can be either an employee ID (demo mode) or an email (SharePoint mode).
function getSubordinateIds(employeeId, employees) {
  const emp = employees.find(e => e.id === employeeId);
  if (!emp) return [];
  const myEmail = emp.email;
  // Direct reports: anyone whose reportsTo matches this employee's ID or email
  const directs = employees.filter(e =>
    e.active && e.id !== employeeId && (e.reportsTo === employeeId || e.reportsTo === myEmail)
  ).map(e => e.id);
  const indirects = directs.flatMap(did => getSubordinateIds(did, employees));
  return [...directs, ...indirects];
}

function getUserAccess(employee, employees) {
  const isAdmin = employee.appRole === "Admin";
  const subordinateIds = getSubordinateIds(employee.id, employees);
  const isManager = subordinateIds.length > 0;
  return { isAdmin, isManager, subordinateIds };
}

// Simulated completions — realistic scenarios for the real team
const DEMO_COMPLETIONS = [
  // ── Brandy Turner (e3) — Admin, fully complete on all her required paths ──
  { id: "comp_b1", employeeId: "e3", courseId: "c1", completedDate: "2025-12-01", score: 96, status: "passed" },
  { id: "comp_b2", employeeId: "e3", courseId: "c2", completedDate: "2025-12-01", score: 100, status: "passed" },
  { id: "comp_b3", employeeId: "e3", courseId: "c3", completedDate: "2025-12-02", score: 98, status: "passed", certExpires: "2026-12-02" },
  { id: "comp_b4", employeeId: "e3", courseId: "c4", completedDate: "2025-12-02", score: 95, status: "passed", certExpires: "2026-12-02" },
  { id: "comp_b5", employeeId: "e3", courseId: "c5", completedDate: "2025-12-03", score: 100, status: "passed", certExpires: "2026-12-03" },
  { id: "comp_b6", employeeId: "e3", courseId: "c6", completedDate: "2025-12-03", score: 92, status: "passed" },
  { id: "comp_b7", employeeId: "e3", courseId: "c7", completedDate: "2025-12-04", score: 95, status: "passed" },
  { id: "comp_b8", employeeId: "e3", courseId: "c8", completedDate: "2025-12-04", score: 100, status: "passed" },
  { id: "comp_b9", employeeId: "e3", courseId: "c9", completedDate: "2025-12-05", score: 90, status: "passed" },
  { id: "comp_b10", employeeId: "e3", courseId: "c10", completedDate: "2025-12-05", score: 88, status: "passed" },
  { id: "comp_b11", employeeId: "e3", courseId: "c14", completedDate: "2025-12-06", score: 92, status: "passed" },
  { id: "comp_b12", employeeId: "e3", courseId: "c15", completedDate: "2025-12-06", score: 96, status: "passed" },
  { id: "comp_b13", employeeId: "e3", courseId: "c16", completedDate: "2025-12-07", score: 88, status: "passed" },
  { id: "comp_b14", employeeId: "e3", courseId: "c17", completedDate: "2025-12-07", score: 92, status: "passed" },
  { id: "comp_b15", employeeId: "e3", courseId: "c18", completedDate: "2025-12-08", score: 90, status: "passed" },
  { id: "comp_b16", employeeId: "e3", courseId: "c19", completedDate: "2025-12-08", score: 95, status: "passed" },
  { id: "comp_b17", employeeId: "e3", courseId: "c20", completedDate: "2025-12-09", score: 100, status: "passed" },
  { id: "comp_b18", employeeId: "e3", courseId: "c21", completedDate: "2025-12-09", score: 92, status: "passed", certExpires: "2026-12-09" },

  // ── Cara Munson (e2) — mostly complete, Fair Housing expiring soon ──
  { id: "comp_c1", employeeId: "e2", courseId: "c1", completedDate: "2025-11-10", score: 90, status: "passed" },
  { id: "comp_c2", employeeId: "e2", courseId: "c2", completedDate: "2025-11-10", score: 88, status: "passed" },
  { id: "comp_c3", employeeId: "e2", courseId: "c3", completedDate: "2025-03-15", score: 92, status: "passed", certExpires: "2026-03-15" },
  { id: "comp_c4", employeeId: "e2", courseId: "c4", completedDate: "2025-03-15", score: 85, status: "passed", certExpires: "2026-03-15" },
  { id: "comp_c5", employeeId: "e2", courseId: "c5", completedDate: "2025-11-11", score: 96, status: "passed", certExpires: "2026-11-11" },
  { id: "comp_c6", employeeId: "e2", courseId: "c6", completedDate: "2025-11-11", score: 88, status: "passed" },
  { id: "comp_c7", employeeId: "e2", courseId: "c7", completedDate: "2025-11-12", score: 90, status: "passed" },
  { id: "comp_c8", employeeId: "e2", courseId: "c8", completedDate: "2025-11-12", score: 92, status: "passed" },
  { id: "comp_c9", employeeId: "e2", courseId: "c9", completedDate: "2025-11-13", score: 88, status: "passed" },
  { id: "comp_c10", employeeId: "e2", courseId: "c19", completedDate: "2025-11-15", score: 90, status: "passed" },
  { id: "comp_c11", employeeId: "e2", courseId: "c20", completedDate: "2025-11-15", score: 86, status: "passed" },

  // ── Amanda Bradshaw (e4) — partially through onboarding, started recently ──
  { id: "comp_a1", employeeId: "e4", courseId: "c1", completedDate: "2025-10-20", score: 88, status: "passed" },
  { id: "comp_a2", employeeId: "e4", courseId: "c2", completedDate: "2025-10-20", score: 84, status: "passed" },
  { id: "comp_a3", employeeId: "e4", courseId: "c3", completedDate: "2025-10-22", score: 92, status: "passed", certExpires: "2026-10-22" },
  { id: "comp_a4", employeeId: "e4", courseId: "c4", completedDate: "2025-10-22", score: 88, status: "passed", certExpires: "2026-10-22" },
  { id: "comp_a5", employeeId: "e4", courseId: "c5", completedDate: "2025-10-23", score: 80, status: "passed", certExpires: "2026-10-23" },

  // ── Leslie Byrd (e8) — long-tenured leasing agent, FH cert EXPIRED ──
  { id: "comp_l1", employeeId: "e8", courseId: "c1", completedDate: "2024-06-15", score: 88, status: "passed" },
  { id: "comp_l2", employeeId: "e8", courseId: "c2", completedDate: "2024-06-15", score: 82, status: "passed" },
  { id: "comp_l3", employeeId: "e8", courseId: "c3", completedDate: "2024-06-16", score: 84, status: "passed", certExpires: "2025-06-16" },
  { id: "comp_l4", employeeId: "e8", courseId: "c4", completedDate: "2024-06-16", score: 80, status: "passed", certExpires: "2025-06-16" },
  { id: "comp_l5", employeeId: "e8", courseId: "c5", completedDate: "2024-06-17", score: 88, status: "passed", certExpires: "2025-06-17" },
  { id: "comp_l6", employeeId: "e8", courseId: "c6", completedDate: "2024-06-17", score: 90, status: "passed" },
  { id: "comp_l7", employeeId: "e8", courseId: "c7", completedDate: "2024-06-18", score: 86, status: "passed" },
  { id: "comp_l8", employeeId: "e8", courseId: "c8", completedDate: "2024-06-18", score: 84, status: "passed" },
  { id: "comp_l9", employeeId: "e8", courseId: "c9", completedDate: "2024-07-01", score: 92, status: "passed" },
  { id: "comp_l10", employeeId: "e8", courseId: "c10", completedDate: "2024-07-02", score: 86, status: "passed" },
  { id: "comp_l11", employeeId: "e8", courseId: "c21", completedDate: "2024-07-05", score: 84, status: "passed", certExpires: "2025-07-05" },

  // ── Gerald Harblin (e9) — Service Manager, mostly done ──
  { id: "comp_g1", employeeId: "e9", courseId: "c1", completedDate: "2025-10-15", score: 90, status: "passed" },
  { id: "comp_g2", employeeId: "e9", courseId: "c2", completedDate: "2025-10-15", score: 86, status: "passed" },
  { id: "comp_g3", employeeId: "e9", courseId: "c3", completedDate: "2025-10-16", score: 88, status: "passed", certExpires: "2026-10-16" },
  { id: "comp_g4", employeeId: "e9", courseId: "c4", completedDate: "2025-10-16", score: 84, status: "passed", certExpires: "2026-10-16" },
  { id: "comp_g5", employeeId: "e9", courseId: "c5", completedDate: "2025-10-17", score: 92, status: "passed", certExpires: "2026-10-17" },
  { id: "comp_g6", employeeId: "e9", courseId: "c6", completedDate: "2025-10-17", score: 88, status: "passed" },
  { id: "comp_g7", employeeId: "e9", courseId: "c7", completedDate: "2025-10-18", score: 86, status: "passed" },
  { id: "comp_g8", employeeId: "e9", courseId: "c8", completedDate: "2025-10-18", score: 90, status: "passed" },
  { id: "comp_g9", employeeId: "e9", courseId: "c11", completedDate: "2025-10-20", score: 92, status: "passed" },
  { id: "comp_g10", employeeId: "e9", courseId: "c12", completedDate: "2025-10-21", score: 88, status: "passed" },
  { id: "comp_g11", employeeId: "e9", courseId: "c13", completedDate: "2025-10-22", score: 86, status: "passed" },

  // ── Alex Said (e10) — maintenance tech, partially through training ──
  { id: "comp_x1", employeeId: "e10", courseId: "c1", completedDate: "2025-10-01", score: 84, status: "passed" },
  { id: "comp_x2", employeeId: "e10", courseId: "c2", completedDate: "2025-10-01", score: 80, status: "passed" },
  { id: "comp_x3", employeeId: "e10", courseId: "c3", completedDate: "2025-10-02", score: 88, status: "passed", certExpires: "2026-10-02" },
  { id: "comp_x4", employeeId: "e10", courseId: "c4", completedDate: "2025-10-02", score: 82, status: "passed", certExpires: "2026-10-02" },
  { id: "comp_x5", employeeId: "e10", courseId: "c11", completedDate: "2025-10-05", score: 90, status: "passed" },

  // ── Moreblessings Mancama (e14) — VA, onboarding in progress ──
  { id: "comp_m1", employeeId: "e14", courseId: "c1", completedDate: "2026-01-25", score: 92, status: "passed" },
  { id: "comp_m2", employeeId: "e14", courseId: "c2", completedDate: "2026-01-25", score: 88, status: "passed" },
  { id: "comp_m3", employeeId: "e14", courseId: "c3", completedDate: "2026-01-27", score: 84, status: "passed", certExpires: "2027-01-27" },

  // ── Failed attempt — Rooky Duncan (e12) failed Fair Housing first try ──
  { id: "comp_f1", employeeId: "e12", courseId: "c3", completedDate: "2026-02-10", score: 60, status: "failed" },
  { id: "comp_f2", employeeId: "e12", courseId: "c1", completedDate: "2026-02-05", score: 80, status: "passed" },
  { id: "comp_f3", employeeId: "e12", courseId: "c2", completedDate: "2026-02-05", score: 84, status: "passed" },
];

const TODAY = new Date().toISOString().split("T")[0];

// ============================================================
// UTILITY FUNCTIONS
// ============================================================
function daysBetween(d1, d2) {
  return Math.round((new Date(d2) - new Date(d1)) / (1000 * 60 * 60 * 24));
}

function getCertStatus(completion, course) {
  if (!completion || completion.status !== "passed") return "incomplete";
  if (!course.recertDays && !completion.certExpires) return "current";
  const expiry = completion.certExpires;
  if (!expiry) return "current";
  const daysLeft = daysBetween(TODAY, expiry);
  if (daysLeft < 0) return "expired";
  if (daysLeft <= 30) return "expiring";
  return "current";
}

function getPathProgress(pathId, employeeId, completions, courses, learningPaths, employeeRole) {
  const path = learningPaths.find(p => p.id === pathId);
  if (!path) return { total: 0, completed: 0, pct: 0 };
  // Filter to Active courses this employee's role qualifies for
  const applicableIds = path.courseIds.filter(cid => {
    const course = courses.find(c => c.id === cid);
    return course && course.status === "Active" && courseMatchesRole(course, employeeRole);
  });
  const total = applicableIds.length;
  if (total === 0) return { total: 0, completed: 0, pct: 100 }; // No active courses = nothing to do = complete
  const completed = applicableIds.filter(cid => {
    const comp = completions.filter(c => c.employeeId === employeeId && c.courseId === cid && c.status === "passed");
    if (comp.length === 0) return false;
    const course = courses.find(c => c.id === cid);
    const latest = comp.sort((a, b) => b.completedDate.localeCompare(a.completedDate))[0];
    return getCertStatus(latest, course) !== "expired";
  }).length;
  return { total, completed, pct: total > 0 ? Math.round((completed / total) * 100) : 100 };
}

// Roles exempt from required training deadlines (can still take courses voluntarily)
const EXEMPT_ROLES = ["Owner/Operator"];

function getEmployeePaths(employee, learningPaths) {
  return learningPaths.filter(p =>
    p.roles.includes("All") || p.roles.includes(employee.role)
  );
}

function isTrainingExempt(employee) {
  return EXEMPT_ROLES.includes(employee.role);
}

// Course-level role filtering: empty roles = all roles, populated = only matching roles
function courseMatchesRole(course, role) {
  if (!course.roles || course.roles.length === 0) return true;
  return course.roles.includes(role);
}

// ============================================================
// ICON COMPONENTS
// ============================================================
const Icons = {
  Play: () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polygon points="5 3 19 12 5 21 5 3"/></svg>,
  Check: () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"/></svg>,
  X: () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>,
  Alert: () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>,
  Clock: () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>,
  Doc: () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>,
  Back: () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="19" y1="12" x2="5" y2="12"/><polyline points="12 19 5 12 12 5"/></svg>,
  Trophy: () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M6 9H4.5a2.5 2.5 0 010-5H6"/><path d="M18 9h1.5a2.5 2.5 0 000-5H18"/><path d="M4 22h16"/><path d="M10 14.66V17c0 .55-.47.98-.97 1.21C7.85 18.75 7 20 7 22"/><path d="M14 14.66V17c0 .55.47.98.97 1.21C16.15 18.75 17 20 17 22"/><path d="M18 2H6v7a6 6 0 0012 0V2z"/></svg>,
  Users: () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87"/><path d="M16 3.13a4 4 0 010 7.75"/></svg>,
  Shield: () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>,
  ChevronRight: () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="9 18 15 12 9 6"/></svg>,
  RefreshCw: () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/><path d="M3.51 9a9 9 0 0114.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0020.49 15"/></svg>,
  Download: () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>,
  BookOpen: () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M2 3h6a4 4 0 014 4v14a3 3 0 00-3-3H2z"/><path d="M22 3h-6a4 4 0 00-4 4v14a3 3 0 013-3h7z"/></svg>,
  Plus: () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>,
};

// ============================================================
// PROGRESS BAR COMPONENT
// ============================================================
function ProgressBar({ pct, color = C.success, height = 8, label = true }) {
  const bg = pct >= 100 ? C.success : pct >= 50 ? C.gold500 : C.warning;
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 10, width: "100%" }}>
      <div style={{ flex: 1, height, borderRadius: 4, background: C.gray100, overflow: "hidden" }}>
        <div style={{ height: "100%", width: `${Math.min(100, pct)}%`, background: color || bg, borderRadius: 4, transition: "width 0.4s ease" }} />
      </div>
      {label && <span style={{ fontSize: 13, fontWeight: 600, color: C.teal700, minWidth: 40 }}>{pct}%</span>}
    </div>
  );
}

// ============================================================
// MAIN APP
// ============================================================
// Helper: get all required course IDs for an employee based on their learning paths
function getRequiredCourseIds(employee, learningPaths, courses) {
  const paths = getEmployeePaths(employee, learningPaths);
  const allCourseIds = paths.flatMap(p => p.courseIds);
  // Filter to courses this employee's role qualifies for
  return [...new Set(allCourseIds.filter(cid => {
    const course = courses?.find(c => c.id === cid);
    return !course || courseMatchesRole(course, employee.role);
  }))];
}

function App() {
  // ── Auth + data loading state ──
  const [authState, setAuthState] = useState("init"); // init | loading | ready | error | demo
  const [authError, setAuthError] = useState(null);
  const [msalAccount, setMsalAccount] = useState(null);
  const tokenRef = useRef(null);

  // ── App data state (populated from SharePoint or demo data) ──
  const [employees, setEmployees] = useState([]);
  const [courses, setCourses] = useState([]);
  const [learningPaths, setLearningPaths] = useState([]);
  const [lessons, setLessons] = useState([]);
  const [quizzes, setQuizzes] = useState({});
  const [completions, setCompletions] = useState([]);
  const [enrollments, setEnrollments] = useState([]);
  const [assignments, setAssignments] = useState([]);
  const [currentUser, setCurrentUser] = useState(null);

  // ── UI state ──
  const [tab, setTab] = useState(0);
  const [view, setView] = useState(null);
  const [mobile, setMobile] = useState(window.innerWidth < 640);

  useEffect(() => {
    const handleResize = () => setMobile(window.innerWidth < 640);
    window.addEventListener("resize", handleResize);
    return () => window.removeEventListener("resize", handleResize);
  }, []);

  // ── Bootstrap: Attempt MSAL login + data load; fallback to demo ──
  useEffect(() => {
    if (!CONFIG.isConfigured) {
      loadDemoData();
      return;
    }
    (async () => {
      setAuthState("loading");
      try {
        const account = await msalLogin();
        if (!account) { loadDemoData(); return; }
        setMsalAccount(account);
        const token = await msalGetToken(account);
        tokenRef.current = token;
        const data = await loadAllData(token);
        setEmployees(data.employees);
        setCourses(data.courses);
        setLearningPaths(data.paths);
        setLessons(data.lessons);
        setQuizzes(data.quizzes);
        setCompletions(data.completions);
        setEnrollments(data.enrollments);
        setAssignments(data.assignments || []);
        // Match logged-in user by email
        const email = account.username.toLowerCase();
        const user = data.employees.find(e => e.email === email);
        if (!user) {
          setAuthError(`Your email (${email}) was not found in the Employees list. Contact your administrator.`);
          setAuthState("error");
          return;
        }
        setCurrentUser(user);
        setAuthState("ready");
      } catch (err) {
        console.error("Auth/load error:", err);
        setAuthError(err.message);
        // Offer demo mode fallback
        loadDemoData();
      }
    })();
  }, []);

  function loadDemoData() {
    setEmployees(DEMO_EMPLOYEES);
    setCourses(DEMO_COURSES);
    setLearningPaths(DEMO_LEARNING_PATHS);
    setLessons(DEMO_LESSONS);
    setQuizzes(DEMO_QUIZZES);
    setCompletions(DEMO_COMPLETIONS);
    setEnrollments([
      { employeeId: "e3", courseId: "c11", enrolledDate: "2026-01-15" },
      { employeeId: "e3", courseId: "c12", enrolledDate: "2026-01-15" },
      { employeeId: "e8", courseId: "c19", enrolledDate: "2026-02-01" },
    ]);
    setCurrentUser(DEMO_EMPLOYEES.find(e => e.id === "e3"));
    setAuthState("demo");
  }

  // ── Refresh token helper ──
  async function getToken() {
    if (!msalAccount) return null;
    try {
      const token = await msalGetToken(msalAccount);
      tokenRef.current = token;
      return token;
    } catch { return tokenRef.current; }
  }

  const isLive = authState === "ready"; // connected to SharePoint
  const [notifBanner, setNotifBanner] = useState(null); // { text, type } for admin notification feedback

  // ── Admin auto-scan: cert expirations + Monday manager report ──
  useEffect(() => {
    if (!isLive || !currentUser) return;
    const isAdmin = currentUser.appRole === "admin";
    (async () => {
      try {
        const token = await getToken();
        if (!token) return;
        // Cert expiration scanner runs for admins
        if (isAdmin) {
          const adminEmails = employees.filter(e => e.active && e.appRole === "admin").map(e => e.email);
          const certsSent = await runCertExpirationScan(token, employees, completions, courses, adminEmails.length > 0 ? adminEmails : [CONFIG.adminEmail]);
          if (certsSent > 0) setNotifBanner({ text: `Sent ${certsSent} certification expiration notification${certsSent > 1 ? "s" : ""}.`, type: "info" });
        }
        // Monday manager report runs for everyone (each manager gets their own)
        const mgrSent = await runMondayManagerReport(token, employees, completions, courses, learningPaths);
        if (mgrSent > 0 && isAdmin) {
          setNotifBanner(prev => prev
            ? { text: `${prev.text} Sent ${mgrSent} Monday manager report${mgrSent > 1 ? "s" : ""}.`, type: "info" }
            : { text: `Sent ${mgrSent} Monday manager report${mgrSent > 1 ? "s" : ""}.`, type: "info" }
          );
        }
        // Auto-dismiss banner after 8 seconds
        if (notifBanner || true) setTimeout(() => setNotifBanner(null), 8000);
      } catch (err) {
        console.error("Notification scan error:", err);
      }
    })();
  }, [isLive, currentUser?.id]); // eslint-disable-line

  // ── Enroll handler ──
  const handleEnroll = async (employeeId, courseId) => {
    if (enrollments.some(e => e.employeeId === employeeId && e.courseId === courseId)) return;
    const emp = employees.find(e => e.id === employeeId);
    const course = courses.find(c => c.id === courseId);
    if (isLive && emp) {
      try {
        const token = await getToken();
        const enrollment = await createEnrollmentSP(token, emp, courseId);
        setEnrollments(prev => [...prev, enrollment]);
        // Enrollment confirmation email (non-blocking)
        if (course) sendEnrollmentEmail(token, emp, course).catch(e => console.error("Enrollment email failed:", e));
      } catch (err) { console.error("Enroll failed:", err); }
    } else {
      setEnrollments(prev => [...prev, { employeeId, courseId, enrolledDate: TODAY }]);
    }
  };

  // ── Unenroll handler ──
  const handleUnenroll = async (employeeId, courseId) => {
    const enrollment = enrollments.find(e => e.employeeId === employeeId && e.courseId === courseId);
    if (!enrollment) return;
    if (isLive && enrollment.spItemId) {
      try {
        const token = await getToken();
        await deleteEnrollmentSP(token, enrollment.spItemId);
      } catch (err) { console.error("Unenroll failed:", err); }
    }
    setEnrollments(prev => prev.filter(e => !(e.employeeId === employeeId && e.courseId === courseId)));
  };

  // ── Assign course handler (supervisor/admin assigns to an employee) ──
  const handleAssignCourse = async ({ employeeId, courseId, dueDate, notes }) => {
    const emp = employees.find(e => e.id === employeeId);
    const course = courses.find(c => c.id === courseId);
    if (!emp || !course) return;
    const assignData = {
      Title: `${emp.name} - ${course.name}`,
      AssignEmployeeEmail: emp.email,
      AssignCourseIDLookupId: parseInt(courseId, 10),
      AssignedByEmail: currentUser.email,
      AssignedDate: new Date().toISOString(),
      AssignDueDate: dueDate || null,
      AssignNotes: notes || "",
      AssignStatus: "Assigned",
    };
    if (isLive) {
      try {
        const token = await getToken();
        const result = await spCreate(token, CONFIG.lists.assignments, assignData);
        const newAssignment = {
          id: String(result.id),
          employeeId: emp.id,
          courseId,
          assignedBy: currentUser.name,
          assignedById: currentUser.id,
          assignedDate: TODAY,
          dueDate: dueDate || null,
          notes: notes || "",
          status: "Assigned",
        };
        setAssignments(prev => [...prev, newAssignment]);
        // Send notification email
        const dueLine = dueDate ? `<p>This course must be completed by <strong>${new Date(dueDate).toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" })}</strong>.</p>` : "";
        const notesLine = notes ? `<p><strong>Notes from your supervisor:</strong> ${notes}</p>` : "";
        const bodyHtml = `<p>Hi ${emp.name.split(" ")[0]},</p>` +
          `<p>${currentUser.name} has assigned you the course <strong>${course.name}</strong> in NewShire University.</p>` +
          `<p>You must complete this course and pass the assessment with a score of ${course.passingScore || CONFIG.passingScore}% or higher.</p>` +
          dueLine + notesLine +
          `<p>Log in to NewShire University to begin.</p>`;
        sendEmail(token, emp.email, `Course Assigned: ${course.name}`, emailTemplate(bodyHtml, `Course Assigned: ${course.name}`))
          .catch(e => console.error("Assignment email failed:", e));
      } catch (err) { console.error("Assign failed:", err); alert("Failed to assign: " + err.message); }
    } else {
      setAssignments(prev => [...prev, { id: `asgn_${Date.now()}`, employeeId, courseId, assignedBy: currentUser.name, assignedById: currentUser.id, assignedDate: TODAY, dueDate, notes, status: "Assigned" }]);
    }
  };

  // ── Quiz submit handler (called from QuizView) ──
  const handleQuizSubmit = async (employee, course, score, passed, answersJson) => {
    const certExpires = passed && course.recertDays
      ? new Date(Date.now() + course.recertDays * 86400000).toISOString().split("T")[0] : null;
    if (isLive) {
      try {
        const token = await getToken();
        const comp = await submitQuizToSP(token, employee, course, score, passed, answersJson);
        setCompletions(prev => [...prev, comp]);
        // Fire notification emails (non-blocking — don't let email failure break the quiz flow)
        const admins = employees.filter(e => e.active && e.appRole === "admin").map(e => e.email);
        sendQuizResultEmail(token, employee, course, score, passed, admins.length > 0 ? admins : [CONFIG.adminEmail]).catch(e => console.error("Quiz email failed:", e));
        // Auto-complete any matching assignments for this employee + course
        if (passed) {
          const openAssignments = assignments.filter(a => a.employeeId === employee.id && a.courseId === course.id && a.status === "Assigned");
          for (const a of openAssignments) {
            try {
              await spUpdate(token, CONFIG.lists.assignments, a.id, { AssignStatus: "Completed" });
              setAssignments(prev => prev.map(x => x.id === a.id ? { ...x, status: "Completed" } : x));
            } catch (e) { console.error("Assignment completion update failed:", e); }
          }
        }
        return comp;
      } catch (err) { console.error("Quiz submit failed:", err); }
    }
    // Demo fallback
    const comp = {
      id: `comp_${Date.now()}`,
      employeeId: employee.id,
      courseId: course.id,
      completedDate: TODAY,
      score,
      status: passed ? "passed" : "failed",
      ...(certExpires ? { certExpires } : {}),
    };
    setCompletions(prev => [...prev, comp]);
    return comp;
  };

  // ── Loading / error screens ──
  if (authState === "init" || authState === "loading") {
    return (
      <div style={{ ...S.page, display: "flex", alignItems: "center", justifyContent: "center", minHeight: "100vh" }}>
        <link href="https://fonts.googleapis.com/css2?family=Source+Sans+3:wght@300;400;500;600;700&display=swap" rel="stylesheet" />
        <div style={{ textAlign: "center" }}>
          <div style={{ fontSize: 18, fontWeight: 600, color: C.teal700, marginBottom: 8 }}>{CONFIG.appName}</div>
          <div style={{ fontSize: 14, color: C.gray400 }}>Signing in and loading training data...</div>
          <div style={{ marginTop: 16, width: 200, height: 4, background: C.gray100, borderRadius: 2, overflow: "hidden", margin: "16px auto" }}>
            <div style={{ width: "60%", height: "100%", background: C.gold500, borderRadius: 2, animation: "pulse 1.5s ease-in-out infinite" }} />
          </div>
          <div style={{ position: "fixed", bottom: 16, left: 0, right: 0, textAlign: "center", fontSize: 11, color: C.gray300 }}>2025 · This application is the intellectual property of NewShire Property Management. Reproduction or use without written permission is prohibited.</div>
        </div>
      </div>
    );
  }

  if (authState === "error") {
    return (
      <div style={{ ...S.page, display: "flex", alignItems: "center", justifyContent: "center", minHeight: "100vh" }}>
        <link href="https://fonts.googleapis.com/css2?family=Source+Sans+3:wght@300;400;500;600;700&display=swap" rel="stylesheet" />
        <div style={{ ...S.card, maxWidth: 480, textAlign: "center" }}>
          <div style={{ fontSize: 18, fontWeight: 600, color: C.error, marginBottom: 8 }}>Unable to Load</div>
          <div style={{ fontSize: 14, color: C.gray600, marginBottom: 16 }}>{authError}</div>
          <button style={S.btnPrimary} onClick={() => window.location.reload()}>Retry</button>
        </div>
      </div>
    );
  }

  if (!currentUser) return null;

  // ── Compute access ──
  const { isAdmin, isManager, subordinateIds } = getUserAccess(currentUser, employees);
  const showComplianceDashboard = isAdmin || isManager;

  const TABS = [];
  TABS.push("My Training");
  if (showComplianceDashboard) TABS.push("Team Compliance");
  TABS.push("Training Library");
  if (isAdmin) TABS.push("Manage");
  const activeTabName = TABS[tab] || TABS[0];

  const complianceEmployeeIds = isAdmin
    ? employees.filter(e => e.active).map(e => e.id)
    : [...subordinateIds, currentUser.id];

  // ── Context value ──
  const ctx = { employees, setEmployees, courses, setCourses, learningPaths, setLearningPaths, lessons, setLessons, quizzes, setQuizzes, completions, enrollments, isLive, getToken: getToken };

  return (
    <DataContext.Provider value={ctx}>
      <div style={S.page}>
        <link href="https://fonts.googleapis.com/css2?family=Source+Sans+3:wght@300;400;500;600;700&family=Source+Code+Pro:wght@400;500;600&display=swap" rel="stylesheet" />

        {/* Demo mode banner */}
        {authState === "demo" && (
          <div style={{ background: C.warningBg, borderBottom: `1px solid ${C.warningBdr}`, padding: "6px 20px", fontSize: 12, color: C.warning, textAlign: "center", fontWeight: 500 }}>
            DEMO MODE — Using sample data. {CONFIG.isConfigured ? "SharePoint auth failed; showing demo." : "Set CONFIG.isConfigured = true for live data."}
          </div>
        )}

        {/* Notification scan results banner */}
        {notifBanner && (
          <div style={{ background: "#EDF4F7", borderBottom: `1px solid ${C.teal100}`, padding: "6px 20px", fontSize: 12, color: C.teal700, textAlign: "center", fontWeight: 500, display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}>
            <span style={{ fontSize: 14 }}>✉</span> {notifBanner.text}
            <button onClick={() => setNotifBanner(null)} style={{ background: "none", border: "none", cursor: "pointer", fontSize: 14, color: C.teal500, padding: "0 4px" }}>×</button>
          </div>
        )}

        {/* Header */}
        <div style={S.header}>
          <div>
            <div style={S.headerTitle}>{CONFIG.appName}</div>
            <div style={S.headerSubtitle}>NewShire Property Management</div>
          </div>
          <div style={S.headerUser}>
            <div style={{ textAlign: "right" }}>
              <div style={{ fontSize: 13, color: "#FFFFFF", fontWeight: 500 }}>{currentUser.name}</div>
              <div style={{ fontSize: 11, color: C.teal100 }}>
                {currentUser.role}
                {isAdmin && " \u00b7 Admin"}
                {!isAdmin && isManager && " \u00b7 Manager"}
              </div>
            </div>
          </div>
        </div>

        {/* Email Paused Banner */}
        {isAdmin && EMAIL_PAUSED && (
          <div style={{ background: "#C44B3B", color: "#FFF", padding: "6px 16px", fontSize: 12, fontWeight: 600, textAlign: "center", letterSpacing: "0.03em" }}>
            EMAILS PAUSED — No notification emails are being sent. Go to Manage → Settings to re-enable.
          </div>
        )}

        {/* Tabs */}
        <div style={S.tabBar}>
          {TABS.map((t, i) => (
            <button key={t} style={S.tab(tab === i)} onClick={() => { setTab(i); setView(null); }}>
              {t}
            </button>
          ))}
        </div>

        {/* Content */}
        <div style={S.content}>
          {activeTabName === "My Training" && (
            <MyTrainingView user={currentUser} completions={completions} setCompletions={setCompletions} enrollments={enrollments} assignments={assignments} onUnenroll={handleUnenroll} onQuizSubmit={handleQuizSubmit} view={view} setView={setView} mobile={mobile} />
          )}
          {activeTabName === "Team Compliance" && showComplianceDashboard && (
            <ComplianceDashboard
              completions={completions}
              enrollments={enrollments}
              visibleEmployeeIds={complianceEmployeeIds}
              isAdmin={isAdmin}
              currentUser={currentUser}
              mobile={mobile}
            />
          )}
          {activeTabName === "Training Library" && (
            <TrainingLibraryView user={currentUser} completions={completions} enrollments={enrollments} assignments={assignments} onEnroll={handleEnroll} onUnenroll={handleUnenroll} onAssign={handleAssignCourse} isManager={isManager} isAdmin={isAdmin} subordinateIds={subordinateIds} setView={(v) => { setView(v); setTab(0); }} mobile={mobile} />
          )}
          {activeTabName === "Manage" && isAdmin && (
            <ManageView mobile={mobile} />
          )}
        </div>
      </div>
    </DataContext.Provider>
  );
}

// ============================================================
// MY TRAINING VIEW
// ============================================================
function MyTrainingView({ user, completions, setCompletions, enrollments, assignments, onUnenroll, onQuizSubmit, view, setView, mobile }) {
  const { courses, learningPaths, lessons, quizzes } = useData();
  const myCompletions = completions.filter(c => c.employeeId === user.id);
  const myAssignments = (assignments || []).filter(a => a.employeeId === user.id && a.status === "Assigned");
  const [collapsedPaths, setCollapsedPaths] = useState([]);

  // Sub-views
  if (view?.type === "course") return <CourseView courseId={view.courseId} user={user} completions={completions} setCompletions={setCompletions} onQuizSubmit={onQuizSubmit} onBack={() => setView(null)} mobile={mobile} />;
  if (view?.type === "quiz") return <QuizView courseId={view.courseId} user={user} completions={completions} setCompletions={setCompletions} onQuizSubmit={onQuizSubmit} onBack={() => setView({ type: "course", courseId: view.courseId })} />;

  const paths = getEmployeePaths(user, learningPaths);
  const expiredCerts = [];
  const expiringCerts = [];
  const overduePaths = [];
  const dueSoonPaths = [];

  // Check path due dates
  paths.forEach(path => {
    const { dueDate, status: dueStatus } = getPathDueStatus(path, user, completions, courses, learningPaths);
    if (dueStatus === "overdue") overduePaths.push({ path, dueDate });
    else if (dueStatus === "due-soon") dueSoonPaths.push({ path, dueDate });
  });

  courses.forEach(course => {
    const latest = myCompletions.filter(c => c.courseId === course.id && c.status === "passed").sort((a, b) => b.completedDate.localeCompare(a.completedDate))[0];
    const status = getCertStatus(latest, course);
    if (status === "expired") expiredCerts.push({ course, completion: latest });
    if (status === "expiring") expiringCerts.push({ course, completion: latest, daysLeft: daysBetween(TODAY, latest.certExpires) });
  });

  const totalRequired = paths.reduce((sum, p) => sum + p.courseIds.length, 0);
  const totalCompleted = paths.reduce((sum, p) => sum + getPathProgress(p.id, user.id, completions, courses, learningPaths, user.role).completed, 0);

  return (
    <div>
      {/* Overdue learning paths alert */}
      {overduePaths.length > 0 && (
        <div style={{ ...S.card, borderLeft: `4px solid ${C.error}`, background: C.errorBg }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, color: C.error, fontWeight: 600, fontSize: 15, marginBottom: 8 }}>
            <Icons.Alert /> Overdue Training — Immediate Action Required
          </div>
          {overduePaths.map(({ path, dueDate }) => {
            const progress = getPathProgress(path.id, user.id, completions, courses, learningPaths, user.role);
            return (
              <div key={path.id} style={{ padding: "8px 0", borderBottom: `1px solid ${C.errorBdr}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div>
                  <span style={{ fontWeight: 600, color: C.teal700 }}>{path.name}</span>
                  <span style={{ fontSize: 13, color: C.error, marginLeft: 8 }}>Due {dueDate} · {progress.completed}/{progress.total} complete</span>
                </div>
              </div>
            );
          })}
        </div>
      )}

      {/* Due-soon learning paths alert */}
      {dueSoonPaths.length > 0 && (
        <div style={{ ...S.card, borderLeft: `4px solid ${C.warning}`, background: C.warningBg }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, color: C.warning, fontWeight: 600, fontSize: 15, marginBottom: 8 }}>
            <Icons.Clock /> Training Due This Week
          </div>
          {dueSoonPaths.map(({ path, dueDate }) => {
            const progress = getPathProgress(path.id, user.id, completions, courses, learningPaths, user.role);
            return (
              <div key={path.id} style={{ padding: "8px 0", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <span style={{ fontWeight: 500, color: C.teal700 }}>{path.name} — <span style={{ color: C.warning }}>Due {dueDate} · {progress.completed}/{progress.total} complete</span></span>
              </div>
            );
          })}
        </div>
      )}

      {/* Alerts */}
      {expiredCerts.length > 0 && (
        <div style={{ ...S.card, borderLeft: `4px solid ${C.error}`, background: C.errorBg }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, color: C.error, fontWeight: 600, fontSize: 15, marginBottom: 8 }}>
            <Icons.Alert /> Expired Certifications — Action Required
          </div>
          {expiredCerts.map(({ course, completion }) => (
            <div key={course.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: `1px solid ${C.errorBdr}` }}>
              <div>
                <span style={{ fontWeight: 600, color: C.teal700 }}>{course.name}</span>
                <span style={{ fontSize: 13, color: C.gray400, marginLeft: 8 }}>Expired {completion.certExpires}</span>
              </div>
              <button style={{ ...S.btnPrimary, ...S.btnSmall, background: C.error }} onClick={() => setView({ type: "course", courseId: course.id })}>
                <Icons.RefreshCw /> Recertify
              </button>
            </div>
          ))}
        </div>
      )}

      {expiringCerts.length > 0 && (
        <div style={{ ...S.card, borderLeft: `4px solid ${C.warning}`, background: C.warningBg }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, color: C.warning, fontWeight: 600, fontSize: 15, marginBottom: 8 }}>
            <Icons.Clock /> Certifications Expiring Soon
          </div>
          {expiringCerts.map(({ course, daysLeft }) => (
            <div key={course.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0" }}>
              <span style={{ fontWeight: 500, color: C.teal700 }}>{course.name} — <span style={{ color: C.warning }}>{daysLeft} days remaining</span></span>
              <button style={{ ...S.btnSecondary, ...S.btnSmall }} onClick={() => setView({ type: "course", courseId: course.id })}>
                Start Review
              </button>
            </div>
          ))}
        </div>
      )}

      {/* KPIs */}
      <div style={{ ...S.row, marginBottom: 24 }}>
        <div style={S.kpiCard}>
          <div style={S.kpiLabel}>Overall Progress</div>
          <div style={S.kpiValue}>{totalRequired > 0 ? Math.round((totalCompleted / totalRequired) * 100) : 0}%</div>
        </div>
        <div style={S.kpiCard}>
          <div style={S.kpiLabel}>Courses Completed</div>
          <div style={S.kpiValue}>{totalCompleted}</div>
        </div>
        <div style={S.kpiCard}>
          <div style={S.kpiLabel}>Remaining</div>
          <div style={S.kpiValue}>{totalRequired - totalCompleted}</div>
        </div>
        <div style={S.kpiCard}>
          <div style={S.kpiLabel}>Expired Certs</div>
          <div style={{ ...S.kpiValue, color: expiredCerts.length > 0 ? C.error : C.success }}>{expiredCerts.length}</div>
        </div>
      </div>

      {/* Assigned Courses (from supervisor) */}
      {myAssignments.length > 0 && (
        <div style={{ ...S.card, borderLeft: `3px solid ${C.error}`, marginBottom: 16 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <Icons.Alert />
              <span style={{ fontSize: 16, fontWeight: 600, color: C.teal700 }}>Assigned Courses</span>
            </div>
            <span style={S.badge("error")}>{myAssignments.length} Assigned</span>
          </div>
          {myAssignments.map(assignment => {
            const course = courses.find(c => c.id === assignment.courseId);
            if (!course) return null;
            const isOverdue = assignment.dueDate && assignment.dueDate < new Date().toISOString().split("T")[0];
            return (
              <div
                key={assignment.id}
                style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "12px 14px", borderBottom: `1px solid ${C.gray100}`, borderRadius: 4, background: isOverdue ? C.errorBg : "transparent" }}
              >
                <div
                  onClick={() => course.status === "Active" && setView({ type: "course", courseId: course.id })}
                  style={{ cursor: course.status === "Active" ? "pointer" : "default", flex: 1 }}
                >
                  <div style={{ fontSize: 14, fontWeight: 600, color: C.teal700, display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
                    {course.name}
                    {isOverdue && <span style={{ ...S.badge("error"), fontSize: 10 }}>OVERDUE</span>}
                  </div>
                  <div style={{ fontSize: 12, color: C.gray400, marginTop: 2 }}>
                    Assigned by {assignment.assignedBy} on {assignment.assignedDate}
                    {assignment.dueDate && <span style={{ color: isOverdue ? C.error : C.warning, fontWeight: 600 }}> · Due {assignment.dueDate}</span>}
                  </div>
                  {assignment.notes && (
                    <div style={{ fontSize: 12, color: C.teal500, marginTop: 4, fontStyle: "italic" }}>"{assignment.notes}"</div>
                  )}
                </div>
                <span onClick={() => course.status === "Active" && setView({ type: "course", courseId: course.id })} style={{ cursor: "pointer", color: C.gray400 }}>
                  <Icons.ChevronRight />
                </span>
              </div>
            );
          })}
        </div>
      )}

      {/* Learning Paths */}
      {paths.map(path => {
        const progress = getPathProgress(path.id, user.id, completions, courses, learningPaths, user.role);
        const isCollapsed = collapsedPaths.includes(path.id);
        const { dueDate, status: dueStatus } = getPathDueStatus(path, user, completions, courses, learningPaths);
        return (
          <div key={path.id} style={S.card}>
            <div
              onClick={() => setCollapsedPaths(prev => prev.includes(path.id) ? prev.filter(p => p !== path.id) : [...prev, path.id])}
              style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", cursor: "pointer", flexWrap: "wrap", gap: 8 }}
            >
              <div style={{ display: "flex", alignItems: "flex-start", gap: 10, flex: 1 }}>
                <span style={{ color: C.gray400, marginTop: 3, transition: "transform 0.2s", transform: isCollapsed ? "rotate(0deg)" : "rotate(90deg)", flexShrink: 0 }}>
                  <Icons.ChevronRight />
                </span>
                <div>
                  <div style={{ fontSize: 17, fontWeight: 600, color: C.teal700, display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                    {path.required && <Icons.Shield />}
                    {path.name}
                    {dueDate && dueStatus === "overdue" && (
                      <span style={{ ...S.badge("error"), fontSize: 11, fontWeight: 600 }}>OVERDUE — was due {dueDate}</span>
                    )}
                    {dueDate && dueStatus === "due-soon" && (
                      <span style={{ ...S.badge("warning"), fontSize: 11, fontWeight: 600 }}>Due {dueDate}</span>
                    )}
                  </div>
                  <div style={{ fontSize: 13, color: C.gray400, marginTop: 2 }}>
                    {path.description}
                    {dueDate && dueStatus === "on-track" && (
                      <span style={{ marginLeft: 8, fontSize: 12, color: C.teal500 }}>· Due {dueDate}</span>
                    )}
                  </div>
                </div>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 10, flexShrink: 0 }}>
                {progress.total === 0 ? (
                  <span style={{ ...S.badge("warning"), background: "#FFF8E8", color: "#B8860B" }}>COMING SOON</span>
                ) : progress.pct === 100 ? (
                  <span style={S.badge("success")}>✓ Complete</span>
                ) : (
                  <span style={S.badge("warning")}>{progress.completed}/{progress.total} Courses</span>
                )}
              </div>
            </div>
            <div style={{ marginTop: 8, marginLeft: 24 }}>
              {progress.total > 0 ? <ProgressBar pct={progress.pct} /> : (
                <div style={{ fontSize: 13, color: C.gold700, fontStyle: "italic", padding: "4px 0" }}>Courses for this path are being developed and will appear here when published.</div>
              )}
            </div>
            {!isCollapsed && (
              <div style={{ marginTop: 12, marginLeft: 24 }}>
                {path.courseIds.filter(cid => { const c = courses.find(x => x.id === cid); return c && courseMatchesRole(c, user.role); }).map(cid => {
                const course = courses.find(c => c.id === cid);
                if (!course) return null;
                const latest = myCompletions.filter(c => c.courseId === cid && c.status === "passed").sort((a, b) => b.completedDate.localeCompare(a.completedDate))[0];
                const certStatus = getCertStatus(latest, course);
                const failed = myCompletions.find(c => c.courseId === cid && c.status === "failed");
                return (
                  <div
                    key={cid}
                    onClick={() => setView({ type: "course", courseId: cid })}
                    style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 12px", borderBottom: `1px solid ${C.gray100}`, cursor: "pointer", borderRadius: 4, transition: "background 0.15s" }}
                    onMouseEnter={e => e.currentTarget.style.background = C.teal50}
                    onMouseLeave={e => e.currentTarget.style.background = "transparent"}
                  >
                    <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                      {certStatus === "current" && <span style={{ color: C.success }}><Icons.Check /></span>}
                      {certStatus === "expired" && <span style={{ color: C.error }}><Icons.Alert /></span>}
                      {certStatus === "expiring" && <span style={{ color: C.warning }}><Icons.Clock /></span>}
                      {certStatus === "incomplete" && <span style={{ color: C.gray300 }}>○</span>}
                      <div>
                        <div style={{ fontSize: 14, fontWeight: 500, color: C.teal700 }}>{course.name}</div>
                        <div style={{ fontSize: 12, color: C.gray400 }}>
                          {course.durationMin} min
                          {latest && ` · Score: ${latest.score}%`}
                          {latest?.certExpires && ` · Expires: ${latest.certExpires}`}
                          {failed && !latest && ` · Last attempt: ${failed.score}% (retry required)`}
                        </div>
                      </div>
                    </div>
                    <Icons.ChevronRight />
                  </div>
                );
              })}
              </div>
            )}
          </div>
        );
      })}

      {/* ── Voluntary Courses ── */}
      {(() => {
        const requiredCourseIds = getRequiredCourseIds(user, learningPaths, courses);
        const myEnrollments = enrollments.filter(e => e.employeeId === user.id && !requiredCourseIds.includes(e.courseId));
        if (myEnrollments.length === 0) return null;

        return (
          <div style={{ ...S.card, borderLeft: `3px solid ${C.teal400}` }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
              <div>
                <div style={{ fontSize: 17, fontWeight: 600, color: C.teal700, display: "flex", alignItems: "center", gap: 8 }}>
                  <Icons.BookOpen /> Voluntary Courses
                </div>
                <div style={{ fontSize: 13, color: C.gray400, marginTop: 2 }}>
                  Self-enrolled courses outside your required learning paths. These do not affect your compliance status.
                </div>
              </div>
              <span style={S.badge("info")}>{myEnrollments.length} Enrolled</span>
            </div>
            <div>
              {myEnrollments.map(enrollment => {
                const course = courses.find(c => c.id === enrollment.courseId);
                if (!course) return null;
                const latest = myCompletions.filter(c => c.courseId === course.id && c.status === "passed").sort((a, b) => b.completedDate.localeCompare(a.completedDate))[0];
                const certStatus = getCertStatus(latest, course);
                const failed = myCompletions.find(c => c.courseId === course.id && c.status === "failed");
                return (
                  <div
                    key={course.id}
                    style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 12px", borderBottom: `1px solid ${C.gray100}`, borderRadius: 4 }}
                  >
                    <div
                      onClick={() => setView({ type: "course", courseId: course.id })}
                      style={{ display: "flex", alignItems: "center", gap: 10, cursor: "pointer", flex: 1 }}
                    >
                      {certStatus === "current" && <span style={{ color: C.success }}><Icons.Check /></span>}
                      {certStatus === "expired" && <span style={{ color: C.error }}><Icons.Alert /></span>}
                      {certStatus === "expiring" && <span style={{ color: C.warning }}><Icons.Clock /></span>}
                      {certStatus === "incomplete" && <span style={{ color: C.gray300 }}>○</span>}
                      <div>
                        <div style={{ fontSize: 14, fontWeight: 500, color: C.teal700 }}>{course.name}</div>
                        <div style={{ fontSize: 12, color: C.gray400 }}>
                          {course.category} · {course.durationMin} min
                          {latest && ` · Score: ${latest.score}%`}
                          {failed && !latest && ` · Last attempt: ${failed.score}% (retry required)`}
                        </div>
                      </div>
                    </div>
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <button
                        onClick={(e) => { e.stopPropagation(); onUnenroll(user.id, course.id); }}
                        style={{ ...S.btnSecondary, ...S.btnSmall, fontSize: 12, padding: "4px 10px", color: C.gray400 }}
                        title="Remove from my voluntary courses"
                      >
                        Unenroll
                      </button>
                      <span onClick={() => setView({ type: "course", courseId: course.id })} style={{ cursor: "pointer" }}>
                        <Icons.ChevronRight />
                      </span>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        );
      })()}
    </div>
  );
}

// ============================================================
// COURSE VIEW (lessons + quiz launcher)
// ============================================================
function CourseView({ courseId, user, completions, setCompletions, onQuizSubmit, onBack, mobile }) {
  const { courses, lessons: allLessons, quizzes } = useData();
  const course = courses.find(c => c.id === courseId);
  if (!course) return <div style={{ padding: 40, textAlign: "center", color: C.gray400 }}>Course not found.</div>;
  const lessons = allLessons.filter(l => l.courseId === courseId).sort((a, b) => a.order - b.order);
  const [watchedLessons, setWatchedLessons] = useState(() => {
    // Restore from sessionStorage on mount
    try {
      const saved = sessionStorage.getItem(`ns_watched_${courseId}_${user?.id}`);
      if (saved) return new Set(JSON.parse(saved));
    } catch {}
    return new Set();
  });
  const [activeLesson, setActiveLesson] = useState(null);
  const [showQuiz, setShowQuiz] = useState(false);
  const myAttempts = completions.filter(c => c.employeeId === user.id && c.courseId === courseId);
  const bestPass = myAttempts.filter(c => c.status === "passed").sort((a, b) => b.score - a.score)[0];
  const allWatched = lessons.length > 0 && lessons.every(l => watchedLessons.has(l.id));

  // Persist watchedLessons to sessionStorage whenever they change
  useEffect(() => {
    if (watchedLessons.size > 0 && user?.id) {
      try { sessionStorage.setItem(`ns_watched_${courseId}_${user.id}`, JSON.stringify([...watchedLessons])); } catch {}
    }
  }, [watchedLessons]);

  // If previously passed (and not expired), mark all as watched
  useEffect(() => {
    if (bestPass) {
      const certStatus = getCertStatus(bestPass, course);
      if (certStatus === "current" || certStatus === "expiring") {
        setWatchedLessons(new Set(lessons.map(l => l.id)));
      }
    }
  }, []);

  return (
    <div>
      <button onClick={onBack} style={{ ...S.btnSecondary, marginBottom: 16, gap: 6 }}><Icons.Back /> Back to My Training</button>

      <div style={S.card}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 8, marginBottom: 16 }}>
          <div>
            <span style={{ ...S.badge("info"), marginBottom: 8 }}>{course.category}</span>
            <div style={{ fontSize: 22, fontWeight: 700, color: C.teal700, marginTop: 8 }}>{course.name}</div>
            <div style={{ fontSize: 14, color: C.gray400, marginTop: 4 }}>{course.description}</div>
            <div style={{ fontSize: 13, color: C.gray400, marginTop: 8, display: "flex", gap: 16, flexWrap: "wrap" }}>
              <span style={{ display: "flex", alignItems: "center", gap: 4 }}><Icons.Clock /> {course.durationMin} min total</span>
              <span style={{ display: "flex", alignItems: "center", gap: 4 }}><Icons.Play /> {lessons.length} lessons</span>
              {course.recertDays && <span style={{ display: "flex", alignItems: "center", gap: 4 }}><Icons.RefreshCw /> Recertification: every {course.recertDays} days</span>}
            </div>
          </div>
          {bestPass && (
            <div style={{ textAlign: "right" }}>
              <span style={S.badge(getCertStatus(bestPass, course) === "expired" ? "error" : "success")}>
                {getCertStatus(bestPass, course) === "expired" ? "EXPIRED" : `PASSED — ${bestPass.score}%`}
              </span>
              {bestPass.certExpires && (
                <div style={{ fontSize: 12, color: C.gray400, marginTop: 4 }}>
                  {getCertStatus(bestPass, course) === "expired" ? "Expired" : "Expires"}: {bestPass.certExpires}
                </div>
              )}
            </div>
          )}
        </div>
      </div>

      {/* Active lesson player */}
      {activeLesson && (
        <div style={{ ...S.card, borderLeft: `4px solid ${C.gold500}` }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
            <div style={{ fontSize: 16, fontWeight: 600, color: C.teal700 }}>
              Lesson {activeLesson.order}: {activeLesson.title}
            </div>
            <button onClick={() => setActiveLesson(null)} style={{ background: "none", border: "none", cursor: "pointer", color: C.gray400 }}><Icons.X /></button>
          </div>
          {/* Video embed */}
          {activeLesson.videoUrl && (() => {
            const url = activeLesson.videoUrl;
            const ytMatch = url.match(/(?:youtube\.com\/watch\?v=|youtu\.be\/|youtube\.com\/embed\/)([^&?#]+)/);
            if (ytMatch) return <div style={{ position: "relative", paddingBottom: mobile ? "56.25%" : "50%", height: 0, borderRadius: 6, overflow: "hidden", marginBottom: 16 }}><iframe src={`https://www.youtube.com/embed/${ytMatch[1]}`} style={{ position: "absolute", top: 0, left: 0, width: "100%", height: "100%", border: "none" }} allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" allowFullScreen /></div>;
            const vmMatch = url.match(/vimeo\.com\/(\d+)/);
            if (vmMatch) return <div style={{ position: "relative", paddingBottom: mobile ? "56.25%" : "50%", height: 0, borderRadius: 6, overflow: "hidden", marginBottom: 16 }}><iframe src={`https://player.vimeo.com/video/${vmMatch[1]}`} style={{ position: "absolute", top: 0, left: 0, width: "100%", height: "100%", border: "none" }} allow="autoplay; fullscreen; picture-in-picture" allowFullScreen /></div>;
            if (url.includes("sharepoint.com") || url.includes("microsoftstream.com")) return <div style={{ position: "relative", paddingBottom: mobile ? "56.25%" : "50%", height: 0, borderRadius: 6, overflow: "hidden", marginBottom: 16 }}><iframe src={url.includes("embed") ? url : url.replace("/video/", "/embed/video/")} style={{ position: "absolute", top: 0, left: 0, width: "100%", height: "100%", border: "none" }} allowFullScreen /></div>;
            if (url.match(/\.(mp4|webm|ogg)($|\?)/i)) return <div style={{ borderRadius: 6, overflow: "hidden", marginBottom: 16 }}><video src={url} controls controlsList="nodownload" onContextMenu={e => e.preventDefault()} style={{ width: "100%", maxHeight: mobile ? 240 : 480, background: C.dark }} /></div>;
            return <div style={{ position: "relative", paddingBottom: mobile ? "56.25%" : "50%", height: 0, borderRadius: 6, overflow: "hidden", marginBottom: 16 }}><iframe src={url} style={{ position: "absolute", top: 0, left: 0, width: "100%", height: "100%", border: "none" }} allowFullScreen /></div>;
          })()}
          {/* PowerPoint — opens in SharePoint viewer (new tab) */}
          {activeLesson.documentUrl && (() => {
            const url = activeLesson.documentUrl;
            // Convert any embed URL to interactivepreview for best viewing experience
            let viewUrl = url;
            if (url.includes("action=embedview")) viewUrl = url.replace("action=embedview", "action=interactivepreview");
            else if (url.includes("action=")) viewUrl = url.replace(/action=\w+/, "action=interactivepreview");
            else viewUrl = url + (url.includes("?") ? "&" : "?") + "action=interactivepreview";
            return (
              <div style={{
                background: `linear-gradient(135deg, ${C.headerBg} 0%, #243F4A 100%)`, borderRadius: 6,
                padding: mobile ? "28px 20px" : "36px 32px", marginBottom: 16,
                display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center",
                border: `1px solid ${C.gold500}20`, position: "relative", overflow: "hidden",
              }}>
                <div style={{ position: "absolute", top: 0, left: 0, right: 0, height: 2, background: C.gold500 }} />
                <div style={{ fontSize: 32, marginBottom: 8, opacity: 0.9 }}>📊</div>
                <div style={{ fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.08em", color: C.gold500, marginBottom: 4 }}>Course Presentation</div>
                <div style={{ fontSize: 15, color: "#FFFFFF", fontWeight: 600, marginBottom: 16, textAlign: "center" }}>{activeLesson.title}</div>
                <button onClick={() => window.open(viewUrl, "_blank")} style={{
                  display: "inline-flex", alignItems: "center", gap: 8, padding: "12px 28px",
                  fontSize: 14, fontWeight: 600, fontFamily: "'Source Sans 3',sans-serif",
                  color: C.headerBg, background: C.gold500, border: "none", borderRadius: 4,
                  cursor: "pointer", transition: "background 0.2s",
                }}>
                  ▶ Open Presentation
                </button>
                <div style={{ fontSize: 11, color: C.teal300, marginTop: 10 }}>Opens in SharePoint • {activeLesson.durationMin || "—"} min</div>
              </div>
            );
          })()}
          {/* No content yet */}
          {!activeLesson.videoUrl && !activeLesson.documentUrl && (
            <div style={{
              background: C.dark, borderRadius: 6, height: mobile ? 120 : 180, display: "flex", flexDirection: "column",
              alignItems: "center", justifyContent: "center", color: C.gray300, marginBottom: 16
            }}>
              <Icons.Clock />
              <div style={{ marginTop: 8, fontSize: 14 }}>Content coming soon</div>
              <div style={{ fontSize: 12, color: C.gray400, marginTop: 4 }}>{activeLesson.durationMin} minutes</div>
            </div>
          )}
          {/* Supplemental documents — opens as read-only view */}
          {activeLesson.supplements && activeLesson.supplements.length > 0 && (
            <div style={{ marginBottom: 16 }}>
              {activeLesson.supplements.length > 1 && <div style={{ fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.06em", color: C.gold700, marginBottom: 6 }}>Supplemental Materials</div>}
              {activeLesson.supplements.map((sup, si) => (
                <div key={si} style={{ display: "flex", alignItems: "center", gap: 10, padding: "12px 16px", background: C.gold50, borderRadius: 4, border: `1px solid ${C.gold100}`, marginBottom: 6 }}>
                  <Icons.Doc />
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 14, fontWeight: 500, color: C.teal700 }}>{sup.title || "Supplemental Material"}</div>
                    <div style={{ fontSize: 12, color: C.gray400 }}>Worksheet / reference document</div>
                  </div>
                  <button onClick={() => {
                    let url = sup.url;
                    if (url.includes("sharepoint.com") && !url.includes("?")) url = url + "?web=1";
                    else if (url.includes("sharepoint.com") && !url.includes("web=1")) url = url + "&web=1";
                    window.open(url, "_blank");
                  }} style={{ ...S.btnSecondary, ...S.btnSmall, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 4, border: `1px solid ${C.gold500}`, color: C.gold700 }}>📄 View Document</button>
                </div>
              ))}
            </div>
          )}
          <button
            onClick={() => {
              setWatchedLessons(prev => new Set([...prev, activeLesson.id]));
              const nextLesson = lessons.find(l => l.order === activeLesson.order + 1);
              if (nextLesson) setActiveLesson(nextLesson);
              else setActiveLesson(null);
            }}
            style={{ ...S.btnPrimary, marginTop: 16 }}
          >
            <Icons.Check /> Mark Complete {lessons.find(l => l.order === activeLesson.order + 1) ? "& Next Lesson" : ""}
          </button>
        </div>
      )}

      {/* Lesson list */}
      <div style={S.card}>
        <div style={S.cardTitle}>Course Lessons</div>
        {lessons.map(lesson => {
          const watched = watchedLessons.has(lesson.id);
          const isActive = activeLesson?.id === lesson.id;
          return (
            <div
              key={lesson.id}
              onClick={() => setActiveLesson(lesson)}
              style={{
                display: "flex", alignItems: "center", gap: 12, padding: "12px", borderRadius: 6, cursor: "pointer",
                background: isActive ? C.teal50 : "transparent", borderBottom: `1px solid ${C.gray100}`,
                transition: "background 0.15s"
              }}
              onMouseEnter={e => { if (!isActive) e.currentTarget.style.background = C.gray100; }}
              onMouseLeave={e => { if (!isActive) e.currentTarget.style.background = "transparent"; }}
            >
              <div style={{
                width: 32, height: 32, borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center",
                background: watched ? C.successBg : C.gray100, color: watched ? C.success : C.gray400, fontWeight: 600, fontSize: 14,
                border: `2px solid ${watched ? C.success : C.gray200}`, flexShrink: 0
              }}>
                {watched ? <Icons.Check /> : lesson.order}
              </div>
              <div style={{ flex: 1 }}>
                <div style={{ fontSize: 14, fontWeight: 500, color: C.teal700 }}>{lesson.title}</div>
                <div style={{ fontSize: 12, color: C.gray400, display: "flex", gap: 10, marginTop: 2 }}>
                  <span>{lesson.durationMin} min</span>
                  {lesson.videoUrl && <span style={{ display: "flex", alignItems: "center", gap: 3 }}><Icons.Play /> Video</span>}
                  {lesson.documentUrl && <span style={{ display: "flex", alignItems: "center", gap: 3 }}><Icons.Doc /> Slides</span>}
                  {lesson.supplements && lesson.supplements.length > 0 && <span style={{ display: "flex", alignItems: "center", gap: 3 }}><Icons.Download /> {lesson.supplements.length > 1 ? `${lesson.supplements.length} Docs` : "Worksheet"}</span>}
                  {!lesson.videoUrl && !lesson.documentUrl && <span style={{ color: C.gold500 }}>Content pending</span>}
                </div>
              </div>
              {isActive ? <span style={{ fontSize: 12, fontWeight: 600, color: C.gold600 }}>PLAYING</span> : <Icons.Play />}
            </div>
          );
        })}
      </div>

      {/* Quiz section */}
      <div style={{ ...S.card, borderLeft: `4px solid ${allWatched ? C.success : C.gray200}` }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
          <div>
            <div style={{ fontSize: 16, fontWeight: 600, color: C.teal700, display: "flex", alignItems: "center", gap: 8 }}>
              <Icons.Trophy /> Course Assessment
            </div>
            <div style={{ fontSize: 13, color: C.gray400, marginTop: 4 }}>
              {allWatched
                ? `All lessons complete. Score ${CONFIG.passingScore}% or higher to pass.`
                : `Complete all ${lessons.length} lessons to unlock the quiz.`
              }
            </div>
            {myAttempts.length > 0 && (
              <div style={{ fontSize: 13, color: C.gray400, marginTop: 4 }}>
                Previous attempts: {myAttempts.map(a => (
                  <span key={a.id} style={{ ...S.badge(a.status === "passed" ? "success" : "error"), marginLeft: 4 }}>{a.score}%</span>
                ))}
              </div>
            )}
          </div>
          <button
            disabled={!allWatched}
            onClick={() => setShowQuiz(true)}
            style={{ ...S.btnPrimary, opacity: allWatched ? 1 : 0.4, cursor: allWatched ? "pointer" : "not-allowed" }}
          >
            {bestPass && getCertStatus(bestPass, course) !== "expired" ? "Retake Quiz" : "Start Quiz"}
          </button>
        </div>
        {allWatched && showQuiz && (
          <div style={{ marginTop: 12 }}>
            <QuizView
              courseId={courseId}
              user={user}
              completions={completions}
              setCompletions={setCompletions}
              onQuizSubmit={onQuizSubmit}
              inline={true}
              autoStart={true}
            />
          </div>
        )}
      </div>
    </div>
  );
}

// ============================================================
// QUIZ VIEW
// ============================================================
function QuizView({ courseId, user, completions, setCompletions, onQuizSubmit, onBack, inline = false, autoStart = false }) {
  const { courses, quizzes } = useData();
  const quiz = quizzes[courseId];
  const course = courses.find(c => c.id === courseId);
  const [answers, setAnswers] = useState({});
  const [submitted, setSubmitted] = useState(false);
  const [score, setScore] = useState(null);
  const [started, setStarted] = useState(autoStart);

  if (!quiz || !course) return <div style={S.card}><div style={{ color: C.gray400 }}>No quiz available for this course.</div></div>;

  const handleSubmit = async () => {
    const total = quiz.questions.length;
    const correct = quiz.questions.filter(q => answers[q.id] === q.correct).length;
    const pct = Math.round((correct / total) * 100);
    setScore(pct);
    setSubmitted(true);

    const passed = pct >= CONFIG.passingScore;
    const answersJson = JSON.stringify(quiz.questions.map(q => ({ qId: q.id, selected: answers[q.id] || null, correct: q.correct })));
    if (onQuizSubmit) {
      await onQuizSubmit(user, course, pct, passed, answersJson);
    } else {
      // Fallback: local state only
      const newCompletion = {
        id: `comp_${Date.now()}`,
        employeeId: user.id,
        courseId: courseId,
        completedDate: TODAY,
        score: pct,
        status: passed ? "passed" : "failed",
        ...(passed && course.recertDays ? { certExpires: new Date(Date.now() + course.recertDays * 86400000).toISOString().split("T")[0] } : {}),
      };
      setCompletions(prev => [...prev, newCompletion]);
    }
  };

  if (!started) {
    return (
      <div style={inline ? {} : S.card}>
        <div style={{ textAlign: "center", padding: "20px 0" }}>
          <div style={{ fontSize: 18, fontWeight: 600, color: C.teal700, marginBottom: 8 }}>
            {course.name} — Assessment
          </div>
          <div style={{ fontSize: 14, color: C.gray400, marginBottom: 4 }}>
            {quiz.questions.length} questions · Passing score: {CONFIG.passingScore}%
          </div>
          <div style={{ fontSize: 13, color: C.gray400, marginBottom: 20 }}>
            You must answer at least {Math.ceil(quiz.questions.length * CONFIG.passingScore / 100)} of {quiz.questions.length} questions correctly.
          </div>
          <button style={S.btnPrimary} onClick={() => setStarted(true)}>Begin Assessment</button>
        </div>
      </div>
    );
  }

  if (submitted) {
    const passed = score >= CONFIG.passingScore;
    return (
      <div style={inline ? {} : S.card}>
        <div style={{ textAlign: "center", padding: "24px 0" }}>
          <div style={{
            width: 80, height: 80, borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center",
            margin: "0 auto 16px", background: passed ? C.successBg : C.errorBg, color: passed ? C.success : C.error,
            border: `3px solid ${passed ? C.success : C.error}`
          }}>
            {passed ? <span style={{ fontSize: 32 }}><Icons.Trophy /></span> : <span style={{ fontSize: 32 }}><Icons.X /></span>}
          </div>
          <div style={{ fontSize: 36, fontWeight: 700, fontFamily: mono, color: passed ? C.success : C.error }}>{score}%</div>
          <div style={{ fontSize: 16, fontWeight: 600, color: C.teal700, marginTop: 8 }}>
            {passed ? "Congratulations! You passed." : "You did not meet the passing score."}
          </div>
          <div style={{ fontSize: 14, color: C.gray400, marginTop: 4 }}>
            {passed
              ? course.recertDays ? `Your certification is valid until ${new Date(Date.now() + course.recertDays * 86400000).toISOString().split("T")[0]}.` : "This course is marked complete."
              : `Required: ${CONFIG.passingScore}%. You may review the material and retake the quiz.`
            }
          </div>
          {!passed && (
            <button style={{ ...S.btnPrimary, marginTop: 16 }} onClick={() => { setAnswers({}); setSubmitted(false); setScore(null); setStarted(false); }}>
              <Icons.RefreshCw /> Retake Quiz
            </button>
          )}

          {/* Show correct/incorrect breakdown */}
          <div style={{ marginTop: 24, textAlign: "left" }}>
            {quiz.questions.map((q, qi) => {
              const userAnswer = answers[q.id];
              const isCorrect = userAnswer === q.correct;
              return (
                <div key={q.id} style={{ padding: "12px 0", borderBottom: `1px solid ${C.gray100}` }}>
                  <div style={{ display: "flex", alignItems: "flex-start", gap: 8 }}>
                    <span style={{ color: isCorrect ? C.success : C.error, marginTop: 2, flexShrink: 0 }}>
                      {isCorrect ? <Icons.Check /> : <Icons.X />}
                    </span>
                    <div>
                      <div style={{ fontSize: 14, fontWeight: 500, color: C.teal700 }}>{qi + 1}. {q.text}</div>
                      {!isCorrect && (
                        <div style={{ fontSize: 13, marginTop: 4 }}>
                          <span style={{ color: C.error }}>Your answer: {q.options[userAnswer]}</span>
                          <span style={{ color: C.success, marginLeft: 12 }}>Correct: {q.options[q.correct]}</span>
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>
    );
  }

  const answeredCount = Object.keys(answers).length;
  const totalQ = quiz.questions.length;

  return (
    <div style={inline ? {} : S.card}>
      {/* Progress */}
      <div style={{ marginBottom: 20 }}>
        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, color: C.gray400, marginBottom: 6 }}>
          <span>Question {answeredCount} of {totalQ}</span>
          <span>Passing: {CONFIG.passingScore}%</span>
        </div>
        <ProgressBar pct={Math.round((answeredCount / totalQ) * 100)} color={C.gold500} />
      </div>

      {/* Questions */}
      {quiz.questions.map((q, qi) => {
        const optionEntries = Array.isArray(q.options) ? q.options.map((o, i) => [String.fromCharCode(65 + i), o]) : Object.entries(q.options);
        return (
        <div key={q.id} style={{ marginBottom: 24, paddingBottom: 20, borderBottom: `1px solid ${C.gray100}` }}>
          <div style={{ fontSize: 15, fontWeight: 600, color: C.teal700, marginBottom: 12 }}>
            {qi + 1}. {q.question || q.text}
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {optionEntries.map(([key, text]) => {
              const selected = answers[q.id] === key;
              return (
                <label
                  key={key}
                  style={{
                    display: "flex", alignItems: "center", gap: 10, padding: "10px 14px", borderRadius: 6, cursor: "pointer",
                    border: `1px solid ${selected ? C.gold500 : C.gray200}`, background: selected ? C.gold50 : C.white,
                    transition: "all 0.15s"
                  }}
                >
                  <input
                    type="radio"
                    name={q.id}
                    checked={selected}
                    onChange={() => setAnswers(prev => ({ ...prev, [q.id]: key }))}
                    style={{ accentColor: C.gold600 }}
                  />
                  <span style={{ fontSize: 14, color: C.teal700 }}>{key}. {text}</span>
                </label>
              );
            })}
          </div>
        </div>
        );
      })}

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
        <span style={{ fontSize: 13, color: answeredCount < totalQ ? C.warning : C.success }}>
          {answeredCount < totalQ ? `${totalQ - answeredCount} unanswered questions` : "All questions answered"}
        </span>
        <button
          disabled={answeredCount < totalQ}
          style={{ ...S.btnPrimary, opacity: answeredCount < totalQ ? 0.4 : 1, cursor: answeredCount < totalQ ? "not-allowed" : "pointer" }}
          onClick={handleSubmit}
        >
          Submit Assessment
        </button>
      </div>
    </div>
  );
}

// ============================================================
// COMPLIANCE DASHBOARD (Admin)
// ============================================================
function ComplianceDashboard({ completions, enrollments, visibleEmployeeIds, isAdmin, currentUser, mobile }) {
  const { employees, courses, learningPaths } = useData();
  const [filterRole, setFilterRole] = useState("All");
  const [filterStatus, setFilterStatus] = useState("All");
  const [expandedEmp, setExpandedEmp] = useState(null);

  // Scope: Admin sees all active employees (except themselves). Manager sees only their subordinates.
  // Exclude training-exempt roles (Owner/Operator) from compliance tracking
  const scopedEmployees = employees.filter(e => e.active && visibleEmployeeIds.includes(e.id) && !isTrainingExempt(e));
  const roles = ["All", ...new Set(scopedEmployees.map(e => e.role))];

  // Build compliance matrix from scoped employees only
  const matrix = scopedEmployees.map(emp => {
    const paths = getEmployeePaths(emp, learningPaths);
    const requiredCourses = [...new Set(paths.flatMap(p => p.courseIds))];
    const empCompletions = completions.filter(c => c.employeeId === emp.id);

    let completed = 0, expired = 0, missing = 0, expiring = 0;
    const courseStatuses = requiredCourses.map(cid => {
      const course = courses.find(c => c.id === cid);
      if (!course) return null;
      if (course.status !== "Active") return null; // Only show Active courses
      const latest = empCompletions.filter(c => c.courseId === cid && c.status === "passed").sort((a, b) => b.completedDate.localeCompare(a.completedDate))[0];
      const status = getCertStatus(latest, course);
      if (status === "current") completed++;
      else if (status === "expired") expired++;
      else if (status === "expiring") { expiring++; completed++; }
      else missing++;
      // Calculate due date for incomplete courses
      let courseDueDate = null;
      if (status === "none" || status === "expired") {
        const matchingPath = paths.find(p => p.required && p.dueDays && p.courseIds.includes(cid));
        if (matchingPath) courseDueDate = getCourseDueDate(course, matchingPath, emp);
      }
      return { course, status, completion: latest, dueDate: courseDueDate };
    }).filter(Boolean);

    // Due date tracking
    const pathDueStatuses = paths.filter(p => p.required).map(p => {
      const { dueDate, status: ds } = getPathDueStatus(p, emp, completions, courses, learningPaths);
      return { path: p, dueDate, dueStatus: ds };
    });
    const overduePaths = pathDueStatuses.filter(d => d.dueStatus === "overdue");
    const dueSoonPaths = pathDueStatuses.filter(d => d.dueStatus === "due-soon");

    const compliancePct = requiredCourses.length > 0 ? Math.round((completed / requiredCourses.length) * 100) : 100;
    const overallStatus = expired > 0 || overduePaths.length > 0 ? "non-compliant" : missing > 0 ? "in-progress" : expiring > 0 || dueSoonPaths.length > 0 ? "expiring" : "compliant";

    // Find who this person reports to for context (reportsTo can be ID or email)
    const manager = employees.find(e => e.id === emp.reportsTo || e.email === emp.reportsTo);

    return { emp, paths, requiredCourses, courseStatuses, completed, expired, missing, expiring, overduePaths, dueSoonPaths, compliancePct, overallStatus, manager };
  });

  const filtered = matrix.filter(m => {
    if (filterRole !== "All" && m.emp.role !== filterRole) return false;
    if (filterStatus === "Compliant" && m.overallStatus !== "compliant") return false;
    if (filterStatus === "Non-Compliant" && m.overallStatus !== "non-compliant") return false;
    if (filterStatus === "In Progress" && m.overallStatus !== "in-progress") return false;
    if (filterStatus === "Expiring" && m.overallStatus !== "expiring") return false;
    return true;
  });

  const totalEmployees = matrix.length;
  const compliant = matrix.filter(m => m.overallStatus === "compliant").length;
  const nonCompliant = matrix.filter(m => m.overallStatus === "non-compliant").length;
  const expiringSoon = matrix.filter(m => m.overallStatus === "expiring").length;

  const statusBadge = (status) => {
    const map = { "compliant": "success", "non-compliant": "error", "in-progress": "warning", "expiring": "info" };
    const labels = { "compliant": "COMPLIANT", "non-compliant": "NON-COMPLIANT", "in-progress": "IN PROGRESS", "expiring": "EXPIRING SOON" };
    return <span style={S.badge(map[status])}>{labels[status]}</span>;
  };

  const certBadge = (status) => {
    const map = { "current": "success", "expired": "error", "expiring": "warning", "incomplete": "neutral" };
    const labels = { "current": "CURRENT", "expired": "EXPIRED", "expiring": "EXPIRING", "incomplete": "NOT STARTED" };
    return <span style={S.badge(map[status])}>{labels[status]}</span>;
  };

  return (
    <div>
      {/* Scope Indicator */}
      <div style={{ ...S.card, borderLeft: `3px solid ${isAdmin ? C.gold500 : C.teal400}`, padding: "12px 20px", marginBottom: 20 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
          <Icons.Users />
          <div>
            <span style={{ fontSize: 14, fontWeight: 600, color: C.teal700 }}>
              {isAdmin ? "Organization View" : "Your Direct & Indirect Reports"}
            </span>
            <span style={{ fontSize: 13, color: C.gray400, marginLeft: 8 }}>
              {isAdmin
                ? `Viewing all ${totalEmployees} active employees`
                : `Viewing ${totalEmployees} report${totalEmployees !== 1 ? "s" : ""} in your chain of command`
              }
            </span>
          </div>
        </div>
      </div>

      {/* KPIs */}
      <div style={{ ...S.row, marginBottom: 24 }}>
        <div style={S.kpiCard}>
          <div style={S.kpiLabel}>{isAdmin ? "Total Staff" : "Your Reports"}</div>
          <div style={S.kpiValue}>{totalEmployees}</div>
        </div>
        <div style={S.kpiCard}>
          <div style={S.kpiLabel}>Fully Compliant</div>
          <div style={{ ...S.kpiValue, color: C.success }}>{compliant}</div>
        </div>
        <div style={S.kpiCard}>
          <div style={S.kpiLabel}>Non-Compliant</div>
          <div style={{ ...S.kpiValue, color: nonCompliant > 0 ? C.error : C.success }}>{nonCompliant}</div>
        </div>
        <div style={S.kpiCard}>
          <div style={S.kpiLabel}>Expiring Soon</div>
          <div style={{ ...S.kpiValue, color: expiringSoon > 0 ? C.warning : C.success }}>{expiringSoon}</div>
        </div>
      </div>

      {/* Compliance Rate */}
      <div style={{ ...S.card, marginBottom: 24 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
          <div style={{ fontSize: 16, fontWeight: 600, color: C.teal700 }}>
            {isAdmin ? "Organization Compliance Rate" : "Team Compliance Rate"}
          </div>
          <span style={{ fontSize: 24, fontWeight: 700, fontFamily: mono, color: compliant === totalEmployees ? C.success : C.warning }}>
            {totalEmployees > 0 ? Math.round((compliant / totalEmployees) * 100) : 0}%
          </span>
        </div>
        <ProgressBar pct={totalEmployees > 0 ? Math.round((compliant / totalEmployees) * 100) : 0} />
      </div>

      {/* Filters */}
      <div style={{ ...S.card }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12, marginBottom: 16 }}>
          <div style={S.cardTitle}>Employee Compliance Matrix</div>
          <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
            <div>
              <label style={{ ...S.label, fontSize: 12 }}>Role</label>
              <select style={{ ...S.select, width: "auto", minWidth: 160 }} value={filterRole} onChange={e => setFilterRole(e.target.value)}>
                {roles.map(r => <option key={r} value={r}>{r}</option>)}
              </select>
            </div>
            <div>
              <label style={{ ...S.label, fontSize: 12 }}>Status</label>
              <select style={{ ...S.select, width: "auto", minWidth: 160 }} value={filterStatus} onChange={e => setFilterStatus(e.target.value)}>
                {["All", "Compliant", "Non-Compliant", "In Progress", "Expiring"].map(s => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>
          </div>
        </div>

        {/* Employee rows */}
        {filtered.map(m => (
          <div key={m.emp.id} style={{ borderBottom: `1px solid ${C.gray100}` }}>
            <div
              onClick={() => setExpandedEmp(expandedEmp === m.emp.id ? null : m.emp.id)}
              style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px 0", cursor: "pointer", flexWrap: "wrap", gap: 8 }}
            >
              <div style={{ display: "flex", alignItems: "center", gap: 12, minWidth: 200 }}>
                <div style={{
                  width: 36, height: 36, borderRadius: "50%", background: C.teal50, display: "flex", alignItems: "center", justifyContent: "center",
                  color: C.teal700, fontWeight: 600, fontSize: 14, flexShrink: 0
                }}>
                  {m.emp.name.split(" ").map(n => n[0]).join("")}
                </div>
                <div>
                  <div style={{ fontSize: 14, fontWeight: 600, color: C.teal700 }}>{m.emp.name}</div>
                  <div style={{ fontSize: 12, color: C.gray400 }}>
                    {m.emp.role} · Hired {m.emp.hireDate}
                    {m.manager && ` · Reports to ${m.manager.name}`}
                  </div>
                </div>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 16, flexWrap: "wrap" }}>
                <div style={{ width: 120 }}>
                  <ProgressBar pct={m.compliancePct} height={6} label={false} />
                  <div style={{ fontSize: 11, color: C.gray400, marginTop: 2, textAlign: "center" }}>{m.compliancePct}% complete</div>
                </div>
                {statusBadge(m.overallStatus)}
                <span style={{ color: C.gray300, transform: expandedEmp === m.emp.id ? "rotate(90deg)" : "rotate(0)", transition: "transform 0.2s" }}>
                  <Icons.ChevronRight />
                </span>
              </div>
            </div>

            {/* Expanded detail */}
            {expandedEmp === m.emp.id && (
              <div style={{ padding: "0 0 16px 48px" }}>
                {/* Overdue/Due Soon Paths */}
                {(m.overduePaths.length > 0 || m.dueSoonPaths.length > 0) && (
                  <div style={{ marginBottom: 12, padding: "10px 14px", borderRadius: 6, background: m.overduePaths.length > 0 ? C.errorBg : C.warningBg, borderLeft: `3px solid ${m.overduePaths.length > 0 ? C.error : C.warning}` }}>
                    <div style={{ fontSize: 12, fontWeight: 600, color: m.overduePaths.length > 0 ? C.error : C.warning, marginBottom: 4 }}>
                      {m.overduePaths.length > 0 ? "OVERDUE TRAINING" : "TRAINING DUE SOON"}
                    </div>
                    {m.overduePaths.map(d => (
                      <div key={d.path.id} style={{ fontSize: 13, color: C.error, padding: "2px 0" }}>
                        {d.path.name} — was due {d.dueDate}
                      </div>
                    ))}
                    {m.dueSoonPaths.map(d => (
                      <div key={d.path.id} style={{ fontSize: 13, color: C.warning, padding: "2px 0" }}>
                        {d.path.name} — due {d.dueDate}
                      </div>
                    ))}
                  </div>
                )}

                {/* Required Courses Header */}
                <div style={{ fontSize: 12, fontWeight: 600, color: C.teal700, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 8 }}>
                  Required Courses
                </div>
                {mobile ? (
                  m.courseStatuses.map(cs => {
                    const isOverdue = cs.dueDate && cs.dueDate < new Date().toISOString().split("T")[0];
                    return (
                    <div key={cs.course.id} style={{ padding: "10px 0", borderBottom: `1px solid ${C.gray100}`, display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
                      <div>
                        <div style={{ fontSize: 13, fontWeight: 500, color: C.teal700 }}>{cs.course.name}</div>
                        <div style={{ fontSize: 12, color: C.gray400 }}>
                          {cs.completion ? `${cs.completion.score}% · ${cs.completion.completedDate}` : "Not started"}
                          {cs.completion?.certExpires && ` · Exp: ${cs.completion.certExpires}`}
                        </div>
                        {cs.dueDate && (
                          <div style={{ fontSize: 11, fontWeight: 600, color: isOverdue ? C.error : C.warning, marginTop: 2 }}>
                            {isOverdue ? `OVERDUE — was due ${cs.dueDate}` : `Due ${cs.dueDate}`}
                          </div>
                        )}
                      </div>
                      {certBadge(cs.status)}
                    </div>
                    );
                  })
                ) : (
                  <table style={{ width: "100%", borderCollapse: "collapse" }}>
                    <thead>
                      <tr>
                        <th style={{ ...S.th, fontSize: 12 }}>Course</th>
                        <th style={{ ...S.th, fontSize: 12 }}>Category</th>
                        <th style={{ ...S.th, fontSize: 12 }}>Score</th>
                        <th style={{ ...S.th, fontSize: 12 }}>Completed</th>
                        <th style={{ ...S.th, fontSize: 12 }}>Expires</th>
                        <th style={{ ...S.th, fontSize: 12 }}>Due</th>
                        <th style={{ ...S.th, fontSize: 12 }}>Status</th>
                      </tr>
                    </thead>
                    <tbody>
                      {m.courseStatuses.map(cs => {
                        const isOverdue = cs.dueDate && cs.dueDate < new Date().toISOString().split("T")[0];
                        return (
                        <tr key={cs.course.id}>
                          <td style={{ ...S.td, fontSize: 13, fontWeight: 500 }}>{cs.course.name}</td>
                          <td style={{ ...S.td, fontSize: 13 }}>{cs.course.category}</td>
                          <td style={{ ...S.td, fontSize: 13 }}>{cs.completion ? `${cs.completion.score}%` : "—"}</td>
                          <td style={{ ...S.td, fontSize: 13 }}>{cs.completion?.completedDate || "—"}</td>
                          <td style={{ ...S.td, fontSize: 13 }}>{cs.completion?.certExpires || "N/A"}</td>
                          <td style={{ ...S.td, fontSize: 13, color: isOverdue ? C.error : cs.dueDate ? C.warning : C.gray400, fontWeight: cs.dueDate ? 600 : 400 }}>{cs.dueDate || "—"}</td>
                          <td style={S.td}>{certBadge(cs.status)}</td>
                        </tr>
                        );
                      })}
                    </tbody>
                  </table>
                )}

                {/* Voluntary Courses for this employee */}
                {(() => {
                  const empRequiredIds = getRequiredCourseIds(m.emp, learningPaths, courses);
                  const empVoluntary = enrollments.filter(e => e.employeeId === m.emp.id && !empRequiredIds.includes(e.courseId));
                  if (empVoluntary.length === 0) return null;

                  const empCompletions = completions.filter(c => c.employeeId === m.emp.id);
                  return (
                    <div style={{ marginTop: 16 }}>
                      <div style={{ fontSize: 12, fontWeight: 600, color: C.teal400, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 8, display: "flex", alignItems: "center", gap: 6 }}>
                        <Icons.BookOpen /> Voluntary Enrollments ({empVoluntary.length})
                      </div>
                      {empVoluntary.map(enrollment => {
                        const course = courses.find(c => c.id === enrollment.courseId);
                        if (!course) return null;
                        const latest = empCompletions.filter(c => c.courseId === course.id && c.status === "passed").sort((a, b) => b.completedDate.localeCompare(a.completedDate))[0];
                        const status = getCertStatus(latest, course);
                        return (
                          <div key={course.id} style={{ padding: "8px 0", borderBottom: `1px solid ${C.gray100}`, display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
                            <div>
                              <div style={{ fontSize: 13, fontWeight: 500, color: C.teal600 }}>{course.name}</div>
                              <div style={{ fontSize: 12, color: C.gray400 }}>
                                {course.category} · Enrolled {enrollment.enrolledDate}
                                {latest && ` · Score: ${latest.score}%`}
                              </div>
                            </div>
                            {certBadge(status)}
                          </div>
                        );
                      })}
                    </div>
                  );
                })()}
              </div>
            )}
          </div>
        ))}

        {filtered.length === 0 && (
          <div style={{ padding: 24, textAlign: "center", color: C.gray400, fontSize: 14 }}>
            No employees match the selected filters.
          </div>
        )}
      </div>
    </div>
  );
}

// ============================================================
// TRAINING LIBRARY VIEW
// ============================================================
function TrainingLibraryView({ user, completions, enrollments, assignments, onEnroll, onUnenroll, onAssign, isManager, isAdmin, subordinateIds, setView, mobile }) {
  const { courses, lessons, quizzes, learningPaths, employees } = useData();
  const [filterCat, setFilterCat] = useState("All");
  const [filterType, setFilterType] = useState("All");
  const [search, setSearch] = useState("");
  const [assignModal, setAssignModal] = useState(null); // { courseId, courseName }
  const canAssign = isAdmin || isManager;
  const categories = ["All", ...new Set(courses.map(c => c.category))];
  const requiredCourseIds = getRequiredCourseIds(user, learningPaths, courses);
  const myEnrollmentIds = enrollments.filter(e => e.employeeId === user.id).map(e => e.courseId);

  const filtered = courses.filter(c => {
    if (c.status === "Archived") return false;
    if (filterCat !== "All" && c.category !== filterCat) return false;
    if (search && !c.name.toLowerCase().includes(search.toLowerCase()) && !c.description.toLowerCase().includes(search.toLowerCase())) return false;
    if (filterType === "Required" && !requiredCourseIds.includes(c.id)) return false;
    if (filterType === "Voluntary" && requiredCourseIds.includes(c.id)) return false;
    if (filterType === "Enrolled" && !requiredCourseIds.includes(c.id) && !myEnrollmentIds.includes(c.id)) return false;
    return true;
  });

  const myCompletions = completions.filter(c => c.employeeId === user.id);

  return (
    <div>
      {/* Explore banner */}
      <div style={{ ...S.card, borderLeft: `3px solid ${C.teal400}`, background: C.teal50, marginBottom: 20 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <Icons.BookOpen />
          <div>
            <div style={{ fontSize: 15, fontWeight: 600, color: C.teal700 }}>Expand Your Knowledge</div>
            <div style={{ fontSize: 13, color: C.gray600 }}>
              Browse courses outside your assigned learning paths and self-enroll to build new skills.
              Voluntary courses appear on your My Training page but do not affect your compliance status.
            </div>
          </div>
        </div>
      </div>

      <div style={{ ...S.card }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12, marginBottom: 16 }}>
          <div style={S.cardTitle}>Training Library</div>
          <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
            <input
              style={{ ...S.input, width: mobile ? "100%" : 240 }}
              placeholder="Search courses..."
              value={search}
              onChange={e => setSearch(e.target.value)}
            />
            <select style={{ ...S.select, width: "auto", minWidth: 140 }} value={filterCat} onChange={e => setFilterCat(e.target.value)}>
              {categories.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
            <select style={{ ...S.select, width: "auto", minWidth: 130 }} value={filterType} onChange={e => setFilterType(e.target.value)}>
              {["All", "Required", "Voluntary", "Enrolled"].map(t => <option key={t} value={t}>{t}</option>)}
            </select>
          </div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: mobile ? "1fr" : "repeat(auto-fill, minmax(340px, 1fr))", gap: 16 }}>
          {filtered.map(course => {
            const latest = myCompletions.filter(c => c.courseId === course.id && c.status === "passed").sort((a, b) => b.completedDate.localeCompare(a.completedDate))[0];
            const certStatus = getCertStatus(latest, course);
            const courseLessons = lessons.filter(l => l.courseId === course.id);
            const isRequired = requiredCourseIds.includes(course.id);
            const isEnrolled = myEnrollmentIds.includes(course.id);
            const isVoluntaryAvailable = !isRequired && !isEnrolled;

            return (
              <div
                key={course.id}
                style={{
                  ...S.card, marginBottom: 0, transition: "box-shadow 0.2s, transform 0.15s",
                  display: "flex", flexDirection: "column", justifyContent: "space-between",
                  ...(course.status === "Coming Soon" ? { borderLeft: `3px solid ${C.gold500}` } : {})
                }}
                onMouseEnter={e => { e.currentTarget.style.boxShadow = "0 4px 12px rgba(28,55,64,0.10)"; e.currentTarget.style.transform = "translateY(-2px)"; }}
                onMouseLeave={e => { e.currentTarget.style.boxShadow = "0 1px 3px rgba(28,55,64,0.06)"; e.currentTarget.style.transform = "translateY(0)"; }}
              >
                <div>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8, marginBottom: 8 }}>
                    <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                      <span style={S.badge("info")}>{course.category}</span>
                      {course.status === "Coming Soon" && <span style={{ ...S.badge("warning"), background: "#FFF8E8", color: "#B8860B" }}>COMING SOON</span>}
                      {isRequired && <span style={S.badge("warning")}>REQUIRED</span>}
                      {isEnrolled && course.status !== "Coming Soon" && <span style={{ ...S.badge("neutral"), color: C.teal400, background: C.teal50 }}>ENROLLED</span>}
                      {isEnrolled && course.status === "Coming Soon" && <span style={{ ...S.badge("neutral"), color: C.gold500, background: "#FFF8E8" }}>PRE-REGISTERED</span>}
                    </div>
                    {certStatus !== "incomplete" && (
                      <span style={S.badge(certStatus === "current" ? "success" : certStatus === "expired" ? "error" : "warning")}>
                        {certStatus === "current" ? "PASSED" : certStatus === "expired" ? "EXPIRED" : "EXPIRING"}
                      </span>
                    )}
                  </div>
                  <div
                    onClick={() => course.status !== "Coming Soon" && setView({ type: "course", courseId: course.id })}
                    style={{ cursor: course.status === "Coming Soon" ? "default" : "pointer" }}
                  >
                    <div style={{ fontSize: 16, fontWeight: 600, color: C.teal700, marginBottom: 6 }}>{course.code && <span style={{ fontSize: 12, color: C.gold500, marginRight: 6 }}>{course.code}</span>}{course.name}</div>
                    <div style={{ fontSize: 13, color: C.gray400, marginBottom: 12, lineHeight: 1.4 }}>{course.description}</div>
                  </div>
                </div>
                <div>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 12, color: C.gray400, borderTop: `1px solid ${C.gray100}`, paddingTop: 10, marginBottom: 10 }}>
                    <span style={{ display: "flex", alignItems: "center", gap: 4 }}><Icons.Clock /> {course.durationMin} min</span>
                    <span style={{ display: "flex", alignItems: "center", gap: 4 }}><Icons.Play /> {courseLessons.length} lessons</span>
                    {course.recertDays && <span style={{ display: "flex", alignItems: "center", gap: 4 }}><Icons.RefreshCw /> Annual</span>}
                  </div>

                  {/* Coming Soon — Pre-register / Already registered */}
                  {course.status === "Coming Soon" && !isEnrolled && !isRequired && (
                    <button
                      onClick={(e) => { e.stopPropagation(); onEnroll(user.id, course.id); }}
                      style={{ ...S.btnSecondary, width: "100%", justifyContent: "center", fontSize: 13, padding: "8px 16px", borderColor: C.gold500, color: "#B8860B" }}
                    >
                      Pre-Register \u2014 Notify Me When Available
                    </button>
                  )}
                  {course.status === "Coming Soon" && isEnrolled && !isRequired && (
                    <div style={{ display: "flex", gap: 8 }}>
                      <div style={{ flex: 1, textAlign: "center", fontSize: 13, color: C.gold500, padding: "8px 16px", background: "#FFF8E8", borderRadius: 4 }}>
                        You'll be notified when this course goes live
                      </div>
                      <button
                        onClick={(e) => { e.stopPropagation(); onUnenroll(user.id, course.id); }}
                        style={{ ...S.btnSecondary, fontSize: 13, padding: "8px 12px", color: C.gray400 }}
                      >
                        Cancel
                      </button>
                    </div>
                  )}
                  {course.status === "Coming Soon" && isRequired && (
                    <div style={{ textAlign: "center", fontSize: 13, color: C.gold500, padding: "8px 16px", background: "#FFF8E8", borderRadius: 4 }}>
                      Required \u2014 You'll be notified when this course is ready
                    </div>
                  )}

                  {/* Active courses — normal enrollment/start buttons */}
                  {course.status !== "Coming Soon" && isVoluntaryAvailable && (
                    <button
                      onClick={(e) => { e.stopPropagation(); onEnroll(user.id, course.id); }}
                      style={{ ...S.btnPrimary, width: "100%", justifyContent: "center", fontSize: 13, padding: "8px 16px" }}
                    >
                      <Icons.Plus /> Enroll \u2014 Add to My Training
                    </button>
                  )}
                  {course.status !== "Coming Soon" && isEnrolled && (
                    <div style={{ display: "flex", gap: 8 }}>
                      <button
                        onClick={() => setView({ type: "course", courseId: course.id })}
                        style={{ ...S.btnPrimary, flex: 1, justifyContent: "center", fontSize: 13, padding: "8px 16px" }}
                      >
                        <Icons.Play /> Continue
                      </button>
                      <button
                        onClick={(e) => { e.stopPropagation(); onUnenroll(user.id, course.id); }}
                        style={{ ...S.btnSecondary, fontSize: 13, padding: "8px 12px", color: C.gray400 }}
                      >
                        Unenroll
                      </button>
                    </div>
                  )}
                  {course.status !== "Coming Soon" && isRequired && certStatus === "incomplete" && (
                    <button
                      onClick={() => setView({ type: "course", courseId: course.id })}
                      style={{ ...S.btnPrimary, width: "100%", justifyContent: "center", fontSize: 13, padding: "8px 16px" }}
                    >
                      <Icons.Play /> Start Course
                    </button>
                  )}
                  {/* Assign button for managers/admins */}
                  {canAssign && course.status === "Active" && (
                    <button
                      onClick={(e) => { e.stopPropagation(); setAssignModal({ courseId: course.id, courseName: course.name }); }}
                      style={{ ...S.btnSecondary, width: "100%", justifyContent: "center", fontSize: 12, padding: "6px 16px", marginTop: 4, color: C.gold700, borderColor: C.gold500 }}
                    >
                      Assign to Employee
                    </button>
                  )}
                </div>
              </div>
            );
          })}
        </div>

        {filtered.length === 0 && (
          <div style={{ padding: 40, textAlign: "center", color: C.gray400 }}>No courses match your search.</div>
        )}
      </div>

      {/* Assign Course Modal */}
      {assignModal && (() => {
        const AssignModal = () => {
          const [selectedEmp, setSelectedEmp] = useState("");
          const [dueDate, setDueDate] = useState("");
          const [notes, setNotes] = useState("");
          const [saving, setSaving] = useState(false);
          // Admins see all active non-exempt employees, managers see subordinates
          const assignableEmployees = employees.filter(e => {
            if (!e.active || isTrainingExempt(e)) return false;
            if (isAdmin) return true;
            return subordinateIds.includes(e.id);
          });
          const handleSubmit = async () => {
            if (!selectedEmp) return alert("Please select an employee.");
            setSaving(true);
            await onAssign({ employeeId: selectedEmp, courseId: assignModal.courseId, dueDate: dueDate || null, notes });
            setSaving(false);
            setAssignModal(null);
          };
          return (
            <Modal title={`Assign: ${assignModal.courseName}`} onClose={() => setAssignModal(null)}>
              <FormField label="Assign To">
                <select style={S.select} value={selectedEmp} onChange={e => setSelectedEmp(e.target.value)}>
                  <option value="">— Select Employee —</option>
                  {assignableEmployees.sort((a,b) => a.name.localeCompare(b.name)).map(e => (
                    <option key={e.id} value={e.id}>{e.name} — {e.role}</option>
                  ))}
                </select>
              </FormField>
              <FormField label="Due Date (optional)" hint="Leave blank for no deadline">
                <input style={S.input} type="date" value={dueDate} onChange={e => setDueDate(e.target.value)} />
              </FormField>
              <FormField label="Notes (optional)" hint="Reason for assignment — visible to the employee">
                <textarea style={{ ...S.input, minHeight: 60, resize: "vertical" }} value={notes} onChange={e => setNotes(e.target.value)} placeholder="e.g. Retake required per coaching conversation on 3/10..." />
              </FormField>
              <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16, paddingTop: 12, borderTop: `1px solid ${C.gray100}` }}>
                <button style={S.btnSecondary} onClick={() => setAssignModal(null)}>Cancel</button>
                <button style={S.btnPrimary} onClick={handleSubmit} disabled={saving}>{saving ? "Assigning..." : "Assign & Notify"}</button>
              </div>
            </Modal>
          );
        };
        return <AssignModal />;
      })()}
    </div>
  );
}

// ============================================================
// MODAL COMPONENT
// ============================================================
function Modal({ title, onClose, width, children }) {
  return (
    <div style={{ position: "fixed", top: 0, left: 0, right: 0, bottom: 0, background: "rgba(28,55,64,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 9999, padding: 16 }} onClick={onClose}>
      <div style={{ background: C.white, borderRadius: 8, boxShadow: "0 8px 32px rgba(28,55,64,0.18)", width: "100%", maxWidth: width || 560, maxHeight: "90vh", overflow: "auto" }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "16px 20px", borderBottom: `1px solid ${C.gray100}` }}>
          <div style={{ fontSize: 17, fontWeight: 600, color: C.teal700 }}>{title}</div>
          <button onClick={onClose} style={{ background: "none", border: "none", fontSize: 20, cursor: "pointer", color: C.gray400, padding: "4px 8px" }}>✕</button>
        </div>
        <div style={{ padding: 20 }}>{children}</div>
      </div>
    </div>
  );
}
function FormField({ label, children, hint }) {
  return (<div style={{ marginBottom: 14 }}><label style={S.label}>{label}</label>{children}{hint && <div style={{ fontSize: 12, color: C.gray400, marginTop: 3 }}>{hint}</div>}</div>);
}
function FormRow({ children }) { return <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>{children}</div>; }
function SaveBar({ saving, onSave, onCancel, onDelete, deleteLabel }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 20, paddingTop: 16, borderTop: `1px solid ${C.gray100}` }}>
      <div>{onDelete && <button onClick={onDelete} style={{ ...S.btnDanger, ...S.btnSmall, opacity: saving ? 0.5 : 1 }} disabled={saving}>{deleteLabel || "Deactivate"}</button>}</div>
      <div style={{ display: "flex", gap: 8 }}>
        <button onClick={onCancel} style={{ ...S.btnSecondary, ...S.btnSmall }} disabled={saving}>Cancel</button>
        <button onClick={onSave} style={{ ...S.btnPrimary, ...S.btnSmall, opacity: saving ? 0.5 : 1 }} disabled={saving}>{saving ? "Saving..." : "Save"}</button>
      </div>
    </div>
  );
}

// ── EMPLOYEE FORM ──
function EmployeeForm({ item, onClose }) {
  const { employees, setEmployees, isLive, getToken } = useData();
  const isEdit = !!item;
  const [form, setForm] = useState({ name: item?.name || "", email: item?.email || "", role: item?.role || "Property Manager", appRole: item?.appRole || "Employee", reportsTo: item?.reportsTo || "", hireDate: item?.hireDate || new Date().toISOString().split("T")[0], active: item?.active !== false });
  const [saving, setSaving] = useState(false);
  const set = (k, v) => setForm(p => ({ ...p, [k]: v }));
  const roles = [...new Set(employees.map(e => e.role).filter(Boolean))].sort();
  const handleSave = async () => {
    if (!form.name.trim() || !form.email.trim()) return alert("Name and email are required.");
    setSaving(true);
    const fields = { Title: form.name.trim(), Email: form.email.trim().toLowerCase(), JobTitle: form.role, AccessLevel: form.appRole, ManagerEmail: form.reportsTo, StartDate: form.hireDate, EmployeeActive: form.active };
    try {
      if (isLive) {
        const token = await getToken();
        if (isEdit) { await spUpdate(token, CONFIG.lists.users, item.id, fields); setEmployees(prev => prev.map(e => e.id === item.id ? { ...e, name: fields.Title, email: fields.Email, role: fields.JobTitle, appRole: fields.AccessLevel, reportsTo: fields.ManagerEmail.toLowerCase(), hireDate: fields.StartDate, active: fields.EmployeeActive } : e)); }
        else { const res = await spCreate(token, CONFIG.lists.users, fields); setEmployees(prev => [...prev, { id: String(res.id), name: fields.Title, email: fields.Email, role: fields.JobTitle, appRole: fields.AccessLevel, reportsTo: fields.ManagerEmail.toLowerCase(), hireDate: fields.StartDate, active: fields.EmployeeActive }]); }
      }
      onClose();
    } catch (err) { alert("Save failed: " + err.message); }
    setSaving(false);
  };
  const handleToggleActive = async () => {
    if (!confirm(`${form.active ? "Deactivate" : "Reactivate"} ${form.name}?`)) return;
    setSaving(true);
    try { if (isLive) { const token = await getToken(); await spUpdate(token, CONFIG.lists.users, item.id, { EmployeeActive: !form.active }); } setEmployees(prev => prev.map(e => e.id === item.id ? { ...e, active: !form.active } : e)); onClose(); } catch (err) { alert("Failed: " + err.message); }
    setSaving(false);
  };
  const handleDelete = async () => {
    if (!confirm(`PERMANENTLY DELETE "${form.name}"? This cannot be undone. All completion records for this employee will be orphaned.`)) return;
    if (!confirm("Are you absolutely sure? This is a permanent delete.")) return;
    setSaving(true);
    try { if (isLive) { const token = await getToken(); await spDelete(token, CONFIG.lists.users, item.id); } setEmployees(prev => prev.filter(e => e.id !== item.id)); onClose(); } catch (err) { alert("Delete failed: " + err.message); }
    setSaving(false);
  };
  return (
    <Modal title={isEdit ? `Edit Employee \u2014 ${item.name}` : "Add Employee"} onClose={onClose}>
      <FormRow><FormField label="Full Name"><input style={S.input} value={form.name} onChange={e => set("name", e.target.value)} /></FormField><FormField label="Email"><input style={S.input} type="email" value={form.email} onChange={e => set("email", e.target.value)} /></FormField></FormRow>
      <FormRow>
        <FormField label="Role"><input style={S.input} list="role-options" value={form.role} onChange={e => set("role", e.target.value)} placeholder="Type or select..." /><datalist id="role-options">{roles.map(r => <option key={r} value={r} />)}</datalist></FormField>
        <FormField label="App Access Level"><select style={S.select} value={form.appRole} onChange={e => set("appRole", e.target.value)}><option value="Employee">Employee</option><option value="Admin">Admin</option></select></FormField>
      </FormRow>
      <FormRow>
        <FormField label="Reports To"><select style={S.select} value={form.reportsTo} onChange={e => set("reportsTo", e.target.value)}><option value="">{"\u2014"} None (Top Level) {"\u2014"}</option>{employees.filter(e => e.active && e.id !== item?.id).map(e => <option key={e.id} value={e.email}>{e.name} ({e.role})</option>)}</select></FormField>
        <FormField label="Hire Date"><input style={S.input} type="date" value={form.hireDate} onChange={e => set("hireDate", e.target.value)} /></FormField>
      </FormRow>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 16, paddingTop: 16, borderTop: `1px solid ${C.gray100}` }}>
        <div style={{ display: "flex", gap: 8 }}>
          {isEdit && <button style={{ ...S.btnSecondary, ...S.btnSmall }} onClick={handleToggleActive} disabled={saving}>{form.active ? "Deactivate" : "Reactivate"}</button>}
          {isEdit && <button style={{ ...S.btnSecondary, ...S.btnSmall, color: "#C44B3B", borderColor: "#C44B3B" }} onClick={handleDelete} disabled={saving}>Permanently Delete</button>}
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button style={S.btnSecondary} onClick={onClose} disabled={saving}>Cancel</button>
          <button style={S.btnPrimary} onClick={handleSave} disabled={saving}>{saving ? "Saving..." : "Save"}</button>
        </div>
      </div>
    </Modal>
  );
}

// ── COURSE FORM ──
function CourseForm({ item, onClose }) {
  const { employees, enrollments, learningPaths, setCourses, isLive, getToken } = useData();
  const isEdit = !!item;
  const wasComingSoon = isEdit && item.status === "Coming Soon";
  const categories = ["Onboarding", "Compliance", "Leasing", "Maintenance", "Operations", "Safety", "Financial", "Management"];
  const defaultRoles = ["Property Manager","Leasing Agent","Maintenance Technician","Service Manager","Area Director","Virtual Assistant"];
  const employeeRoles = [...new Set(employees.map(e => e.role).filter(Boolean))];
  const allRoles = [...new Set([...defaultRoles, ...employeeRoles, ...(item?.roles || [])])].sort();
  const [form, setForm] = useState({ name: item?.name || "", code: item?.code || "", description: item?.description || "", category: item?.category || "Onboarding", durationMin: item?.durationMin || 30, recertDays: item?.recertDays || "", passingScore: item?.passingScore || CONFIG.passingScore, sortOrder: item?.sortOrder || 999, status: item?.status || "Active", roles: item?.roles || [] });
  const [saving, setSaving] = useState(false);
  const set = (k, v) => setForm(p => ({ ...p, [k]: v }));
  const toggleCourseRole = (role) => { set("roles", form.roles.includes(role) ? form.roles.filter(r => r !== role) : [...form.roles, role]); };

  // Send "Course Now Available" email to pre-registered and learning path employees
  const sendGoLiveNotifications = async (token, courseId, courseName) => {
    // 1. Pre-registered employees (voluntary enrollments for this course)
    const preRegistered = enrollments.filter(e => e.courseId === courseId).map(e => employees.find(emp => emp.id === e.employeeId)).filter(Boolean);
    // 2. Employees whose learning path includes this course
    const pathsWithCourse = learningPaths.filter(p => p.courseIds.includes(courseId));
    const pathEmployees = [];
    for (const path of pathsWithCourse) {
      for (const emp of employees.filter(e => e.active)) {
        const empRoles = [emp.role];
        if (path.roles.includes("All") || path.roles.some(r => empRoles.includes(r))) {
          if (!pathEmployees.some(pe => pe.id === emp.id)) pathEmployees.push(emp);
        }
      }
    }
    // Combine and deduplicate
    const allRecipients = [...preRegistered];
    for (const emp of pathEmployees) { if (!allRecipients.some(r => r.id === emp.id)) allRecipients.push(emp); }
    // Send individual emails
    let sent = 0;
    for (const emp of allRecipients) {
      const isPreReg = preRegistered.some(pr => pr.id === emp.id);
      const empPaths = pathsWithCourse.filter(p => p.roles.includes("All") || p.roles.includes(emp.role));
      const exempt = isTrainingExempt(emp);
      // Calculate due date for this course based on the first matching required path
      let dueLine = "";
      if (!exempt && empPaths.length > 0) {
        const requiredPath = empPaths.find(p => p.required && p.dueDays);
        if (requiredPath && emp.hireDate) {
          const course = courses.find(c => c.id === courseId);
          if (course) {
            const dueDate = getCourseDueDate(course, requiredPath, emp);
            if (dueDate) dueLine = `<p>This course is due by <strong>${new Date(dueDate).toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" })}</strong>.</p>`;
          }
        }
      }
      const bodyHtml = `<p>Hi ${emp.name.split(" ")[0]},</p>` +
        `<p><strong>${courseName}</strong> is now available in NewShire University!</p>` +
        (isPreReg ? `<p>You pre-registered for this course and it's now ready for you to begin.</p>` : "") +
        (empPaths.length > 0 ? `<p>This course is part of your ${empPaths.map(p=>p.name).join(", ")} learning path${empPaths.length>1?"s":""}.</p>` : "") +
        dueLine +
        `<p>Log in to NewShire University to start the course.</p>`;
      try {
        await sendEmail(token, emp.email, `Course Now Available: ${courseName}`, emailTemplate(bodyHtml, `Course Now Available: ${courseName}`));
        sent++;
      } catch (e) { console.error(`Go-live email failed for ${emp.email}:`, e); }
    }
    if (sent > 0) alert(`Course published! ${sent} notification${sent > 1 ? "s" : ""} sent.`);
  };

  const handleSave = async () => {
    if (!form.name.trim()) return alert("Course name is required.");
    const goingLive = wasComingSoon && form.status === "Active";
    setSaving(true);
    const fields = { Title: form.name.trim(), CourseCode: form.code.trim(), CourseDescription: form.description, Category: form.category, DurationMin: parseInt(form.durationMin,10)||0, RecertDays: parseInt(form.recertDays,10)||0, PassingScore: parseInt(form.passingScore,10)||80, SortOrder: parseInt(form.sortOrder,10)||999, CourseActive: form.status !== "Archived", CourseStatus: form.status, CourseRoles: form.roles.join(",") };
    try {
      if (isLive) {
        const token = await getToken();
        if (isEdit) {
          await spUpdate(token, CONFIG.lists.courses, item.id, fields);
          setCourses(prev => prev.map(c => c.id === item.id ? { ...c, name: fields.Title, code: fields.CourseCode, description: fields.CourseDescription, category: fields.Category, durationMin: fields.DurationMin, recertDays: fields.RecertDays||null, passingScore: fields.PassingScore, sortOrder: fields.SortOrder, status: fields.CourseStatus, roles: form.roles } : c));
          // Fire go-live notifications if status changed from Coming Soon → Active
          if (goingLive) { sendGoLiveNotifications(token, item.id, fields.Title).catch(e => console.error("Go-live notifications failed:", e)); }
        }
        else { const res = await spCreate(token, CONFIG.lists.courses, fields); setCourses(prev => [...prev, { id: String(res.id), name: fields.Title, code: fields.CourseCode, description: fields.CourseDescription, category: fields.Category, durationMin: fields.DurationMin, recertDays: fields.RecertDays||null, passingScore: fields.PassingScore, sortOrder: fields.SortOrder, status: fields.CourseStatus, roles: form.roles }].sort((a,b) => a.sortOrder - b.sortOrder)); }
      }
      onClose();
    } catch (err) { alert("Save failed: " + err.message); }
    setSaving(false);
  };
  const handleDelete = async () => {
    if (!confirm(`PERMANENTLY DELETE "${form.name}"? This cannot be undone. All lessons and quiz questions for this course will be orphaned.`)) return;
    if (!confirm("Are you absolutely sure? This is a permanent delete.")) return;
    setSaving(true);
    try { if (isLive) { const token = await getToken(); await spDelete(token, CONFIG.lists.courses, item.id); } setCourses(prev => prev.filter(c => c.id !== item.id)); onClose(); } catch (err) { alert("Delete failed: " + err.message); }
    setSaving(false);
  };
  return (
    <Modal title={isEdit ? `Edit Course \u2014 ${item.name}` : "Add Course"} onClose={onClose}>
      <FormField label="Course Name"><input style={S.input} value={form.name} onChange={e => set("name", e.target.value)} /></FormField>
      <FormRow><FormField label="Course Code" hint="e.g. FHC 101, MNT 202"><input style={S.input} value={form.code} onChange={e => set("code", e.target.value)} placeholder="FHC 101" /></FormField><FormField label="Category"><select style={S.select} value={form.category} onChange={e => set("category", e.target.value)}>{categories.map(c => <option key={c} value={c}>{c}</option>)}</select></FormField></FormRow>
      <FormField label="Description"><textarea style={{ ...S.input, minHeight: 60 }} value={form.description} onChange={e => set("description", e.target.value)} /></FormField>
      <FormRow><FormField label="Duration (minutes)"><input style={S.input} type="number" value={form.durationMin} onChange={e => set("durationMin", e.target.value)} /></FormField><FormField label="Passing Score (%)" hint="Leave at 80 for default"><input style={S.input} type="number" value={form.passingScore} onChange={e => set("passingScore", e.target.value)} /></FormField></FormRow>
      <FormRow><FormField label="Recert Period (days)" hint="0 or blank = no recert"><input style={S.input} type="number" value={form.recertDays} onChange={e => set("recertDays", e.target.value)} placeholder="e.g. 365" /></FormField><FormField label="Sort Order" hint="Lower = first"><input style={S.input} type="number" value={form.sortOrder} onChange={e => set("sortOrder", e.target.value)} /></FormField></FormRow>
        <FormField label="Status">{wasComingSoon && form.status === "Active" && <div style={{fontSize:12,color:C.gold500,marginBottom:4}}>Changing to Active will notify all pre-registered employees and those with this course in their learning path.</div>}<select style={S.select} value={form.status} onChange={e => set("status", e.target.value)}><option value="Active">Active</option><option value="Coming Soon">Coming Soon</option><option value="Archived">Archived</option></select></FormField>
      <FormField label="Role Restrictions" hint={form.roles.length === 0 ? "No restrictions \u2014 all roles will see this course" : `${form.roles.length} role${form.roles.length > 1 ? "s" : ""} selected \u2014 only these roles will see this course in their learning path`}>
        <div style={{display:"flex",gap:6,flexWrap:"wrap",marginTop:4}}>
          <button onClick={() => set("roles", [])} style={{padding:"5px 12px",fontSize:13,borderRadius:9999,cursor:"pointer",fontFamily:"inherit",border:`1px solid ${form.roles.length===0?C.teal700:C.gray200}`,background:form.roles.length===0?C.teal50:C.white,color:form.roles.length===0?C.teal700:C.gray400,fontWeight:form.roles.length===0?600:400}}>{form.roles.length===0?"\u2713 ":""}All Roles</button>
          {allRoles.map(role => (
            <button key={role} onClick={()=>toggleCourseRole(role)} style={{padding:"5px 12px",fontSize:13,borderRadius:9999,cursor:"pointer",fontFamily:"inherit",border:`1px solid ${form.roles.includes(role)?C.teal700:C.gray200}`,background:form.roles.includes(role)?C.teal50:C.white,color:form.roles.includes(role)?C.teal700:C.gray400,fontWeight:form.roles.includes(role)?600:400}}>{form.roles.includes(role)?"\u2713 ":""}{role}</button>
          ))}
        </div>
      </FormField>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 16, paddingTop: 16, borderTop: `1px solid ${C.gray100}` }}>
        <div style={{ display: "flex", gap: 8 }}>
          {isEdit && <button style={{ ...S.btnSecondary, ...S.btnSmall, color: "#C44B3B", borderColor: "#C44B3B" }} onClick={handleDelete} disabled={saving}>Permanently Delete</button>}
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button style={{ ...S.btnSecondary }} onClick={onClose} disabled={saving}>Cancel</button>
          <button style={S.btnPrimary} onClick={handleSave} disabled={saving}>{saving ? (wasComingSoon && form.status === "Active" ? "Publishing..." : "Saving...") : (wasComingSoon && form.status === "Active" ? "Publish & Notify" : "Save")}</button>
        </div>
      </div>
    </Modal>
  );
}

// ── LEARNING PATH FORM ──
function PathForm({ item, onClose }) {
  const { employees, courses, setLearningPaths, isLive, getToken } = useData();
  const isEdit = !!item;
  // Dynamic roles: pull from employee list + hardcoded defaults + any roles already on this path
  const defaultRoles = ["All","Property Manager","Leasing Agent","Maintenance Technician","Service Manager","Area Director","Virtual Assistant"];
  const employeeRoles = [...new Set(employees.map(e => e.role).filter(Boolean))];
  const existingPathRoles = item?.roles || [];
  const allRoles = [...new Set([...defaultRoles, ...employeeRoles, ...existingPathRoles])].sort((a,b) => a === "All" ? -1 : b === "All" ? 1 : a.localeCompare(b));
  const [customRole, setCustomRole] = useState("");
  const [form, setForm] = useState({ name: item?.name||"", description: item?.description||"", roles: item?.roles||["All"], courseIds: item?.courseIds||[], required: item?.required!==false, dueDays: item?.dueDays||"" });
  const [saving, setSaving] = useState(false);
  const [roleList, setRoleList] = useState(allRoles);
  const set = (k,v) => setForm(p => ({...p,[k]:v}));
  const toggleRole = (role) => { if (role==="All"){set("roles",form.roles.includes("All")?[]:["All"]);return;} let next=form.roles.filter(r=>r!=="All"); next=next.includes(role)?next.filter(r=>r!==role):[...next,role]; set("roles",next.length===0?["All"]:next); };
  const addCustomRole = () => {
    const r = customRole.trim();
    if (!r) return;
    if (!roleList.includes(r)) setRoleList(prev => [...prev, r].sort((a,b) => a === "All" ? -1 : b === "All" ? 1 : a.localeCompare(b)));
    let next = form.roles.filter(x => x !== "All");
    if (!next.includes(r)) next.push(r);
    set("roles", next);
    setCustomRole("");
  };
  const toggleCourse = (cid) => { set("courseIds", form.courseIds.includes(cid)?form.courseIds.filter(c=>c!==cid):[...form.courseIds,cid]); };
  const handleSave = async () => {
    if (!form.name.trim()) return alert("Path name is required.");
    if (form.courseIds.length===0) return alert("Select at least one course.");
    setSaving(true);
    const fields = { Title: form.name.trim(), PathDescription: form.description, Roles: form.roles.join(","), CourseIDs: form.courseIds.join(","), Required: form.required, DueDays: parseInt(form.dueDays,10)||0, PathActive: true };
    try {
      if (isLive) {
        const token = await getToken();
        if (isEdit) { await spUpdate(token, CONFIG.lists.paths, item.id, fields); setLearningPaths(prev => prev.map(p => p.id===item.id ? {...p, name:fields.Title, description:fields.PathDescription, roles:form.roles, courseIds:form.courseIds, required:fields.Required, dueDays:fields.DueDays||null} : p)); }
        else { const res = await spCreate(token, CONFIG.lists.paths, fields); setLearningPaths(prev => [...prev, {id:String(res.id), name:fields.Title, description:fields.PathDescription, roles:form.roles, courseIds:form.courseIds, required:fields.Required, dueDays:fields.DueDays||null}]); }
      }
      onClose();
    } catch (err) { alert("Save failed: " + err.message); }
    setSaving(false);
  };
  const handleDelete = async () => {
    if (!confirm(`PERMANENTLY DELETE path "${form.name}"? This cannot be undone.`)) return;
    if (!confirm("Are you absolutely sure? This is a permanent delete.")) return;
    setSaving(true);
    try { if (isLive) { const token = await getToken(); await spDelete(token, CONFIG.lists.paths, item.id); } setLearningPaths(prev => prev.filter(p => p.id!==item.id)); onClose(); } catch (err) { alert("Delete failed: " + err.message); }
    setSaving(false);
  };
  return (
    <Modal title={isEdit ? `Edit Path \u2014 ${item.name}` : "Add Learning Path"} onClose={onClose} width={640}>
      <FormField label="Path Name"><input style={S.input} value={form.name} onChange={e => set("name", e.target.value)} /></FormField>
      <FormField label="Description"><textarea style={{...S.input,minHeight:60}} value={form.description} onChange={e => set("description", e.target.value)} /></FormField>
      <FormField label="Assigned Roles" hint="Select which roles this path applies to, or add a new role">
        <div style={{display:"flex",gap:6,flexWrap:"wrap",marginTop:4}}>{roleList.map(role => (
          <button key={role} onClick={()=>toggleRole(role)} style={{padding:"5px 12px",fontSize:13,borderRadius:9999,cursor:"pointer",fontFamily:"inherit",border:`1px solid ${form.roles.includes(role)?C.teal700:C.gray200}`,background:form.roles.includes(role)?C.teal50:C.white,color:form.roles.includes(role)?C.teal700:C.gray400,fontWeight:form.roles.includes(role)?600:400}}>{form.roles.includes(role)?"\u2713 ":""}{role}</button>
        ))}</div>
        <div style={{display:"flex",gap:6,marginTop:8}}>
          <input style={{...S.input,flex:1,marginTop:0}} value={customRole} onChange={e => setCustomRole(e.target.value)} placeholder="Add a new role..." onKeyDown={e => e.key==="Enter" && (e.preventDefault(), addCustomRole())} />
          <button style={{...S.btnSecondary,...S.btnSmall}} onClick={addCustomRole}>Add Role</button>
        </div>
      </FormField>
      <FormField label="Courses in Path" hint={`${form.courseIds.length} selected \u2014 click to toggle, order matches selection order`}>
        <div style={{maxHeight:200,overflowY:"auto",border:`1px solid ${C.gray200}`,borderRadius:4,marginTop:4}}>{courses.filter(c => c.status !== "Archived").map(course => (
          <div key={course.id} onClick={()=>toggleCourse(course.id)} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 12px",cursor:"pointer",borderBottom:`1px solid ${C.gray100}`,background:form.courseIds.includes(course.id)?C.teal50:C.white}}>
            <span style={{width:18,height:18,borderRadius:3,border:`2px solid ${form.courseIds.includes(course.id)?C.teal700:C.gray200}`,background:form.courseIds.includes(course.id)?C.teal700:C.white,color:C.white,display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,flexShrink:0}}>{form.courseIds.includes(course.id)?"\u2713":""}</span>
            <div><div style={{fontSize:14,fontWeight:500,color:C.teal700}}>{course.code && <span style={{fontSize:12,color:C.gold500,marginRight:6}}>{course.code}</span>}{course.name}</div><div style={{fontSize:12,color:C.gray400}}>{course.category} \u00b7 {course.durationMin} min{course.status === "Coming Soon" ? " \u00b7 Coming Soon" : ""}</div></div>
          </div>
        ))}</div>
      </FormField>
      <FormRow>
        <FormField label="Required?"><select style={S.select} value={form.required?"yes":"no"} onChange={e => set("required",e.target.value==="yes")}><option value="yes">Yes \u2014 Required for assigned roles</option><option value="no">No \u2014 Optional / Voluntary</option></select></FormField>
        <FormField label="Due Days from Hire" hint="0 or blank = no deadline"><input style={S.input} type="number" value={form.dueDays} onChange={e => set("dueDays", e.target.value)} placeholder="e.g. 30" /></FormField>
      </FormRow>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 16, paddingTop: 16, borderTop: `1px solid ${C.gray100}` }}>
        <div>{isEdit && <button style={{ ...S.btnSecondary, ...S.btnSmall, color: "#C44B3B", borderColor: "#C44B3B" }} onClick={handleDelete} disabled={saving}>Permanently Delete</button>}</div>
        <div style={{ display: "flex", gap: 8 }}>
          <button style={S.btnSecondary} onClick={onClose} disabled={saving}>Cancel</button>
          <button style={S.btnPrimary} onClick={handleSave} disabled={saving}>{saving ? "Saving..." : "Save"}</button>
        </div>
      </div>
    </Modal>
  );
}

// ── LESSON FORM ──
function LessonForm({ item, courseId, onClose }) {
  const { courses, lessons, setLessons, isLive, getToken } = useData();
  const isEdit = !!item;
  // Initialize supplements from item
  const initSupplements = () => {
    if (item?.supplements && item.supplements.length > 0) return item.supplements.map(s => ({...s}));
    return [];
  };
  const [form, setForm] = useState({ title: item?.title||"", courseId: item?.courseId||courseId||"", order: item?.order||(lessons.filter(l=>l.courseId===(item?.courseId||courseId)).length+1), durationMin: item?.durationMin||10, videoUrl: item?.videoUrl||"", documentUrl: item?.documentUrl||"", documentTitle: item?.documentTitle||"" });
  const [supplements, setSupplements] = useState(initSupplements);
  const [saving, setSaving] = useState(false);
  const set = (k,v) => setForm(p => ({...p,[k]:v}));
  const addSupplement = () => setSupplements(prev => [...prev, { title: "", url: "" }]);
  const updateSupplement = (i, k, v) => setSupplements(prev => prev.map((s, idx) => idx === i ? { ...s, [k]: v } : s));
  const removeSupplement = (i) => setSupplements(prev => prev.filter((_, idx) => idx !== i));
  const handleSave = async () => {
    if (!form.title.trim()||!form.courseId) return alert("Title and course are required.");
    setSaving(true);
    const fields = { Title: form.title.trim(), CourseIDLookupId: parseInt(form.courseId,10), LessonSortOrder: parseInt(form.order,10)||1, LessonDurationMin: parseInt(form.durationMin,10)||0, DocumentTitle: form.documentTitle || "" };
    if (form.videoUrl.trim()) fields.VideoURL = form.videoUrl.trim();
    if (form.documentUrl.trim()) {
      let cleanUrl = form.documentUrl.trim();
      const srcMatch = cleanUrl.match(/src=["']([^"']+)["']/);
      if (srcMatch) cleanUrl = srcMatch[1];
      cleanUrl = cleanUrl.replace(/&amp;/g, "&");
      fields.DocumentURL = cleanUrl;
    }
    // Save supplements as JSON array
    const validSupps = supplements.filter(s => s.url && s.url.trim());
    if (validSupps.length > 0) {
      fields.SupplementURL = JSON.stringify(validSupps.map(s => ({ title: s.title || "Supplemental Material", url: s.url.trim() })));
      fields.SupplementTitle = validSupps[0].title || "";
    } else {
      fields.SupplementURL = "";
      fields.SupplementTitle = "";
    }
    try {
      if (isLive) {
        const token = await getToken();
        const lessonData = { title: fields.Title, courseId: String(fields.CourseIDLookupId), order: fields.LessonSortOrder, durationMin: fields.LessonDurationMin, videoUrl: form.videoUrl || null, documentUrl: form.documentUrl || null, documentTitle: form.documentTitle || null, supplements: validSupps };
        if (isEdit) { await spUpdate(token, CONFIG.lists.lessons, item.id, fields); setLessons(prev => prev.map(l => l.id === item.id ? { ...l, ...lessonData } : l).sort((a, b) => a.order - b.order)); }
        else { const res = await spCreate(token, CONFIG.lists.lessons, fields); setLessons(prev => [...prev, { id: String(res.id), ...lessonData }].sort((a, b) => a.order - b.order)); }
      }
      onClose();
    } catch (err) { alert("Save failed: " + err.message); }
    setSaving(false);
  };
  const handleDelete = async () => {
    if (!confirm(`Delete lesson "${form.title}"? This cannot be undone.`)) return;
    setSaving(true);
    try { if (isLive) { const token = await getToken(); await spDelete(token, CONFIG.lists.lessons, item.id); } setLessons(prev => prev.filter(l => l.id!==item.id)); onClose(); } catch (err) { alert("Failed: " + err.message); }
    setSaving(false);
  };
  return (
    <Modal title={isEdit ? `Edit Lesson — ${item.title}` : "Add Lesson"} onClose={onClose}>
      <FormField label="Lesson Title"><input style={S.input} value={form.title} onChange={e => set("title", e.target.value)} /></FormField>
      <FormRow><FormField label="Course"><select style={S.select} value={form.courseId} onChange={e => set("courseId", e.target.value)}><option value="">— Select —</option>{courses.map(c => <option key={c.id} value={c.id}>{c.code ? `${c.code} — ` : ""}{c.name}</option>)}</select></FormField><FormField label="Sort Order"><input style={S.input} type="number" value={form.order} onChange={e => set("order", e.target.value)} /></FormField></FormRow>
      <FormField label="Duration (minutes)"><input style={S.input} type="number" value={form.durationMin} onChange={e => set("durationMin", e.target.value)} /></FormField>
      <FormField label="Video URL" hint="YouTube, Vimeo, SharePoint Stream, or direct video link"><input style={S.input} type="url" value={form.videoUrl} onChange={e => set("videoUrl", e.target.value)} placeholder="https://..." /></FormField>
      <FormField label="Presentation URL" hint="Paste the SharePoint embed code or URL — iframe tags and formatting are cleaned automatically"><input style={S.input} value={form.documentUrl} onChange={e => {
        let v = e.target.value;
        const srcMatch = v.match(/src=["']([^"']+)["']/);
        if (srcMatch) v = srcMatch[1];
        v = v.replace(/&amp;/g, "&");
        set("documentUrl", v);
      }} placeholder="Paste embed code or URL from SharePoint..." /></FormField>
      <FormField label="Document Title" hint="Display name for the presentation"><input style={S.input} value={form.documentTitle} onChange={e => set("documentTitle", e.target.value)} /></FormField>
      <div style={{ borderTop: `1px solid ${C.gray100}`, marginTop: 12, paddingTop: 12 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
          <div style={{ fontSize: 13, fontWeight: 600, color: C.teal400 }}>Supplemental Documents</div>
          <button style={{ ...S.btnSecondary, ...S.btnSmall, fontSize: 12 }} onClick={addSupplement}>+ Add Document</button>
        </div>
        {supplements.length === 0 && <div style={{ fontSize: 13, color: C.gray300, padding: "8px 0" }}>No supplemental documents attached. Click "Add Document" to attach worksheets, handouts, or reference materials.</div>}
        {supplements.map((sup, i) => (
          <div key={i} style={{ background: C.gold50, border: `1px solid ${C.gold100}`, borderRadius: 6, padding: "10px 12px", marginBottom: 8 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
              <span style={{ fontSize: 11, fontWeight: 700, color: C.gold700, textTransform: "uppercase", letterSpacing: "0.05em" }}>Document {i + 1}</span>
              <button onClick={() => removeSupplement(i)} style={{ background: "none", border: "none", color: C.error, cursor: "pointer", fontSize: 16, lineHeight: 1 }}>×</button>
            </div>
            <FormField label="Title" hint="e.g. Protected Classes Quick Reference"><input style={S.input} value={sup.title} onChange={e => updateSupplement(i, "title", e.target.value)} placeholder="Document title..." /></FormField>
            <FormField label="URL" hint="Direct SharePoint link to file"><input style={S.input} value={sup.url} onChange={e => updateSupplement(i, "url", e.target.value)} placeholder="https://vanrockre.sharepoint.com/..." /></FormField>
          </div>
        ))}
      </div>
      <SaveBar saving={saving} onSave={handleSave} onCancel={onClose} onDelete={isEdit ? handleDelete : null} deleteLabel="Delete Lesson" />
    </Modal>
  );
}

// ── QUIZ QUESTION FORM ──
function QuizForm({ item, courseId, onClose }) {
  const { courses, quizzes, setQuizzes, isLive, getToken } = useData();
  const isEdit = !!item;
  const defaultCourseId = isEdit ? Object.keys(quizzes).find(cid => quizzes[cid].questions.some(q => q.id===item.id)) || courseId || "" : courseId || "";
  const [form, setForm] = useState({ question: item?.question||"", courseId: defaultCourseId, optA: item?.options?.A||"", optB: item?.options?.B||"", optC: item?.options?.C||"", optD: item?.options?.D||"", correct: item?.correct||"A" });
  const [saving, setSaving] = useState(false);
  const set = (k,v) => setForm(p => ({...p,[k]:v}));
  const handleSave = async () => {
    if (!form.question.trim()||!form.courseId) return alert("Question and course are required.");
    if (!form.optA||!form.optB) return alert("At least options A and B are required.");
    setSaving(true);
    const fields = { Title: form.question.trim(), QuizCourseIDLookupId: parseInt(form.courseId,10), OptionA: form.optA, OptionB: form.optB, OptionC: form.optC, OptionD: form.optD, CorrectAnswer: form.correct, QuizSortOrder: 0 };
    let savedId = item?.id;
    try {
      if (isLive) {
        const token = await getToken();
        if (isEdit) { await spUpdate(token, CONFIG.lists.quizzes, item.id, fields); }
        else { const res = await spCreate(token, CONFIG.lists.quizzes, fields); savedId = String(res.id); }
      } else { savedId = savedId || "new_" + Date.now(); }
      setQuizzes(prev => {
        const next = {...prev}; const cid = String(fields.QuizCourseIDLookupId);
        const qObj = {id:savedId, question:fields.Title, options:{A:fields.OptionA,B:fields.OptionB,C:fields.OptionC,D:fields.OptionD}, correct:fields.CorrectAnswer};
        if (!next[cid]) next[cid] = {questions:[]};
        if (isEdit) { next[cid] = {questions: next[cid].questions.map(q => q.id===item.id ? qObj : q)}; }
        else { next[cid] = {questions: [...next[cid].questions, qObj]}; }
        return next;
      });
      onClose();
    } catch (err) { alert("Save failed: " + err.message); }
    setSaving(false);
  };
  const handleDelete = async () => {
    if (!confirm("Delete this quiz question?")) return;
    setSaving(true);
    try {
      if (isLive) { const token = await getToken(); await spDelete(token, CONFIG.lists.quizzes, item.id); }
      setQuizzes(prev => { const next = {...prev}; for (const cid of Object.keys(next)) { next[cid] = {questions: next[cid].questions.filter(q => q.id!==item.id)}; } return next; });
      onClose();
    } catch (err) { alert("Failed: " + err.message); }
    setSaving(false);
  };
  return (
    <Modal title={isEdit ? "Edit Quiz Question" : "Add Quiz Question"} onClose={onClose} width={600}>
      <FormField label="Course"><select style={S.select} value={form.courseId} onChange={e => set("courseId", e.target.value)} disabled={isEdit}><option value="">— Select —</option>{courses.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}</select></FormField>
      <FormField label="Question"><textarea style={{...S.input,minHeight:60}} value={form.question} onChange={e => set("question", e.target.value)} /></FormField>
      <FormField label="Option A"><input style={S.input} value={form.optA} onChange={e => set("optA", e.target.value)} /></FormField>
      <FormField label="Option B"><input style={S.input} value={form.optB} onChange={e => set("optB", e.target.value)} /></FormField>
      <FormField label="Option C"><input style={S.input} value={form.optC} onChange={e => set("optC", e.target.value)} placeholder="Optional" /></FormField>
      <FormField label="Option D"><input style={S.input} value={form.optD} onChange={e => set("optD", e.target.value)} placeholder="Optional" /></FormField>
      <FormField label="Correct Answer"><select style={S.select} value={form.correct} onChange={e => set("correct", e.target.value)}>{["A","B","C","D"].map(o => <option key={o} value={o}>{o}{form["opt"+o] ? ` — ${form["opt"+o].substring(0,40)}` : ""}</option>)}</select></FormField>
      <SaveBar saving={saving} onSave={handleSave} onCancel={onClose} onDelete={isEdit ? handleDelete : null} deleteLabel="Delete Question" />
    </Modal>
  );
}

// ============================================================
// ============================================================
// MANAGE VIEW (Admin — Full CRUD)
// ============================================================
function ManageView({ mobile }) {
  const { employees, courses, learningPaths, lessons, quizzes, isLive, getToken } = useData();
  const [subTab, setSubTab] = useState("employees");
  const [modal, setModal] = useState(null);
  const [expandedCourse, setExpandedCourse] = useState(null);
  const [showInactive, setShowInactive] = useState(false);
  const [courseFilter, setCourseFilter] = useState(null);
  const [emailsPaused, setEmailsPaused] = useState(EMAIL_PAUSED);

  const closeModal = () => setModal(null);

  return (
    <div>
      {/* Sub-tabs */}
      <div style={{ display: "flex", gap: 8, marginBottom: 20, flexWrap: "wrap" }}>
        {[["employees", "Employees"], ["paths", "Learning Paths"], ["courses", "Courses"], ["config", "Settings"]].map(([key, label]) => (
          <button
            key={key}
            onClick={() => { setSubTab(key); setExpandedCourse(null); }}
            style={{
              ...S.btnSecondary,
              ...(subTab === key ? { background: C.teal50, borderColor: C.gold500, color: C.teal700 } : {})
            }}
          >
            {label}
          </button>
        ))}
      </div>

      {/* ── EMPLOYEES TAB ── */}
      {subTab === "employees" && (
        <div style={S.card}>
          <div style={{ ...S.cardTitle, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
            <span>Employees</span>
            <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
              <label style={{ fontSize: 12, color: C.gray400, display: "flex", alignItems: "center", gap: 4, cursor: "pointer" }}>
                <input type="checkbox" checked={showInactive} onChange={e => setShowInactive(e.target.checked)} /> Show inactive
              </label>
              <button style={{ ...S.btnPrimary, ...S.btnSmall }} onClick={() => setModal({ type: "employee" })}>+ Add Employee</button>
            </div>
          </div>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                <th style={S.th}>Name</th>
                {!mobile && <th style={S.th}>Email</th>}
                <th style={S.th}>Role</th>
                {!mobile && <th style={S.th}>Reports To</th>}
                {!mobile && <th style={S.th}>Access</th>}
                {!mobile && <th style={S.th}>Hire Date</th>}
                <th style={S.th}>Status</th>
                <th style={S.th}></th>
              </tr>
            </thead>
            <tbody>
              {employees.filter(e => showInactive || e.active).map(emp => {
                const mgr = employees.find(e => e.id === emp.reportsTo || e.email === emp.reportsTo);
                const subs = employees.filter(e => e.active && (e.reportsTo === emp.id || e.reportsTo === emp.email));
                return (
                  <tr key={emp.id} style={{ opacity: emp.active ? 1 : 0.5 }}>
                    <td style={{ ...S.td, fontWeight: 500 }}>{emp.name}</td>
                    {!mobile && <td style={S.td}>{emp.email}</td>}
                    <td style={S.td}>{emp.role}</td>
                    {!mobile && <td style={S.td}>{mgr ? mgr.name : "\u2014"}</td>}
                    {!mobile && <td style={S.td}>
                      {emp.appRole === "Admin" && <span style={S.badge("warning")}>ADMIN</span>}
                      {emp.appRole !== "Admin" && subs.length > 0 && <span style={S.badge("info")}>MANAGER ({subs.length})</span>}
                      {emp.appRole !== "Admin" && subs.length === 0 && <span style={S.badge("neutral")}>EMPLOYEE</span>}
                    </td>}
                    {!mobile && <td style={S.td}>{emp.hireDate}</td>}
                    <td style={S.td}><span style={S.badge(emp.active ? "success" : "neutral")}>{emp.active ? "Active" : "Inactive"}</span></td>
                    <td style={S.td}><button style={{ ...S.btnSecondary, ...S.btnSmall }} onClick={() => setModal({ type: "employee", item: emp })}>Edit</button></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {/* ── LEARNING PATHS TAB ── */}
      {subTab === "paths" && (
        <div style={S.card}>
          <div style={{ ...S.cardTitle, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span>Learning Paths</span>
            <button style={{ ...S.btnPrimary, ...S.btnSmall }} onClick={() => setModal({ type: "path" })}>+ Add Path</button>
          </div>
          {learningPaths.map(path => (
            <div key={path.id} style={{ padding: "14px 0", borderBottom: `1px solid ${C.gray100}` }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 8 }}>
                <div>
                  <div style={{ fontSize: 15, fontWeight: 600, color: C.teal700 }}>{path.name}</div>
                  <div style={{ fontSize: 13, color: C.gray400, marginTop: 2 }}>{path.description}</div>
                  <div style={{ display: "flex", gap: 8, marginTop: 6, flexWrap: "wrap" }}>
                    <span style={S.badge("info")}>{path.courseIds.length} courses</span>
                    {path.required && <span style={S.badge("warning")}>REQUIRED</span>}
                    <span style={S.badge("neutral")}>{path.roles.join(", ")}</span>
                    {path.dueDays && <span style={S.badge("info")}>Due within {path.dueDays}d of hire</span>}
                  </div>
                  <div style={{ fontSize: 12, color: C.gray400, marginTop: 6 }}>
                    Courses: {path.courseIds.map(cid => courses.find(c => c.id === cid)?.name || cid).join(" \u2192 ")}
                  </div>
                </div>
                <button style={{ ...S.btnSecondary, ...S.btnSmall }} onClick={() => setModal({ type: "path", item: path })}>Edit</button>
              </div>
            </div>
          ))}
        </div>
      )}

      {/* ── COURSES TAB ── */}
      {subTab === "courses" && !expandedCourse && (
        <div style={S.card}>
          <div style={{ ...S.cardTitle, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
            <span>Courses</span>
            <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
              <select style={{ ...S.select, fontSize: 12, padding: "4px 8px", width: "auto" }} value={courseFilter || "all"} onChange={e => setCourseFilter(e.target.value === "all" ? null : e.target.value)}>
                <option value="all">All statuses</option>
                <option value="Active">Active</option>
                <option value="Coming Soon">Coming Soon</option>
                <option value="Archived">Archived</option>
              </select>
              <button style={{ ...S.btnPrimary, ...S.btnSmall }} onClick={() => setModal({ type: "course" })}>+ Add Course</button>
            </div>
          </div>
          {(() => { const filtered = courseFilter ? courses.filter(c => c.status === courseFilter) : courses; return mobile ? (
            filtered.map(course => {
              const cLessons = lessons.filter(l => l.courseId === course.id);
              const cQuiz = quizzes[course.id]?.questions || [];
              const statusBadge = course.status === "Coming Soon" ? "info" : course.status === "Archived" ? "neutral" : "success";
              return (
                <div key={course.id} style={{ padding: "12px 0", borderBottom: `1px solid ${C.gray100}`, cursor: "pointer", opacity: course.status === "Archived" ? 0.5 : 1 }} onClick={() => setExpandedCourse(course)}>
                  <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                    {course.code && <span style={{ fontSize: 12, fontWeight: 600, color: C.gold500 }}>{course.code}</span>}
                    <span style={{ fontSize: 14, fontWeight: 600, color: C.teal700 }}>{course.name}</span>
                    <span style={S.badge(statusBadge)}>{course.status}</span>
                  </div>
                  <div style={{ fontSize: 13, color: C.gray400, marginTop: 2 }}>{course.category} \u00b7 {course.durationMin} min \u00b7 {cLessons.length} lessons \u00b7 {cQuiz.length} quiz Qs</div>
                  <div style={{ display: "flex", gap: 6, marginTop: 6 }}>
                    {course.recertDays && <span style={S.badge("warning")}>Recert: {course.recertDays}d</span>}
                    <span style={{ fontSize: 12, color: C.gold500 }}>Tap to manage \u2192</span>
                  </div>
                </div>
              );
            })
          ) : (
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead>
                <tr>
                  <th style={S.th}>Code</th>
                  <th style={S.th}>Course</th>
                  <th style={S.th}>Status</th>
                  <th style={S.th}>Category</th>
                  <th style={S.th}>Duration</th>
                  <th style={S.th}>Lessons</th>
                  <th style={S.th}>Quiz Qs</th>
                  <th style={S.th}>Recert</th>
                  <th style={S.th}>Actions</th>
                </tr>
              </thead>
              <tbody>
                {filtered.map(course => {
                  const cLessons = lessons.filter(l => l.courseId === course.id);
                  const cQuiz = quizzes[course.id]?.questions || [];
                  const statusBadge = course.status === "Coming Soon" ? "info" : course.status === "Archived" ? "neutral" : "success";
                  return (
                    <tr key={course.id} style={{ cursor: "pointer", opacity: course.status === "Archived" ? 0.5 : 1 }} onClick={() => setExpandedCourse(course)}>
                      <td style={{ ...S.td, fontWeight: 500, color: C.gold500, whiteSpace: "nowrap", fontSize: 12 }}>{course.code || "\u2014"}</td>
                      <td style={{ ...S.td, fontWeight: 500, color: C.teal700 }}>{course.name}</td>
                      <td style={S.td}><span style={S.badge(statusBadge)}>{course.status}</span></td>
                      <td style={S.td}>{course.category}</td>
                      <td style={S.td}>{course.durationMin} min</td>
                      <td style={S.td}>{cLessons.length}</td>
                      <td style={S.td}>{cQuiz.length}</td>
                      <td style={S.td}>{course.recertDays ? `${course.recertDays}d` : "\u2014"}</td>
                      <td style={S.td} onClick={e => e.stopPropagation()}>
                        <div style={{ display: "flex", gap: 6 }}>
                          <button style={{ ...S.btnSecondary, ...S.btnSmall }} onClick={() => setModal({ type: "course", item: course })}>Edit</button>
                          <button style={{ ...S.btnSecondary, ...S.btnSmall, color: C.teal700 }} onClick={() => setExpandedCourse(course)}>Manage \u2192</button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          ); })()}
        </div>
      )}

      {/* ── COURSE DRILL-DOWN (Lessons + Quiz Questions) ── */}
      {subTab === "courses" && expandedCourse && (() => {
        const course = expandedCourse;
        const cLessons = lessons.filter(l => l.courseId === course.id).sort((a,b) => a.order - b.order);
        const cQuiz = quizzes[course.id]?.questions || [];
        return (
          <div>
            <button style={{ ...S.btnSecondary, ...S.btnSmall, marginBottom: 16 }} onClick={() => setExpandedCourse(null)}>{"\u2190"} Back to Courses</button>
            <div style={{ ...S.card, marginBottom: 20 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 8 }}>
                <div>
                  <div style={{ fontSize: 18, fontWeight: 700, color: C.teal700 }}>{course.name}</div>
                  <div style={{ fontSize: 13, color: C.gray400, marginTop: 2 }}>{course.description}</div>
                  <div style={{ display: "flex", gap: 8, marginTop: 8, flexWrap: "wrap" }}>
                    <span style={S.badge("info")}>{course.category}</span>
                    <span style={S.badge("neutral")}>{course.durationMin} min</span>
                    <span style={S.badge("neutral")}>Pass: {course.passingScore}%</span>
                    {course.recertDays && <span style={S.badge("warning")}>Recert: {course.recertDays}d</span>}
                  </div>
                </div>
                <button style={{ ...S.btnSecondary, ...S.btnSmall }} onClick={() => setModal({ type: "course", item: course })}>Edit Course</button>
              </div>
            </div>

            {/* Lessons */}
            <div style={{ ...S.card, marginBottom: 20 }}>
              <div style={{ ...S.cardTitle, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <span>Lessons ({cLessons.length})</span>
                <button style={{ ...S.btnPrimary, ...S.btnSmall }} onClick={() => setModal({ type: "lesson", courseId: course.id })}>+ Add Lesson</button>
              </div>
              {cLessons.length === 0 && <div style={{ padding: 16, color: C.gray400, fontSize: 13 }}>No lessons yet. Add one to get started.</div>}
              {cLessons.map((lesson, idx) => (
                <div key={lesson.id} style={{ padding: "10px 0", borderBottom: idx < cLessons.length-1 ? `1px solid ${C.gray100}` : "none", display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
                  <div>
                    <div style={{ fontSize: 14, fontWeight: 500, color: C.teal700 }}>
                      <span style={{ color: C.gray300, fontSize: 12, marginRight: 8 }}>#{lesson.order}</span>
                      {lesson.title}
                    </div>
                    <div style={{ fontSize: 12, color: C.gray400, marginTop: 2, display: "flex", gap: 8 }}>
                      {lesson.durationMin > 0 && <span>{lesson.durationMin} min</span>}
                      {lesson.videoUrl && <span style={{ color: C.gold500 }}>Video</span>}
                      {lesson.documentUrl && <span style={{ color: C.gold500 }}>{lesson.documentTitle || "Document"}</span>}
                    </div>
                  </div>
                  <button style={{ ...S.btnSecondary, ...S.btnSmall }} onClick={() => setModal({ type: "lesson", item: lesson, courseId: course.id })}>Edit</button>
                </div>
              ))}
            </div>

            {/* Quiz Questions */}
            <div style={S.card}>
              <div style={{ ...S.cardTitle, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <span>Quiz Questions ({cQuiz.length})</span>
                <button style={{ ...S.btnPrimary, ...S.btnSmall }} onClick={() => setModal({ type: "quiz", courseId: course.id })}>+ Add Question</button>
              </div>
              {cQuiz.length === 0 && <div style={{ padding: 16, color: C.gray400, fontSize: 13 }}>No quiz questions yet. Add at least one for employees to complete this course.</div>}
              {cQuiz.map((q, idx) => (
                <div key={q.id} style={{ padding: "10px 0", borderBottom: idx < cQuiz.length-1 ? `1px solid ${C.gray100}` : "none", display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8 }}>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 14, fontWeight: 500, color: C.teal700 }}>{q.question}</div>
                    <div style={{ fontSize: 12, color: C.gray400, marginTop: 4, display: "grid", gridTemplateColumns: "1fr 1fr", gap: "2px 16px" }}>
                      {["A","B","C","D"].filter(o => q.options[o]).map(o => (
                        <div key={o} style={{ color: o === q.correct ? "#2E7D5B" : C.gray400 }}>
                          {o === q.correct ? "\u2713" : " "} {o}. {q.options[o]}
                        </div>
                      ))}
                    </div>
                  </div>
                  <button style={{ ...S.btnSecondary, ...S.btnSmall }} onClick={() => setModal({ type: "quiz", item: q, courseId: course.id })}>Edit</button>
                </div>
              ))}
            </div>
          </div>
        );
      })()}

      {/* ── SETTINGS TAB ── */}
      {subTab === "config" && (
        <div style={S.card}>
          <div style={S.cardTitle}>Training Settings</div>

          {/* Email Pause Toggle */}
          <div style={{ marginBottom: 20, padding: 16, background: emailsPaused ? "#FFF0F0" : "#F0FFF4", border: `1px solid ${emailsPaused ? "#C44B3B" : "#2E7D5B"}`, borderRadius: 6 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div>
                <div style={{ fontSize: 14, fontWeight: 700, color: emailsPaused ? "#C44B3B" : "#2E7D5B" }}>
                  {emailsPaused ? "Emails are PAUSED" : "Emails are ACTIVE"}
                </div>
                <div style={{ fontSize: 12, color: C.gray400, marginTop: 2 }}>
                  {emailsPaused ? "No notification emails will be sent. All other app functions continue normally. Power Automate flows are not affected by this toggle." : "The app will send notification emails for quiz results, enrollments, and course launches."}
                </div>
              </div>
              <button
                onClick={async () => {
                  const next = !emailsPaused;
                  setEmailsPaused(next);
                  EMAIL_PAUSED = next;
                  // Persist to AppConfig in SharePoint
                  try {
                    const token = await getToken();
                    const existing = await spGet(token, CONFIG.lists.config, { filter: "fields/Title eq 'EmailsPaused'" });
                    if (existing.length > 0) {
                      await spUpdate(token, CONFIG.lists.config, existing[0].id, { Value: String(next) });
                    } else {
                      await spCreate(token, CONFIG.lists.config, { Title: "EmailsPaused", Value: String(next) });
                    }
                  } catch(e) { console.error("Failed to save email pause state:", e); }
                }}
                style={{
                  ...S.btnSecondary, ...S.btnSmall, minWidth: 120,
                  color: emailsPaused ? "#2E7D5B" : "#C44B3B",
                  borderColor: emailsPaused ? "#2E7D5B" : "#C44B3B",
                }}
              >
                {emailsPaused ? "Enable Emails" : "Pause Emails"}
              </button>
            </div>
          </div>

          <div style={{ display: "grid", gridTemplateColumns: mobile ? "1fr" : "1fr 1fr", gap: 16 }}>
            <div>
              <label style={S.label}>Default Passing Score (%)</label>
              <input style={S.input} type="number" defaultValue={CONFIG.passingScore} />
            </div>
            <div>
              <label style={S.label}>Default Recertification Period (days)</label>
              <input style={S.input} type="number" defaultValue={365} />
            </div>
            <div>
              <label style={S.label}>Expiration Warning Window (days before)</label>
              <input style={S.input} type="number" defaultValue={30} />
            </div>
            <div>
              <label style={S.label}>Max Quiz Retakes Per Course</label>
              <input style={S.input} type="number" defaultValue={0} placeholder="0 = unlimited" />
            </div>
          </div>
          <div style={{ marginTop: 20, padding: "16px", background: C.teal50, borderRadius: 6 }}>
            <div style={{ fontSize: 14, fontWeight: 600, color: C.teal700, marginBottom: 8 }}>Notification Settings</div>
            <div style={{ fontSize: 13, color: C.gray600, lineHeight: 1.6 }}>
              Cert expiration reminders and Monday manager reports are sent by Power Automate and are not affected by the pause toggle above.
              To pause those, disable the flows directly in Power Automate.
            </div>
          </div>
        </div>
      )}

      {/* ── FOOTER ── */}
      <div style={{ textAlign: "center", padding: "24px 20px 16px", fontSize: 11, color: C.gray300, borderTop: `1px solid ${C.gray100}`, marginTop: 24 }}>
        2025 · This application is the intellectual property of NewShire Property Management. Reproduction or use without written permission is prohibited.
      </div>

      {/* ── MODALS ── */}
      {modal?.type === "employee" && <EmployeeForm item={modal.item} onClose={closeModal} />}
      {modal?.type === "course" && <CourseForm item={modal.item} onClose={closeModal} />}
      {modal?.type === "path" && <PathForm item={modal.item} onClose={closeModal} />}
      {modal?.type === "lesson" && <LessonForm item={modal.item} courseId={modal.courseId} onClose={closeModal} />}
      {modal?.type === "quiz" && <QuizForm item={modal.item} courseId={modal.courseId} onClose={closeModal} />}
    </div>
  );
}
