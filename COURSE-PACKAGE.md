# Course Packages — SOP → Course Pipeline

This is the repeatable workflow for turning a finalized SOP into a live NewShire University course
with **minimal manual work**. You never hand-write training content or hand-enter lessons.

## The loop (every new SOP)

1. **You:** "Make a course from the [name] SOP." (Point me at it in SharePoint, or paste it.)
2. **Me (Claude):** Read the SOP, generate a **course package** (`.json`) — course metadata, written
   lessons (the SOP's trainer scripts + steps become readable in-app lessons), and a quiz bank.
   For roles that need different training, I generate one package per role (auto-tailored).
3. **You:** Open **Manage → Import Course**, paste the JSON, click **Validate**. In the
   **Import settings** panel set the status, recertification period, visible-to roles, and the
   learning path (none / add to an existing path / create a new Required path with a due date) —
   the path assignment is what makes the course show as **Required**. Then click **Import** and the
   app creates the Course + Lessons + Quiz (and the path link) in SharePoint in one step. Done.

The generated `.json` files live in `course-packages/` so we have a versioned record of every course.

## One-time setup

Lessons now support **written content** (an in-app readable body — no video or slide deck required).
That content is stored in a `LessonBody` column on the `TrainingLessons` list. Run this once:

```powershell
pwsh ./scripts/ensure-lessonbody-column.ps1
```

It signs you in (browser), reports which training lists exist, and adds the `LessonBody` column if
it's missing. Safe to re-run — it skips anything already present.

## Package format

```jsonc
{
  "course": {
    "name": "Application Processing",        // required
    "code": "LEAS 110",                       // optional, shown as a badge; must be unique
    "description": "One–two sentence summary shown on the course card.",
    "category": "Leasing",                    // Onboarding|Compliance|Leasing|Maintenance|Operations|Safety|Financial|Management
    "durationMin": 32,                        // total; usually the sum of lesson minutes
    "recertDays": 0,                          // 0 = no recertification; 365 = annual
    "passingScore": 80,                       // % to pass the quiz
    "sortOrder": 110,                         // lower = appears earlier in the library
    "status": "Active",                       // Active | Coming Soon | Archived
    "roles": ["Leasing Agent", "Property Manager"]  // [] = visible to everyone
  },
  "lessons": [
    {
      "title": "Verify Application Completeness",
      "order": 2,                             // 1-based; optional (defaults to array order)
      "durationMin": 5,
      "body": "<h4>...</h4><p>...</p><ul><li>...</li></ul><div class='callout'>...</div>",
      "supplements": [                         // optional attached docs
        { "title": "Missing Documents Template", "url": "https://vanrockre.sharepoint.com/..." }
      ]
      // videoUrl / documentUrl are also supported but optional — written body is the default
    }
  ],
  "quiz": {
    "questions": [
      {
        "question": "How long does an applicant have to provide missing documents?",
        "A": "24 hours", "B": "48 hours", "C": "72 hours", "D": "7 days",
        "correct": "C"                        // must be "A" | "B" | "C" | "D"
      }
    ]
  },
  "pathName": "Leasing Certification",        // optional — pre-selects this path in the import screen (you can override there)
  "source": { "sopName": "Application Processing SOP", "sopUrl": "https://vanrockre.sharepoint.com/..." }
}
```

## Two course tracks

| Track | Lesson type | When |
| --- | --- | --- |
| **SOP courses** | Written (`body` HTML) | Internal procedures that change often — instant, in-app editable |
| **Industry-standard courses** | Slide deck (`documentUrl`) | Static topics with a polished deck — Fair Housing, harassment prevention, safety, etc. |

Both run through the same Import Course tool. A single course can even mix lesson types.

### Slide-based (PowerPoint) lesson

Host the deck in SharePoint, then put its link in `documentUrl`. The app embeds it in the SharePoint
viewer inside the lesson. No `body` is needed for a slide lesson.

```jsonc
{
  "title": "Protected Classes Under Federal Law",
  "order": 1,
  "durationMin": 15,
  "documentUrl": "https://vanrockre.sharepoint.com/.../FairHousing.pptx?web=1",  // SharePoint share/embed link
  "documentTitle": "Fair Housing — Protected Classes"
}
```

To get the link: open the deck in SharePoint → **Share** (or **Copy link**) → paste here. The importer
auto-strips any surrounding `<iframe>` embed code, so pasting the embed snippet works too. (You can also
add a slide lesson by hand later via **Manage → Courses → Add Lesson → Presentation URL**.)

## Lesson body HTML

The `body` is HTML rendered directly in the lesson player. Supported building blocks (styled in `index.html`):

| Markup | Use |
| --- | --- |
| `<h3>` / `<h4>` | Section headings |
| `<p>` | Narration / explanation |
| `<ul><li>` / `<ol><li>` | Steps, checklists, criteria |
| `<strong>` | Emphasis / labels |
| `<div class="callout">` | Gold "key takeaway" box (lead with `<strong>Key takeaway</strong>`) |
| `<div class="script">` | Teal italic "trainer script" box (great for the SOP's spoken script lines) |

Keep bodies self-contained — no external CSS or `<script>`. Content is admin-authored and internal.

## Role variants

When one SOP needs different training per position, I produce multiple packages (e.g.
`application-processing-leasing.json`, `application-processing-manager.json`), each with its own
`roles`, emphasis, and quiz. Import each one. Role-scoped learning paths then route the right
version to the right people.
