# PDF Generation — Design Spec

**Date:** 2026-05-16
**Status:** Approved

## Overview

Add PDF export to both the Plans and Pipeline reports. Each upload card gains two side-by-side buttons — "Generate Excel" (existing, renamed) and "Generate PDF" (new). The PDF is branded with Diana Silver Grants Strategy's blue color scheme and is designed to be shared directly with clients.

---

## Architecture

### New files

| File | Purpose |
|---|---|
| `reports/plans_pdf.py` | `generate(csv_bytes) -> list[tuple[str, bytes]]` — same signature as Excel version |
| `reports/pipeline_pdf.py` | `generate(csv_bytes, filename) -> tuple[str, bytes]` — same signature as Excel version |
| `templates/plans_pdf.html` | Jinja2 template for Plans PDF |
| `templates/pipeline_pdf.html` | Jinja2 template for Pipeline PDF |
| `tests/test_plans_pdf.py` | Tests for Plans PDF generation |
| `tests/test_pipeline_pdf.py` | Tests for Pipeline PDF generation |

### Changed files

| File | Change |
|---|---|
| `reports/plans.py` | Extract `_load_data(csv_bytes)` private helper; `generate()` calls it |
| `reports/pipeline.py` | Extract `_load_data(csv_bytes, filename)` private helper; `generate()` calls it |
| `app.py` | Two new routes: `POST /generate/plans/pdf` and `POST /generate/pipeline/pdf` |
| `templates/index.html` | "Generate Report" → two side-by-side buttons; `initUploadArea` extended |
| `Dockerfile` | Add `apt-get install -y libpango-1.0-0 libcairo2 libgdk-pixbuf2.0-0` |
| `requirements.txt` | Add `weasyprint` |

### PDF rendering

PDF modules use Jinja2 directly (not Flask's `render_template`) to avoid coupling the report layer to Flask. Templates are resolved via `Path(__file__).parent.parent / "templates"`. The rendered HTML string is passed to `weasyprint.HTML(string=html).write_pdf()`.

### Data loading refactor

`plans.py` and `pipeline.py` each expose a private `_load_data()` function that handles CSV parsing, validation, column mapping, and sorting. Both the existing Excel `generate()` and the new PDF `generate()` in the corresponding PDF module call this helper. This eliminates duplicated parsing logic without adding abstraction beyond what the task requires.

---

## Flask routes

Two new routes, mirroring the existing Excel routes exactly:

```
POST /generate/plans/pdf
POST /generate/pipeline/pdf
```

- Use the same `_get_csv_file()` helper
- `ValueError` → 400; unexpected exceptions → logged + 500
- WeasyPrint render failures treated as unexpected errors (logged, 500)
- Plans PDF: single client → `{client}_report.pdf`; multiple clients → `reports.zip` (same zip logic as Excel)
- Pipeline PDF: `Pipeline_{client}_{date}.pdf`

---

## PDF designs

### Shared: header and footer

**Header** (full width, blue `#4472C4` background):
- Left: "Diana Silver Grants Strategy" (white, bold) + "diana@dsilvergrants.com" (light blue, small)
- Right: "Generated {date}" (light blue, small)

**Footer** (light grey `#f8f9fa` background, top border):
- Left: "Diana Silver Grants Strategy · diana@dsilvergrants.com" (grey, small)
- Right: "Page {n}" (grey, small)

No banner image for now — placeholder for future artwork.

---

### Plans PDF

**Title block** (below header):
- Client full name (large, dark, bold)
- "Plans Report" (medium, blue `#4472C4`, bold)

**Year sections:** The data is split into two sections — `2026` (current year) and `2027+` (future). Each section has a bold year label with a blue bottom rule. If no future rows exist, the 2027+ section is omitted.

**Within each year section — grouped by status:**

Rows are grouped under status labels matching `STATUS_ORDER` from `plans.py`. Only status groups that have rows are shown.

- Status label: small, uppercase, bold, blue (`#4472C4`), letter-spaced
- Each grant row is a card with a left accent bar and light blue background. Accent bar color varies by priority:
  - Solid blue `#4472C4` — Awarded (Active or Closed)
  - Medium blue `#93afd4` — Submitted / In Progress
  - Light blue `#c5d5ee` — Planned / Researching / Declined / Abandoned
- Row layout:
  - Line 1: **Funder — Fund** (bold, left) + `Req: $X  Awd: $X` (right)
  - Line 2: Year · Purpose · Notif Expected · Notif Received · Next Task/Deadline (grey, small)
  - Blank/null fields are omitted from line 2 rather than shown as "—"

**Columns included:** Year, Funder, Fund, Purpose, Request, Award, Notif Expected, Notif Received, Next Task/Deadline. Status is the section header, not a column.

---

### Pipeline PDF

**Title block:** Client full name + "Pipeline Report" (same pattern as Plans).

**Single plain table**, sorted by Due Date ascending, with alternating row shading (white / light blue `#f0f4ff`).

| Column | Notes |
|---|---|
| Year | |
| Project | |
| Funder | |
| Task Type | |
| Task Name | |
| Due Date | MM/DD/YYYY |

Blue header row (`#4472C4`, white bold text), 1px `#e5e7eb` row borders.

---

## UI changes

### Button layout

Each upload card gets two equal-width side-by-side buttons replacing the single "Generate Report" button:

- **Generate Excel** — blue `#4472C4` (existing style)
- **Generate PDF** — teal `#0d9488`

Both buttons are disabled until a file is selected. Both buttons share a single status message and download link area per card — whichever button was last clicked updates that shared area. The existing `initUploadArea()` function is extended with optional `pdfBtnId` and `pdfEndpoint` parameters; when provided, the PDF button is wired up to the same status/download elements alongside the Excel button.

---

## Tests

`tests/test_plans_pdf.py` and `tests/test_pipeline_pdf.py` follow the same pattern as the existing Excel test files:
- Synthetic inline CSV strings as fixtures
- `date.today()` mocked to a fixed date
- Assert returned bytes start with `%PDF-` (PDF magic bytes)
- Same error-case coverage as the Excel tests: missing columns, bad CSV, empty file, invalid filename (pipeline)

---

## Dockerfile change

```dockerfile
RUN apt-get update && apt-get install -y --no-install-recommends \
    libpango-1.0-0 libcairo2 libgdk-pixbuf2.0-0 \
    && rm -rf /var/lib/apt/lists/*
```

Add before the `pip install` step so system deps are in place before WeasyPrint is installed.
