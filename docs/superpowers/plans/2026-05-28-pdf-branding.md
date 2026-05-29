# PDF Branding (Logo + Fonts) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add DS_Logo_Alt.png in a split header and apply Merriweather/Roboto fonts to both PDF templates.

**Architecture:** Download 4 TTF font files into `templates/fonts/`. Pass `base_url=str(_TEMPLATES_DIR)` to WeasyPrint's `HTML()` call in both PDF modules so relative paths in templates resolve correctly. Update both HTML templates with new `@font-face` declarations, a split header, and Merriweather assigned to heading elements.

**Tech Stack:** WeasyPrint, Jinja2, Merriweather (Google Fonts), Roboto (Google Fonts)

---

## File Map

| File | Action | What changes |
|---|---|---|
| `templates/fonts/` | Create dir + 4 TTF files | New font assets |
| `reports/plans_pdf.py` | Modify line 117 | Add `base_url` to `HTML()` call |
| `reports/pipeline_pdf.py` | Modify line 54 | Add `base_url` to `HTML()` call |
| `templates/plans_pdf.html` | Rewrite | Split header + `@font-face` + Merriweather headings |
| `templates/pipeline_pdf.html` | Rewrite | Split header + `@font-face` + Merriweather headings |

> **Note on Roboto weights:** Roboto has no static SemiBold (600) TTF. The spec listed `Roboto-SemiBold` but this plan uses only Regular (400) and Bold (700). Pipeline table headers use `font-weight: 700` instead of 600.

---

### Task 1: Download font files into templates/fonts/

**Files:**
- Create: `templates/fonts/Roboto-Regular.ttf`
- Create: `templates/fonts/Roboto-Bold.ttf`
- Create: `templates/fonts/Merriweather-Regular.ttf`
- Create: `templates/fonts/Merriweather-Bold.ttf`

- [ ] **Step 1: Create the directory and download all four font files**

```bash
mkdir -p templates/fonts

curl -fL -o templates/fonts/Roboto-Regular.ttf \
  "https://raw.githubusercontent.com/google/fonts/main/apache/roboto/static/Roboto-Regular.ttf"

curl -fL -o templates/fonts/Roboto-Bold.ttf \
  "https://raw.githubusercontent.com/google/fonts/main/apache/roboto/static/Roboto-Bold.ttf"

curl -fL -o templates/fonts/Merriweather-Regular.ttf \
  "https://raw.githubusercontent.com/google/fonts/main/ofl/merriweather/Merriweather-Regular.ttf"

curl -fL -o templates/fonts/Merriweather-Bold.ttf \
  "https://raw.githubusercontent.com/google/fonts/main/ofl/merriweather/Merriweather-Bold.ttf"
```

- [ ] **Step 2: Verify the files are real font data (not a 404 redirect page)**

```bash
ls -lh templates/fonts/
```

Expected: 4 files, each several hundred KB. If any file is only a few hundred bytes it is an HTML error page — find the correct raw URL by browsing https://github.com/google/fonts, download the TTF manually, and place it in `templates/fonts/` with the exact filename shown above.

- [ ] **Step 3: Commit**

```bash
git add templates/fonts/
git commit -m "feat: add Roboto and Merriweather TTF files for PDF rendering"
```

---

### Task 2: Pass base_url to WeasyPrint HTML() in both PDF modules

**Files:**
- Modify: `reports/plans_pdf.py` (line 117)
- Modify: `reports/pipeline_pdf.py` (line 54)

- [ ] **Step 1: Update plans_pdf.py**

In `reports/plans_pdf.py`, find the `generate()` function and change the `HTML()` call:

```python
# Before
pdf_bytes = HTML(string=html).write_pdf()
# After
pdf_bytes = HTML(string=html, base_url=str(_TEMPLATES_DIR)).write_pdf()
```

- [ ] **Step 2: Update pipeline_pdf.py**

In `reports/pipeline_pdf.py`, find the `generate()` function and change the `HTML()` call:

```python
# Before
pdf_bytes = HTML(string=html).write_pdf()
# After
pdf_bytes = HTML(string=html, base_url=str(_TEMPLATES_DIR)).write_pdf()
```

- [ ] **Step 3: Run the PDF tests to confirm nothing broke**

```bash
.venv/bin/pytest tests/test_plans_pdf.py tests/test_pipeline_pdf.py -v
```

Expected: all tests pass.

- [ ] **Step 4: Commit**

```bash
git add reports/plans_pdf.py reports/pipeline_pdf.py
git commit -m "feat: pass base_url to WeasyPrint HTML() for font and image resolution"
```

---

### Task 3: Update templates/plans_pdf.html

**Files:**
- Modify: `templates/plans_pdf.html`

- [ ] **Step 1: Replace the full file content**

```html
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
@font-face { font-family: Roboto; src: url('./fonts/Roboto-Regular.ttf'); font-weight: 400; }
@font-face { font-family: Roboto; src: url('./fonts/Roboto-Bold.ttf'); font-weight: 700; }
@font-face { font-family: Merriweather; src: url('./fonts/Merriweather-Regular.ttf'); font-weight: 400; }
@font-face { font-family: Merriweather; src: url('./fonts/Merriweather-Bold.ttf'); font-weight: 700; }

* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: Roboto, sans-serif; font-size: 10pt; color: #1a1a2e; }

@page {
  size: letter;
  margin: 0 0 1.8cm 0;
  @bottom-left {
    content: "Diana Silver Grants Strategy \B7 diana@dsilvergrants.com";
    font-size: 8pt; color: #9ca3af; font-family: Roboto, sans-serif; padding-bottom: 0.3cm;
  }
  @bottom-right {
    content: "Page " counter(page);
    font-size: 8pt; color: #9ca3af; font-family: Roboto, sans-serif; padding-bottom: 0.3cm;
  }
}

.page-header { background: #4472C4; padding: 0; }
.header-table { width: 100%; border-collapse: collapse; }
.header-logo-cell { background: #ffffff; padding: 8px 20px; vertical-align: middle; width: 1%; white-space: nowrap; }
.header-logo { height: 40px; display: block; }
.header-date { color: #c5d5ee; font-size: 9pt; text-align: right; vertical-align: middle; padding: 8px 20px; }

.title-block { padding: 18px 24px 10px; }
.client-name { font-family: Merriweather, serif; font-size: 18pt; font-weight: 700; color: #1a1a2e; }
.report-type { font-family: Merriweather, serif; font-size: 12pt; color: #4472C4; font-weight: 700; margin-top: 3px; }

.year-section { padding: 6px 24px 12px; }
.year-label {
  font-family: Merriweather, serif; font-size: 12pt; font-weight: 700; color: #1a1a2e;
  border-bottom: 2px solid #4472C4; padding-bottom: 4px; margin-bottom: 10px;
}

.status-group { margin-bottom: 10px; }
.status-label {
  font-size: 8pt; font-weight: 700; color: #4472C4;
  text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 5px;
}

.grant-row {
  padding: 6px 10px; background: #f0f4ff; margin-bottom: 3px;
  border-radius: 0 3px 3px 0;
}
.grant-row-table { width: 100%; border-collapse: collapse; }
.grant-title { font-weight: 700; font-size: 10pt; }
.grant-amounts { font-size: 9pt; color: #374151; text-align: right; white-space: nowrap; }
.grant-details { font-size: 8pt; color: #6b7280; margin-top: 2px; }
</style>
</head>
<body>

<div class="page-header">
  <table class="header-table">
    <tr>
      <td class="header-logo-cell">
        <img src="./DS_Logo_Alt.png" class="header-logo" alt="">
      </td>
      <td class="header-date">Generated {{ today }}</td>
    </tr>
  </table>
</div>

<div class="title-block">
  <div class="client-name">{{ client }}</div>
  <div class="report-type">Plans Report</div>
</div>

{% for section in sections %}
<div class="year-section">
  <div class="year-label">{{ section.label }}</div>
  {% for group in section.groups %}
  <div class="status-group">
    <div class="status-label">{{ group.status }}</div>
    {% for row in group.rows %}
    <div class="grant-row" style="border-left: 3px solid {{ group.accent }};">
      <table class="grant-row-table">
        <tr>
          <td class="grant-title">
            {{ row.funder }}{% if row.fund %} &mdash; {{ row.fund }}{% endif %}
          </td>
          <td class="grant-amounts">
            Req: {{ row.request if row.request else "—" }}&nbsp;&nbsp;Awd: {{ row.award if row.award else "—" }}
          </td>
        </tr>
      </table>
      {% if row.details %}
      <div class="grant-details">{{ row.details }}</div>
      {% endif %}
    </div>
    {% endfor %}
  </div>
  {% endfor %}
</div>
{% endfor %}

</body>
</html>
```

- [ ] **Step 2: Run the plans PDF tests**

```bash
.venv/bin/pytest tests/test_plans_pdf.py -v
```

Expected: all tests pass.

- [ ] **Step 3: Commit**

```bash
git add templates/plans_pdf.html
git commit -m "feat: add logo and brand fonts to plans PDF template"
```

---

### Task 4: Update templates/pipeline_pdf.html

**Files:**
- Modify: `templates/pipeline_pdf.html`

- [ ] **Step 1: Replace the full file content**

```html
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
@font-face { font-family: Roboto; src: url('./fonts/Roboto-Regular.ttf'); font-weight: 400; }
@font-face { font-family: Roboto; src: url('./fonts/Roboto-Bold.ttf'); font-weight: 700; }
@font-face { font-family: Merriweather; src: url('./fonts/Merriweather-Regular.ttf'); font-weight: 400; }
@font-face { font-family: Merriweather; src: url('./fonts/Merriweather-Bold.ttf'); font-weight: 700; }

* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: Roboto, sans-serif; font-size: 10pt; color: #1a1a2e; }

@page {
  size: letter landscape;
  margin: 0 0 1.8cm 0;
  @bottom-left {
    content: "Diana Silver Grants Strategy \B7 diana@dsilvergrants.com";
    font-size: 8pt; color: #9ca3af; font-family: Roboto, sans-serif; padding-bottom: 0.3cm;
  }
  @bottom-right {
    content: "Page " counter(page);
    font-size: 8pt; color: #9ca3af; font-family: Roboto, sans-serif; padding-bottom: 0.3cm;
  }
}

.page-header { background: #4472C4; padding: 0; }
.header-table { width: 100%; border-collapse: collapse; }
.header-logo-cell { background: #ffffff; padding: 8px 20px; vertical-align: middle; width: 1%; white-space: nowrap; }
.header-logo { height: 40px; display: block; }
.header-date { color: #c5d5ee; font-size: 9pt; text-align: right; vertical-align: middle; padding: 8px 20px; }

.title-block { padding: 18px 24px 14px; }
.client-name { font-family: Merriweather, serif; font-size: 18pt; font-weight: 700; color: #1a1a2e; }
.report-type { font-family: Merriweather, serif; font-size: 12pt; color: #4472C4; font-weight: 700; margin-top: 3px; }

.data-table-wrap { padding: 0 24px 20px; }
.data-table { width: 100%; border-collapse: collapse; }
.data-table th {
  background: #4472C4; color: #ffffff; padding: 7px 8px;
  text-align: left; font-weight: 700; font-size: 9pt;
}
.data-table td { padding: 6px 8px; border-bottom: 1px solid #e5e7eb; font-size: 9pt; }
.data-table tr.even td { background: #f0f4ff; }
</style>
</head>
<body>

<div class="page-header">
  <table class="header-table">
    <tr>
      <td class="header-logo-cell">
        <img src="./DS_Logo_Alt.png" class="header-logo" alt="">
      </td>
      <td class="header-date">Generated {{ today }}</td>
    </tr>
  </table>
</div>

<div class="title-block">
  <div class="client-name">{{ client }}</div>
  <div class="report-type">Pipeline Report</div>
</div>

<div class="data-table-wrap">
  <table class="data-table">
    <thead>
      <tr>
        <th>Year</th><th>Project</th><th>Funder</th>
        <th>Task Type</th><th>Task Name</th><th>Due Date</th>
      </tr>
    </thead>
    <tbody>
      {% for row in rows %}
      <tr class="{{ 'even' if loop.index is even else 'odd' }}">
        <td>{{ row.year }}</td>
        <td>{{ row.project }}</td>
        <td>{{ row.funder }}</td>
        <td>{{ row.task_type }}</td>
        <td>{{ row.task_name }}</td>
        <td>{{ row.due_date if row.due_date else '' }}</td>
      </tr>
      {% endfor %}
    </tbody>
  </table>
</div>

</body>
</html>
```

- [ ] **Step 2: Run the pipeline PDF tests**

```bash
.venv/bin/pytest tests/test_pipeline_pdf.py -v
```

Expected: all tests pass.

- [ ] **Step 3: Run the full test suite**

```bash
.venv/bin/pytest tests/ -v
```

Expected: all tests pass.

- [ ] **Step 4: Commit**

```bash
git add templates/pipeline_pdf.html
git commit -m "feat: add logo and brand fonts to pipeline PDF template"
```
