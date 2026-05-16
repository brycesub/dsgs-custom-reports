# PDF Generation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add WeasyPrint-based PDF export to both the Plans and Pipeline reports, alongside the existing Excel export.

**Architecture:** Extract shared CSV-parsing logic from `plans.py` and `pipeline.py` into private `_load_data()` helpers, then build new `plans_pdf.py` and `pipeline_pdf.py` modules that call those helpers and render Jinja2 HTML templates to PDF via WeasyPrint. Two new Flask routes and a UI button split complete the feature.

**Tech Stack:** WeasyPrint (HTML→PDF), Jinja2 (templates), Flask (routes), existing pandas/openpyxl stack unchanged.

---

## File Map

| Action | Path | Responsibility |
|---|---|---|
| Modify | `requirements.txt` | Add `weasyprint` |
| Modify | `Dockerfile` | Add system deps for WeasyPrint |
| Modify | `reports/plans.py` | Extract `_load_data()` helper; `generate()` calls it |
| Modify | `reports/pipeline.py` | Extract `_load_data()` helper; `generate()` calls it |
| Create | `reports/plans_pdf.py` | `generate(csv_bytes) -> list[tuple[str, bytes]]` |
| Create | `reports/pipeline_pdf.py` | `generate(csv_bytes, filename) -> tuple[str, bytes]` |
| Create | `templates/plans_pdf.html` | Jinja2 HTML template for Plans PDF |
| Create | `templates/pipeline_pdf.html` | Jinja2 HTML template for Pipeline PDF |
| Modify | `app.py` | Two new routes + two new imports |
| Modify | `templates/index.html` | Split button into Excel + PDF per card |
| Modify | `tests/test_app.py` | Tests for new PDF routes |
| Create | `tests/test_plans_pdf.py` | Tests for Plans PDF generator |
| Create | `tests/test_pipeline_pdf.py` | Tests for Pipeline PDF generator |
| Modify | `.gitignore` | Add `.superpowers/` |

---

## Task 1: Install WeasyPrint

**Files:**
- Modify: `requirements.txt`
- Modify: `Dockerfile`

- [ ] **Step 1: Add weasyprint to requirements.txt**

  Open `requirements.txt` and append:
  ```
  weasyprint
  ```
  Final file:
  ```
  flask==3.1.3
  gunicorn==26.0.0
  pandas==3.0.3
  openpyxl==3.1.5
  weasyprint
  ```

- [ ] **Step 2: Add system deps to Dockerfile**

  The current `Dockerfile` is:
  ```dockerfile
  FROM python:3.12-slim
  WORKDIR /app
  COPY requirements.txt .
  RUN pip install --no-cache-dir -r requirements.txt
  COPY . .
  EXPOSE 5000
  CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--workers", "2", "app:app"]
  ```

  Replace with:
  ```dockerfile
  FROM python:3.12-slim
  WORKDIR /app
  RUN apt-get update && apt-get install -y --no-install-recommends \
      libpango-1.0-0 libcairo2 libgdk-pixbuf2.0-0 \
      && rm -rf /var/lib/apt/lists/*
  COPY requirements.txt .
  RUN pip install --no-cache-dir -r requirements.txt
  COPY . .
  EXPOSE 5000
  CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--workers", "2", "app:app"]
  ```

- [ ] **Step 3: Install into the virtualenv**

  Run: `.venv/bin/pip install weasyprint`

- [ ] **Step 4: Verify import**

  Run: `.venv/bin/python -c "from weasyprint import HTML; print('ok')"`
  Expected: `ok`

- [ ] **Step 5: Update .gitignore**

  Add `.superpowers/` to `.gitignore`:
  ```
  # Brainstorming / superpowers session files
  .superpowers/
  ```

- [ ] **Step 6: Commit**

  ```bash
  git add requirements.txt Dockerfile .gitignore
  git commit -m "feat: add weasyprint dep and system deps for PDF generation"
  ```

---

## Task 2: Refactor plans.py — extract _load_data()

**Files:**
- Modify: `reports/plans.py`

The goal is to pull all CSV-parsing, validation, and sorting logic into a private `_load_data()` function so that both the Excel `generate()` and the upcoming `plans_pdf.py` can reuse it without duplication.

- [ ] **Step 1: Run existing tests to confirm baseline**

  Run: `.venv/bin/pytest tests/test_plans.py -v`
  Expected: all tests PASS

- [ ] **Step 2: Refactor reports/plans.py**

  Replace the entire file with:
  ```python
  import io
  import re
  from datetime import date, datetime

  import openpyxl
  import pandas as pd

  from reports._utils import autofit_columns, parse_date, write_headers

  STATUS_ORDER = [
      "Awarded - Active",
      "Awarded - Closed",
      "Application Submitted",
      "LOI Submitted",
      "Application In Progress",
      "LOI In Progress",
      "Planned",
      "Researching",
      "Declined",
      "Abandoned",
  ]

  COLUMN_MAP = {
      "Year": "Year",
      "Funder name": "Funder",
      "Project": "Fund",
      "Request Purpose": "Purpose",
      "Status": "Status",
      "Amount requested": "Request",
      "Amount awarded": "Award",
      "Expected notification date": "Notif Expected",
      "Notification date": "Notif Received",
      "Next task deadline": "Next Task/Deadline",
  }

  OUTPUT_COLUMNS = [
      "Year",
      "Funder",
      "Fund",
      "Purpose",
      "Status",
      "Request",
      "Award",
      "Notif Expected",
      "Notif Received",
      "Next Task/Deadline",
  ]


  def _parse_amount(val):
      if pd.isna(val) or not str(val).strip():
          return None
      s = re.sub(r"[^\d.]", "", str(val))
      try:
          return float(s) if s else None
      except ValueError:
          return None


  def _build_worksheet(ws, df):
      write_headers(ws, OUTPUT_COLUMNS)

      for row_idx, (_, row) in enumerate(df.iterrows(), 2):
          for col_idx, col_name in enumerate(OUTPUT_COLUMNS, 1):
              val = row.get(col_name)
              cell = ws.cell(row=row_idx, column=col_idx)

              if col_name in ("Notif Expected", "Notif Received"):
                  if isinstance(val, datetime):
                      cell.value = val
                      cell.number_format = "MM/DD/YYYY"
                  else:
                      cell.value = None
              elif col_name in ("Request", "Award"):
                  cell.value = val
                  if val is not None:
                      cell.number_format = "#,##0"
              else:
                  cell.value = None if (isinstance(val, float) and pd.isna(val)) else val

      autofit_columns(ws, OUTPUT_COLUMNS)


  def _load_data(csv_bytes: bytes) -> list[tuple[str, pd.DataFrame, pd.DataFrame]]:
      """Parse and validate a Plans CSV.

      Returns a list of (client, cur_year_df, future_df) tuples, one per unique
      client, where both DataFrames have OUTPUT_COLUMNS columns and rows are
      sorted by STATUS_ORDER rank then Fund name.
      """
      try:
          df = pd.read_csv(io.BytesIO(csv_bytes))
      except (pd.errors.ParserError, UnicodeDecodeError, pd.errors.EmptyDataError) as e:
          raise ValueError(f"This file couldn't be read as a CSV: {e}")

      if df.empty:
          raise ValueError("The uploaded file has no data rows.")

      missing = [c for c in list(COLUMN_MAP.keys()) + ["Client"] if c not in df.columns]
      if missing:
          raise ValueError(f"CSV is missing expected columns: {', '.join(missing)}")

      clients = df["Client"].fillna("Unknown")
      out = df[list(COLUMN_MAP.keys())].rename(columns=COLUMN_MAP).copy()
      out["_client"] = clients.values

      out["Request"] = out["Request"].apply(_parse_amount)
      out["Award"] = out["Award"].apply(_parse_amount)
      out["Notif Expected"] = out["Notif Expected"].apply(parse_date)
      out["Notif Received"] = out["Notif Received"].apply(parse_date)

      out["Year"] = pd.to_numeric(out["Year"], errors="coerce")
      bad_year_count = int(out["Year"].isna().sum())
      if bad_year_count:
          raise ValueError(
              f"{bad_year_count} row(s) have non-numeric Year values — please check the CSV."
          )
      out["Year"] = out["Year"].astype(int)

      status_rank = {s: i for i, s in enumerate(STATUS_ORDER)}
      out["_status_rank"] = out["Status"].map(status_rank).fillna(len(STATUS_ORDER))
      out = out.sort_values(["_status_rank", "Fund"], na_position="last")
      out = out[out["Year"] >= date.today().year]
      if out.empty:
          raise ValueError(
              f"No rows match {date.today().year} or later — "
              "please check the file contains current data."
          )

      current_year = date.today().year
      results = []
      for client, client_df in out.groupby("_client"):
          cur_df = client_df[client_df["Year"] == current_year][OUTPUT_COLUMNS].reset_index(drop=True)
          fut_df = client_df[client_df["Year"] > current_year][OUTPUT_COLUMNS].reset_index(drop=True)
          results.append((client, cur_df, fut_df))
      return results


  def generate(csv_bytes: bytes) -> list[tuple[str, bytes]]:
      results = []
      current_year = date.today().year
      future_label = f"{current_year + 1}+"

      for client, cur_df, fut_df in _load_data(csv_bytes):
          wb = openpyxl.Workbook()
          wb.remove(wb.active)

          ws_cur = wb.create_sheet(title=str(current_year))
          _build_worksheet(ws_cur, cur_df)

          if not fut_df.empty:
              ws_fut = wb.create_sheet(title=future_label)
              _build_worksheet(ws_fut, fut_df)

          wb.active = wb[str(current_year)]

          xlsx_buf = io.BytesIO()
          wb.save(xlsx_buf)
          results.append((client, xlsx_buf.getvalue()))

      return results
  ```

- [ ] **Step 3: Run tests to confirm refactor didn't break anything**

  Run: `.venv/bin/pytest tests/test_plans.py -v`
  Expected: all tests PASS (same count as Step 1)

- [ ] **Step 4: Commit**

  ```bash
  git add reports/plans.py
  git commit -m "refactor: extract _load_data() from plans.generate() for PDF reuse"
  ```

---

## Task 3: Refactor pipeline.py — extract _load_data()

**Files:**
- Modify: `reports/pipeline.py`

- [ ] **Step 1: Run existing tests to confirm baseline**

  Run: `.venv/bin/pytest tests/test_pipeline.py -v`
  Expected: all tests PASS

- [ ] **Step 2: Refactor reports/pipeline.py**

  Replace the entire file with:
  ```python
  import io
  import re
  from datetime import date, datetime

  import openpyxl
  import pandas as pd

  from reports._utils import autofit_columns, parse_date, write_headers

  COLUMN_MAP = {
      "Year": "Year",
      "Project": "Project",
      "Funder name": "Funder",
      "Task type": "Task Type",
      "Task Title": "Task Name",
      "Due Date": "Due Date",
  }

  OUTPUT_COLUMNS = ["Year", "Project", "Funder", "Task Type", "Task Name", "Due Date"]

  _PIPELINE_FILENAME_RE = re.compile(r"^Pipeline \d{8} .+\.csv$", re.IGNORECASE)


  def _extract_client(filename: str) -> str:
      if not _PIPELINE_FILENAME_RE.match(filename):
          raise ValueError(
              f"Filename must match 'Pipeline YYYYMMDD ClientName.csv' — got: {filename!r}"
          )
      stem = filename[:-4]
      parts = stem.split(" ", 2)
      return parts[2]


  def _load_data(csv_bytes: bytes, filename: str) -> tuple[str, pd.DataFrame]:
      """Parse and validate a Pipeline CSV.

      Returns (client_name, df) where df has OUTPUT_COLUMNS columns sorted by Due Date.
      """
      try:
          df = pd.read_csv(io.BytesIO(csv_bytes))
      except (pd.errors.ParserError, UnicodeDecodeError, pd.errors.EmptyDataError) as e:
          raise ValueError(f"This file couldn't be read as a CSV: {e}")

      if df.empty:
          raise ValueError("The uploaded file has no data rows.")

      missing = [c for c in COLUMN_MAP if c not in df.columns]
      if missing:
          raise ValueError(f"CSV is missing expected columns: {', '.join(missing)}")

      client = _extract_client(filename)
      out = df[list(COLUMN_MAP.keys())].rename(columns=COLUMN_MAP).copy()
      out["Due Date"] = out["Due Date"].apply(parse_date)
      out = out.sort_values("Due Date", na_position="last")
      return client, out


  def generate(csv_bytes: bytes, filename: str) -> tuple[str, bytes]:
      client, out = _load_data(csv_bytes, filename)
      sheet_name = f"Pipeline - {date.today().isoformat()}"

      wb = openpyxl.Workbook()
      wb.remove(wb.active)
      ws = wb.create_sheet(title=sheet_name)

      write_headers(ws, OUTPUT_COLUMNS)

      for row_idx, (_, row) in enumerate(out.iterrows(), 2):
          for col_idx, col_name in enumerate(OUTPUT_COLUMNS, 1):
              val = row.get(col_name)
              cell = ws.cell(row=row_idx, column=col_idx)
              if col_name == "Due Date":
                  if isinstance(val, datetime):
                      cell.value = val
                      cell.number_format = "MM/DD/YYYY"
                  else:
                      cell.value = None
              else:
                  cell.value = None if (isinstance(val, float) and pd.isna(val)) else val

      autofit_columns(ws, OUTPUT_COLUMNS)

      xlsx_buf = io.BytesIO()
      wb.save(xlsx_buf)
      return client, xlsx_buf.getvalue()
  ```

- [ ] **Step 3: Run tests to confirm refactor didn't break anything**

  Run: `.venv/bin/pytest tests/test_pipeline.py -v`
  Expected: all tests PASS

- [ ] **Step 4: Commit**

  ```bash
  git add reports/pipeline.py
  git commit -m "refactor: extract _load_data() from pipeline.generate() for PDF reuse"
  ```

---

## Task 4: Plans PDF — template + generator

**Files:**
- Create: `templates/plans_pdf.html`
- Create: `reports/plans_pdf.py`
- Create: `tests/test_plans_pdf.py`

- [ ] **Step 1: Write tests/test_plans_pdf.py**

  ```python
  from datetime import date
  from unittest.mock import MagicMock, patch

  import pytest

  from reports.plans_pdf import generate

  FIXED_TODAY = date(2026, 1, 15)

  _HEADER = (
      "Year,Client,Funder name,Project,Request Purpose,Status,"
      "Amount requested,Amount awarded,Expected notification date,"
      "Notification date,Next task deadline"
  )


  def _csv(*rows):
      return ("\n".join([_HEADER] + list(rows))).encode()


  def _row(
      year=2026,
      client="KCRep",
      funder="Test Funder",
      project="Test Fund",
      purpose="General",
      status="Planned",
      request="10000",
      award="",
      notif_expected="",
      notif_received="",
      next_deadline="",
  ):
      return (
          f"{year},{client},{funder},{project},{purpose},{status},"
          f"{request},{award},{notif_expected},{notif_received},{next_deadline}"
      )


  @pytest.fixture(autouse=True)
  def freeze_today():
      mock = MagicMock()
      mock.today.return_value = FIXED_TODAY
      with patch("reports.plans.date", mock):
          yield


  class TestGeneratePlansPdf:
      def test_returns_pdf_bytes(self):
          _, pdf_bytes = generate(_csv(_row()))[0]
          assert pdf_bytes[:4] == b"%PDF"

      def test_returns_one_result_per_client(self):
          results = generate(_csv(_row(client="KCRep"), _row(client="GOH")))
          assert {r[0] for r in results} == {"KCRep", "GOH"}

      def test_client_name_in_result(self):
          client, _ = generate(_csv(_row(client="WEN")))[0]
          assert client == "WEN"

      def test_future_rows_included(self):
          results = generate(_csv(_row(year=2026), _row(year=2027)))
          assert len(results) == 1
          _, pdf_bytes = results[0]
          assert pdf_bytes[:4] == b"%PDF"

      def test_raises_for_missing_column(self):
          with pytest.raises(ValueError, match="missing expected columns"):
              generate(b"Year,Client\n2026,KCRep\n")

      def test_raises_for_header_only_csv(self):
          with pytest.raises(ValueError, match="no data rows"):
              generate((_HEADER + "\n").encode())

      def test_raises_for_all_past_years(self):
          with pytest.raises(ValueError, match="No rows match"):
              generate(_csv(_row(year=2024)))
  ```

- [ ] **Step 2: Run tests to confirm they fail with ImportError**

  Run: `.venv/bin/pytest tests/test_plans_pdf.py -v`
  Expected: FAIL — `ModuleNotFoundError: No module named 'reports.plans_pdf'`

- [ ] **Step 3: Create templates/plans_pdf.html**

  ```html
  <!DOCTYPE html>
  <html>
  <head>
  <meta charset="UTF-8">
  <style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: Arial, Helvetica, sans-serif; font-size: 10pt; color: #1a1a2e; }

  @page {
    size: letter;
    margin: 0 0 1.8cm 0;
    @bottom-left {
      content: "Diana Silver Grants Strategy · diana@dsilvergrants.com";
      font-size: 8pt; color: #9ca3af; font-family: Arial, sans-serif; padding-bottom: 0.3cm;
    }
    @bottom-right {
      content: "Page " counter(page);
      font-size: 8pt; color: #9ca3af; font-family: Arial, sans-serif; padding-bottom: 0.3cm;
    }
  }

  .page-header { background: #4472C4; padding: 14px 24px; }
  .header-table { width: 100%; border-collapse: collapse; }
  .header-name { color: #ffffff; font-size: 15pt; font-weight: 700; }
  .header-email { color: #c5d5ee; font-size: 9pt; margin-top: 2px; }
  .header-date { color: #c5d5ee; font-size: 9pt; text-align: right; vertical-align: top; }

  .title-block { padding: 18px 24px 10px; }
  .client-name { font-size: 18pt; font-weight: 700; color: #1a1a2e; }
  .report-type { font-size: 12pt; color: #4472C4; font-weight: 600; margin-top: 3px; }

  .year-section { padding: 6px 24px 12px; }
  .year-label {
    font-size: 12pt; font-weight: 700; color: #1a1a2e;
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
        <td>
          <div class="header-name">Diana Silver Grants Strategy</div>
          <div class="header-email">diana@dsilvergrants.com</div>
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
              Req: {{ row.request or "—" }}&nbsp;&nbsp;Awd: {{ row.award or "—" }}
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

- [ ] **Step 4: Create reports/plans_pdf.py**

  ```python
  from datetime import date, datetime
  from pathlib import Path

  import pandas as pd
  from jinja2 import Environment, FileSystemLoader
  from weasyprint import HTML

  from reports.plans import STATUS_ORDER, _load_data

  _TEMPLATES_DIR = Path(__file__).parent.parent / "templates"
  _jinja_env = Environment(loader=FileSystemLoader(str(_TEMPLATES_DIR)), autoescape=True)

  _ACCENT = {
      "Awarded - Active": "#4472C4",
      "Awarded - Closed": "#4472C4",
      "Application Submitted": "#93afd4",
      "LOI Submitted": "#93afd4",
      "Application In Progress": "#93afd4",
      "LOI In Progress": "#93afd4",
      "Planned": "#c5d5ee",
      "Researching": "#c5d5ee",
      "Declined": "#c5d5ee",
      "Abandoned": "#c5d5ee",
  }


  def _fmt_amount(val) -> str | None:
      if val is None or (isinstance(val, float) and pd.isna(val)):
          return None
      return f"${int(val):,}"


  def _fmt_date(val) -> str | None:
      if not isinstance(val, datetime):
          return None
      return val.strftime("%m/%d/%Y")


  def _safe_str(val) -> str | None:
      if val is None or (isinstance(val, float) and pd.isna(val)):
          return None
      return str(val)


  def _build_context(client: str, cur_df: pd.DataFrame, fut_df: pd.DataFrame) -> dict:
      current_year = date.today().year

      def _groups(df: pd.DataFrame) -> list[dict]:
          groups = []
          for status in STATUS_ORDER:
              status_df = df[df["Status"] == status]
              if status_df.empty:
                  continue
              rows = []
              for _, r in status_df.iterrows():
                  parts = []
                  if not pd.isna(r["Year"]):
                      parts.append(str(int(r["Year"])))
                  purpose = _safe_str(r.get("Purpose"))
                  if purpose:
                      parts.append(purpose)
                  exp = _fmt_date(r.get("Notif Expected"))
                  if exp:
                      parts.append(f"Exp: {exp}")
                  rec = _fmt_date(r.get("Notif Received"))
                  if rec:
                      parts.append(f"Rcvd: {rec}")
                  nxt = _safe_str(r.get("Next Task/Deadline"))
                  if nxt:
                      parts.append(nxt)
                  rows.append({
                      "funder": _safe_str(r["Funder"]),
                      "fund": _safe_str(r["Fund"]),
                      "request": _fmt_amount(r.get("Request")),
                      "award": _fmt_amount(r.get("Award")),
                      "details": " · ".join(parts),
                  })
              groups.append({
                  "status": status,
                  "accent": _ACCENT.get(status, "#c5d5ee"),
                  "rows": rows,
              })
          return groups

      sections = []
      if not cur_df.empty:
          sections.append({"label": str(current_year), "groups": _groups(cur_df)})
      if not fut_df.empty:
          sections.append({"label": f"{current_year + 1}+", "groups": _groups(fut_df)})

      return {
          "client": client,
          "today": date.today().strftime("%B %d, %Y"),
          "sections": sections,
      }


  def generate(csv_bytes: bytes) -> list[tuple[str, bytes]]:
      results = []
      for client, cur_df, fut_df in _load_data(csv_bytes):
          ctx = _build_context(client, cur_df, fut_df)
          html = _jinja_env.get_template("plans_pdf.html").render(**ctx)
          pdf_bytes = HTML(string=html).write_pdf()
          results.append((client, pdf_bytes))
      return results
  ```

- [ ] **Step 5: Run tests — they should pass**

  Run: `.venv/bin/pytest tests/test_plans_pdf.py -v`
  Expected: all 7 tests PASS

  If WeasyPrint logs warnings to stderr, that's fine — only assert on the return value.

- [ ] **Step 6: Run full test suite to confirm nothing broken**

  Run: `.venv/bin/pytest tests/ -v`
  Expected: all tests PASS

- [ ] **Step 7: Commit**

  ```bash
  git add reports/plans_pdf.py templates/plans_pdf.html tests/test_plans_pdf.py
  git commit -m "feat: add Plans PDF generator with WeasyPrint"
  ```

---

## Task 5: Pipeline PDF — template + generator

**Files:**
- Create: `templates/pipeline_pdf.html`
- Create: `reports/pipeline_pdf.py`
- Create: `tests/test_pipeline_pdf.py`

- [ ] **Step 1: Write tests/test_pipeline_pdf.py**

  ```python
  from datetime import date
  from unittest.mock import MagicMock, patch

  import pytest

  from reports.pipeline_pdf import generate

  FIXED_TODAY = date(2026, 5, 8)
  _VALID_FILENAME = "Pipeline 20260508 HS.csv"
  _HEADER = "Year,Project,Funder name,Task type,Task Title,Due Date"


  def _csv(*rows):
      return ("\n".join([_HEADER] + list(rows))).encode()


  def _row(
      year=2026,
      project="Test Project",
      funder="Test Funder",
      task_type="Report",
      task_title="My Task",
      due_date="2026-05-01",
  ):
      return f"{year},{project},{funder},{task_type},{task_title},{due_date}"


  @pytest.fixture(autouse=True)
  def freeze_today():
      mock = MagicMock()
      mock.today.return_value = FIXED_TODAY
      with patch("reports.pipeline.date", mock):
          yield


  class TestGeneratePipelinePdf:
      def test_returns_client_and_pdf_bytes(self):
          client, pdf_bytes = generate(_csv(_row()), _VALID_FILENAME)
          assert client == "HS"
          assert pdf_bytes[:4] == b"%PDF"

      def test_multi_word_client(self):
          filename = "Pipeline 20260508 Hartford Stage.csv"
          client, _ = generate(_csv(_row()), filename)
          assert client == "Hartford Stage"

      def test_raises_for_bad_filename(self):
          with pytest.raises(ValueError, match="Filename must match"):
              generate(_csv(_row()), "badfile.csv")

      def test_raises_for_missing_column(self):
          with pytest.raises(ValueError, match="missing expected columns"):
              generate(b"Year,Project\n2026,P\n", _VALID_FILENAME)

      def test_raises_for_header_only_csv(self):
          with pytest.raises(ValueError, match="no data rows"):
              generate((_HEADER + "\n").encode(), _VALID_FILENAME)
  ```

- [ ] **Step 2: Run tests to confirm they fail with ImportError**

  Run: `.venv/bin/pytest tests/test_pipeline_pdf.py -v`
  Expected: FAIL — `ModuleNotFoundError: No module named 'reports.pipeline_pdf'`

- [ ] **Step 3: Create templates/pipeline_pdf.html**

  ```html
  <!DOCTYPE html>
  <html>
  <head>
  <meta charset="UTF-8">
  <style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: Arial, Helvetica, sans-serif; font-size: 10pt; color: #1a1a2e; }

  @page {
    size: letter landscape;
    margin: 0 0 1.8cm 0;
    @bottom-left {
      content: "Diana Silver Grants Strategy · diana@dsilvergrants.com";
      font-size: 8pt; color: #9ca3af; font-family: Arial, sans-serif; padding-bottom: 0.3cm;
    }
    @bottom-right {
      content: "Page " counter(page);
      font-size: 8pt; color: #9ca3af; font-family: Arial, sans-serif; padding-bottom: 0.3cm;
    }
  }

  .page-header { background: #4472C4; padding: 14px 24px; }
  .header-table { width: 100%; border-collapse: collapse; }
  .header-name { color: #ffffff; font-size: 15pt; font-weight: 700; }
  .header-email { color: #c5d5ee; font-size: 9pt; margin-top: 2px; }
  .header-date { color: #c5d5ee; font-size: 9pt; text-align: right; vertical-align: top; }

  .title-block { padding: 18px 24px 14px; }
  .client-name { font-size: 18pt; font-weight: 700; color: #1a1a2e; }
  .report-type { font-size: 12pt; color: #4472C4; font-weight: 600; margin-top: 3px; }

  .data-table-wrap { padding: 0 24px 20px; }
  .data-table { width: 100%; border-collapse: collapse; }
  .data-table th {
    background: #4472C4; color: #ffffff; padding: 7px 8px;
    text-align: left; font-weight: 600; font-size: 9pt;
  }
  .data-table td { padding: 6px 8px; border-bottom: 1px solid #e5e7eb; font-size: 9pt; }
  .data-table tr.even td { background: #f0f4ff; }
  </style>
  </head>
  <body>

  <div class="page-header">
    <table class="header-table">
      <tr>
        <td>
          <div class="header-name">Diana Silver Grants Strategy</div>
          <div class="header-email">diana@dsilvergrants.com</div>
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
          <td>{{ row.due_date or '' }}</td>
        </tr>
        {% endfor %}
      </tbody>
    </table>
  </div>

  </body>
  </html>
  ```

- [ ] **Step 4: Create reports/pipeline_pdf.py**

  ```python
  from datetime import date, datetime
  from pathlib import Path

  import pandas as pd
  from jinja2 import Environment, FileSystemLoader
  from weasyprint import HTML

  from reports.pipeline import _load_data

  _TEMPLATES_DIR = Path(__file__).parent.parent / "templates"
  _jinja_env = Environment(loader=FileSystemLoader(str(_TEMPLATES_DIR)), autoescape=True)


  def _fmt_date(val) -> str | None:
      if not isinstance(val, datetime):
          return None
      return val.strftime("%m/%d/%Y")


  def _safe(val) -> str:
      if val is None or (isinstance(val, float) and pd.isna(val)):
          return ""
      return str(val)


  def generate(csv_bytes: bytes, filename: str) -> tuple[str, bytes]:
      client, out = _load_data(csv_bytes, filename)
      rows = [
          {
              "year": _safe(r["Year"]),
              "project": _safe(r["Project"]),
              "funder": _safe(r["Funder"]),
              "task_type": _safe(r["Task Type"]),
              "task_name": _safe(r["Task Name"]),
              "due_date": _fmt_date(r["Due Date"]),
          }
          for _, r in out.iterrows()
      ]
      ctx = {
          "client": client,
          "today": date.today().strftime("%B %d, %Y"),
          "rows": rows,
      }
      html = _jinja_env.get_template("pipeline_pdf.html").render(**ctx)
      pdf_bytes = HTML(string=html).write_pdf()
      return client, pdf_bytes
  ```

- [ ] **Step 5: Run tests — they should pass**

  Run: `.venv/bin/pytest tests/test_pipeline_pdf.py -v`
  Expected: all 5 tests PASS

- [ ] **Step 6: Run full test suite**

  Run: `.venv/bin/pytest tests/ -v`
  Expected: all tests PASS

- [ ] **Step 7: Commit**

  ```bash
  git add reports/pipeline_pdf.py templates/pipeline_pdf.html tests/test_pipeline_pdf.py
  git commit -m "feat: add Pipeline PDF generator with WeasyPrint"
  ```

---

## Task 6: Flask routes + app route tests

**Files:**
- Modify: `tests/test_app.py` (add new test classes)
- Modify: `app.py` (add imports + two routes)

- [ ] **Step 1: Add route tests to tests/test_app.py**

  Append to the end of the existing `tests/test_app.py`:

  ```python
  # ---------------------------------------------------------------------------
  # /generate/plans/pdf
  # ---------------------------------------------------------------------------


  class TestPlansPdfRoute:
      def test_no_file_returns_400(self, client):
          assert client.post("/generate/plans/pdf").status_code == 400

      def test_non_csv_returns_400(self, client):
          assert _upload(client, "/generate/plans/pdf", b"data", "file.txt").status_code == 400

      def test_single_client_returns_pdf(self, client):
          with patch("app.generate_plans_pdf", return_value=[("KCRep", b"fake pdf")]):
              resp = _upload(client, "/generate/plans/pdf", b"csv", "plans.csv")
          assert resp.status_code == 200
          assert resp.content_type == "application/pdf"
          assert resp.data == b"fake pdf"

      def test_single_client_download_name(self, client):
          with patch("app.generate_plans_pdf", return_value=[("My Client", b"bytes")]):
              resp = _upload(client, "/generate/plans/pdf", b"csv", "plans.csv")
          assert "My Client_report.pdf" in resp.headers.get("Content-Disposition", "")

      def test_multi_client_returns_zip(self, client):
          with patch("app.generate_plans_pdf", return_value=[("A", b"a"), ("B", b"b")]):
              resp = _upload(client, "/generate/plans/pdf", b"csv", "plans.csv")
          assert resp.status_code == 200
          assert resp.content_type == "application/zip"

      def test_multi_client_zip_contains_pdfs(self, client):
          with patch("app.generate_plans_pdf", return_value=[("A", b"a"), ("B", b"b")]):
              resp = _upload(client, "/generate/plans/pdf", b"csv", "plans.csv")
          zf = zipfile.ZipFile(io.BytesIO(resp.data))
          assert set(zf.namelist()) == {"A_report.pdf", "B_report.pdf"}

      def test_value_error_returns_400(self, client):
          with patch("app.generate_plans_pdf", side_effect=ValueError("bad column")):
              resp = _upload(client, "/generate/plans/pdf", b"csv", "plans.csv")
          assert resp.status_code == 400
          assert b"bad column" in resp.data

      def test_unexpected_error_returns_500(self, client):
          with patch("app.generate_plans_pdf", side_effect=RuntimeError("boom")):
              resp = _upload(client, "/generate/plans/pdf", b"csv", "plans.csv")
          assert resp.status_code == 500


  # ---------------------------------------------------------------------------
  # /generate/pipeline/pdf
  # ---------------------------------------------------------------------------


  class TestPipelinePdfRoute:
      def test_no_file_returns_400(self, client):
          assert client.post("/generate/pipeline/pdf").status_code == 400

      def test_non_csv_returns_400(self, client):
          assert _upload(client, "/generate/pipeline/pdf", b"data", "file.txt").status_code == 400

      def test_valid_csv_returns_pdf(self, client):
          with patch("app.generate_pipeline_pdf", return_value=("HS", b"fake pdf")):
              resp = _upload(client, "/generate/pipeline/pdf", b"csv", "Pipeline 20260508 HS.csv")
          assert resp.status_code == 200
          assert resp.content_type == "application/pdf"
          assert resp.data == b"fake pdf"

      def test_download_name_includes_client_and_date(self, client):
          with (
              patch("app.generate_pipeline_pdf", return_value=("HS", b"bytes")),
              patch("app.date") as mock_date,
          ):
              mock_date.today.return_value = date(2026, 5, 8)
              resp = _upload(client, "/generate/pipeline/pdf", b"csv", "Pipeline 20260508 HS.csv")
          assert "Pipeline_HS_2026-05-08.pdf" in resp.headers.get("Content-Disposition", "")

      def test_value_error_returns_400(self, client):
          with patch("app.generate_pipeline_pdf", side_effect=ValueError("bad filename")):
              resp = _upload(client, "/generate/pipeline/pdf", b"csv", "Pipeline 20260508 HS.csv")
          assert resp.status_code == 400
          assert b"bad filename" in resp.data

      def test_unexpected_error_returns_500(self, client):
          with patch("app.generate_pipeline_pdf", side_effect=RuntimeError("boom")):
              resp = _upload(client, "/generate/pipeline/pdf", b"csv", "Pipeline 20260508 HS.csv")
          assert resp.status_code == 500
  ```

- [ ] **Step 2: Run new route tests to confirm they fail**

  Run: `.venv/bin/pytest tests/test_app.py::TestPlansPdfRoute tests/test_app.py::TestPipelinePdfRoute -v`
  Expected: FAIL — routes do not exist yet (404s or AttributeError on the patch target)

- [ ] **Step 3: Update app.py**

  Replace the entire file with:
  ```python
  import io
  import zipfile
  from datetime import date

  from flask import Flask, render_template, request, send_file

  from reports.pipeline import generate as generate_pipeline
  from reports.pipeline_pdf import generate as generate_pipeline_pdf
  from reports.plans import generate as generate_plans
  from reports.plans_pdf import generate as generate_plans_pdf

  app = Flask(__name__)
  app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB upload limit


  def _get_csv_file():
      f = request.files.get("csv_file")
      if not f or not f.filename:
          return None, ("No file uploaded.", 400)
      if not f.filename.lower().endswith(".csv"):
          return None, ("Please upload a CSV file.", 400)
      return f, None


  @app.route("/")
  def index():
      return render_template("index.html")


  @app.route("/generate/plans", methods=["POST"])
  def generate_plans_report():
      f, err = _get_csv_file()
      if err:
          return err

      try:
          results = generate_plans(f.read())
      except ValueError as e:
          return str(e), 400
      except Exception:
          app.logger.exception("Unexpected error generating plans report")
          return "An unexpected error occurred. Please try again.", 500

      if len(results) == 1:
          client, xlsx_bytes = results[0]
          return send_file(
              io.BytesIO(xlsx_bytes),
              mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              as_attachment=True,
              download_name=f"{client}_report.xlsx",
          )

      zip_buf = io.BytesIO()
      with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
          for client, xlsx_bytes in results:
              zf.writestr(f"{client}_report.xlsx", xlsx_bytes)
      zip_buf.seek(0)
      return send_file(
          zip_buf,
          mimetype="application/zip",
          as_attachment=True,
          download_name="reports.zip",
      )


  @app.route("/generate/plans/pdf", methods=["POST"])
  def generate_plans_pdf_report():
      f, err = _get_csv_file()
      if err:
          return err

      try:
          results = generate_plans_pdf(f.read())
      except ValueError as e:
          return str(e), 400
      except Exception:
          app.logger.exception("Unexpected error generating plans PDF report")
          return "An unexpected error occurred. Please try again.", 500

      if len(results) == 1:
          client, pdf_bytes = results[0]
          return send_file(
              io.BytesIO(pdf_bytes),
              mimetype="application/pdf",
              as_attachment=True,
              download_name=f"{client}_report.pdf",
          )

      zip_buf = io.BytesIO()
      with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
          for client, pdf_bytes in results:
              zf.writestr(f"{client}_report.pdf", pdf_bytes)
      zip_buf.seek(0)
      return send_file(
          zip_buf,
          mimetype="application/zip",
          as_attachment=True,
          download_name="reports.zip",
      )


  @app.route("/generate/pipeline", methods=["POST"])
  def generate_pipeline_report():
      f, err = _get_csv_file()
      if err:
          return err

      try:
          client, xlsx_bytes = generate_pipeline(f.read(), f.filename)
      except ValueError as e:
          return str(e), 400
      except Exception:
          app.logger.exception("Unexpected error generating pipeline report")
          return "An unexpected error occurred. Please try again.", 500

      today = date.today().isoformat()
      return send_file(
          io.BytesIO(xlsx_bytes),
          mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          as_attachment=True,
          download_name=f"Pipeline_{client}_{today}.xlsx",
      )


  @app.route("/generate/pipeline/pdf", methods=["POST"])
  def generate_pipeline_pdf_report():
      f, err = _get_csv_file()
      if err:
          return err

      try:
          client, pdf_bytes = generate_pipeline_pdf(f.read(), f.filename)
      except ValueError as e:
          return str(e), 400
      except Exception:
          app.logger.exception("Unexpected error generating pipeline PDF report")
          return "An unexpected error occurred. Please try again.", 500

      today = date.today().isoformat()
      return send_file(
          io.BytesIO(pdf_bytes),
          mimetype="application/pdf",
          as_attachment=True,
          download_name=f"Pipeline_{client}_{today}.pdf",
      )


  if __name__ == "__main__":
      app.run(debug=True)
  ```

- [ ] **Step 4: Run all tests**

  Run: `.venv/bin/pytest tests/ -v`
  Expected: all tests PASS

- [ ] **Step 5: Commit**

  ```bash
  git add app.py tests/test_app.py
  git commit -m "feat: add /generate/plans/pdf and /generate/pipeline/pdf routes"
  ```

---

## Task 7: UI — two-button layout

**Files:**
- Modify: `templates/index.html`

- [ ] **Step 1: Update templates/index.html**

  Replace the entire file with the content below. Key changes:
  - Add `.btn-group` flex container CSS + `.btn-pdf` teal variant
  - Rename `btnId`/`endpoint` params to `excelBtnId`/`excelEndpoint` in `initUploadArea`
  - Add optional `pdfBtnId`/`pdfEndpoint` params
  - Extract shared fetch handler into inner `handleGenerate(endpoint, triggerBtn)` function
  - HTML: replace each single `<button>` with a `<div class="btn-group">` containing Excel + PDF buttons

  ```html
  <!DOCTYPE html>
  <html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Custom Reports</title>
    <style>
      *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

      body {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        background: #f0f2f5;
        min-height: 100vh;
        display: flex;
        align-items: flex-start;
        justify-content: center;
        padding: 48px 24px;
      }

      .page {
        display: flex;
        gap: 24px;
        flex-wrap: wrap;
        justify-content: center;
        width: 100%;
      }

      .card {
        background: #fff;
        border-radius: 12px;
        box-shadow: 0 2px 16px rgba(0,0,0,0.10);
        padding: 40px 48px;
        width: 100%;
        max-width: 520px;
      }

      h1 {
        font-size: 1.4rem;
        font-weight: 600;
        color: #1a1a2e;
        margin-bottom: 8px;
      }

      .subtitle {
        font-size: 0.875rem;
        color: #6b7280;
        margin-bottom: 28px;
      }

      .drop-zone {
        border: 2px dashed #d1d5db;
        border-radius: 8px;
        padding: 36px 20px;
        text-align: center;
        cursor: pointer;
        transition: border-color 0.2s, background 0.2s;
        position: relative;
      }

      .drop-zone.drag-over {
        border-color: #4472C4;
        background: #eef2fb;
      }

      .drop-zone.has-file {
        border-color: #16a34a;
        background: #f0fdf4;
      }

      .drop-zone input[type="file"] {
        position: absolute;
        inset: 0;
        opacity: 0;
        cursor: pointer;
        width: 100%;
        height: 100%;
      }

      .drop-icon {
        font-size: 2rem;
        margin-bottom: 10px;
        color: #9ca3af;
      }

      .drop-zone.has-file .drop-icon { color: #16a34a; }

      .drop-label {
        font-size: 0.9rem;
        color: #374151;
        font-weight: 500;
      }

      .drop-hint {
        font-size: 0.8rem;
        color: #9ca3af;
        margin-top: 4px;
      }

      .file-name {
        font-size: 0.85rem;
        color: #16a34a;
        font-weight: 500;
        margin-top: 6px;
        word-break: break-all;
      }

      .btn-group {
        display: flex;
        gap: 8px;
        margin-top: 20px;
      }

      .btn {
        flex: 1;
        padding: 12px;
        background: #4472C4;
        color: #fff;
        font-size: 0.95rem;
        font-weight: 600;
        border: none;
        border-radius: 8px;
        cursor: pointer;
        transition: background 0.2s;
      }

      .btn:hover:not(:disabled) { background: #3461b0; }
      .btn:disabled { background: #9ca3af; cursor: not-allowed; }

      .btn-pdf { background: #0d9488; }
      .btn-pdf:hover:not(:disabled) { background: #0f766e; }

      .status {
        margin-top: 16px;
        font-size: 0.875rem;
        text-align: center;
        min-height: 1.4em;
      }

      .status.error { color: #dc2626; }
      .status.loading { color: #6b7280; }

      .download-link {
        display: none;
        margin-top: 16px;
        padding: 12px;
        background: #f0fdf4;
        border: 1px solid #86efac;
        border-radius: 8px;
        text-align: center;
        font-size: 0.9rem;
        color: #15803d;
        font-weight: 500;
        text-decoration: none;
        transition: background 0.2s;
      }

      .download-link:hover { background: #dcfce7; }
      .download-link.visible { display: block; }

      .subtitle code {
        font-family: ui-monospace, "SFMono-Regular", Menlo, monospace;
        font-size: 0.8rem;
        background: #f3f4f6;
        border-radius: 4px;
        padding: 1px 5px;
      }
    </style>
  </head>
  <body>
    <div class="page">

      <div class="card">
        <h1>Plans Report</h1>
        <p class="subtitle">Upload a Plans CSV export to generate the report.</p>

        <div class="drop-zone" id="plansDropZone">
          <input type="file" id="plansFileInput" accept=".csv" />
          <div class="drop-icon" id="plansDropIcon">📄</div>
          <div class="drop-label" id="plansDropLabel">Drag &amp; drop a CSV file here</div>
          <div class="drop-hint">or click to browse</div>
          <div class="file-name" id="plansFileName"></div>
        </div>

        <div class="btn-group">
          <button class="btn" id="plansExcelBtn" disabled>Generate Excel</button>
          <button class="btn btn-pdf" id="plansPdfBtn" disabled>Generate PDF</button>
        </div>
        <div class="status" id="plansStatus"></div>
        <a class="download-link" id="plansDownload" href="#"></a>
      </div>

      <div class="card">
        <h1>Pipeline Report</h1>
        <p class="subtitle">Upload a Pipeline CSV export to generate the report. Filename must be in the format <code>Pipeline YYYYMMDD ClientName.csv</code>.</p>

        <div class="drop-zone" id="pipelineDropZone">
          <input type="file" id="pipelineFileInput" accept=".csv" />
          <div class="drop-icon" id="pipelineDropIcon">📄</div>
          <div class="drop-label" id="pipelineDropLabel">Drag &amp; drop a CSV file here</div>
          <div class="drop-hint">or click to browse</div>
          <div class="file-name" id="pipelineFileName"></div>
        </div>

        <div class="btn-group">
          <button class="btn" id="pipelineExcelBtn" disabled>Generate Excel</button>
          <button class="btn btn-pdf" id="pipelinePdfBtn" disabled>Generate PDF</button>
        </div>
        <div class="status" id="pipelineStatus"></div>
        <a class="download-link" id="pipelineDownload" href="#"></a>
      </div>

    </div>

    <script>
      function initUploadArea({
        dropZoneId, fileInputId, dropIconId, dropLabelId, fileNameId,
        excelBtnId, excelEndpoint, pdfBtnId, pdfEndpoint,
        statusId, downloadId
      }) {
        const dropZone = document.getElementById(dropZoneId);
        const fileInput = document.getElementById(fileInputId);
        const dropIcon = document.getElementById(dropIconId);
        const dropLabel = document.getElementById(dropLabelId);
        const fileName = document.getElementById(fileNameId);
        const excelBtn = document.getElementById(excelBtnId);
        const pdfBtn = pdfBtnId ? document.getElementById(pdfBtnId) : null;
        const status = document.getElementById(statusId);
        const downloadLink = document.getElementById(downloadId);

        let selectedFile = null;

        function setFile(file) {
          if (!file || !file.name.toLowerCase().endsWith(".csv")) {
            status.textContent = "Please select a .csv file.";
            status.className = "status error";
            return;
          }
          selectedFile = file;
          fileName.textContent = file.name;
          dropIcon.textContent = "✅";
          dropLabel.textContent = "File selected";
          dropZone.classList.add("has-file");
          excelBtn.disabled = false;
          if (pdfBtn) pdfBtn.disabled = false;
          status.textContent = "";
          status.className = "status";
          downloadLink.classList.remove("visible");
        }

        fileInput.addEventListener("change", () => {
          if (fileInput.files[0]) setFile(fileInput.files[0]);
        });

        dropZone.addEventListener("dragover", (e) => {
          e.preventDefault();
          dropZone.classList.add("drag-over");
        });

        dropZone.addEventListener("dragleave", () => {
          dropZone.classList.remove("drag-over");
        });

        dropZone.addEventListener("drop", (e) => {
          e.preventDefault();
          dropZone.classList.remove("drag-over");
          const file = e.dataTransfer.files[0];
          if (file) setFile(file);
        });

        async function handleGenerate(endpoint, triggerBtn) {
          if (!selectedFile) return;

          excelBtn.disabled = true;
          if (pdfBtn) pdfBtn.disabled = true;
          downloadLink.classList.remove("visible");
          status.textContent = "Generating report…";
          status.className = "status loading";

          const formData = new FormData();
          formData.append("csv_file", selectedFile);

          try {
            const res = await fetch(endpoint, { method: "POST", body: formData });

            if (!res.ok) {
              const msg = await res.text();
              status.textContent = msg || "Something went wrong.";
              status.className = "status error";
              excelBtn.disabled = false;
              if (pdfBtn) pdfBtn.disabled = false;
              return;
            }

            const disposition = res.headers.get("Content-Disposition") || "";
            const match = disposition.match(/filename="([^"]+)"/);
            const filename = match ? match[1] : "report";

            const blob = await res.blob();
            const url = URL.createObjectURL(blob);
            downloadLink.href = url;
            downloadLink.download = filename;
            downloadLink.textContent = "⬇ Download " + filename;
            downloadLink.classList.add("visible");
            status.textContent = "";
            status.className = "status";
          } catch (err) {
            status.textContent = "Network error — is the server running?";
            status.className = "status error";
          }

          excelBtn.disabled = false;
          if (pdfBtn) pdfBtn.disabled = false;
        }

        excelBtn.addEventListener("click", () => handleGenerate(excelEndpoint, excelBtn));
        if (pdfBtn) pdfBtn.addEventListener("click", () => handleGenerate(pdfEndpoint, pdfBtn));
      }

      initUploadArea({
        dropZoneId: "plansDropZone",
        fileInputId: "plansFileInput",
        dropIconId: "plansDropIcon",
        dropLabelId: "plansDropLabel",
        fileNameId: "plansFileName",
        excelBtnId: "plansExcelBtn",
        excelEndpoint: "/generate/plans",
        pdfBtnId: "plansPdfBtn",
        pdfEndpoint: "/generate/plans/pdf",
        statusId: "plansStatus",
        downloadId: "plansDownload",
      });

      initUploadArea({
        dropZoneId: "pipelineDropZone",
        fileInputId: "pipelineFileInput",
        dropIconId: "pipelineDropIcon",
        dropLabelId: "pipelineDropLabel",
        fileNameId: "pipelineFileName",
        excelBtnId: "pipelineExcelBtn",
        excelEndpoint: "/generate/pipeline",
        pdfBtnId: "pipelinePdfBtn",
        pdfEndpoint: "/generate/pipeline/pdf",
        statusId: "pipelineStatus",
        downloadId: "pipelineDownload",
      });
    </script>
  </body>
  </html>
  ```

- [ ] **Step 2: Run full test suite to confirm nothing broken**

  Run: `.venv/bin/pytest tests/ -v`
  Expected: all tests PASS (the HTML change doesn't affect automated tests, but confirms no other breakage)

- [ ] **Step 3: Manual smoke test**

  Start the server: `.venv/bin/python app.py`

  Open `http://localhost:5000` and verify:
  - Each card shows two side-by-side buttons: "Generate Excel" (blue) and "Generate PDF" (teal)
  - Both buttons are disabled before a file is selected
  - After selecting a CSV, both buttons enable
  - "Generate Excel" downloads a `.xlsx` as before
  - "Generate PDF" downloads a `.pdf` with the correct client name

- [ ] **Step 4: Commit**

  ```bash
  git add templates/index.html
  git commit -m "feat: split Generate Report button into Generate Excel + Generate PDF"
  ```
