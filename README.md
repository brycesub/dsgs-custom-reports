# Custom Reports

A small Flask web app that converts CSV exports from a grant-tracking system into formatted Excel reports. Upload a CSV through the browser; download a ready-to-use `.xlsx` (or `.zip` for multi-client plans).

Two reports are supported:

| Report | Input | Output |
|---|---|---|
| **Plans** | Plans CSV export | One `.xlsx` per client (zipped if multiple), filtered to current fiscal year+, one tab per FY |
| **Pipeline** | Pipeline CSV export | Single `.xlsx` sorted by due date |

---

## Running locally (development)

**Prerequisites:** Python 3.12+

```bash
# Create and activate a virtualenv
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
pip install -r requirements-dev.txt

# Install pre-commit hooks (runs ruff on every commit)
pre-commit install

# Start the dev server
python app.py
```

Open [http://localhost:5000](http://localhost:5000).

---

## Running with Docker (LAN deployment)

```bash
docker compose up --build
```

The app is served by Gunicorn on port 5000 with 2 workers. `restart: unless-stopped` keeps it running across reboots. A health check polls `GET /` every 30 seconds.

To expose it on a specific address, edit the `ports` line in `docker-compose.yml`:

```yaml
ports:
  - "192.168.1.10:5000:5000"
```

---

## Reports

### Plans report

**Endpoint:** `POST /generate/plans`

**Input:** A CSV export with at least these columns:

| Column | Notes |
|---|---|
| `Year` | Numeric (e.g. `2027`; `FY 2027` prefixes tolerated) |
| `Client` | Used to split output into per-client files |
| `Funder name` | |
| `Project` | Appears as "Fund" in the output |
| `Request Purpose` | |
| `Status` | See status order below |
| `Amount requested` | Dollar amounts; `$` and commas stripped automatically |
| `Amount awarded` | Same |
| `Expected notification date` | Accepts `Mon DD, YYYY`, `MM/DD/YYYY`, or `YYYY-MM-DD` |
| `Notification date` | Same |
| `Next task deadline` | Same |

Extra columns in the CSV are ignored.

**Processing:**
- Rows tagged with a fiscal year before the current one are filtered out. Fiscal years run July 1 – June 30, so in August 2026 the current fiscal year is FY 2027.
- Rows are split into one sheet per fiscal year — **`FY 2027`** (left-most and active), then **`FY 2028`**, etc. — each omitted if it has no rows.
- Within each sheet, rows are sorted first by status rank, then alphabetically by Fund:

| Rank | Status |
|---|---|
| 1 | Awarded - Active |
| 2 | Awarded - Closed |
| 3 | Application Submitted |
| 4 | LOI Submitted |
| 5 | Application In Progress |
| 6 | LOI In Progress |
| 7 | Planned |
| 8 | Researching |
| 9 | Declined |
| 10 | Abandoned |
| — | Any other status (sorted last) |

**Output:**
- Single client → `{Client}_report.xlsx`
- Multiple clients → `reports.zip` containing one `.xlsx` per client
- `Request`/`Award` amounts use Excel's **Currency** format (`$175,000`, whole dollars); `Notif Expected`/`Notif Received` render as `m/d/yy` (e.g. `6/30/26`).

---

### Pipeline report

**Endpoint:** `POST /generate/pipeline`

**Filename convention:** The upload filename must match `Pipeline YYYYMMDD ClientName.csv` (case-insensitive extension). The client name is parsed from the filename — everything after the date token.

Examples of valid filenames:
```
Pipeline 20260508 HS.csv
Pipeline 20260508 GOH.csv
```

**Input:** A CSV export with at least these columns:

| Column | Notes |
|---|---|
| `Year` | |
| `Project` | |
| `Funder name` | |
| `Task type` | |
| `Task Title` | Appears as "Task Name" in output |
| `Due Date` | Accepts `Mon DD, YYYY`, `MM/DD/YYYY`, or `YYYY-MM-DD` |

Extra columns in the CSV are ignored. All rows are included (no date filtering).

**Processing:** Rows are sorted by Due Date ascending; rows with no due date appear last.

**Output:** `Pipeline_{client}_{YYYY-MM-DD}.xlsx` — a single sheet named `Pipeline - YYYY-MM-DD`.

---

## PDF rendering

Preferred fonts for rendered PDFs are **Merriweather** (serif) and **Roboto** (sans-serif).

Logo assets live in `templates/`:
- `DS_Logo_Primary.png`
- `DS_Logo_Alt.png`
- `DS_Logo_Secondary2.png`

---

## Testing

```bash
.venv/bin/pytest tests/ -v
```

69 tests covering `reports/_utils.py`, `reports/plans.py`, `reports/pipeline.py`, and the Flask routes. Tests use synthetic inline CSV fixtures — no external files required.

---

## Project structure

```
app.py                  # Flask routes (HTTP concerns only)
reports/
    _utils.py           # Shared helpers: parse_date, write_headers, autofit_columns
    plans.py            # Plans report: generate(csv_bytes) → [(client, xlsx_bytes), ...]
    pipeline.py         # Pipeline report: generate(csv_bytes, filename) → (client, xlsx_bytes)
templates/
    index.html          # Single-page UI with drag-and-drop upload cards
tests/
    test_utils.py       # Unit tests for _utils helpers
    test_plans.py       # Unit tests for plans report logic
    test_pipeline.py    # Unit tests for pipeline report logic
    test_app.py         # Integration tests for Flask routes
Dockerfile
docker-compose.yml
requirements.txt        # Runtime dependencies
requirements-dev.txt    # ruff, pre-commit, pytest
pyproject.toml          # Ruff config (line length 100, E/F/I rules, double quotes)
```

---

## Adding a new report

1. Create `reports/<name>.py` with a `generate()` function. Import shared helpers from `reports/_utils.py`:
   - `parse_date(val)` — parses date strings to `datetime` or `None`
   - `write_headers(ws, columns)` — writes a styled header row and freezes pane at A2
   - `autofit_columns(ws, columns)` — sets column widths based on content (capped at 50)

2. Add a `POST /generate/<name>` route in `app.py` following the existing pattern.

3. Add an upload card to `templates/index.html` and wire it to `initUploadArea()` with the new endpoint.

4. Add a test file `tests/test_<name>.py` with synthetic inline CSV fixtures.

---

## Development

**Lint and format:**

```bash
# Check
.venv/bin/ruff check .
.venv/bin/ruff format --check .

# Auto-fix
.venv/bin/ruff check --fix . && .venv/bin/ruff format .
```

Pre-commit runs both automatically on every commit.

**Dependencies:** runtime deps are pinned in `requirements.txt`. Update with `pip install -U <package>` and re-pin with `pip freeze`.
