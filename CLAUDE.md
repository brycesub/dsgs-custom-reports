# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Environment

Always use the `.venv` virtualenv:

```bash
.venv/bin/python   # run Python
.venv/bin/flask    # Flask CLI
```

Install runtime deps: `pip install -r requirements.txt`
Install dev deps: `pip install -r requirements-dev.txt`
Install pre-commit hooks: `.venv/bin/pre-commit install`

## Common commands

```bash
# Run dev server
.venv/bin/python app.py

# Lint + format check
.venv/bin/ruff check .
.venv/bin/ruff format --check .

# Auto-fix lint and format
.venv/bin/ruff check --fix . && .venv/bin/ruff format .

# Run tests
.venv/bin/pytest tests/ -v

# Docker (LAN deployment)
docker compose up --build
```

Tests live in `tests/`. Fixtures are synthetic inline CSVs — the real customer files in `input/` (gitignored) are for manual smoke-testing only. Date-sensitive logic is tested by mocking `date.today()` to a fixed date.

## Architecture

Flask app (`app.py`) with two routes: `POST /generate/plans` and `POST /generate/pipeline`. Each accepts a `csv_file` form field and returns an `.xlsx` directly (or `.zip` for multi-client plans).

Report logic lives in `reports/plans.py` and `reports/pipeline.py`. Each module exposes a single `generate()` function that takes `csv_bytes` (and `filename` for pipeline) and returns the processed data. `app.py` only handles HTTP concerns; all CSV parsing, sorting, and Excel building is in the report modules.

Shared Excel utilities (header styling, column autofit, date parsing) live in `reports/_utils.py` — import from there rather than duplicating.

### plans report
- Input CSV must have a `Client` column — one `.xlsx` is produced per unique client value. If multiple clients, `app.py` zips them.
- Rows filtered to the current fiscal year (July–June) and forward; split into one sheet per fiscal year labeled `FY {year}`, with the current fiscal year left-most and active.
- Sorted by `STATUS_ORDER` rank, then alphabetically by `Fund` within each group.

### pipeline report
- Single-client; client name is parsed from the upload filename (`Pipeline YYYYMMDD {Client}.csv`).
- All rows included (no date filtering), single sheet named `Pipeline - YYYY-MM-DD`, sorted by Due Date ascending.
- Download filename: `Pipeline_{client}_{date}.xlsx`.

### PDF rendering
Preferred fonts for rendered PDFs are Merriweather (serif) and Roboto (sans-serif).

Logo assets for use in PDFs live in `templates/`:
- `DS_Logo_Primary.png`
- `DS_Logo_Alt.png`
- `DS_Logo_Secondary2.png`

### adding a new report
1. Create `reports/<name>.py` with a `generate()` function; import shared helpers from `reports/_utils.py`.
2. Add a `POST /generate/<name>` route in `app.py`.
3. Add an upload card to `templates/index.html` and wire it to `initUploadArea()` with the new endpoint.

## Linting

Ruff enforces E, F, I rules at line-length 100 with double quotes. Pre-commit runs `ruff --fix` and `ruff-format` on every commit.
