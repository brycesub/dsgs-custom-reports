# Changelog

All notable changes to this project will be documented in this file.

## [1.0.0] - 2026-05-08

Initial release.

### Added

- **Plans report** — uploads a Plans CSV export and produces a formatted `.xlsx` per client (zipped when multiple clients); rows filtered to current year and forward, split across a current-year sheet and a future-year sheet, sorted by grant status rank then alphabetically by fund name
- **Pipeline report** — uploads a Pipeline CSV export (filename must match `Pipeline YYYYMMDD ClientName.csv`) and produces a single `.xlsx` sorted by due date ascending
- Drag-and-drop upload UI with inline error messages and a direct download link on success
- Docker + Gunicorn setup for LAN deployment (`docker compose up --build`)
- CSV input validation with descriptive user-facing error messages for missing columns, unparseable files, and bad filenames
- Shared Excel helpers (`parse_date`, `write_headers`, `autofit_columns`) in `reports/_utils.py`
- 69-test pytest suite covering all report logic and Flask routes
- Ruff linting and formatting enforced via pre-commit hooks
