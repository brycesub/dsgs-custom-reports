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


def generate(csv_bytes: bytes, filename: str) -> tuple[str, bytes]:
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
    sheet_name = f"Pipeline - {date.today().isoformat()}"

    out = df[list(COLUMN_MAP.keys())].rename(columns=COLUMN_MAP).copy()
    out["Due Date"] = out["Due Date"].apply(parse_date)
    out = out.sort_values("Due Date", na_position="last")

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
