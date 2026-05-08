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


def generate(csv_bytes: bytes) -> list[tuple[str, bytes]]:
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

    results = []
    current_year = date.today().year
    future_label = f"{current_year + 1}+"

    for client, client_df in out.groupby("_client"):
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        cur_df = client_df[client_df["Year"] == current_year]
        fut_df = client_df[client_df["Year"] > current_year]

        ws_cur = wb.create_sheet(title=str(current_year))
        _build_worksheet(ws_cur, cur_df[OUTPUT_COLUMNS])

        if not fut_df.empty:
            ws_fut = wb.create_sheet(title=future_label)
            _build_worksheet(ws_fut, fut_df[OUTPUT_COLUMNS])

        wb.active = wb[str(current_year)]

        xlsx_buf = io.BytesIO()
        wb.save(xlsx_buf)
        results.append((client, xlsx_buf.getvalue()))

    return results
