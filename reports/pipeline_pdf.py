from datetime import date, datetime
from pathlib import Path

import pandas as pd
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML

from reports.pipeline import _load_data

_TEMPLATES_DIR = Path(__file__).parent.parent / "templates"
_jinja_env = Environment(loader=FileSystemLoader(str(_TEMPLATES_DIR)), autoescape=True)


def _fmt_year(val) -> str:
    if pd.isna(val):
        return ""
    try:
        return str(int(val))
    except (ValueError, TypeError):
        return str(val)


def _fmt_date(val) -> str | None:
    if not isinstance(val, datetime) or pd.isnull(val):
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
            "year": _fmt_year(r["Year"]),
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
