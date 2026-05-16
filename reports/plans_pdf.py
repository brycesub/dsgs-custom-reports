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
                rows.append(
                    {
                        "funder": _safe_str(r["Funder"]),
                        "fund": _safe_str(r["Fund"]),
                        "request": _fmt_amount(r.get("Request")),
                        "award": _fmt_amount(r.get("Award")),
                        "details": " · ".join(parts),
                    }
                )
            groups.append(
                {
                    "status": status,
                    "accent": _ACCENT.get(status, "#c5d5ee"),
                    "rows": rows,
                }
            )
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
