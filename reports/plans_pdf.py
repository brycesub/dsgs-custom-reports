from datetime import date, datetime
from pathlib import Path

import pandas as pd
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML

from reports.plans import STATUS_ORDER, _load_data

_TEMPLATES_DIR = Path(__file__).parent.parent / "templates"
_jinja_env = Environment(loader=FileSystemLoader(str(_TEMPLATES_DIR)), autoescape=True)

_ACCENT = {
    "Awarded - Active": "#1580BD",
    "Awarded - Closed": "#1580BD",
    "Application Submitted": "#7FB6DD",
    "LOI Submitted": "#7FB6DD",
    "Application In Progress": "#7FB6DD",
    "LOI In Progress": "#7FB6DD",
    "Planned": "#CFE3F2",
    "Researching": "#CFE3F2",
    "Declined": "#CFE3F2",
    "Abandoned": "#CFE3F2",
}


def _fmt_year(val) -> str:
    if pd.isna(val):
        return ""
    try:
        return str(int(val))
    except (ValueError, TypeError):
        return str(val)


def _fmt_amount(val) -> str | None:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    return f"${int(val):,}"


def _fmt_date(val) -> str | None:
    if not isinstance(val, datetime) or pd.isnull(val):
        return None
    return val.strftime("%m/%d/%Y")


def _safe_str(val) -> str | None:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    return str(val)


def _build_context(client: str, fy_map: dict[int, pd.DataFrame]) -> dict:
    def _groups(df: pd.DataFrame) -> list[dict]:
        groups = []
        for status in STATUS_ORDER:
            status_df = df[df["Status"] == status]
            if status_df.empty:
                continue
            rows = []
            for _, r in status_df.iterrows():
                parts = []
                year_str = _fmt_year(r["Year"])
                if year_str:
                    parts.append(year_str)
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
                    "accent": _ACCENT.get(status, "#CFE3F2"),
                    "rows": rows,
                }
            )
        return groups

    sections = [
        {"label": f"FY {year}", "groups": _groups(year_df)} for year, year_df in fy_map.items()
    ]
    return {
        "client": client,
        "today": date.today().strftime("%B %d, %Y"),
        "sections": sections,
    }


def generate(csv_bytes: bytes) -> list[tuple[str, bytes]]:
    results = []
    for client, fy_map in _load_data(csv_bytes):
        ctx = _build_context(client, fy_map)
        html = _jinja_env.get_template("plans_pdf.html").render(**ctx)
        pdf_bytes = HTML(string=html, base_url=str(_TEMPLATES_DIR)).write_pdf()
        results.append((client, pdf_bytes))
    return results
