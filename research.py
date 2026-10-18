"""Discover the stand-alone HTML deliverables of arts-funder-research studies.

Only each study's ``report.html`` and its network page (``viz/*.html`` other than
``template.html``) are ever exposed; everything else in the study folder is private.
"""

from dataclasses import dataclass
from datetime import date
from pathlib import Path

REPORT = "report.html"


@dataclass
class Study:
    slug: str
    title: str
    updated: date
    report: str
    network: str | None


def _title(study_dir: Path) -> str:
    md = study_dir / "report.md"
    if md.is_file():
        with md.open(encoding="utf-8") as f:
            for line in f:
                if line.startswith("# "):
                    return line[2:].strip()
    return study_dir.name


def _network(study_dir: Path) -> str | None:
    pages = sorted(p for p in (study_dir / "viz").glob("*.html") if p.name != "template.html")
    return f"viz/{pages[0].name}" if pages else None


def list_studies(root: str | Path | None) -> list[Study]:
    if not root:
        return []
    studies_dir = Path(root) / "studies"
    if not studies_dir.is_dir():
        return []

    studies = []
    for study_dir in studies_dir.iterdir():
        report = study_dir / REPORT
        if not report.is_file():
            continue
        studies.append(
            Study(
                slug=study_dir.name,
                title=_title(study_dir),
                updated=date.fromtimestamp(report.stat().st_mtime),
                report=REPORT,
                network=_network(study_dir),
            )
        )
    studies.sort(key=lambda s: s.title)
    studies.sort(key=lambda s: s.updated, reverse=True)
    return studies


def resolve_artifact(root: str | Path | None, slug: str, path: str) -> Path | None:
    """Return the file for ``path`` only if it is one of the study's whitelisted artifacts."""
    for study in list_studies(root):
        if study.slug == slug and path in (study.report, study.network):
            return Path(root) / "studies" / slug / path
    return None
