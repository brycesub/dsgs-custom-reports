import os
from datetime import date

import pytest

from app import app as flask_app
from research import list_studies, resolve_artifact


@pytest.fixture
def research_dir(tmp_path):
    studies = tmp_path / "studies"

    kc = studies / "kansas-city-kc-rep"
    (kc / "viz").mkdir(parents=True)
    (kc / "research_notes").mkdir()
    (kc / "report.md").write_text("Intro\n\n# Kansas City arts funders\n\n## Section\n")
    (kc / "report.html").write_text("<html>kc report</html>")
    (kc / "viz" / "template.html").write_text("<html>template</html>")
    (kc / "viz" / "kc-network.html").write_text("<html>kc network</html>")
    (kc / "research_notes" / "notes.md").write_text("private notes")
    os.utime(kc / "report.html", (1_780_000_000, 1_780_000_000))  # 2026-05-28

    hfd = studies / "hartford-hartford-stage"
    hfd.mkdir()
    (hfd / "report.html").write_text("<html>hartford report</html>")
    os.utime(hfd / "report.html", (1_790_000_000, 1_790_000_000))  # 2026-09-21

    (studies / "in-progress").mkdir()
    (studies / "in-progress" / "report.md").write_text("# Not built yet\n")

    return tmp_path


@pytest.fixture
def client(research_dir):
    flask_app.config["TESTING"] = True
    flask_app.config["RESEARCH_DIR"] = str(research_dir)
    with flask_app.test_client() as c:
        yield c
    flask_app.config["RESEARCH_DIR"] = None


# ---------------------------------------------------------------------------
# list_studies / resolve_artifact
# ---------------------------------------------------------------------------


def test_list_studies_skips_dirs_without_report(research_dir):
    slugs = [s.slug for s in list_studies(research_dir)]
    assert slugs == ["hartford-hartford-stage", "kansas-city-kc-rep"]  # newest first


def test_list_studies_title_from_report_md(research_dir):
    kc = next(s for s in list_studies(research_dir) if s.slug == "kansas-city-kc-rep")
    assert kc.title == "Kansas City arts funders"
    assert kc.updated == date.fromtimestamp(1_780_000_000)


def test_list_studies_title_falls_back_to_slug(research_dir):
    hfd = next(s for s in list_studies(research_dir) if s.slug == "hartford-hartford-stage")
    assert hfd.title == "hartford-hartford-stage"


def test_list_studies_network_excludes_template(research_dir):
    studies = {s.slug: s for s in list_studies(research_dir)}
    assert studies["kansas-city-kc-rep"].network == "viz/kc-network.html"
    assert studies["hartford-hartford-stage"].network is None


@pytest.mark.parametrize("root", [None, "", "/nonexistent/path"])
def test_list_studies_unconfigured(root):
    assert list_studies(root) == []


def test_resolve_artifact_only_whitelisted(research_dir):
    assert resolve_artifact(research_dir, "kansas-city-kc-rep", "report.html") is not None
    assert resolve_artifact(research_dir, "kansas-city-kc-rep", "report.md") is None
    assert resolve_artifact(None, "kansas-city-kc-rep", "report.html") is None


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------


def test_research_index_lists_studies(client):
    resp = client.get("/research")
    assert resp.status_code == 200
    body = resp.get_data(as_text=True)
    assert "Kansas City arts funders" in body
    assert "hartford-hartford-stage" in body
    assert "/research/kansas-city-kc-rep/report.html" in body
    assert "/research/kansas-city-kc-rep/viz/kc-network.html" in body
    assert "Not built yet" not in body


def test_research_index_empty_state():
    flask_app.config["TESTING"] = True
    flask_app.config["RESEARCH_DIR"] = None
    with flask_app.test_client() as c:
        resp = c.get("/research")
    assert resp.status_code == 200
    assert "No studies found" in resp.get_data(as_text=True)


def test_index_links_to_research(client):
    assert "/research" in client.get("/").get_data(as_text=True)


@pytest.mark.parametrize(
    "url, body",
    [
        ("/research/kansas-city-kc-rep/report.html", "kc report"),
        ("/research/kansas-city-kc-rep/viz/kc-network.html", "kc network"),
        ("/research/hartford-hartford-stage/report.html", "hartford report"),
    ],
)
def test_research_artifact_served(client, url, body):
    resp = client.get(url)
    assert resp.status_code == 200
    assert resp.mimetype == "text/html"
    assert resp.get_data(as_text=True) == f"<html>{body}</html>"


@pytest.mark.parametrize(
    "url",
    [
        "/research/kansas-city-kc-rep/report.md",
        "/research/kansas-city-kc-rep/research_notes/notes.md",
        "/research/kansas-city-kc-rep/viz/template.html",
        "/research/no-such-study/report.html",
        "/research/in-progress/report.md",
        "/research/kansas-city-kc-rep/../hartford-hartford-stage/report.html",
        "/research/kansas-city-kc-rep/viz/../../hartford-hartford-stage/report.html",
    ],
)
def test_research_artifact_not_whitelisted_404(client, url):
    assert client.get(url).status_code == 404


def test_research_artifact_unconfigured_404():
    flask_app.config["TESTING"] = True
    flask_app.config["RESEARCH_DIR"] = None
    with flask_app.test_client() as c:
        assert c.get("/research/kansas-city-kc-rep/report.html").status_code == 404
