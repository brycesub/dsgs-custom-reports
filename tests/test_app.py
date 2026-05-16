import io
import zipfile
from datetime import date
from unittest.mock import patch

import pytest

from app import app as flask_app


@pytest.fixture
def client():
    flask_app.config["TESTING"] = True
    with flask_app.test_client() as c:
        yield c


def _upload(client, endpoint, data, filename):
    return client.post(
        endpoint,
        data={"csv_file": (io.BytesIO(data), filename)},
        content_type="multipart/form-data",
    )


# ---------------------------------------------------------------------------
# Index
# ---------------------------------------------------------------------------


def test_index_returns_200(client):
    resp = client.get("/")
    assert resp.status_code == 200


# ---------------------------------------------------------------------------
# /generate/plans
# ---------------------------------------------------------------------------


class TestPlansRoute:
    def test_no_file_returns_400(self, client):
        assert client.post("/generate/plans").status_code == 400

    def test_non_csv_returns_400(self, client):
        assert _upload(client, "/generate/plans", b"data", "file.txt").status_code == 400

    def test_single_client_returns_xlsx(self, client):
        with patch("app.generate_plans", return_value=[("Acme", b"fake xlsx")]):
            resp = _upload(client, "/generate/plans", b"csv", "plans.csv")
        assert resp.status_code == 200
        assert "spreadsheetml" in resp.content_type
        assert resp.data == b"fake xlsx"

    def test_single_client_download_name(self, client):
        with patch("app.generate_plans", return_value=[("My Client", b"bytes")]):
            resp = _upload(client, "/generate/plans", b"csv", "plans.csv")
        assert "My Client_report.xlsx" in resp.headers.get("Content-Disposition", "")

    def test_multi_client_returns_zip(self, client):
        with patch("app.generate_plans", return_value=[("A", b"aaa"), ("B", b"bbb")]):
            resp = _upload(client, "/generate/plans", b"csv", "plans.csv")
        assert resp.status_code == 200
        assert resp.content_type == "application/zip"
        zf = zipfile.ZipFile(io.BytesIO(resp.data))
        assert set(zf.namelist()) == {"A_report.xlsx", "B_report.xlsx"}

    def test_multi_client_zip_download_name(self, client):
        with patch("app.generate_plans", return_value=[("A", b"a"), ("B", b"b")]):
            resp = _upload(client, "/generate/plans", b"csv", "plans.csv")
        assert "reports.zip" in resp.headers.get("Content-Disposition", "")

    def test_value_error_returns_400_with_message(self, client):
        with patch("app.generate_plans", side_effect=ValueError("bad column")):
            resp = _upload(client, "/generate/plans", b"csv", "plans.csv")
        assert resp.status_code == 400
        assert b"bad column" in resp.data

    def test_unexpected_error_returns_500(self, client):
        with patch("app.generate_plans", side_effect=RuntimeError("boom")):
            resp = _upload(client, "/generate/plans", b"csv", "plans.csv")
        assert resp.status_code == 500


# ---------------------------------------------------------------------------
# /generate/pipeline
# ---------------------------------------------------------------------------


class TestPipelineRoute:
    def test_no_file_returns_400(self, client):
        assert client.post("/generate/pipeline").status_code == 400

    def test_non_csv_returns_400(self, client):
        assert _upload(client, "/generate/pipeline", b"data", "file.txt").status_code == 400

    def test_valid_csv_returns_xlsx(self, client):
        with patch("app.generate_pipeline", return_value=("HS", b"fake xlsx")):
            resp = _upload(client, "/generate/pipeline", b"csv", "Pipeline 20260508 HS.csv")
        assert resp.status_code == 200
        assert "spreadsheetml" in resp.content_type
        assert resp.data == b"fake xlsx"

    def test_download_name_includes_client_and_date(self, client):
        with (
            patch("app.generate_pipeline", return_value=("HS", b"bytes")),
            patch("app.date") as mock_date,
        ):
            mock_date.today.return_value = date(2026, 5, 8)
            resp = _upload(client, "/generate/pipeline", b"csv", "Pipeline 20260508 HS.csv")
        assert "Pipeline_HS_2026-05-08.xlsx" in resp.headers.get("Content-Disposition", "")

    def test_value_error_returns_400_with_message(self, client):
        with patch("app.generate_pipeline", side_effect=ValueError("bad filename")):
            resp = _upload(client, "/generate/pipeline", b"csv", "Pipeline 20260508 HS.csv")
        assert resp.status_code == 400
        assert b"bad filename" in resp.data

    def test_unexpected_error_returns_500(self, client):
        with patch("app.generate_pipeline", side_effect=RuntimeError("boom")):
            resp = _upload(client, "/generate/pipeline", b"csv", "Pipeline 20260508 HS.csv")
        assert resp.status_code == 500


# ---------------------------------------------------------------------------
# /generate/plans/pdf
# ---------------------------------------------------------------------------


class TestPlansPdfRoute:
    def test_no_file_returns_400(self, client):
        assert client.post("/generate/plans/pdf").status_code == 400

    def test_non_csv_returns_400(self, client):
        assert _upload(client, "/generate/plans/pdf", b"data", "file.txt").status_code == 400

    def test_single_client_returns_pdf(self, client):
        with patch("app.generate_plans_pdf", return_value=[("KCRep", b"fake pdf")]):
            resp = _upload(client, "/generate/plans/pdf", b"csv", "plans.csv")
        assert resp.status_code == 200
        assert resp.content_type == "application/pdf"
        assert resp.data == b"fake pdf"

    def test_single_client_download_name(self, client):
        with patch("app.generate_plans_pdf", return_value=[("My Client", b"bytes")]):
            resp = _upload(client, "/generate/plans/pdf", b"csv", "plans.csv")
        assert "My Client_report.pdf" in resp.headers.get("Content-Disposition", "")

    def test_multi_client_returns_zip(self, client):
        with patch("app.generate_plans_pdf", return_value=[("A", b"a"), ("B", b"b")]):
            resp = _upload(client, "/generate/plans/pdf", b"csv", "plans.csv")
        assert resp.status_code == 200
        assert resp.content_type == "application/zip"

    def test_multi_client_zip_contains_pdfs(self, client):
        with patch("app.generate_plans_pdf", return_value=[("A", b"a"), ("B", b"b")]):
            resp = _upload(client, "/generate/plans/pdf", b"csv", "plans.csv")
        zf = zipfile.ZipFile(io.BytesIO(resp.data))
        assert set(zf.namelist()) == {"A_report.pdf", "B_report.pdf"}

    def test_value_error_returns_400(self, client):
        with patch("app.generate_plans_pdf", side_effect=ValueError("bad column")):
            resp = _upload(client, "/generate/plans/pdf", b"csv", "plans.csv")
        assert resp.status_code == 400
        assert b"bad column" in resp.data

    def test_unexpected_error_returns_500(self, client):
        with patch("app.generate_plans_pdf", side_effect=RuntimeError("boom")):
            resp = _upload(client, "/generate/plans/pdf", b"csv", "plans.csv")
        assert resp.status_code == 500


# ---------------------------------------------------------------------------
# /generate/pipeline/pdf
# ---------------------------------------------------------------------------


class TestPipelinePdfRoute:
    def test_no_file_returns_400(self, client):
        assert client.post("/generate/pipeline/pdf").status_code == 400

    def test_non_csv_returns_400(self, client):
        assert _upload(client, "/generate/pipeline/pdf", b"data", "file.txt").status_code == 400

    def test_valid_csv_returns_pdf(self, client):
        with patch("app.generate_pipeline_pdf", return_value=("HS", b"fake pdf")):
            resp = _upload(client, "/generate/pipeline/pdf", b"csv", "Pipeline 20260508 HS.csv")
        assert resp.status_code == 200
        assert resp.content_type == "application/pdf"
        assert resp.data == b"fake pdf"

    def test_download_name_includes_client_and_date(self, client):
        with (
            patch("app.generate_pipeline_pdf", return_value=("HS", b"bytes")),
            patch("app.date") as mock_date,
        ):
            mock_date.today.return_value = date(2026, 5, 8)
            resp = _upload(client, "/generate/pipeline/pdf", b"csv", "Pipeline 20260508 HS.csv")
        assert "Pipeline_HS_2026-05-08.pdf" in resp.headers.get("Content-Disposition", "")

    def test_value_error_returns_400(self, client):
        with patch("app.generate_pipeline_pdf", side_effect=ValueError("bad filename")):
            resp = _upload(client, "/generate/pipeline/pdf", b"csv", "Pipeline 20260508 HS.csv")
        assert resp.status_code == 400
        assert b"bad filename" in resp.data

    def test_unexpected_error_returns_500(self, client):
        with patch("app.generate_pipeline_pdf", side_effect=RuntimeError("boom")):
            resp = _upload(client, "/generate/pipeline/pdf", b"csv", "Pipeline 20260508 HS.csv")
        assert resp.status_code == 500
