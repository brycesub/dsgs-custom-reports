from datetime import date
from unittest.mock import MagicMock, patch

import pytest

from reports.pipeline_pdf import generate

FIXED_TODAY = date(2026, 5, 8)
_VALID_FILENAME = "Pipeline 20260508 HS.csv"
_HEADER = "Year,Project,Funder name,Task type,Task Title,Due Date"


def _csv(*rows):
    return ("\n".join([_HEADER] + list(rows))).encode()


def _row(
    year=2026,
    project="Test Project",
    funder="Test Funder",
    task_type="Report",
    task_title="My Task",
    due_date="2026-05-01",
):
    return f"{year},{project},{funder},{task_type},{task_title},{due_date}"


@pytest.fixture(autouse=True)
def freeze_today():
    mock = MagicMock()
    mock.today.return_value = FIXED_TODAY
    with patch("reports.pipeline.date", mock), patch("reports.pipeline_pdf.date", mock):
        yield


class TestGeneratePipelinePdf:
    def test_returns_client_and_pdf_bytes(self):
        client, pdf_bytes = generate(_csv(_row()), _VALID_FILENAME)
        assert client == "HS"
        assert pdf_bytes[:4] == b"%PDF"

    def test_multi_word_client(self):
        filename = "Pipeline 20260508 Hartford Stage.csv"
        client, _ = generate(_csv(_row()), filename)
        assert client == "Hartford Stage"

    def test_raises_for_bad_filename(self):
        with pytest.raises(ValueError, match="Filename must match"):
            generate(_csv(_row()), "badfile.csv")

    def test_raises_for_missing_column(self):
        with pytest.raises(ValueError, match="missing expected columns"):
            generate(b"Year,Project\n2026,P\n", _VALID_FILENAME)

    def test_raises_for_header_only_csv(self):
        with pytest.raises(ValueError, match="no data rows"):
            generate((_HEADER + "\n").encode(), _VALID_FILENAME)

    def test_nat_due_date_does_not_raise(self):
        # A column with mixed real dates and blanks produces pd.NaT, which must not crash strftime
        csv = _csv(_row(due_date="2026-05-01"), _row(due_date=""))
        client, pdf_bytes = generate(csv, _VALID_FILENAME)
        assert pdf_bytes[:4] == b"%PDF"
