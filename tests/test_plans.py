import io
from datetime import date
from unittest.mock import MagicMock, patch

import openpyxl
import pytest

from reports.plans import generate

# All tests freeze today to 2026-01-15 so year-filter logic is deterministic.
FIXED_TODAY = date(2026, 1, 15)
CURRENT_YEAR = 2026
FUTURE_YEAR = 2027

_HEADER = (
    "Year,Client,Funder name,Project,Request Purpose,Status,"
    "Amount requested,Amount awarded,Expected notification date,"
    "Notification date,Next task deadline"
)


def _csv(*rows):
    return ("\n".join([_HEADER] + list(rows))).encode()


def _row(
    year=CURRENT_YEAR,
    client="TestClient",
    funder="Test Funder",
    project="Test Project",
    purpose="General",
    status="Planned",
    request="10000",
    award="",
    notif_expected="",
    notif_received="",
    next_deadline="",
):
    return (
        f"{year},{client},{funder},{project},{purpose},{status},"
        f"{request},{award},{notif_expected},{notif_received},{next_deadline}"
    )


@pytest.fixture(autouse=True)
def freeze_today():
    mock = MagicMock()
    mock.today.return_value = FIXED_TODAY
    with patch("reports.plans.date", mock):
        yield


# ---------------------------------------------------------------------------
# Happy path
# ---------------------------------------------------------------------------


class TestGenerateHappyPath:
    def test_returns_one_result_per_client(self):
        results = generate(_csv(_row(client="A"), _row(client="B")))
        assert {r[0] for r in results} == {"A", "B"}

    def test_result_is_non_empty_bytes(self):
        _, xlsx_bytes = generate(_csv(_row()))[0]
        assert isinstance(xlsx_bytes, bytes) and len(xlsx_bytes) > 0

    def test_client_name_in_result(self):
        client, _ = generate(_csv(_row(client="Acme Corp")))[0]
        assert client == "Acme Corp"

    def test_current_year_sheet_always_created(self):
        _, xlsx_bytes = generate(_csv(_row(year=CURRENT_YEAR)))[0]
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        assert str(CURRENT_YEAR) in wb.sheetnames

    def test_future_sheet_created_when_future_rows_exist(self):
        _, xlsx_bytes = generate(_csv(_row(year=CURRENT_YEAR), _row(year=FUTURE_YEAR)))[0]
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        assert f"{FUTURE_YEAR}+" in wb.sheetnames

    def test_no_future_sheet_when_no_future_rows(self):
        _, xlsx_bytes = generate(_csv(_row(year=CURRENT_YEAR)))[0]
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        assert len(wb.sheetnames) == 1

    def test_current_year_sheet_is_active(self):
        _, xlsx_bytes = generate(_csv(_row(year=CURRENT_YEAR), _row(year=FUTURE_YEAR)))[0]
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        assert wb.active.title == str(CURRENT_YEAR)

    def test_headers_written_to_row_1(self):
        _, xlsx_bytes = generate(_csv(_row()))[0]
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        ws = wb[str(CURRENT_YEAR)]
        # OUTPUT_COLUMNS: Year, Funder, Fund, Purpose, Status, Request, Award, ...
        assert ws.cell(1, 1).value == "Year"
        assert ws.cell(1, 2).value == "Funder"
        assert ws.cell(1, 3).value == "Fund"
        assert ws.cell(1, 5).value == "Status"

    def test_data_written_to_row_2(self):
        _, xlsx_bytes = generate(_csv(_row(funder="Big Foundation", project="My Fund")))[0]
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        ws = wb[str(CURRENT_YEAR)]
        assert ws.cell(2, 2).value == "Big Foundation"  # Funder
        assert ws.cell(2, 3).value == "My Fund"  # Fund


# ---------------------------------------------------------------------------
# Year filtering
# ---------------------------------------------------------------------------


class TestYearFiltering:
    def test_past_year_rows_excluded(self):
        # Two rows: one past (2025), one current (2026)
        _, xlsx_bytes = generate(_csv(_row(year=2025), _row(year=CURRENT_YEAR)))[0]
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        ws = wb[str(CURRENT_YEAR)]
        assert ws.max_row == 2  # header + 1 row (2025 filtered out)

    def test_raises_when_all_rows_in_past(self):
        with pytest.raises(ValueError, match="No rows match"):
            generate(_csv(_row(year=2024), _row(year=2025)))

    def test_future_rows_go_to_future_sheet(self):
        _, xlsx_bytes = generate(_csv(_row(year=FUTURE_YEAR)))[0]
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        assert f"{FUTURE_YEAR}+" in wb.sheetnames
        assert wb[f"{FUTURE_YEAR}+"].max_row == 2  # header + 1 row


# ---------------------------------------------------------------------------
# Sort order
# ---------------------------------------------------------------------------


class TestSortOrder:
    def test_status_rank_applied(self):
        # "Awarded - Active" ranks 0; "Planned" ranks 6 — should appear first
        csv_bytes = _csv(
            _row(status="Planned", project="ZZZ Fund"),
            _row(status="Awarded - Active", project="AAA Fund"),
        )
        _, xlsx_bytes = generate(csv_bytes)[0]
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        ws = wb[str(CURRENT_YEAR)]
        assert ws.cell(2, 5).value == "Awarded - Active"
        assert ws.cell(3, 5).value == "Planned"

    def test_alphabetical_fund_within_same_status(self):
        csv_bytes = _csv(
            _row(status="Planned", project="Zoo Fund"),
            _row(status="Planned", project="Alpha Fund"),
        )
        _, xlsx_bytes = generate(csv_bytes)[0]
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        ws = wb[str(CURRENT_YEAR)]
        assert ws.cell(2, 3).value == "Alpha Fund"  # Fund col
        assert ws.cell(3, 3).value == "Zoo Fund"


# ---------------------------------------------------------------------------
# Amount parsing
# ---------------------------------------------------------------------------


class TestAmountParsing:
    def test_dollar_sign_and_commas_stripped(self):
        _, xlsx_bytes = generate(_csv(_row(request="$10500", award="$3000")))[0]
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        ws = wb[str(CURRENT_YEAR)]
        assert ws.cell(2, 6).value == pytest.approx(10500.0)  # Request
        assert ws.cell(2, 7).value == pytest.approx(3000.0)  # Award

    def test_empty_amount_written_as_none(self):
        _, xlsx_bytes = generate(_csv(_row(request="", award="")))[0]
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        ws = wb[str(CURRENT_YEAR)]
        assert ws.cell(2, 6).value is None
        assert ws.cell(2, 7).value is None


# ---------------------------------------------------------------------------
# Validation errors
# ---------------------------------------------------------------------------


class TestValidationErrors:
    def test_raises_for_missing_column(self):
        csv_bytes = b"Year,Funder name,Project\n2026,Some Funder,Some Project"
        with pytest.raises(ValueError, match="missing expected columns"):
            generate(csv_bytes)

    def test_raises_for_header_only_csv(self):
        with pytest.raises(ValueError, match="no data rows"):
            generate(_HEADER.encode() + b"\n")

    def test_raises_for_non_numeric_year(self):
        with pytest.raises(ValueError, match="non-numeric Year"):
            generate(_csv(_row(year="FY 2026")))
