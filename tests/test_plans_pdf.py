from datetime import date
from unittest.mock import MagicMock, patch

import pytest

from reports.plans_pdf import generate

FIXED_TODAY = date(2026, 1, 15)

_HEADER = (
    "Year,Client,Funder name,Project,Request Purpose,Status,"
    "Amount requested,Amount awarded,Expected notification date,"
    "Notification date,Next task deadline"
)


def _csv(*rows):
    return ("\n".join([_HEADER] + list(rows))).encode()


def _row(
    year=2026,
    client="KCRep",
    funder="Test Funder",
    project="Test Fund",
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
    with patch("reports.plans.date", mock), patch("reports.plans_pdf.date", mock):
        yield


class TestGeneratePlansPdf:
    def test_returns_pdf_bytes(self):
        _, pdf_bytes = generate(_csv(_row()))[0]
        assert pdf_bytes[:4] == b"%PDF"

    def test_returns_one_result_per_client(self):
        results = generate(_csv(_row(client="KCRep"), _row(client="GOH")))
        assert {r[0] for r in results} == {"KCRep", "GOH"}

    def test_client_name_in_result(self):
        client, _ = generate(_csv(_row(client="WEN")))[0]
        assert client == "WEN"

    def test_future_rows_included(self):
        results = generate(_csv(_row(year=2026), _row(year=2027)))
        assert len(results) == 1
        _, pdf_bytes = results[0]
        assert pdf_bytes[:4] == b"%PDF"

    def test_raises_for_missing_column(self):
        with pytest.raises(ValueError, match="missing expected columns"):
            generate(b"Year,Client\n2026,KCRep\n")

    def test_raises_for_header_only_csv(self):
        with pytest.raises(ValueError, match="no data rows"):
            generate((_HEADER + "\n").encode())

    def test_raises_for_all_past_years(self):
        with pytest.raises(ValueError, match="No rows match"):
            generate(_csv(_row(year=2024)))

    def test_nat_notification_dates_do_not_raise(self):
        # Mixed real dates and blanks produce pd.NaT, which must not crash strftime
        csv = _csv(
            _row(notif_expected="2026-03-01", notif_received="2026-04-01"),
            _row(notif_expected="", notif_received=""),
        )
        _, pdf_bytes = generate(csv)[0]
        assert pdf_bytes[:4] == b"%PDF"
