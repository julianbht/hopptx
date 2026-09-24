from types import SimpleNamespace

import pandas as pd
import pytest

from hopptx import loader


def test_load_reports_directory_filters_companies_inline(
    monkeypatch: pytest.MonkeyPatch, tmp_path
) -> None:
    (tmp_path / "a.xlsx").touch()
    (tmp_path / "b.xlsx").touch()

    def fake_load_report_file(path):
        if path.name == "a.xlsx":
            return "Technical Skills Corp", pd.DataFrame({"x": [1]})
        return "Business Skills Inc", pd.DataFrame({"x": [2]})

    monkeypatch.setattr(loader, "load_report_file", fake_load_report_file)
    config = SimpleNamespace(
        input=SimpleNamespace(type="directory", path=str(tmp_path)),
        companies_inline=["business skills inc"],
    )

    reports, skipped = loader.load_reports(config)

    assert [company for company, _ in reports] == ["Business Skills Inc"]
    assert skipped == []


def test_load_reports_directory_raises_when_inline_company_missing(
    monkeypatch: pytest.MonkeyPatch, tmp_path
) -> None:
    (tmp_path / "a.xlsx").touch()

    monkeypatch.setattr(
        loader,
        "load_report_file",
        lambda _path: ("Technical Skills Corp", pd.DataFrame({"x": [1]})),
    )
    config = SimpleNamespace(
        input=SimpleNamespace(type="directory", path=str(tmp_path)),
        companies_inline=["Missing Company"],
    )

    with pytest.raises(ValueError, match="Requested companies not found"):
        loader.load_reports(config)


def _write_report(path, rows: list[dict], sheet_name: str = "Acme") -> None:
    pd.DataFrame(rows).to_excel(path, sheet_name=sheet_name, index=False)


def _row(**overrides) -> dict:
    row = {
        "Event Type Name": "AI 101 - Q1",
        "Start Date & Time": "2026-01-02 10:00",
        "Event Created Date & Time": "2026-01-01 09:00",
        "Canceled": False,
        "Invitee Email": "a@example.com",
    }
    row.update(overrides)
    return row


def test_load_report_file_needs_only_used_columns_and_tolerates_header_spaces(tmp_path) -> None:
    path = tmp_path / "report.xlsx"
    pd.DataFrame([_row()]).rename(columns={"Canceled": " Canceled "}).to_excel(
        path, sheet_name="Acme", index=False
    )

    company, df = loader.load_report_file(path)

    assert company == "Acme"
    assert list(df.columns) == list(loader.REQUIRED_COLUMNS)


def test_load_report_file_reports_missing_columns(tmp_path) -> None:
    path = tmp_path / "report.xlsx"
    row = _row()
    del row["Invitee Email"]
    _write_report(path, [row])

    with pytest.raises(ValueError, match="missing required column.*Invitee Email"):
        loader.load_report_file(path)


def test_load_report_file_drops_empty_rows_and_parses_text_canceled(tmp_path) -> None:
    path = tmp_path / "report.xlsx"
    empty = {column: None for column in loader.REQUIRED_COLUMNS}
    _write_report(path, [_row(Canceled="Yes"), empty, _row(Canceled="FALSE")])

    _, df = loader.load_report_file(path)

    assert df["Canceled"].tolist() == [True, False]


@pytest.mark.parametrize(
    ("overrides", "message"),
    [
        ({"Event Type Name": " "}, r"'Event Type Name' column is empty in row\(s\) 2"),
        ({"Start Date & Time": "soon"}, "'Start Date & Time' column has values that are not dates"),
        ({"Canceled": "maybe"}, "'Canceled' column must contain TRUE/FALSE"),
    ],
)
def test_load_report_file_reports_bad_cells_with_excel_row(tmp_path, overrides, message) -> None:
    path = tmp_path / "report.xlsx"
    _write_report(path, [_row(**overrides)])

    with pytest.raises(ValueError, match=message):
        loader.load_report_file(path)


def test_load_reports_skips_bad_files_and_excel_lock_files(tmp_path) -> None:
    _write_report(tmp_path / "good.xlsx", [_row()])
    _write_report(tmp_path / "bad.xlsx", [_row(Canceled="maybe")])
    (tmp_path / "~$good.xlsx").touch()
    config = SimpleNamespace(
        input=SimpleNamespace(type="directory", path=str(tmp_path)),
        companies_inline=None,
    )

    reports, skipped = loader.load_reports(config)

    assert [company for company, _ in reports] == ["Acme"]
    assert [s.name for s in skipped] == ["bad.xlsx"]


@pytest.mark.parametrize("sheet_name", ["Sheet1", "Tabelle1", "sheet 2"])
def test_load_report_file_uses_unknown_company_for_default_sheet_names(
    tmp_path, sheet_name
) -> None:
    path = tmp_path / "report.xlsx"
    _write_report(path, [_row()], sheet_name=sheet_name)

    company, _ = loader.load_report_file(path)

    assert company == loader.UNKNOWN_COMPANY
