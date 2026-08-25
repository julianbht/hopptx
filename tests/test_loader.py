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

    reports = loader.load_reports(config)

    assert [company for company, _ in reports] == ["Business Skills Inc"]


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
