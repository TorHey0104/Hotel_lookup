"""Tests for automatic Excel selection in the setup tab."""

from __future__ import annotations

from pathlib import Path

import pytest

tkinter = pytest.importorskip("tkinter")
pytest.importorskip("openpyxl")

from openpyxl import Workbook  # type: ignore

from spirit_lookup.config import AppConfig
from spirit_lookup.controller import SpiritLookupController
from spirit_lookup.excel_helper_config import ExcelHelperConfigStore
from spirit_lookup.models import SpiritRecord
from spirit_lookup.providers import BaseDataProvider, RecordNotFoundError
from spirit_lookup.ui import SpiritLookupApp


class DummyProvider(BaseDataProvider):
    """Simple provider returning no data for the UI tests."""

    def list_records(
        self,
        query: str = "",
        *,
        page: int = 0,
        page_size: int = 50,
    ) -> tuple[list[SpiritRecord], bool]:
        return [], False

    def get_record(self, spirit_code: str) -> SpiritRecord:
        raise RecordNotFoundError("not implemented")


def _create_excel_file(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Hotels"
    sheet.append(["Spirit Code", "Hotel Name"])
    sheet.append(["ABC123", "Sample Hotel"])
    workbook.save(path)


def test_setup_tab_restores_last_excel_selection(tmp_path: Path) -> None:
    data_dir = tmp_path / "data"
    data_dir.mkdir()
    config_path = data_dir / "excel_helper_config.json"

    excel_path = tmp_path / "fixture.xlsx"
    _create_excel_file(excel_path)

    store = ExcelHelperConfigStore(config_path)
    store.save_entry(excel_path, ["Spirit Code", "Hotel Name"], ["Hotel Name"])

    provider = DummyProvider()
    controller = SpiritLookupController(provider)
    app_config = AppConfig(fixture_path=data_dir / "spirit_fixture.json")

    root = tkinter.Tk()
    root.withdraw()
    try:
        app = SpiritLookupApp(root, controller, app_config)
        assert app.setup_excel_path == excel_path
        assert app.setup_sheet_var.get() == "Hotels"
        assert app.setup_convert_button["state"] == tkinter.NORMAL
    finally:
        root.destroy()
