"""Tests for automatic Excel selection in the setup tab."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

tkinter = pytest.importorskip("tkinter")
pytest.importorskip("openpyxl")

from tkinter import messagebox

from openpyxl import Workbook  # type: ignore

from spirit_lookup.config import AppConfig
from spirit_lookup.controller import SpiritLookupController
from spirit_lookup.excel_helper_config import ExcelHelperConfigStore
from spirit_lookup.models import SpiritRecord
from spirit_lookup.providers import BaseDataProvider, RecordNotFoundError
from spirit_lookup.ui import SpiritLookupApp


def _create_root_or_skip() -> tkinter.Tk:
    try:
        root = tkinter.Tk()
    except tkinter.TclError:
        pytest.skip("Tkinter benötigt eine Display-Umgebung")
    root.withdraw()
    return root


class DummyProvider(BaseDataProvider):
    """Simple provider returning no data for the UI tests."""

    def __init__(self) -> None:
        self.reload_calls = 0

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

    def reload(self) -> None:
        self.reload_calls += 1


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

    root = _create_root_or_skip()
    try:
        app = SpiritLookupApp(root, controller, app_config)
        assert app.setup_excel_path == excel_path
        assert app.setup_sheet_var.get() == "Hotels"
        assert app.setup_convert_button["state"] == tkinter.NORMAL
        assert app.setup_generate_display_button["state"] == tkinter.NORMAL
    finally:
        root.destroy()


def test_generate_display_config(tmp_path: Path, monkeypatch) -> None:
    data_dir = tmp_path / "data"
    data_dir.mkdir()
    config_path = data_dir / "excel_helper_config.json"

    excel_path = tmp_path / "fixture.xlsx"
    _create_excel_file(excel_path)

    store = ExcelHelperConfigStore(config_path)
    store.save_entry(
        excel_path,
        ["Spirit Code", "Hotel Name", "Kontakt 1 Email"],
        ["Kontakt 1 Email"],
    )

    provider = DummyProvider()
    controller = SpiritLookupController(provider)
    app_config = AppConfig(fixture_path=data_dir / "spirit_fixture.json")

    root = _create_root_or_skip()
    try:
        app = SpiritLookupApp(root, controller, app_config)

        monkeypatch.setattr(messagebox, "showinfo", lambda *args, **kwargs: None)
        monkeypatch.setattr(messagebox, "showerror", lambda *args, **kwargs: None)

        app._setup_generate_display_config()

        assert app.display_config_path.exists()
        payload = json.loads(app.display_config_path.read_text(encoding="utf-8"))
        assert payload["fields"][0]["label"] == "Spirit Code"
        assert payload["fields"][2]["isEmail"] is True
        assert len(app.display_definitions) == 3
    finally:
        root.destroy()


def test_auto_apply_fixture_updates_files(tmp_path: Path, monkeypatch) -> None:
    data_dir = tmp_path / "data"
    data_dir.mkdir()
    app_config = AppConfig(fixture_path=data_dir / "spirit_fixture.json")

    excel_path = tmp_path / "fixture.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Spirit Code", "Hotel Name", "Contact1 Email", "Notes"])
    sheet.append(["ABC123", "Sample Hotel", "info@example.com", "Test"])
    workbook.save(excel_path)

    provider = DummyProvider()
    controller = SpiritLookupController(provider)

    root = _create_root_or_skip()
    try:
        app = SpiritLookupApp(root, controller, app_config)

        app.setup_excel_path = excel_path
        app.setup_sheet_var.set(workbook.active.title)

        monkeypatch.setattr(messagebox, "showinfo", lambda *args, **kwargs: None)
        monkeypatch.setattr(messagebox, "showerror", lambda *args, **kwargs: None)
        monkeypatch.setattr(messagebox, "showwarning", lambda *args, **kwargs: None)

        app._setup_convert_excel()
        assert app.setup_apply_button["state"] == tkinter.NORMAL

        app._setup_apply_fixture()

        assert provider.reload_calls == 1
        assert app_config.fixture_path.exists()
        records = json.loads(app_config.fixture_path.read_text(encoding="utf-8"))
        assert records[0]["spiritCode"] == "ABC123"
        display_payload = json.loads(app.display_config_path.read_text(encoding="utf-8"))
        labels = [item["label"] for item in display_payload["fields"]]
        assert "Spirit Code" in labels
        assert any(item.get("isEmail") for item in display_payload["fields"])
    finally:
        root.destroy()
