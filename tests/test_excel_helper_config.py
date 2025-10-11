from __future__ import annotations

import json
from pathlib import Path

from spirit_lookup.excel_helper_config import (
    ExcelHelperConfigStore,
    detect_email_headers,
)


def test_detect_email_headers_identifies_contact_columns() -> None:
    headers = [
        "Spirit Code",
        "Contact1 Email",
        "Kontakt2 E-Mail",
        "Hotel Name",
        "SupportMail",
    ]

    detected = detect_email_headers(headers)

    assert set(detected) == {"Contact1 Email", "Kontakt2 E-Mail", "SupportMail"}


def test_config_store_roundtrip(tmp_path: Path) -> None:
    config_path = tmp_path / "excel_helper_config.json"
    store = ExcelHelperConfigStore(config_path)

    excel_file = tmp_path / "sample.xlsx"
    excel_file.write_text("dummy")

    store.save_entry(excel_file, ["Spirit Code", "Hotel Name"], ["Contact1 Email"])

    other_store = ExcelHelperConfigStore(config_path)
    loaded = other_store.get_entry(excel_file)

    assert loaded is not None
    assert loaded.selected_columns == ["Spirit Code", "Hotel Name"]
    assert loaded.email_columns == ["Contact1 Email"]

    pretty = other_store.to_pretty_json(excel_file)
    assert "Spirit Code" in pretty
    assert "Contact1 Email" in pretty


def test_config_store_lists_entries_and_last_used(tmp_path: Path) -> None:
    config_path = tmp_path / "excel_helper_config.json"
    store = ExcelHelperConfigStore(config_path)

    first = (tmp_path / "first.xlsx").resolve()
    second = (tmp_path / "second.xlsx").resolve()
    first.write_text("dummy")
    second.write_text("dummy")

    store.save_entry(first, ["Hotel"], [])
    store.save_entry(second, ["Spirit"], ["Contact"])

    reloaded = ExcelHelperConfigStore(config_path)

    entries = reloaded.list_entries()
    assert entries == sorted([first, second])
    assert reloaded.get_last_used_path() == second


def test_config_store_handles_legacy_payload_without_last_used(tmp_path: Path) -> None:
    config_path = tmp_path / "excel_helper_config.json"
    legacy_content = {
        "files": {
            str((tmp_path / "legacy.xlsx").resolve()): {
                "selectedColumns": ["A"],
                "emailColumns": ["B"],
            }
        }
    }
    config_path.write_text(json.dumps(legacy_content), encoding="utf-8")

    store = ExcelHelperConfigStore(config_path)

    entries = store.list_entries()
    assert len(entries) == 1
    assert store.get_last_used_path() is None


def test_config_store_reload_reflects_external_changes(tmp_path: Path) -> None:
    config_path = tmp_path / "excel_helper_config.json"
    store = ExcelHelperConfigStore(config_path)

    first_excel = (tmp_path / "first.xlsx").resolve()
    second_excel = (tmp_path / "second.xlsx").resolve()
    first_excel.write_text("dummy")
    second_excel.write_text("dummy")

    store.save_entry(first_excel, ["Spirit"], [])
    assert store.get_last_used_path() == first_excel

    other_store = ExcelHelperConfigStore(config_path)
    other_store.save_entry(second_excel, ["Hotel"], ["Contact"])

    # Without reloading the instance still exposes the stale configuration
    assert store.get_last_used_path() == first_excel

    store.reload()

    assert store.get_last_used_path() == second_excel
    reloaded_entry = store.get_entry(second_excel)
    assert reloaded_entry is not None
    assert reloaded_entry.selected_columns == ["Hotel"]
