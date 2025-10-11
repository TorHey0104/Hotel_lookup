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
