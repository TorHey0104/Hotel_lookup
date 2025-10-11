import json
from pathlib import Path

import pytest

from spirit_lookup.excel_helper_config import ExcelHelperConfigStore
from tools.excel_to_display import generate_display_json


def _prepare_store(config_path: Path, excel_path: Path) -> None:
    store = ExcelHelperConfigStore(config_path)
    store.save_entry(
        excel_path,
        ["Spirit Code", "Hotel Name", "Kontakt 1 Email"],
        ["Kontakt 1 Email"],
    )


def test_generate_display_json_creates_expected_payload(tmp_path: Path) -> None:
    data_dir = tmp_path / "data"
    data_dir.mkdir()
    config_path = data_dir / "excel_helper_config.json"
    excel_path = tmp_path / "fixture.xlsx"
    excel_path.write_bytes(b"")

    _prepare_store(config_path, excel_path)

    output_path = data_dir / "display.json"
    generate_display_json(config_path=config_path, output_path=output_path, excel_path=excel_path)

    payload = json.loads(output_path.read_text(encoding="utf-8"))
    assert payload["fields"] == [
        {"label": "Spirit Code", "isEmail": False},
        {"label": "Hotel Name", "isEmail": False},
        {"label": "Kontakt 1 Email", "isEmail": True},
    ]


def test_generate_display_json_uses_last_used_entry(tmp_path: Path) -> None:
    data_dir = tmp_path / "data"
    data_dir.mkdir()
    config_path = data_dir / "excel_helper_config.json"
    excel_path = tmp_path / "fixture.xlsx"
    excel_path.write_bytes(b"")

    _prepare_store(config_path, excel_path)

    output_path = data_dir / "display.json"
    generate_display_json(config_path=config_path, output_path=output_path)

    payload = json.loads(output_path.read_text(encoding="utf-8"))
    labels = [field["label"] for field in payload["fields"]]
    assert labels[0] == "Spirit Code"


def test_generate_display_json_raises_for_missing_entry(tmp_path: Path) -> None:
    data_dir = tmp_path / "data"
    data_dir.mkdir()
    config_path = data_dir / "excel_helper_config.json"
    output_path = data_dir / "display.json"

    with pytest.raises(ValueError):
        generate_display_json(config_path=config_path, output_path=output_path)

