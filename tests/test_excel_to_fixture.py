"""Tests für den Excel-zu-Fixture-Helfer."""

from __future__ import annotations

from pathlib import Path

import json
import pytest

openpyxl = pytest.importorskip("openpyxl")
from openpyxl import Workbook  # type: ignore

from tools.excel_to_fixture import convert_excel_to_fixture, main, write_fixture


def create_workbook(tmp_path: Path) -> Path:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Hotels"
    sheet.append(
        [
            "Spirit Code",
            "Hotel Name",
            "Region",
            "Status",
            "City",
            "Country",
            "Address",
            "Contact1 Role",
            "Contact1 Name",
            "Contact1 Email",
            "Contact2 Role",
            "Contact2 Name",
            "Contact2 Phone",
            "Meta.launchYear",
            "Meta Notes",
        ]
    )
    sheet.append(
        [
            "ZRH001",
            "Hyatt Regency Zurich",
            "EAME",
            "Operating",
            "Zürich",
            "Schweiz",
            "Flughafenstrasse 1",
            "Project Director",
            "Max Muster",
            "max@hyatt.com",
            "Design Lead",
            "Anna Beispiel",
            "+41 44 123 45 67",
            2024,
            "Go-Live nach Umbau",
        ]
    )
    sheet.append(
        [
            1001,
            "Sample Hotel",
            "AMER",
            None,
            "Chicago",
            "USA",
            None,
            None,
            "Chris Cooper",
            None,
            None,
            None,
            None,
            "true",
            None,
        ]
    )

    excel_path = tmp_path / "fixture.xlsx"
    workbook.save(excel_path)
    return excel_path


def test_convert_excel_to_fixture(tmp_path: Path) -> None:
    excel_path = create_workbook(tmp_path)

    fixture, warnings = convert_excel_to_fixture(excel_path)

    assert warnings == []
    assert isinstance(fixture, dict)

    config = fixture.get("config", {})
    assert config.get("fieldMapping", {}).get("spiritCode") == "Spirit Code"
    assert config.get("fieldMapping", {}).get("displayField") == "Hotel Name"

    records = fixture.get("records", [])
    assert len(records) == 2

    first = records[0]
    assert first["spiritCode"] == "ZRH001"
    assert first["displayValue"] == "Hyatt Regency Zurich"
    assert first.get("emails", {}) == {"Contact1 Email": "max@hyatt.com"}

    fields = first["fields"]
    assert fields["Spirit Code"] == "ZRH001"
    assert fields["Hotel Name"] == "Hyatt Regency Zurich"
    assert fields["Region"] == "EAME"
    assert fields["Status"] == "Operating"
    assert fields["City"] == "Zürich"
    assert fields["Country"] == "Schweiz"
    assert fields["Address"] == "Flughafenstrasse 1"
    assert fields["Contact1 Name"] == "Max Muster"
    assert fields["Meta.launchYear"] == "2024"
    assert fields["Meta Notes"] == "Go-Live nach Umbau"

    second = records[1]
    assert second["spiritCode"] == "1001"
    assert second["displayValue"] == "Sample Hotel"
    second_fields = second["fields"]
    assert second_fields["Spirit Code"] == "1001"
    assert second_fields["Hotel Name"] == "Sample Hotel"
    assert second_fields["Region"] == "AMER"
    assert second_fields["City"] == "Chicago"
    assert second_fields["Contact1 Name"] == "Chris Cooper"
    assert second_fields["Meta.launchYear"] == "true"


def test_cli_roundtrip(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    excel_path = create_workbook(tmp_path)
    output_path = tmp_path / "output.json"

    exit_code = main([str(excel_path), str(output_path)])

    assert exit_code == 0
    captured = capsys.readouterr()
    assert "2 Datensätze" in captured.out

    data = json.loads(output_path.read_text(encoding="utf-8"))
    assert isinstance(data, dict)
    assert len(data.get("records", [])) == 2


def test_fail_on_warning(tmp_path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Spirit Code", "Hotel Name", "Unknown Column"])
    sheet.append(["X001", "Hotel", "unexpected"])
    excel_path = tmp_path / "warning.xlsx"
    workbook.save(excel_path)

    fixture, warnings = convert_excel_to_fixture(excel_path)
    assert fixture["records"][0]["spiritCode"] == "X001"
    assert warnings == []


def test_write_fixture(tmp_path: Path) -> None:
    target = tmp_path / "fixture.json"
    fixture = {
        "config": {
            "selectedColumns": ["Spirit Code", "Hotel"],
            "emailColumns": [],
            "fieldMapping": {"spiritCode": "Spirit Code", "displayField": "Hotel"},
        },
        "records": [
            {
                "spiritCode": "ABC123",
                "displayValue": "Test",
                "fields": {"Spirit Code": "ABC123", "Hotel": "Test"},
            }
        ],
    }

    write_fixture(fixture, target, indent=4)

    content = json.loads(target.read_text(encoding="utf-8"))
    assert content == fixture
