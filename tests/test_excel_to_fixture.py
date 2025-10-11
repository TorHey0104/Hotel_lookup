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

    records, warnings = convert_excel_to_fixture(excel_path)

    assert warnings == []
    assert len(records) == 2

    first = records[0]
    assert first["spiritCode"] == "ZRH001"
    assert first["hotelName"] == "Hyatt Regency Zurich"
    assert first["region"] == "EAME"
    assert first["status"] == "Operating"
    assert first["location"] == {
        "city": "Zürich",
        "country": "Schweiz",
        "address": "Flughafenstrasse 1",
    }
    assert first["contacts"] == [
        {
            "role": "Project Director",
            "name": "Max Muster",
            "email": "max@hyatt.com",
        },
        {
            "role": "Design Lead",
            "name": "Anna Beispiel",
            "phone": "+41 44 123 45 67",
        },
    ]
    assert first["meta"] == {
        "launchYear": 2024,
        "notes": "Go-Live nach Umbau",
    }

    second = records[1]
    # numerische Codes werden als Strings formatiert
    assert second["spiritCode"] == "1001"
    assert second["hotelName"] == "Sample Hotel"
    assert second["region"] == "AMER"
    assert "status" not in second
    assert second.get("location", {}).get("address") is None
    assert second.get("contacts", []) == [
        {
            "name": "Chris Cooper",
        }
    ]
    assert second["meta"] == {"launchYear": True}


def test_cli_roundtrip(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    excel_path = create_workbook(tmp_path)
    output_path = tmp_path / "output.json"

    exit_code = main([str(excel_path), str(output_path)])

    assert exit_code == 0
    captured = capsys.readouterr()
    assert "2 Datensätze" in captured.out

    data = json.loads(output_path.read_text(encoding="utf-8"))
    assert isinstance(data, list)
    assert len(data) == 2


def test_fail_on_warning(tmp_path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Spirit Code", "Hotel Name", "Unknown Column"])
    sheet.append(["X001", "Hotel", "unexpected"])
    excel_path = tmp_path / "warning.xlsx"
    workbook.save(excel_path)

    records, warnings = convert_excel_to_fixture(excel_path)
    assert records[0]["spiritCode"] == "X001"
    assert warnings and "Unknown Column" in warnings[0]


def test_write_fixture(tmp_path: Path) -> None:
    target = tmp_path / "fixture.json"
    records = [
        {
            "spiritCode": "ABC123",
            "hotelName": "Test",
        }
    ]

    write_fixture(records, target, indent=4)

    content = json.loads(target.read_text(encoding="utf-8"))
    assert content == records
