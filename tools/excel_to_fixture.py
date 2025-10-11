"""Hilfsprogramm zur Konvertierung einer Excel-Tabelle in die Spirit-Fixture."""

from __future__ import annotations

import argparse
import json
import math
import re
from collections import defaultdict
from pathlib import Path
from typing import Any, Iterable, List, Sequence, Tuple

try:
    from openpyxl import load_workbook
    from openpyxl.worksheet.worksheet import Worksheet
except ModuleNotFoundError as exc:  # pragma: no cover - wird in Tests übersprungen
    raise ModuleNotFoundError(
        "Für den Excel-Import wird 'openpyxl' benötigt. Installiere das Paket z. B. mit `pip install openpyxl`."
    ) from exc


_CONTACT_PATTERN = re.compile(
    r"(?P<prefix>contact|kontakt)(?P<index>\d+)(?P<field>role|name|email|phone)",
    re.IGNORECASE,
)


def normalize_key(value: str) -> str:
    """Normalisiere Spaltenüberschriften für Vergleiche."""

    return re.sub(r"[^0-9a-z]", "", value.lower())


def is_empty(value: Any) -> bool:
    """Prüfe, ob eine Zelle leer oder nur Whitespace enthält."""

    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    if isinstance(value, float):
        return not math.isfinite(value)
    return False


def format_primary_cell(value: Any) -> str | None:
    """Bereite Zelleninhalte für Hauptfelder (Strings) auf."""

    if is_empty(value):
        return None
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if not math.isfinite(value):
            return None
        if value.is_integer():
            return str(int(value))
        return format(value, "g")
    return str(value).strip() or None


def format_meta_value(value: Any) -> Any:
    """Bereite Meta-Felder so auf, dass der Datentyp bestmöglich erhalten bleibt."""

    if is_empty(value):
        return None
    if isinstance(value, str):
        stripped = value.strip()
        lowered = stripped.lower()
        if lowered in {"true", "yes", "ja"}:
            return True
        if lowered in {"false", "no", "nein"}:
            return False
        return stripped
    if isinstance(value, bool):
        return value
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        if not math.isfinite(value):
            return None
        if value.is_integer():
            return int(value)
        return value
    return value


def derive_meta_key(header: str) -> str:
    """Erzeuge einen Feldnamen für Meta-Daten aus der Spaltenüberschrift."""

    stripped = header.strip()
    stripped = re.sub(r"(?i)^meta[\s_.-]*", "", stripped)
    if not stripped:
        stripped = "metaField"
    parts = [part for part in re.split(r"[^0-9A-Za-z]+", stripped) if part]
    if not parts:
        return "metaField"
    first, *rest = parts
    camel = first[:1].lower() + first[1:]
    for piece in rest:
        camel += piece[:1].upper() + piece[1:]
    return camel


def load_rows(sheet: Worksheet) -> Tuple[List[str], List[List[Any]]]:
    """Lese die Header und Datenzeilen aus einem Worksheet."""

    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return [], []

    headers = [str(cell).strip() if cell is not None else "" for cell in rows[0]]
    data_rows = [list(row) for row in rows[1:]]
    return headers, data_rows


def convert_row(headers: Sequence[str], values: Sequence[Any]) -> Tuple[dict[str, Any], List[str]]:
    """Konvertiere eine einzelne Excel-Zeile in die Fixture-Struktur."""

    items: List[Tuple[str, Any]] = []
    for idx, header in enumerate(headers):
        if not header:
            continue
        cell_value = values[idx] if idx < len(values) else None
        items.append((header, cell_value))

    used_headers: set[str] = set()

    def pop_alias(label: str, aliases: Iterable[str], *, required: bool = False) -> str | None:
        normalized_aliases = {alias for alias in aliases}
        for header, value in items:
            if header in used_headers:
                continue
            if normalize_key(header) in normalized_aliases:
                used_headers.add(header)
                result = format_primary_cell(value)
                if required and not result:
                    raise ValueError(f"Erforderliche Spalte '{label}' ist leer.")
                return result
        if required:
            raise ValueError(f"Erforderliche Spalte '{label}' wurde nicht gefunden.")
        return None

    spirit_code = pop_alias("spiritCode", {"spiritcode"}, required=True)
    hotel_name = pop_alias("hotelName", {"hotelname"}, required=True)

    region = pop_alias("region", {"region"})
    status = pop_alias("status", {"status"})

    city = pop_alias("city", {"city", "locationcity"})
    country = pop_alias("country", {"country", "locationcountry"})
    address = pop_alias("address", {"address", "street", "locationaddress"})

    contact_groups: dict[int, dict[str, str]] = defaultdict(dict)
    for header, value in items:
        if header in used_headers:
            continue
        match = _CONTACT_PATTERN.match(normalize_key(header))
        if not match:
            continue
        used_headers.add(header)
        formatted = format_primary_cell(value)
        if not formatted:
            continue
        index = int(match.group("index"))
        field = match.group("field").lower()
        contact_groups[index][field] = formatted

    contacts = []
    for index in sorted(contact_groups):
        contact = contact_groups[index]
        if any(contact.get(key) for key in ("role", "name", "email", "phone")):
            contacts.append(contact)

    meta: dict[str, Any] = {}
    for header, value in items:
        if header in used_headers:
            continue
        if normalize_key(header).startswith("meta"):
            used_headers.add(header)
            meta_value = format_meta_value(value)
            if meta_value is not None:
                meta_key = derive_meta_key(header)
                meta[meta_key] = meta_value

    record: dict[str, Any] = {
        "spiritCode": spirit_code,
        "hotelName": hotel_name,
    }
    if region:
        record["region"] = region
    if status:
        record["status"] = status

    location: dict[str, str] = {}
    if city:
        location["city"] = city
    if country:
        location["country"] = country
    if address:
        location["address"] = address
    if location:
        record["location"] = location

    if contacts:
        record["contacts"] = contacts
    if meta:
        record["meta"] = meta

    unused_columns = [
        header
        for header, value in items
        if header not in used_headers and not is_empty(value)
    ]
    return record, unused_columns


def convert_excel_to_fixture(
    excel_path: Path, *, sheet_name: str | None = None
) -> Tuple[List[dict[str, Any]], List[str]]:
    """Konvertiere eine Excel-Datei in Fixture-Struktur."""

    workbook = load_workbook(excel_path, data_only=True)
    if sheet_name:
        if sheet_name not in workbook.sheetnames:
            raise ValueError(
                f"Arbeitsblatt '{sheet_name}' existiert nicht. Verfügbare Blätter: {', '.join(workbook.sheetnames)}"
            )
        worksheet = workbook[sheet_name]
    else:
        worksheet = workbook.active

    headers, rows = load_rows(worksheet)
    records: List[dict[str, Any]] = []
    warnings: List[str] = []

    for offset, row in enumerate(rows, start=2):
        if all(is_empty(value) for value in row):
            continue
        try:
            record, unused = convert_row(headers, row)
        except ValueError as exc:
            raise ValueError(f"Zeile {offset}: {exc}") from exc
        records.append(record)
        if unused:
            warnings.append(
                f"Zeile {offset}: Die Spalten {', '.join(sorted(unused))} konnten nicht zugeordnet werden."
            )

    return records, warnings


def write_fixture(records: Sequence[dict[str, Any]], output_path: Path, *, indent: int = 2) -> None:
    """Schreibe die Datensätze in eine JSON-Datei."""

    output_path.write_text(
        json.dumps(records, indent=indent, ensure_ascii=False),
        encoding="utf-8",
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Konvertiert eine Excel-Datei in die Spirit Lookup JSON-Fixture",
    )
    parser.add_argument("excel_path", type=Path, help="Pfad zur Excel-Datei (.xlsx)")
    parser.add_argument(
        "output_path",
        type=Path,
        nargs="?",
        help="Zieldatei für die JSON-Fixture (Standard: gleicher Pfad mit .json)",
    )
    parser.add_argument(
        "--sheet",
        dest="sheet_name",
        help="Name des Arbeitsblatts, das konvertiert werden soll (Standard: erstes Blatt)",
    )
    parser.add_argument(
        "--indent",
        type=int,
        default=2,
        help="Einrückung im JSON-Output (Standard: 2)",
    )
    parser.add_argument(
        "--fail-on-warning",
        action="store_true",
        help="Beende mit Fehler, wenn Spalten nicht zugeordnet werden konnten",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    excel_path: Path = args.excel_path
    if not excel_path.exists():
        parser.error(f"Die Datei '{excel_path}' wurde nicht gefunden.")

    output_path: Path = args.output_path or excel_path.with_suffix(".json")

    try:
        records, warnings = convert_excel_to_fixture(excel_path, sheet_name=args.sheet_name)
    except ValueError as exc:
        parser.error(str(exc))
        return 2

    write_fixture(records, output_path, indent=args.indent)

    message_lines = [
        f"✅ {len(records)} Datensätze nach '{output_path}' geschrieben.",
    ]
    if warnings:
        warning_block = "\n".join(f"⚠️  {warning}" for warning in warnings)
        if args.fail_on_warning:
            parser.error(f"Es wurden Spalten nicht zugeordnet:\n{warning_block}")
        message_lines.append("Warnungen:")
        message_lines.append(warning_block)

    print("\n".join(message_lines))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
