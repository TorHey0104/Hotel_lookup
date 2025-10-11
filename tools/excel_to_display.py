"""Generate a display configuration JSON based on the Excel helper selection."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from spirit_lookup.display_config import DisplayFieldDefinition
from spirit_lookup.excel_helper_config import ExcelHelperConfigStore


def _resolve_entry(store: ExcelHelperConfigStore, excel_path: Path | None) -> tuple[Path, list[str], list[str]]:
    """Return the selected and email columns for the desired Excel file."""

    if excel_path is None:
        excel_path = store.get_last_used_path()
        if excel_path is None:
            raise ValueError(
                "Keine zuletzt genutzte Excel-Datei gefunden. Bitte den Excel Helper verwenden, um eine Auswahl zu speichern."
            )

    excel_path = excel_path.expanduser().resolve()
    entry = store.get_entry(excel_path)
    if entry is None:
        raise ValueError(
            "Für die angegebene Excel-Datei existiert keine gespeicherte Konfiguration."
        )
    if not entry.selected_columns:
        raise ValueError("Die gespeicherte Konfiguration enthält keine ausgewählten Spalten.")
    return excel_path, entry.selected_columns, entry.email_columns


def generate_display_json(
    *,
    config_path: Path,
    output_path: Path,
    excel_path: Path | None = None,
) -> Path:
    """Create the display.json file and return the written path."""

    store = ExcelHelperConfigStore(config_path)
    _excel_path, selected_columns, email_columns = _resolve_entry(store, excel_path)

    definitions = [
        DisplayFieldDefinition(label=column, is_email=column in email_columns)
        for column in selected_columns
    ]

    if not output_path.parent.exists():
        output_path.parent.mkdir(parents=True, exist_ok=True)

    payload = {"fields": [definition.to_dict() for definition in definitions]}
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Erzeuge eine display.json aus der Excel-Helper-Konfiguration.")
    parser.add_argument(
        "--config",
        type=Path,
        default=Path("data/excel_helper_config.json"),
        help="Pfad zur excel_helper_config.json",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Ausgabedatei für display.json (Standard: data/display.json neben der Konfiguration)",
    )
    parser.add_argument(
        "--excel",
        type=Path,
        default=None,
        help="Optionale explizite Excel-Datei, deren Konfiguration genutzt werden soll.",
    )

    args = parser.parse_args()

    config_path = args.config.expanduser()
    if not config_path.exists():
        raise SystemExit("Die Konfigurationsdatei existiert nicht: " + str(config_path))

    if args.output is None:
        output_path = config_path.parent / "display.json"
    else:
        output_path = args.output.expanduser()

    excel_path = args.excel.expanduser() if args.excel is not None else None

    try:
        generate_display_json(config_path=config_path, output_path=output_path, excel_path=excel_path)
    except ValueError as exc:
        raise SystemExit(str(exc)) from exc


if __name__ == "__main__":  # pragma: no cover - CLI entry point
    main()

