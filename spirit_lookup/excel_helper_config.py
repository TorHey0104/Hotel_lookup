"""Configuration helpers for the Excel helper GUI."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Sequence


_EMAIL_PATTERN = re.compile(r"email", re.IGNORECASE)
_CONTACT_PATTERN = re.compile(
    r"(?P<prefix>contact|kontakt)[\s_.-]*(?P<index>\d+)[\s_.-]*(?P<field>email)",
    re.IGNORECASE,
)


def normalize_header(header: str) -> str:
    """Return a normalized representation of a header."""

    return re.sub(r"[^0-9a-z]", "", header.lower())


def detect_email_headers(headers: Sequence[str]) -> List[str]:
    """Detect headers that likely contain email information."""

    detected: List[str] = []
    for header in headers:
        if not header:
            continue
        normalized = normalize_header(header)
        if _EMAIL_PATTERN.search(header) or _CONTACT_PATTERN.match(header):
            detected.append(header)
            continue
        # Fallback: look for known aliases
        if "mail" in normalized:
            detected.append(header)
    return detected


@dataclass
class ExcelHelperConfigEntry:
    """Represents the stored configuration for a specific Excel file."""

    selected_columns: List[str]
    email_columns: List[str]

    def to_dict(self) -> Dict[str, List[str]]:
        return {
            "selectedColumns": list(self.selected_columns),
            "emailColumns": list(self.email_columns),
        }


class ExcelHelperConfigStore:
    """Persist and restore Excel helper configuration selections."""

    def __init__(self, config_path: Path):
        self.config_path = config_path
        self._data: Dict[str, Dict[str, List[str]]] = {"files": {}}
        self._loaded = False

    def _ensure_loaded(self) -> None:
        if self._loaded:
            return
        if self.config_path.exists():
            try:
                raw = json.loads(self.config_path.read_text(encoding="utf-8"))
            except json.JSONDecodeError:
                raw = {}
            if isinstance(raw, dict) and isinstance(raw.get("files"), dict):
                files = raw["files"]
                valid_entries: Dict[str, Dict[str, List[str]]] = {}
                for key, value in files.items():
                    if not isinstance(value, dict):
                        continue
                    selected = value.get("selectedColumns")
                    emails = value.get("emailColumns")
                    if isinstance(selected, list) and isinstance(emails, list):
                        valid_entries[key] = {
                            "selectedColumns": [str(item) for item in selected],
                            "emailColumns": [str(item) for item in emails],
                        }
                self._data["files"] = valid_entries
        self._loaded = True

    def load(self) -> Dict[str, Dict[str, List[str]]]:
        """Return the raw configuration dictionary."""

        self._ensure_loaded()
        return self._data["files"].copy()

    def get_entry(self, excel_path: Path) -> ExcelHelperConfigEntry | None:
        """Retrieve the stored entry for a given Excel path."""

        self._ensure_loaded()
        key = str(excel_path.resolve())
        stored = self._data["files"].get(key)
        if not stored:
            return None
        return ExcelHelperConfigEntry(
            selected_columns=list(stored.get("selectedColumns", [])),
            email_columns=list(stored.get("emailColumns", [])),
        )

    def save_entry(
        self,
        excel_path: Path,
        selected_columns: Iterable[str],
        email_columns: Iterable[str],
    ) -> ExcelHelperConfigEntry:
        """Persist the configuration for the provided Excel file."""

        self._ensure_loaded()
        key = str(excel_path.resolve())
        entry = ExcelHelperConfigEntry(
            selected_columns=list(selected_columns),
            email_columns=list(email_columns),
        )
        if not self.config_path.parent.exists():
            self.config_path.parent.mkdir(parents=True, exist_ok=True)
        self._data.setdefault("files", {})[key] = entry.to_dict()
        self.config_path.write_text(
            json.dumps(self._data, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        return entry

    def to_pretty_json(self, excel_path: Path) -> str:
        """Return a pretty JSON representation of the entry for display."""

        entry = self.get_entry(excel_path)
        if not entry:
            return "{}"
        payload = {
            "excelPath": str(excel_path),
            **entry.to_dict(),
        }
        return json.dumps(payload, indent=2, ensure_ascii=False)
