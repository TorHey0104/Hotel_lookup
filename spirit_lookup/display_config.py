"""Helpers for managing dynamic display configuration."""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List


@dataclass(frozen=True)
class DisplayFieldDefinition:
    """Describe a single field that should be rendered in the UI."""

    label: str
    is_email: bool = False

    @staticmethod
    def from_dict(payload: dict) -> "DisplayFieldDefinition":
        label = payload.get("label")
        if not isinstance(label, str) or not label.strip():
            raise ValueError("Display field requires a non-empty label")
        is_email = bool(payload.get("isEmail"))
        return DisplayFieldDefinition(label=label.strip(), is_email=is_email)

    def to_dict(self) -> dict:
        return {"label": self.label, "isEmail": self.is_email}


class DisplayConfig:
    """Persist and restore the fields that should be shown in the UI."""

    def __init__(self, path: Path):
        self.path = path
        self.fields: List[DisplayFieldDefinition] = []

    def load(self) -> None:
        if not self.path.exists():
            self.fields = []
            return
        try:
            payload = json.loads(self.path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            self.fields = []
            return
        fields_data = payload.get("fields") if isinstance(payload, dict) else None
        if not isinstance(fields_data, list):
            self.fields = []
            return
        parsed: List[DisplayFieldDefinition] = []
        for raw in fields_data:
            if isinstance(raw, dict):
                try:
                    parsed.append(DisplayFieldDefinition.from_dict(raw))
                except ValueError:
                    continue
        self.fields = parsed

    def save(self, fields: Iterable[DisplayFieldDefinition]) -> None:
        self.fields = list(fields)
        if not self.path.parent.exists():
            self.path.parent.mkdir(parents=True, exist_ok=True)
        payload = {"fields": [field.to_dict() for field in self.fields]}
        self.path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")

