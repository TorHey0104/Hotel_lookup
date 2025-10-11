"""Data models for Spirit Lookup."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional


@dataclass(slots=True)
class Contact:
    role: str
    name: str
    email: Optional[str] = None
    phone: Optional[str] = None


@dataclass(slots=True)
class SpiritRecord:
    spirit_code: str
    hotel_name: str
    region: Optional[str] = None
    status: Optional[str] = None
    location_city: Optional[str] = None
    location_country: Optional[str] = None
    address: Optional[str] = None
    contacts: List[Contact] = field(default_factory=list)
    meta: Dict[str, object] = field(default_factory=dict)
    fields: Dict[str, Optional[str]] = field(default_factory=dict)
    field_order: List[str] = field(default_factory=list)
    email_fields: Dict[str, str] = field(default_factory=dict)

    def display_label(self) -> str:
        """Return a label for dropdown entries."""

        location = ", ".join(
            part for part in (self.location_city, self.location_country) if part
        )
        if location:
            return f"{self.spirit_code} • {self.hotel_name} ({location})"
        return f"{self.spirit_code} • {self.hotel_name}"
