"""Controller logic for Spirit Lookup."""

from __future__ import annotations

from dataclasses import dataclass
from typing import List

from .models import SpiritRecord
from .providers import BaseDataProvider, RecordNotFoundError


@dataclass
class LookupResult:
    records: List[SpiritRecord]
    has_more: bool
    page: int


class SpiritLookupController:
    """High-level operations shared between UI and tests."""

    def __init__(self, provider: BaseDataProvider, *, page_size: int = 50) -> None:
        self.provider = provider
        self.page_size = page_size

    def list_records(self, query: str = "", *, page: int = 0) -> LookupResult:
        records, has_more = self.provider.list_records(query, page=page, page_size=self.page_size)
        return LookupResult(records=records, has_more=has_more, page=page)

    def get_record(self, spirit_code: str) -> SpiritRecord:
        return self.provider.get_record(spirit_code)

    def search_by_input(
        self,
        *,
        spirit_code: str | None,
        selected_label: str | None,
        cached_records: List[SpiritRecord],
    ) -> SpiritRecord:
        """Determine the desired spirit record from user input."""

        if spirit_code:
            return self.get_record(spirit_code)

        if selected_label:
            for record in cached_records:
                if record.display_label() == selected_label:
                    return record

        raise RecordNotFoundError("Es wurde kein Spirit Code ausgewählt.")
