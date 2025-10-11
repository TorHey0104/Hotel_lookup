"""Application configuration utilities."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class AppConfig:
    """Configuration container for the Spirit Lookup application."""

    data_source: str = "fixture"
    fixture_path: Path = Path("data/spirit_fixture.json")
    page_size: int = 50
    debounce_ms: int = 250
    draft_email_enabled: bool = True

    @property
    def use_sharepoint(self) -> bool:
        return self.data_source.lower() == "sharepoint"


def load_config(base_dir: Path | None = None) -> AppConfig:
    """Load configuration from environment variables."""

    base_dir = base_dir or Path.cwd()
    data_source = os.getenv("DATA_SOURCE", "fixture")

    fixture_override = os.getenv("SPIRIT_FIXTURE_PATH")
    if fixture_override:
        fixture_path = Path(fixture_override)
    else:
        fixture_path = base_dir / "data" / "spirit_fixture.json"

    draft_email_flag = os.getenv("DRAFT_EMAIL_ENABLED", "true").lower() in {"1", "true", "yes"}

    page_size_env = os.getenv("SPIRIT_PAGE_SIZE")
    try:
        page_size = int(page_size_env) if page_size_env else 50
    except ValueError:
        page_size = 50

    debounce_ms_env = os.getenv("SPIRIT_DEBOUNCE_MS")
    try:
        debounce_ms = int(debounce_ms_env) if debounce_ms_env else 250
    except ValueError:
        debounce_ms = 250

    return AppConfig(
        data_source=data_source,
        fixture_path=fixture_path,
        page_size=page_size,
        debounce_ms=debounce_ms,
        draft_email_enabled=draft_email_flag,
    )
