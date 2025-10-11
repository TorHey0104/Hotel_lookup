from __future__ import annotations

import pytest

from spirit_lookup.config import AppConfig
from spirit_lookup.providers import FixtureDataProvider


@pytest.fixture()
def fixture_provider() -> FixtureDataProvider:
    fixture_path = AppConfig().fixture_path
    return FixtureDataProvider(fixture_path)
