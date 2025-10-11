from __future__ import annotations

import os
from pathlib import Path

import pytest

from spirit_lookup.config import AppConfig, load_config
from spirit_lookup.providers import (
    DataProviderError,
    FixtureDataProvider,
    RecordNotFoundError,
    create_provider,
)


def test_fixture_provider_full_cycle(tmp_path):
    fixture_path = Path("data/spirit_fixture.json")
    provider = FixtureDataProvider(fixture_path)
    records, has_more = provider.list_records("z")
    assert records
    assert isinstance(has_more, bool)

    with pytest.raises(RecordNotFoundError):
        provider.get_record("unknown")


def test_create_provider_missing_env(monkeypatch):
    monkeypatch.setenv("DATA_SOURCE", "sharepoint")
    for key in ["SP_TENANT_ID", "SP_CLIENT_ID", "SP_CLIENT_SECRET", "SP_SITE_ID", "SP_LIST_ID"]:
        monkeypatch.delenv(key, raising=False)
    config = load_config(Path.cwd())
    with pytest.raises(DataProviderError):
        create_provider(config)


def test_create_provider_fixture_default(monkeypatch):
    monkeypatch.delenv("DATA_SOURCE", raising=False)
    config = AppConfig()
    provider = create_provider(config)
    assert isinstance(provider, FixtureDataProvider)
