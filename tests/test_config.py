from __future__ import annotations

import os
from pathlib import Path

from spirit_lookup.config import AppConfig, load_config


def test_load_config_defaults(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    config = load_config(tmp_path)
    assert config.data_source == "fixture"
    assert config.fixture_path == tmp_path / "data" / "spirit_fixture.json"
    assert config.page_size == 50
    assert config.debounce_ms == 250
    assert config.draft_email_enabled is True


def test_load_config_env(monkeypatch, tmp_path):
    monkeypatch.setenv("DATA_SOURCE", "sharepoint")
    monkeypatch.setenv("SPIRIT_FIXTURE_PATH", str(tmp_path / "custom.json"))
    monkeypatch.setenv("SPIRIT_PAGE_SIZE", "75")
    monkeypatch.setenv("SPIRIT_DEBOUNCE_MS", "400")
    monkeypatch.setenv("DRAFT_EMAIL_ENABLED", "0")
    config = load_config(tmp_path)
    assert config.use_sharepoint is True
    assert config.fixture_path == tmp_path / "custom.json"
    assert config.page_size == 75
    assert config.debounce_ms == 400
    assert config.draft_email_enabled is False


def test_app_config_use_sharepoint_property():
    cfg = AppConfig(data_source="sharepoint")
    assert cfg.use_sharepoint is True
    cfg2 = AppConfig(data_source="fixture")
    assert cfg2.use_sharepoint is False


def test_load_config_invalid_numbers(monkeypatch, tmp_path):
    monkeypatch.setenv("SPIRIT_PAGE_SIZE", "notanumber")
    monkeypatch.setenv("SPIRIT_DEBOUNCE_MS", "oops")
    config = load_config(tmp_path)
    assert config.page_size == 50
    assert config.debounce_ms == 250

    monkeypatch.delenv("SPIRIT_PAGE_SIZE", raising=False)
    monkeypatch.delenv("SPIRIT_DEBOUNCE_MS", raising=False)

