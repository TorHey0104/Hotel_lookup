from __future__ import annotations

import os
import sys
from unittest import mock

from spirit_lookup.config import AppConfig
from spirit_lookup.controller import SpiritLookupController
from spirit_lookup.mail import open_mail_client
from spirit_lookup.providers import FixtureDataProvider


def test_e2e_draft_flow(monkeypatch):
    config = AppConfig()
    provider = FixtureDataProvider(config.fixture_path)
    controller = SpiritLookupController(provider, page_size=2)

    result = controller.list_records("zrh", page=0)
    assert result.records, "Es werden Treffer erwartet"

    record = controller.search_by_input(
        spirit_code=None,
        selected_label=result.records[0].display_label(),
        cached_records=result.records,
    )
    assert record.contacts, "Kontaktinformationen werden erwartet"

    monkeypatch.setattr(os, "name", "posix")
    monkeypatch.setattr(sys, "platform", "linux")
    with mock.patch("subprocess.run") as run:
        open_mail_client("mailto:")
        run.assert_called_once()
        args, kwargs = run.call_args
        assert "mailto:" in args[0][-1]
