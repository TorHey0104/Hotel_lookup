"""Tests for the CLI entry point."""

from __future__ import annotations

import builtins
from types import ModuleType

import pytest

import main


def test_main_exits_with_helpful_message(monkeypatch: pytest.MonkeyPatch) -> None:
    """Ensure a helpful error is shown when the Tk UI cannot be imported."""

    original_import = builtins.__import__

    def fake_import(name, globals=None, locals=None, fromlist=(), level=0):  # type: ignore[override]
        if name == "spirit_lookup.ui":
            raise ImportError("DISPLAY is required")
        return original_import(name, globals, locals, fromlist, level)

    monkeypatch.setattr(builtins, "__import__", fake_import)

    with pytest.raises(SystemExit) as excinfo:
        main.main()

    assert "grafische Umgebung" in str(excinfo.value)


def test_main_exits_when_tkinter_fails(monkeypatch: pytest.MonkeyPatch) -> None:
    """The UI should surface a friendly message when Tk initialisation fails."""

    dummy_module = ModuleType("spirit_lookup.ui")

    def fake_run_app(config, controller):  # type: ignore[no-untyped-def]
        raise main.TkinterError("no display")

    dummy_module.run_app = fake_run_app  # type: ignore[attr-defined]

    original_import = builtins.__import__

    def fake_import(name, globals=None, locals=None, fromlist=(), level=0):  # type: ignore[override]
        if name == "spirit_lookup.ui":
            return dummy_module
        return original_import(name, globals, locals, fromlist, level)

    monkeypatch.setattr(builtins, "__import__", fake_import)
    monkeypatch.setattr(main, "create_provider", lambda _config: object())

    class DummyController:  # pragma: no cover - simple stub
        def __init__(self, provider, page_size):  # type: ignore[no-untyped-def]
            self.provider = provider
            self.page_size = page_size

    monkeypatch.setattr(main, "SpiritLookupController", DummyController)

    with pytest.raises(SystemExit) as excinfo:
        main.main()

    assert "grafische Umgebung" in str(excinfo.value)
