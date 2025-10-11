"""Tests for the CLI entry point."""

from __future__ import annotations

import builtins

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
