"""Entry point for the Spirit Lookup Tkinter application."""

from __future__ import annotations

from pathlib import Path

from spirit_lookup import SpiritLookupController, create_provider, load_config

try:  # pragma: no cover - imported for exception handling
    from tkinter import TclError as TkinterError
except Exception:  # pragma: no cover - tkinter might be unavailable in headless tests
    TkinterError = RuntimeError  # type: ignore[assignment]

_TKINTER_ERROR_MESSAGE = (
    "Die Tkinter-Oberfläche konnte nicht gestartet werden. "
    "Stellen Sie sicher, dass eine grafische Umgebung (DISPLAY) verfügbar ist."
)


def main() -> None:
    base_dir = Path(__file__).parent
    config = load_config(base_dir)
    provider = create_provider(config)
    controller = SpiritLookupController(provider, page_size=config.page_size)
    try:
        from spirit_lookup.ui import run_app
    except ImportError as exc:  # pragma: no cover - defensive guard
        raise SystemExit(_TKINTER_ERROR_MESSAGE) from exc

    try:
        run_app(config, controller)
    except TkinterError as exc:
        raise SystemExit(_TKINTER_ERROR_MESSAGE) from exc


if __name__ == "__main__":
    main()
