"""Entry point for the Spirit Lookup Tkinter application."""

from __future__ import annotations

from pathlib import Path

from spirit_lookup import SpiritLookupController, create_provider, load_config


def main() -> None:
    base_dir = Path(__file__).parent
    config = load_config(base_dir)
    provider = create_provider(config)
    controller = SpiritLookupController(provider, page_size=config.page_size)
    try:
        from spirit_lookup.ui import run_app
    except ImportError as exc:  # pragma: no cover - defensive guard
        raise SystemExit(
            "Die Tkinter-Oberfläche konnte nicht gestartet werden. "
            "Stellen Sie sicher, dass eine grafische Umgebung (DISPLAY) verfügbar ist."
        ) from exc

    run_app(config, controller)


if __name__ == "__main__":
    main()
