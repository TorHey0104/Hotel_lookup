"""Entry point for the Spirit Lookup Tkinter application."""

from __future__ import annotations

from pathlib import Path

from spirit_lookup import SpiritLookupController, create_provider, load_config
from spirit_lookup.ui import run_app


def main() -> None:
    base_dir = Path(__file__).parent
    config = load_config(base_dir)
    provider = create_provider(config)
    controller = SpiritLookupController(provider, page_size=config.page_size)
    run_app(config, controller)


if __name__ == "__main__":
    main()
