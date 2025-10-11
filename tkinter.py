"""Compatibility shim for tkinter to support headless test environments."""

from __future__ import annotations

import importlib
import os
import sys
from types import ModuleType


if not os.environ.get("DISPLAY"):
    raise ImportError("tkinter requires an available DISPLAY environment variable.")


def _load_real_tkinter() -> ModuleType:
    current_dir = os.path.dirname(__file__)
    real_module: ModuleType | None = None
    original_entry = sys.modules.pop("tkinter", None)
    original_path = list(sys.path)
    try:
        sys.path = [
            entry
            for entry in original_path
            if os.path.realpath(entry) != os.path.realpath(current_dir)
        ]
        real_module = importlib.import_module("tkinter")
    finally:
        sys.path = original_path
        if real_module is not None:
            sys.modules["tkinter"] = real_module
        elif original_entry is not None:
            sys.modules["tkinter"] = original_entry
    if real_module is None:
        raise ImportError("tkinter could not be loaded from the standard library.")
    return real_module


_real_tk = _load_real_tkinter()
globals().update(_real_tk.__dict__)

