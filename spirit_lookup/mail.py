"""Helpers to open a draft email in the default mail client."""

from __future__ import annotations

import os
import subprocess
import sys
from typing import Protocol


class MailClientError(RuntimeError):
    """Raised when the system cannot open the default mail client."""


class MailOpener(Protocol):
    def __call__(self, uri: str) -> None:
        ...


def _open_macos(uri: str) -> None:
    subprocess.run(["open", uri], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def _open_windows(uri: str) -> None:
    os.startfile(uri)  # type: ignore[attr-defined]


def _open_linux(uri: str) -> None:
    subprocess.run(["xdg-open", uri], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def open_mail_client(uri: str = "mailto:") -> None:
    """Open the user's default mail client with the given URI."""

    opener: MailOpener
    try:
        if sys.platform.startswith("darwin"):
            opener = _open_macos
        elif os.name == "nt":
            opener = _open_windows
        else:
            opener = _open_linux
        opener(uri)
    except Exception as exc:
        raise MailClientError("Der Standard-Mailclient konnte nicht geöffnet werden.") from exc
