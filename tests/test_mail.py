from __future__ import annotations

import os
import sys
from unittest import mock

import pytest

import spirit_lookup.mail as mail
from spirit_lookup.mail import MailClientError, open_mail_client


def test_open_mail_client_windows(monkeypatch):
    monkeypatch.setattr(os, "name", "nt")
    with mock.patch("spirit_lookup.mail.os.startfile", create=True) as startfile:  # type: ignore[attr-defined]
        open_mail_client("mailto:")
        startfile.assert_called_once_with("mailto:")


def test_open_mail_client_macos(monkeypatch):
    monkeypatch.setattr(sys, "platform", "darwin")
    with mock.patch("spirit_lookup.mail.subprocess.run") as run:
        open_mail_client("mailto:")
        run.assert_called_once()


def test_open_mail_client_failure(monkeypatch):
    monkeypatch.setattr(os, "name", "posix")
    monkeypatch.setattr(sys, "platform", "linux")
    with mock.patch("spirit_lookup.mail.subprocess.run", side_effect=OSError("boom")):
        with pytest.raises(MailClientError):
            open_mail_client("mailto:")


def test_platform_specific_helpers(monkeypatch):
    with mock.patch("spirit_lookup.mail.subprocess.run") as run_macos:
        mail._open_macos("mailto:")
        run_macos.assert_called_once()
    with mock.patch("spirit_lookup.mail.subprocess.run") as run_linux:
        mail._open_linux("mailto:")
        run_linux.assert_called_once()

    with mock.patch("spirit_lookup.mail.os.startfile", create=True) as startfile:
        mail._open_windows("mailto:")
        startfile.assert_called_once_with("mailto:")
