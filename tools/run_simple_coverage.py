"""Compute coverage for the spirit_lookup package using sys.settrace."""

from __future__ import annotations

import sys
import threading
from collections import defaultdict
from pathlib import Path

import pytest

PACKAGE_DIR = Path(__file__).resolve().parent.parent / "spirit_lookup"


def _line_tracer(frame, event, arg):
    if event == "line":
        filename = Path(frame.f_code.co_filename)
        if not str(filename).startswith(str(PACKAGE_DIR)):
            return _line_tracer
        EXECUTED_LINES[filename].add(frame.f_lineno)
    return _line_tracer


EXECUTED_LINES: dict[Path, set[int]] = defaultdict(set)


def main() -> int:
    sys.settrace(_line_tracer)
    threading.settrace(_line_tracer)
    try:
        exit_code = pytest.main()
    finally:
        sys.settrace(None)
        threading.settrace(None)

    total_lines = 0
    executed_lines = 0
    summaries = []

    excluded = {"spirit_lookup/ui.py", "spirit_lookup/__init__.py"}
    for file_path in sorted(PACKAGE_DIR.rglob("*.py")):
        rel = file_path.relative_to(PACKAGE_DIR.parent)
        if rel.as_posix() in excluded:
            continue
        with file_path.open("r", encoding="utf-8") as handle:
            lines = list(handle)
        relevant_lines = []
        ignore_block = False
        for idx, content in enumerate(lines, start=1):
            stripped = content.strip()
            if "coverage: ignore start" in stripped:
                ignore_block = True
                continue
            if "coverage: ignore end" in stripped:
                ignore_block = False
                continue
            if (
                stripped
                and not stripped.startswith("#")
                and "pragma: no cover" not in stripped
                and not ignore_block
            ):
                relevant_lines.append(idx)
        executed = len(EXECUTED_LINES.get(file_path, set()) & set(relevant_lines))
        total = len(relevant_lines)
        total_lines += total
        executed_lines += executed
        percent = (executed / total * 100) if total else 100.0
        summaries.append((percent, executed, total, file_path))

    print("Coverage Report (sys.settrace)")
    for percent, executed, total, file_path in summaries:
        rel_path = file_path.relative_to(PACKAGE_DIR.parent)
        if rel_path.as_posix() in excluded:
            continue
        print(f"{percent:6.2f}% {executed:5d}/{total:5d} {rel_path}")

    overall = (executed_lines / total_lines * 100) if total_lines else 100.0
    print(f"Overall coverage: {overall:.2f}%")
    if overall < 80.0:
        print("ERROR: Coverage below 80%", file=sys.stderr)
        return max(exit_code, 1)
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
