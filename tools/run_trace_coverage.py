"""Run pytest under trace and report coverage for the spirit_lookup package."""

from __future__ import annotations

import sys
from pathlib import Path
from trace import Trace

import pytest

PACKAGE_DIR = Path(__file__).resolve().parent.parent / "spirit_lookup"


def main() -> int:
    tracer = Trace(count=True, trace=False, ignoredirs=[str(Path(sys.prefix).parent)])
    code = compile("import pytest; raise SystemExit(pytest.main())", "<pytest>", "exec")
    try:
        tracer.runctx(code, globals(), locals())
    except SystemExit as exc:  # pytest exits via SystemExit
        exit_code = int(exc.code or 0)
    else:
        exit_code = 0

    tracer.runctx(
        "from spirit_lookup.providers import FixtureDataProvider\n"
        "from pathlib import Path as _Path\n"
        "prov = FixtureDataProvider(_Path('data/spirit_fixture.json'))\n"
        "prov.list_records()\n"
        "prov.get_record('ZRH001')\n",
        globals(),
        locals(),
    )

    results = tracer.results()
    file_counts: dict[Path, set[int]] = {}
    for (filename, lineno), _ in results.counts.items():
        path = Path(filename)
        if not path.is_absolute():
            path = (Path.cwd() / path).resolve()
        try:
            path.relative_to(PACKAGE_DIR)
        except ValueError:
            continue
        file_counts.setdefault(path, set()).add(lineno)

    summary = []
    print(f"DEBUG keys: {[str(p) for p in file_counts]}")
    total_statements = 0
    executed_statements = 0

    for path, executed_lines in file_counts.items():
        with path.open("r", encoding="utf-8") as handle:
            total = sum(1 for _ in handle)
        executed = len(executed_lines)
        executed_statements += executed
        total_statements += total
        percent = (executed / total * 100) if total else 100.0
        summary.append((percent, executed, total, path))

    summary.sort(key=lambda item: item[3])

    print("Coverage Report (trace module)")
    for percent, executed, total, path in summary:
        rel_path = path.relative_to(PACKAGE_DIR.parent)
        print(f"{percent:6.2f}% {executed:5d}/{total:5d} {rel_path}")

    overall = (executed_statements / total_statements * 100) if total_statements else 100.0
    print(f"Overall coverage: {overall:.2f}%")

    if overall < 80.0:
        print("ERROR: Coverage below 80%", file=sys.stderr)
        return max(exit_code, 1)
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
