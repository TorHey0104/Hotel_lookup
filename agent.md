# Development Notes for Hyatt EAME Hotel Lookup

Baseline: version 4.2.2

## Scope
- Tkinter desktop app for hotel lookup and Outlook draft creation.
- Supports single and multi-hotel workflows, role-based routing, placeholders, and optional signatures.
- Configurable data columns and visible filtered columns; config saved/loaded via JSON.

## How to Run
```bash
python "Hotel_lookup_interactive v4_2_2.py"
```

## Important Files
- `Hotel_lookup_interactive v4_2_2.py` – main application.
- `README.md` – user instructions.
- `agent.md` – this file (dev notes).
- `hyatt_logo.png` – optional splash logo (place next to script).

## Manual Test Checklist
- Splash appears centered, shows version/author/file status, dismissible by button or timeout.
- Default data loads; filters populate; visible columns setting applies.
- Multi-email: select hotels, role routing To/CC/BCC works, placeholders render, signatures insert.
- Single-hotel: detail panel, role selection, draft creation with placeholders/signature.
- Config save/load restores data path, column mappings, role routing, visible columns.

## Future Development Considerations
- Add automated tests (UI harness) if feasible.
- Improve logging for load/config errors.
- Consider bundling dependencies to avoid first-run installs.
