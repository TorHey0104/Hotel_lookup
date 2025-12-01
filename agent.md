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
- When cutting a new version, bump `VERSION` and `VERSION_DATE` near the top of `Hotel_lookup_interactive vX.py` so the splash and status text show the correct release info.
- As part of the release checklist, open the app once to confirm the splash displays the new version string and date.
- Attachments (v5+): support Common and Spirit-code-specific attachments under a user-selected root; folders named `Common` and `Spirit/<SpiritCode>` are expected.
- Forward assist (v5+): capture an Outlook email (selected in Outlook or browsed by subject in Inbox/Sent). Captured subject is prefixed with `FW:`; body (HTML or text) and attachments are reused in drafts ahead of app attachments.
