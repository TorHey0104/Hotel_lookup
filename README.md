# Hyatt EAME Hotel Lookup and Multi E-Mail Tool

Version: 4.2.2

## Requirements
- Python 3.x on Windows
- `pandas`, `openpyxl`, `pywin32` (the script will attempt to install `pandas`/`openpyxl` automatically)
- Microsoft Outlook installed (for drafting emails)

## Running
```bash
python "Hotel_lookup_interactive v4_2_2.py"
```

## Key Features
- Single hotel lookup and multi-hotel selection with filters.
- Role-based recipient routing (To/CC/BCC) for GM, Engineering, DOF, AVP, MD, Regional Eng Specialist.
- Outlook draft creation with placeholder-enabled subject/body and optional Outlook signature insertion.
- Configurable column mappings and visible columns for the filtered list (saved/loaded via JSON config).
- Splash screen showing version/author/file status while loading.

## Configuration
- Save/load configuration via Datei → Konfiguration speichern/laden.
- Config stores data file path, column mappings, role routing, and visible columns for the filtered list.

## Placeholders (subject/body)
`{hotel}, {spirit_code}, {city}, {relationship}, {brand}, {brand_band}, {region}, {country}, {owner}, {rooms}`

## Tips
- Place `hyatt_logo.png` next to the script to show the logo on the splash.
- Use About → About / Splash to reopen the splash info.

## Testing
No automated tests are included. Manual checks recommended:
- Launch app, ensure splash appears and closes (button or timeout).
- Load default data; verify filters populate.
- Try multi-email drafts with role routing and signatures.
- Save config, reload, and confirm mappings/visible columns persist.
