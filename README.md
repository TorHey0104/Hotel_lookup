# Hyatt EAME Hotel Lookup and Multi E-Mail Tool

Version: 5.2.0

## Requirements
- Windows with Python 3.x
- `pandas`, `openpyxl`, `pywin32` (auto-install attempted for pandas/openpyxl)
- Microsoft Outlook (for drafting emails)

## Running
```bash
python "Hotel_lookup_interactive v5_1_2.py"
```

## What It Does
- **Lookup tab**: pick a hotel and compose a single Outlook draft inline (subject/body placeholders, signature picker, per-recipient To/CC/BCC, “Insert Link…” helper). Separate attachment root for single emails.
- **Multi-Email tab**: filter by Brand/Brand Band/Relationship/Region/Country, Hyatt date modes, quick Spirit filter; move filtered → selected and draft many emails. Role routing (AVP/MD/GM/Engineering/DOF/Regional Eng Specialist) with N/A filtering.
- **Forward assist (multi)**: browse Outlook Inbox/Sent to reuse subject/body (prefixed `FW:`) and attachments; your note/signature sit above forwarded content.
- **Attachments**: root folder with `Common` subfolder (all emails) plus per-Spirit Code subfolders; multi-email and single-email attachment pickers are independent.
- **Configs**: column mappings, role routing, visible columns, attachment settings, and data file path stored in JSON; recent configs offered at startup. Splash shows version/author/status.
- **Friendly links**: use `[label](url)` or “Insert Link…”; missing schemes are auto-prefixed with `https://`.

## Configuration
- On launch, choose a config (or skip). Later: Datei → Konfiguration laden/speichern.
- Config stores data file path, columns, roles, visible columns, and attachment roots.

## Placeholders (subject/body)
`{hotel}, {spirit_code}, {city}, {relationship}, {brand}, {brand_band}, {region}, {country}, {owner}, {rooms}`

## Tips
- Place `hyatt_logo.png` next to the script to show the logo on the splash.
- Use About → About / Splash to reopen the splash info.
- For attachments, create `Common` and per-Spirit-Code subfolders under the chosen root.
- Use “Insert Link…” to add friendly links; `[Google](www.google.com)` becomes clickable.

## Manual Checkpoints
- Splash appears, then config prompt; loading a config sets columns/roles/attachments.
- Filters populate; moving hotels between filtered/selected works.
- Single and multi-email drafts honor routing, signatures, links, and attachments.
- Saving/loading config restores mappings, visible columns, and attachment roots.
