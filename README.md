# Spirit Lookup

Eine Tkinter-Desktopanwendung zur Recherche von Hotels anhand ihres **Spirit Codes**. Nutzende können den Code entweder direkt eingeben oder ein Hotel über ein Dropdown mit Typeahead und Pagination auswählen. Nach der Auswahl öffnet sich ein Dialog mit Schlüssel- und Kontaktinformationen sowie der Möglichkeit, – nach Bestätigung eines Hinweises – einen leeren E-Mail-Entwurf im Standard-Mailclient zu öffnen.

## Inhalt

- [Funktionen](#funktionen)
- [Systemvoraussetzungen](#systemvoraussetzungen)
- [Installation](#installation)
- [Konfiguration](#konfiguration)
- [Entwicklung](#entwicklung)
- [Tests](#tests)
- [Mockdaten](#mockdaten)
- [Troubleshooting](#troubleshooting)

## Funktionen

- Spirit-Code-Eingabe oder Auswahl via Dropdown mit 250 ms Debounce und Lazy-Loading ab 50 Treffern pro Seite.
- Detaildialog mit Schlüssel- und Kontaktinformationen, Copy-to-Clipboard-Aktionen und Tastatursteuerung (Esc schließt, Fokus bleibt im Dialog).
- Checkbox-Absicherung vor dem Öffnen eines leeren E-Mail-Entwurfs (`mailto:`) im Standard-Mailclient – inklusive Fehlermeldung, falls kein Client verfügbar ist.
- Datenquelle über Environment wählbar: SharePoint (Microsoft Graph) oder lokale JSON-Fixture.

## Systemvoraussetzungen

- Python 3.11 oder 3.12 (getestet in CI auf Ubuntu & Windows)
- Optional: Für den SharePoint-Modus einen registrierten Azure AD App-Client mit Zugriff auf die gewünschte Liste.

## Installation

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
pip install -r requirements-dev.txt  # Entwickler:innen
```

## Konfiguration

Legen Sie eine `.env` oder `.env.local` an (siehe `.env.example`). Wichtige Variablen:

| Variable | Beschreibung | Standard |
|----------|--------------|----------|
| `DATA_SOURCE` | `fixture` oder `sharepoint` | `fixture` |
| `SPIRIT_FIXTURE_PATH` | Optionaler Pfad zur Fixture-Datei | `data/spirit_fixture.json` |
| `SPIRIT_PAGE_SIZE` | Anzahl Ergebnisse pro Seite im Dropdown | `50` |
| `SPIRIT_DEBOUNCE_MS` | Debounce in Millisekunden | `250` |
| `DRAFT_EMAIL_ENABLED` | Aktiviert den Draft-E-Mail-Button | `true` |

Für den SharePoint-Modus werden zusätzlich `SP_TENANT_ID`, `SP_CLIENT_ID`, `SP_CLIENT_SECRET`, `SP_SITE_ID` und `SP_LIST_ID` benötigt.

## Entwicklung

```bash
python main.py
```

Der Legacy-Starter `Hotel_lookup_interactive v3.py` leitet automatisch auf `main.py` weiter.

## Tests

```bash
pytest
python tools/run_simple_coverage.py
```

Die Test-Suite umfasst Unit-, Integrations- und E2E-nahe Szenarien gegen die Fixture. Das Coverage-Skript basiert auf `sys.settrace` und ignoriert optional den SharePoint-spezifischen Teil; die angestrebte Abdeckung liegt bei ≥ 80 % für die Kernmodule.
## Mockdaten

Die Datei `data/spirit_fixture.json` enthält drei Beispiel-Hotels (ZRH001, LON123, DXB777). Für lokale Tests können weitere Einträge ergänzt werden. Die Struktur entspricht dem Interface `SpiritRecord` aus `spirit_lookup/models.py`.

### Excel-Helfer

Über den Button **„Excel Helper“** in der Spirit-Lookup-Anwendung lässt sich eine grafische Oberfläche öffnen, in der die gewünschte Excel-Datei ausgewählt und relevante Spalten markiert werden können. Die Auswahl wird als JSON-Konfiguration (`data/excel_helper_config.json`) gespeichert, automatisch wieder geladen und kann über die Auswahlliste im Dialog jederzeit erneut geöffnet werden. Nach dem Speichern weist der Dialog auf das Zielverzeichnis hin; nutzen Sie anschließend das Skript `tools/excel_to_fixture.py`, um auf Basis derselben Excel-Datei eine JSON-Fixture zu erzeugen.

Wer lieber auf der Kommandozeile arbeitet, kann weiterhin das Skript `tools/excel_to_fixture.py` nutzen:

```bash
python tools/excel_to_fixture.py meine_hotels.xlsx data/meine_hotels.json
```

Unterstützte Spaltenüberschriften (Groß-/Kleinschreibung egal, Leerzeichen erlaubt):

| Pflichtspalten | Optionale Spalten | Kontakte | Meta-Daten |
|----------------|-------------------|----------|------------|
| `Spirit Code`, `Hotel Name` | `Region`, `Status`, `City`, `Country`, `Address` | `Contact1 Role`, `Contact1 Name`, `Contact1 Email`, `Contact1 Phone` (für weitere Kontakte `Contact2 …`, `Contact3 …` usw.) | Spalten, die mit `Meta` beginnen, z. B. `Meta.launchYear` oder `Meta Notes` |

Das Skript liest standardmäßig das erste Tabellenblatt, unterstützt die Option `--sheet` zur Auswahl eines anderen Blatts und warnt, falls Spalten nicht zugeordnet werden konnten. Das Resultat lässt sich direkt als Fixture-Datei verwenden, indem `SPIRIT_FIXTURE_PATH` auf den erzeugten JSON-Pfad zeigt.

## Troubleshooting

- **Tkinter-Fehler „no display name“**: Auf Linux muss ggf. `sudo apt-get install python3-tk` oder `xvfb` installiert werden.
- **SharePoint-Authentifizierung schlägt fehl**: Prüfen Sie Client-ID, Secret und List-ID. Nutzen Sie `DATA_SOURCE=fixture` für lokale/offline Entwicklung.
- **Mailclient öffnet sich nicht**: Stellen Sie sicher, dass ein Standard-Mailprogramm registriert ist. Andernfalls erscheint eine Fehlermeldung.
