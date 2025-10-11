# Architekturüberblick

Die Anwendung ist in drei Schichten organisiert:

1. **UI (`spirit_lookup/ui.py`)** – Enthält die Tkinter-Oberfläche mit Fokus auf Zugänglichkeit (Tastaturkürzel, Fokusfallen), Lade-/Fehlerzuständen sowie dem Info-Dialog mit Checkbox-gesteuertem Draft-E-Mail-Button.
2. **Controller (`spirit_lookup/controller.py`)** – Liefert eine dünne Fassade über den Datenprovidern und stellt wiederverwendbare Suchlogik bereit. Diese Schicht wird in Unit- und Integrations-Tests adressiert und kann unabhängig von Tkinter verwendet werden.
3. **Datenprovider (`spirit_lookup/providers`)** – Implementiert sowohl den Fixture- als auch den SharePoint-Zugriff. Die Provider liefern `SpiritRecord`-Objekte aus `spirit_lookup/models.py`.

## Datenfluss

```mermaid
graph TD
    UI -->|Suche| Controller
    Controller -->|list_records|getProvider[DataProvider]
    DataProvider -->|SpiritRecord[]| Controller
    Controller -->|Record| UI
    UI -->|mailto:| MailClient
```

## Tests

- **Unit-Tests**: Prüfen die Mapper und Provider (Fixture, Controller, Mail-Helper).
- **Integration**: `test_controller_flow` verbindet Controller und Fixture-Provider.
- **E2E-nah**: `test_e2e_draft_flow` simuliert den kompletten Nutzungsfluss inkl. Stub für `mailto:`.

## Erweiterbarkeit

- Weitere Datenquellen lassen sich durch Implementieren von `BaseDataProvider` ergänzen.
- UI-Texte sind zentral im UI-Modul definiert und können für Internationalisierung extrahiert werden.
- Die Draft-E-Mail-Funktion lässt sich über den Feature-Flag `DRAFT_EMAIL_ENABLED` deaktivieren.
