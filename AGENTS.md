# AGENTS.md – Operator Instructions (DE)

> **Zweck:** Dieses Dokument definiert klare, wiederholbare **Operator‑Anweisungen** für **AI‑Agenten (CODEX)** sowie menschliche Maintainer:innen, um Features, Tests, PRs und Releases im verknüpften Repo sicher, konsistent und prüfbar durchzuführen.

---

## 1) Scope & Zielbild

* **Produkt:** *Spirit Code Info Tool*

  * Eingabe **Spirit Code** (Tippfeld **oder** Dropdown mit Typeahead) → Anzeige **Key‑Informationen** & **Kontaktinformationen** in Modal/Drawer.
  * **Checkbox‑Gate**: Erst wenn die Checkbox aktiv ist, darf der Button **„Draft E‑Mail“** klickbar sein.
  * **Draft E‑Mail** öffnet **einen leeren Mail‑Entwurf** im Standard‑Mailclient (**ohne** To/CC/BCC/Betreff/Body Auto‑Füllung), via `mailto:`/OS‑Open.
* **Agent:** **CODEX** (AI Dev Agent) mit Repo- und CI‑Zugriff.
* **Ergebnis:** Fertige PRs inkl. Tests, Doku, CI‑Checks, A11y‑Basics.

---

## 2) Rollen & Verantwortlichkeiten

* **Product Owner:** Torsten (Engineering Ops EAME/HDS) – fachliche Priorisierung, Abnahme.
* **Maintainer:in:** Review, Qualitätsgate, Merge; Guardrails durchsetzen.
* **CODEX (AI Agent):** Implementierung, Tests, Doku, PR‑Erstellung gemäß SOP.
* **CI/Bot:** Lint/Test/Build; Required‑Checks durchsetzen; optional Auto‑Merge.

---

## 3) Grundsätze (Guardrails)

1. **Security/Privacy**

   * Keine **Secrets** im Repo; `.env` lokal, CI‑Secrets via GitHub Actions.
   * **Keine PII** in Logs; Kontakte nur anzeigen, nicht exportieren.
2. **Determinismus & Reproduzierbarkeit**

   * Klare Versionsangaben; Lockfiles commiten.
   * Tests → deterministisch & stabil.
3. **A11y & UX‑Resilienz**

   * Tastaturbedienung, Fokus‑Management, ARIA Labels für Dialog/Buttons.
   * Fehler/Leer/Lade‑Zustände mit verständlichen Hinweisen + Retry.
4. **Compliance mit Projektkonventionen**

   * Halte dich an bestehenden **Stack** und Linters/Formatter.
   * **Conventional Commits** und PR‑Template verwenden.

---

## 4) Architektur & Konfiguration

* **Datenquellenstrategie**

  * **Primär:** SharePoint/Graph/List (wenn Env‑Variablen vorhanden).
  * **Fallback:** `data/spirit-fixture.json` (statische Testdaten) für lokale/E2E‑Tests.
* **Konfig (Beispiele – je nach Stack einsetzen):**

  * `DATA_SOURCE=sharepoint|fixture`
  * `SP_TENANT_ID`, `SP_CLIENT_ID`, `SP_CLIENT_SECRET` (nicht im Repo), `SP_SITE_ID`, `SP_LIST_ID`
* **E‑Mail‑Entwurf (strict):** nur `mailto:` **ohne** Parameter. Fehlerfall → Nutzermeldung + Troubleshooting‑Hinweis.

---

## 5) Standard Operating Procedure (SOP)

### 5.1 Intake & Planung

1. Tickets/Issues lesen, **DoD** checken (siehe §7).
2. Falls unklare Anforderungen → **kleinen RFC ins PR‑Body** schreiben (Kontext, Annahmen, Risiken).

### 5.2 Branching & Commits

* Branch: `feature/spirit-info-modal-email-draft` oder gemäß Issue → `feat/<kurz‑slug>`
* **Conventional Commits**: `feat:`, `fix:`, `docs:`, `test:`, `refactor:` …

### 5.3 Implementierung (MVP → vollständig)

* Komponenten: `SpiritSelect` (Suche/Dropdown), `SpiritInfoModal` (Key/Contacts), `DraftEmailSection` (Checkbox + Button).
* DataProvider‑Strategie: `sharepoint` **oder** `fixture`; Timeouts/Retry; defensive Fehlerbehandlung.
* UI‑Zustände: Laden | Leer | Fehler | Erfolgsfall.
* **A11y**: Fokus‑Trap im Modal, Esc schließt, Enter bestätigt, `aria-*` korrekt.

### 5.4 Tests

* **Unit/Integration**: Parsing, Mapping, Stores, DataProvider.
* **E2E**:

  1. Select‑Flow (Tippen → Auswahl → Modal zeigt korrekte Daten)
  2. Error‑Flow (ungültiger Code)
  3. Network‑Error (500/Timeout) + Retry/Fallback
  4. Draft‑Flow (Checkbox an → Button aktiv → `mailto:` Aufruf **abfangen**/stuben)
  5. A11y‑Smoke (Tab‑Order, Esc, Fokus)
* **Coverage‑Ziel:** ≥ 80 % in Kernmodulen.

### 5.5 Dokumentation

* `README.md`: Setup, Run, Tests, Env, Mocks, Troubleshooting.
* `CHANGELOG.md`: Eintrag mit Feature/Fixed.
* Optional `docs/ARCHITECTURE.md`– Kurzüberblick.

### 5.6 Pull Request

* Titel: `feat(spirit): Info‑Modal & leerer Draft‑E‑Mail‑Flow`
* Body: siehe Template unten.
* Screens/GIFs: Select/Modal/Checkbox/Draft/Fehler.
* Labels: `feature`, `frontend`, `tests`, `ready-for-review`.
* Reviewer: Maintainer‑Team.
* **Auto‑Merge nur wenn**: alle Checks **grün** & Mindest‑Reviews erfüllt.

### 5.7 Post‑Merge

* Optional Release‑Tag (semver, falls relevant).
* Prüfen, ob CI‑Deploys/Staging grün; Rollback‑Plan dokumentieren.

---

## 6) Testumgebungen & CI

* **Node/TS**: Node 18 & 20; Jest/Vitest; Playwright/Cypress.
* **Python**: 3.11 & 3.12; pytest; ggf. `pytest-qt`/Electron Runner.
* **OS‑Matrix**: ubuntu‑latest, windows‑latest (optional macos‑latest).
* **Jobs**: `lint`, `test-unit`, `test-e2e` (nightly möglich), `build`.

**Minimaler GitHub‑Actions‑Workflow (Beispiel):**

```yaml
name: ci
on: [push, pull_request]
jobs:
  build-test:
    runs-on: ubuntu-latest
    strategy:
      matrix:
        node: [18, 20]
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with: { node-version: ${{ matrix.node }}, cache: 'npm' }
      - run: npm ci
      - run: npm run lint --if-present
      - run: npm test --if-present -- --ci
      - run: npm run build --if-present
```

---

## 7) Definition of Done (DoD)

* [ ] Suche + Dropdown performant (Debounce 200–300 ms, Pagination ab >100).
* [ ] Modal zeigt korrekte Key‑Infos & Kontakte (Fixture **und** SharePoint).
* [ ] Fehler/Leer/Lade‑Zustände vorhanden; klare UX‑Meldungen + Retry.
* [ ] Checkbox‑Gate vor „Draft E‑Mail“.
* [ ] Klick auf „Draft E‑Mail“ öffnet **leeren** Entwurf (Web/Electron/Python Desktop), Fehlerfälle abgefangen.
* [ ] A11y‑Basics erfüllt (Fokus, ARIA, Tastatur).
* [ ] Tests ≥ 80 % + E2E grün; Mailto‑Call stub/verified.
* [ ] CI grün; Required Checks aktiv.
* [ ] README/Docs/Changelog aktualisiert.
* [ ] PR‑Template vollständig; Screens/GIFs angehängt.

---

## 8) PR‑Template (Snippet)

```md
## Beschreibung
Kurz: Spirit‑Code Auswahl → Info‑Modal (Key/Contacts) + Checkbox‑Gate + "Draft E‑Mail" öffnet leeren Entwurf.

## Änderungen
- Komponenten: SpiritSelect, SpiritInfoModal, DraftEmailSection
- DataProvider: sharepoint|fixture + Fehler/Retry
- Tests: Unit/Integration/E2E (mailto stub)
- Docs: README, CHANGELOG

## Wie getestet
- Unit/Integration: Befehle + Screenshots/Logs
- E2E: Szenarien 1–5, alle grün

## Akzeptanzkriterien (DoD)
- [x] Suche/Dropdown performant
- [x] Modal korrekt (Fixture/SharePoint)
- [x] Checkbox‑Gate & leerer Draft
- [x] A11y‑Smoke
- [x] CI grün

## Risiken & Rollback
- Risiken: …
- Rollback: Revert PR #<id>

## Sonstiges
Screens/GIFs anbei
```

---

## 9) Commit‑Konventionen

* Beispiele:

  * `feat(spirit): add InfoModal with contacts + checkbox gate`
  * `test(e2e): stub mailto and verify openExternal`
  * `docs(readme): add setup and troubleshooting`

---

## 10) Troubleshooting

* **Mailclient öffnet nicht** → Prüfen: OS‑Handler für `mailto:` gesetzt? Electron `shell.openExternal`/Browser‑Popup geblockt?
* **SharePoint 401/403** → Env‑Variablen & App‑Permissions überprüfen; bei CI **Mocks** verwenden.
* **Leere Ergebnisse** → Fixture fallback aktivieren (`DATA_SOURCE=fixture`).
* **Flaky E2E** → Wartezeiten/`await` sauber; Netzwerkmocks stabilisieren.

---

## 11) Telemetrie & Logging

* Nur **nicht‑personenbezogene** Ereignisse (z. B. UI‑Fehlercodes, Latenzen).
* Keine E‑Mail‑Adressen/Telefonnummern in Logs.
* Log‑Level: `error`/`warn` standard; Debug nur lokal.

---

## 12) Glossar

* **Spirit Code**: Standort‑Identifikator.
* **Draft E‑Mail**: Neuer, **leer**er E‑Mail‑Entwurf im Standard‑Mailclient.
* **Fixture**: Statische Testdaten im Repo.
* **DoD**: Definition of Done, Abnahmekriterien.

---

## 13) Anhänge (Copy‑Paste‑Bausteine)

**.env.example (Beispiel)**

```env
DATA_SOURCE=fixture
SP_TENANT_ID=
SP_CLIENT_ID=
SP_CLIENT_SECRET=
SP_SITE_ID=
SP_LIST_ID=
```

**Node Scripts (Beispiel) – package.json**

```json
{
  "scripts": {
    "dev": "vite",
    "lint": "eslint .",
    "test": "vitest run",
    "e2e": "playwright test",
    "build": "vite build"
  }
}
```

**mailto‑Open (Web/Electron/Python) – Beispiele**

```ts
// Web/Electron
const openDraftEmail = () => window.open('mailto:');
```

```py
# Python Desktop
import os, sys, subprocess
url = 'mailto:'
if sys.platform.startswith('win'): subprocess.run(['start', url], shell=True)
elif sys.platform == 'darwin': subprocess.run(['open', url])
else: subprocess.run(['xdg-open', url])
```

---

**Hinweis:** Dieses Dokument ist die verbindliche **Operator‑Instruction**. Anpassungen (z. B. Reviewer‑Anzahl, Branch‑Protection, Release‑Prozess) bitte als PR gegen `AGENTS.md` einbringen und versionieren.
