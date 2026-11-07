# Changelog

Alle nennenswerten Änderungen des Projekts werden hier dokumentiert.
Format angelehnt an [Keep a Changelog](https://keepachangelog.com/de/1.1.0/),
Versionierung nach [Semantic Versioning](https://semver.org/lang/de/).

## [Unreleased]

### Added
- Git-Hook-Datenschutzschleuse unter `.githooks/` (`pre-commit`, `commit-msg`, `pre-push`; gemeinsame Regelquelle `lib.sh`): gesperrte E-Mail-Domains, persönliche Datenpfade (data/, privat/) und die Bank-Kennung werden nie Teil eines Commits oder Pushes; Diagnosen drucken gefundene Werte nie aus
- GitHub-Actions-Workflow `.github/workflows/datenschutz-pruefung.yml`, der dieselben Hook-Regeln auf der GitHub-Seite für alle neuen Commits ausführt
- Hook-README mit Installations- und Notfallregel (dokumentierter Bypass `RELEASE_ALLOW=1`)
- GitHub-Actions-Workflow `.github/workflows/actionlint.yml`, der die Workflow-Dateien selbst statisch prüft (Syntax, Ausdruckskontexte, Action-Inputs) über das offizielle Image `rhysd/actionlint:1.7.12` samt shellcheck und pyflakes

### Changed
- **Pre-Push-Gate (07.11.2026):** `pre-push` führt nach den Datenschutzprüfungen die 19 Unit-Tests (`python -m unittest discover -s tests -v`) und Ruff (`python -m ruff check .`) aus; beide blockieren den Push bei Fehlern.
- **Ruff-Policy (07.11.2026):** `pyproject.toml` aktiviert eine dreistufige Regelauswahl: Korrektheit (`E4`, `E7`, `E9`, `F`), Struktur und moderne Python-Nutzung (`I`, `UP`, `SIM`, `PTH`, `RUF`) sowie Robustheit und Sicherheit (`B`, `BLE`, `TRY`, `S`, `DTZ`). `S101` ist nur für `tests/*` ausgenommen; im Anwendungscode bleibt die Regel aktiv. Damit hängt der Prüfumfang nicht mehr von einer globalen `RUFF_CONFIG` mit `select = ["ALL"]` ab.
- **Agentenregeln (06.11.2026):** `AGENTS.md` neu angelegt und auf Aufruf, Tests, Ruff und die Datenschutz-Schleuse reduziert. Allgemeine Arbeits-, Datenschutz- und Git-Regeln stehen in `$HOME/AGENTS.md`.
- Privater Ordner heißt jetzt `privat/` (davor `docs/`, zuvor `backup/`) — eindeutiger Name, damit die Konvention "docs/ = öffentliche Projektdoku" frei bleibt; `.gitignore` und Hook-Regeln synchron angepasst

### Fixed
- **Ruff-Befunde (07.11.2026):** Alle 12 Befunde der dreistufigen Auswahl behoben: Importreihenfolge, `pathlib` statt `os.path`, Ternary statt `if`/`else`, eindeutige Kommentar- und Docstring-Zeichen. `DTZ007` ist an der Datumskonvertierung eng begründet (reine Kalenderdaten ohne Zeitzone), `BLE001` an der PDF-Batch-Grenze (eine defekte Datei darf den Lauf nicht abbrechen); beide Ausnahmen stehen im Code, nicht global in der Konfiguration.
- **Verifikation (07.11.2026):** `python -m ruff check .` und `python -m ruff format --check .` ohne Befunde; alle 19 Unit-Tests bestanden.

## [1.0.0] – 2026-10-21

### Added
- Erste öffentliche Version des Kontoauszug-Parser (deutsche und englische README, GPL-3.0)
- README-Abschnitt "Parser-Logik": zwei UND-verknüpfte Bedingungen der Buchungserkennung (Datum am Zeilenanfang UND Betrag am Zeilenende), Ignorierregeln für Fortführungs-/Saldozeilen
- Unit-Tests für die Parser-Kernfunktionen (ohne echte PDFs), ausführbar mit `python -m unittest discover -s tests`
- Zentrale Projekt- und Toolkonfiguration in `pyproject.toml`

### Changed
- Excel-Ausgabe schreibt Buchungsdatum und Betrag als echte Excel-Typen (Datum und Zahl mit festem Format) statt als Text
- Datei-Suche über `--prefix` nutzt `fnmatch`; Muster ohne Stern behalten Prefix-Semantik
- Standard-Suchmuster auf "Girokonto_*" erweitert – so werden auch Kontoauszüge mit Datums-Suffix gefunden (z. B. ..._Kontoauszug_20250131.pdf)

### Fixed
- Die erste Buchungszeile einer Seite geht nicht mehr verloren, wenn unter der Spaltenüberschrift keine "Valuta"-Zeile folgt
- Fußnotenzeichen hinter dem Betrag werden auch mit Leerzeichen davor korrekt ignoriert (z. B. "300,00 1")
- Buchungserkennung vom Layout entkoppelt: Der extract_text()-Text von pdfplumber enthält keine Leerzeilen zwischen den Blöcken. Maßgeblich ist jetzt die erste Blockzeile (Datum + Verwendungszweck + Betrag in einer Zeile); Valuta-/Fortführungszeilen (kein Betrag am Zeilenende) und der rechnerische Saldo werden ignoriert – Live-Test: alle 49 Buchungen identisch zur manuellen Zählung (7/10/9/9/9/5 je Seite)
