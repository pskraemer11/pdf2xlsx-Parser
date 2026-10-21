# Changelog

Alle nennenswerten Änderungen des Projekts werden hier dokumentiert.
Format angelehnt an [Keep a Changelog](https://keepachangelog.com/de/1.1.0/),
Versionierung nach [Semantic Versioning](https://semver.org/lang/de/).

## [Unreleased]

### Added
- Git-Hook-Datenschutzschleuse unter `.githooks/` (`pre-commit`, `commit-msg`, `pre-push`; gemeinsame Regelquelle `lib.sh`): gesperrte E-Mail-Domains, persönliche Datenpfade (data/, privat/) und die Bank-Kennung werden nie Teil eines Commits oder Pushes; Diagnosen drucken gefundene Werte nie aus
- GitHub-Actions-Workflow `.github/workflows/datenschutz-pruefung.yml`, der dieselben Hook-Regeln auf der GitHub-Seite für alle neuen Commits ausführt
- Hook-README mit Installations- und Notfallregel (dokumentierter Bypass `RELEASE_ALLOW=1`)

### Changed
- Privater Ordner heißt jetzt `privat/` (davor `docs/`, zuvor `backup/`) — eindeutiger Name, damit die Konvention "docs/ = öffentliche Projektdoku" frei bleibt; `.gitignore` und Hook-Regeln synchron angepasst

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
