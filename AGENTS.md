# AGENTS.md – pdf2xlsx-Parser

Stand: 07.11.2026

## Zweck und Aufruf

- Kein installierbares Paket; das Skript wird direkt ausgeführt:
  `python pdf2xlsx-Parser.py <pdf-ordner>`.
- Abhängigkeiten: `pdfplumber` und `openpyxl`; Python 3.10+.
- `data/` und `privat/` enthalten persönliche Daten und werden nie committet.

## Tests und Prüfungen

- Unit-Tests ohne echte PDFs: `python -m unittest discover -s tests -v`.
- Ruff: `python -m ruff check .` (Konfiguration in `pyproject.toml`, Zeilenlänge 100).
- CI: `.github/workflows/datenschutz-pruefung.yml` und `.github/workflows/actionlint.yml`.

## Datenschutz-Schleuse

- Die verbindlichen Regeln stehen in `.githooks/lib.sh`; Aktivierung mit `git config core.hooksPath .githooks`.
- `pre-commit` prüft Identität, neue Textzeilen und persönliche Dateipfade im Staging.
- `commit-msg` prüft die Commit-Nachricht; `pre-push` prüft alle neuen Commits, Dateiinhalte und Datenpfade.
- Nach den Datenschutzprüfungen führt `pre-push` zusätzlich die Unit-Tests (`python -m unittest discover -s tests -v`) und Ruff (`python -m ruff check .`) aus; beide blockieren den Push bei Fehlern.
- Gesperrte E-Mail-Domains, persönliche Daten unter `data/`, `privat/` und lokale `COMPACTIONS.md` erreichen nie einen Commit oder Push.
- `RELEASE_ALLOW=1` ist nur für dokumentierte Ausnahmen zulässig, nie als Routine.

## Entscheidungen (ADR)

Dauerhafte Entscheidungen liegen unter `docs/adr/NNNN-slug.md` (nächste freie
Nummer verwenden). Kriterien und Format stehen in
`$HOME/.agents/references/compaction-workflow.md`, Abschnitt
„Promoting a decision to an ADR“; Beispiel ist
`docs/adr/0001-ruff-regeln-projektlokal-in-drei-stufen.md`. Das Skill
`domain-modeling` ist im aktuellen Harness nicht verfügbar; der Schritt wird
manuell ausgeführt.
