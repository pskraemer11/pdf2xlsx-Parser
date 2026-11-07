---
status: accepted
date: 2026-11-07
---

# Ruff-Regeln werden projektlokal in drei Stufen gepflegt

`pyproject.toml` legt die Ruff-Auswahl selbst fest: Stufe 1 prüft mit
`E4`, `E7`, `E9` und `F` Syntax-, Namens- und Importfehler; Stufe 2 ergänzt
`I`, `UP`, `SIM`, `PTH` und `RUF` für Struktur und moderne Python-Nutzung;
Stufe 3 ergänzt `B`, `BLE`, `TRY`, `S` und `DTZ` für Robustheit und Sicherheit.
`S101` ist nur für `tests/*` ausgenommen; im Anwendungscode bleibt die Regel
aktiv. `DTZ007` und `BLE001` sind nur an den betroffenen Stellen im Code
ausgenommen, nicht global in der Konfiguration; die Begründung steht im
`CHANGELOG.md`. Grund: Eine globale
`RUFF_CONFIG` mit `select = ["ALL"]` würde den Prüfumfang dieses Projekts
verändern und den Pre-Push-Hook unabhängig von der Projektkonfiguration
scheitern lassen. Die dreistufige Auswahl hält den Einstieg nachvollziehbar und
macht jede Erweiterung zu einer bewussten Entscheidung.

## Considered Options

- **`select = ["ALL"]`:** maximal streng, aber der Prüfumfang wäre nicht mehr
  projektlokal kontrolliert; eine globale Änderung könnte den Hook beeinflussen.
- **Nur `E4`, `E7`, `E9`, `F`:** wenig Reibung, aber Robustheit, Sicherheit und
  moderne Python-Nutzung blieben ungeprüft.
- **Globale Ruff-Konfiguration verwenden:** weniger Duplikation, aber Gate und
  CI würden nicht mehr dieselbe projekteigene Regelauswahl prüfen.

## Notes

- Die drei Stufen sind keine Rangfolge der Wichtigkeit, sondern eine
  Lesereihenfolge: Korrektheit zuerst, dann Struktur, dann Robustheit.
- Die 12 Befunde aus der Einführung wurden im Code behoben, nicht durch weitere
  Ausnahmen. Die beiden verbleibenden `noqa`-Kommentare sind lokal begründet.
- Kriterium „schwer umkehrbar“ ist nur schwach erfüllt – die Konfiguration ist
  eine Liste. Die Entscheidung ist ohne Kontext überraschend und war ein echter
  Trade-off zwischen strengerer Prüfung und Reibung.
