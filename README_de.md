## README Kontoauszug-Parser pdf2xlsx - Mai 2026

Python-Skript zur Extraktion von Kontobewegungen aus PDF-Kontoauszügen.

Dieses Python-Skript wurde entwickelt, um mehrseitige Kontoauszüge in Excel-Dateien umzuwandeln, die dann in Excel oder OpenOffice-Calc eingelesen und weiter verarbeitet werden können.

Hierfür wird keine Adobe-Installation benötigt.

## Voraussetzungen
- Python 3.10+
- `pdfplumber` und `openpyxl`

## Installation der Abhängigkeiten
```bash
pip install pdfplumber openpyxl
```

## Parser-Logik

Eine einzelne Buchung erstreckt sich im PDF über mehrere Zeilen ("Block"); relevant ist nur die erste Zeile, weil nur sie den Betrag in der rechten Spalte trägt. 
Eine Zeile wird als Buchung erkannt, wenn **beide** Bedingungen erfüllt sind (logisches UND):

1. die Zeile beginnt mit einem Datum (TT.MM.JJJJ), **und**
2. die Zeile endet mit einer Dezimalzahl mit zwei Nachkommastellen (ggf. mit Tausender-Gruppierung, Vorzeichen und Fußnotenzeichen).

Fortführungszeilen einer Buchung (weiterer Verwendungszweck) und Zeilen ohne Betrag (z. B. der rechnerische Saldo) werden ignoriert.

Die Erkennungsparameter sind Konfigurationskonstanten am Skriptanfang (`HEADER_MARKER`, `BUCHUNG_REGEX`), so lassen sich kleinere Layout-Änderungen dort anpassen. 
Eine vollständige strukturierte Extraktion über die Tabellen-Geometrie des PDFs ist ein mögliches Folgeprojekt.

## Standard-Verwendung / Usage
```bash
python pdf2xlsx-Parser.py /pfad/zu/pdf-ordner

# Mit benutzerdefiniertem Ausgabename
python pdf2xlsx-Parser.py /pfad/zu/pdf-ordner --output mein_ergebnis.xlsx

# Mit benutzerdefiniertem Suchmuster
python pdf2xlsx-Parser.py /pfad/zu/pdf-ordner --prefix "Kontoauszug_*"
```

## Tests
Die Parser-Funktionen sind mit Unit-Tests abgedeckt, die ohne echte PDF-Dateien auskommen (nur synthetische Texte). 
Extra-Testframework ist nicht nötig (Stdlib `unittest`):

```bash
python -m unittest discover -s tests -v
```
