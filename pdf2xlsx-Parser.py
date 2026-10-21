"""
Kontoauszug-Parser
==================
Extrahiert Kontobewegungen aus PDF-Kontoauszügen und speichert
die Daten als Excel-Datei.

Voraussetzungen:
    pip install pdfplumber openpyxl

Verwendung:
    python pdf2xlsx-Parser.py /pfad/zu/pdf-ordner
    python pdf2xlsx-Parser.py /pfad/zu/pdf-ordner --output ergebnis.xlsx
    python pdf2xlsx-Parser.py /pfad/zu/pdf-ordner --prefix "Girokonto_*_Kontoauszug"

Autor: OpenCode AI
Lizenz: GPL-3.0
"""

import argparse
import fnmatch
import os
import re
import sys
from datetime import date, datetime

import openpyxl
from openpyxl.styles import Font
import pdfplumber


# ============================================================================
# KONFIGURATION (anpassbar)
# ============================================================================

# Standard-Suchmuster für PDF-Dateinamen
# (Alle Dateien, deren Name mit "Girokonto_" beginnt und auf ".pdf" endet -
# damit greifen auch Dateien mit Suffix wie "..._Kontoauszug_20250131.pdf")
DEFAULT_FILE_PREFIX = "Girokonto_*"

# Standard-Ausgabedatei (wird im Quellordner abgelegt)
DEFAULT_OUTPUT_FILENAME = "Kontoauszuege_Zusammenfassung.xlsx"

# Spaltenüberschrift, die den Beginn der Buchungsdaten markiert
# (muss in jeder PDF-Seite vorkommen, die Buchungen enthält)
HEADER_MARKER = ["Buchung", "Buchung / Verwendungszweck", "Betrag (EUR)"]

# Kennzeichen der Hilfszeile direkt unter der Spaltenüberschrift
VALUTA_MARKER = "Valuta"

# Regex-Muster zur Erkennung von Buchungszeilen:
#   Gruppe 1: Datum (DD.MM.YYYY)
#   Gruppe 2: Verwendungszweck (beliebiger Text)
#   Gruppe 3: Betrag (Dezimalzahlen mit optionalen Tausendertrennzeichen)
#   Optionale Fußnotenzeichen (¹, ², ³, 1, 2, 3) werden am Ende ignoriert -
#   mit oder ohne Leerzeichen vor dem Zeichen (z. B. "300,00 1")
BUCHUNG_REGEX = re.compile(
    r"(\d{2}\.\d{2}\.\d{4})\s+(.*?)\s+([-+]?\d{1,3}(?:\.\d{3})*(?:,\d{2}))(?:\s?[¹²³1-3])?\s*$"
)

# Erkennung, ob eine Zeile überhaupt wie eine Buchungszeile beginnt (Datum am
# Zeilenanfang). Alle übrigen nicht-leeren Zeilen sind Fortführungszeilen des
# Verwendungszwecks und werden ignoriert - sie dürfen NICHT als Fehler zählen,
# denn jeder Buchungsblock besteht aus mehreren Zeilen, von denen nur die
# erste ausgewertet wird.
DATUM_ANFANG_REGEX = re.compile(r"^\d{2}\.\d{2}\.\d{4}")

# Zeilen, die NUR aus einem Datum bestehen: Im Auszugs-Layout steht in der
# linken Spalte unter dem Buchungsdatum noch das "Valuta"-Datum. pdfplumber
# liefert dieses zweite Datum als eigene Zeile ohne Verwendungszweck und
# zonder Betrag - solche Zeilen sind KEINE Buchungen und werden ignoriert.
DATUM_ALLEIN_REGEX = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")


# ============================================================================
# HILFSFUNKTIONEN
# ============================================================================


def finde_pdf_dateien(ordner: str, muster: str) -> list[str]:
    """
    Sucht im angegebenen Ordner nach PDF-Dateien, die dem Namensmuster entsprechen.

    Args:
        ordner: Pfad zum Ordner mit den PDF-Dateien
        muster: Suchmuster für Dateinamen, konsolidiert über fnmatch.
                Stern (*) steht für beliebigen Text. Ein Muster ohne Stern
                wird als Präfix interpretiert und um "*" ergänzt
                (z. B. "Girokonto_" ⇔ "Girokonto_*").

    Returns:
        Sortierte Liste der gefundenen PDF-Dateinamen
    """
    # Präfix-Semantik beibehalten: Muster ohne "*" bedeutet "beginnt mit"
    if "*" not in muster:
        muster = muster + "*"

    gefundene = [f for f in os.listdir(ordner) if fnmatch.fnmatch(f, muster) and f.endswith(".pdf")]

    return sorted(gefundene)


def extrahiere_seite(text: str, seiten_nummer: int) -> list[tuple[date, str, float]]:
    """
    Extrahiert Buchungsdaten aus einer einzelnen PDF-Seite (zeilengenau).

    Args:
        text: Der extrahierte Text der Seite
        seiten_nummer: Nummer der aktuellen Seite (für Ausgaben am Terminal)

    Returns:
        Buchungen der Seite (Buchungsdatum, Verwendungszweck, Betrag)
    """
    zeilen = text.split("\n")
    start_index: int | None = None

    # --- 1. Spaltenüberschrift finden ---
    for i, zeile in enumerate(zeilen):
        # Prüfe ob alle Bestandteile der Überschrift vorhanden sind
        if all(marker in zeile for marker in HEADER_MARKER):
            # Direkt unter der Überschrift folgt im Auszugs-Layout eine reine
            # "Valuta"-Zeile. Falls sie auf dieser Seite fehlt, direkt hinter
            # der Überschrift weitermachen, sonst würde die erste Buchung der
            # Seite stillschweigend verloren gehen.
            # Zeile nach der Überschrift ermitteln (am Seitenende kann sie fehlen)
            naechste_zeile = ""
            if i + 1 < len(zeilen):
                naechste_zeile = zeilen[i + 1]

            if VALUTA_MARKER in naechste_zeile:
                # Reine "Valuta"-Zeile → auch dieser Blockkopf wird übersprungen
                start_index = i + 2
            else:
                # Erste Buchung steht direkt unter der Überschrift
                start_index = i + 1
            print(f"   → Seite {seiten_nummer}: Spaltenüberschrift in Zeile {i + 1} gefunden.")
            break

    if start_index is None:
        # Keine Überschrift gefunden → Seite enthält keine Buchungsdaten
        print(f"   → Seite {seiten_nummer}: Keine Spaltenüberschrift gefunden. Überspringe Seite.")
        return []

    # --- 2. Buchungszeilen zeilengenau auswerten ---
    # Der von pdfplumber gelieferte Text enthält zwischen den Buchungsblöcken
    # KEINE Leerzeilen (im Original-Layout sind die Blöcke nur visuell getrennt).
    # Kennzeichen der maßgeblichen ersten Blockzeile ist der Betrag in der
    # rechten Spalte: NUR sie endet auf einen Betrag. Deshalb gilt pro Zeile:
    #   a) kein Datum am Zeilenanfang           → Fortführungszeile, ignorieren
    #   b) nur ein Datum                        → "Valuta"-Zeile, ignorieren
    #   c) Datum ohne Betrag am Zeilenende      → Valuta-Zeile mit Zweck-
    #      Fortführung bzw. Hinweiszeile (z. B. rechnerischer Saldo), ignorieren
    #   d) Datum + Verwendungszweck + Betrag    → Buchung, übernehmen
    treffer: list[tuple[date, str, float]] = []

    for zeile in zeilen[start_index:]:
        zeile = zeile.strip()

        if not zeile or not DATUM_ANFANG_REGEX.match(zeile):
            # (a) Leerzeile oder Fortführungszeile
            continue

        if DATUM_ALLEIN_REGEX.match(zeile):
            # (b) reine Valuta-Zeile
            continue

        buchung = verarbeite_zeile(zeile)
        if buchung is not None:
            # (d) vollständige Buchungszeile
            treffer.append(buchung)
        # Sonst: (c) Datum am Anfang, aber kein vollständiges Muster
        # (kein Betrag am Zeilenende) → Block-Fortführung, ignorieren

    return treffer


def verarbeite_zeile(zeile: str) -> tuple[date, str, float] | None:
    """
    Verarbeitet eine einzelne Buchungszeile und extrahiert Datum,
    Verwendungszweck und Betrag.

    Eine Buchung besteht im Auszugs-Layout aus einem Block mit mehreren
    Zeilen; ausgewertet wird aber nur die erste Zeile, denn sie enthält
    Datum, Verwendungszweck und Betrag in derselben Zeile.

    Args:
        zeile: Die erste (maßgebliche) Zeile eines Buchungsblocks

    Returns:
        (Buchungsdatum, Verwendungszweck, Betrag) oder None bei Fehler
    """
    erste_zeile = zeile.strip()

    # Regex-Muster auf die Zeile anwenden
    match = BUCHUNG_REGEX.match(erste_zeile)

    if not match:
        # Typischer Fall hier: Valuta-Zeile mit Zweck-Fortführung (kein Betrag
        # am Zeilenende) – keine Buchung, wird vom Aufrufer ignoriert.
        return None

    # Datum echt konvertieren ("01.03.2023" → date-Objekt)
    datum = datetime.strptime(match.group(1), "%d.%m.%Y").date()
    verwendungszweck = match.group(2).strip()

    # Betrag umwandeln: "1.234,56" → 1234.56 (Punkt entfernt, Komma → Punkt, zu Float konvertieren)
    betrag = float(match.group(3).replace(".", "").replace(",", "."))

    return datum, verwendungszweck, betrag


def speichere_excel(daten: list[tuple[date, str, float]], ausgabe_datei: str) -> None:
    """
    Speichert die extrahierten Buchungsdaten in einer Excel-Datei.

    Datum und Betrag werden als echte Excel-Typen geschrieben (Datum + Zahl
    mit festem Format), damit die Spalten in Excel sinnvoll sortiert,
    gefiltert und summiert werden können.

    Args:
        daten: Liste der Buchungen (Buchungsdatum, Verwendungszweck, Betrag)
        ausgabe_datei: Pfad zur Ausgabedatei (.xlsx)
    """
    arbeitsbuch = openpyxl.Workbook()
    tabelle = arbeitsbuch.active
    tabelle.title = "Buchungen"

    # --- Überschrift schreiben ---
    spaltennamen = ["Buchungsdatum", "Verwendungszweck", "Betrag (EUR)"]
    tabelle.append(spaltennamen)

    # Überschrift fett formatieren
    for zelle in tabelle[1]:
        zelle.font = Font(bold=True)

    # --- Buchungsdaten schreiben (mit echten Excel-Typen und Formatangeben) ---
    for datum, verwendungszweck, betrag in daten:
        tabelle.append([datum, verwendungszweck, betrag])

        aktuelle_zeile = tabelle.max_row
        datumszelle = tabelle.cell(row=aktuelle_zeile, column=1)
        datumszelle.number_format = "DD.MM.YYYY"
        betragzelle = tabelle.cell(row=aktuelle_zeile, column=3)
        betragzelle.number_format = "0.00"

    # --- Spaltenbreiten anpassen ---
    tabelle.column_dimensions["A"].width = 14  # Datum
    tabelle.column_dimensions["B"].width = 55  # Verwendungszweck
    tabelle.column_dimensions["C"].width = 16  # Betrag

    # --- Datei speichern ---
    arbeitsbuch.save(ausgabe_datei)


# ============================================================================
# HAUPTPROGRAMM
# ============================================================================


def main() -> None:
    """Hauptfunktion: Liest PDFs, extrahiert Daten und speichert Excel-Datei."""

    # --- Kommandozeilenargumente parsen ---
    parser = argparse.ArgumentParser(
        description="Extrahiert Kontobewegungen aus PDF-Kontoauszügen."
    )
    parser.add_argument("ordner", help="Pfad zum Ordner mit den PDF-Kontoauszügen")
    parser.add_argument(
        "--output",
        "-o",
        default=None,
        help=f"Name der Ausgabedatei (Standard: {DEFAULT_OUTPUT_FILENAME} im Quellordner)",
    )
    parser.add_argument(
        "--prefix",
        "-p",
        default=DEFAULT_FILE_PREFIX,
        help=f'Suchmuster für PDF-Dateinamen (Standard: "{DEFAULT_FILE_PREFIX}")',
    )

    args = parser.parse_args()

    # Ordner prüfen
    if not os.path.isdir(args.ordner):
        print(f"❌ Fehler: Ordner '{args.ordner}' existiert nicht.")
        sys.exit(1)

    # Ausgabedatei bestimmen
    if args.output:
        ausgabe_datei = os.path.abspath(args.output)
    else:
        ausgabe_datei = os.path.join(args.ordner, DEFAULT_OUTPUT_FILENAME)

    print("=" * 60)
    print("  Kontoauszug-Parser")
    print("=" * 60)
    print(f"  Quellordner:   {args.ordner}")
    print(f"  Ausgabedatei:  {ausgabe_datei}")
    print(f"  Suchmuster:    {args.prefix}")
    print("=" * 60)

    # --- PDF-Dateien suchen ---
    pdf_dateien = finde_pdf_dateien(args.ordner, args.prefix)

    if not pdf_dateien:
        print(f"\n⚠️  Keine PDF-Dateien mit Muster '{args.prefix}' gefunden.")
        sys.exit(0)

    print(f"\n📂 {len(pdf_dateien)} PDF-Datei(en) gefunden.\n")

    # --- Alle PDFs verarbeiten ---
    alle_buchungen: list[tuple[date, str, float]] = []

    for i, dateiname in enumerate(pdf_dateien, 1):
        dateipfad = os.path.join(args.ordner, dateiname)
        print(f"🔍 [{i}/{len(pdf_dateien)}] Verarbeite: {dateiname}")

        try:
            with pdfplumber.open(dateipfad) as pdf:
                for seiten_nummer, seite in enumerate(pdf.pages, 1):
                    text = seite.extract_text()

                    if not text:
                        continue  # Leere Seite überspringen

                    buchungen = extrahiere_seite(text, seiten_nummer)
                    alle_buchungen.extend(buchungen)
        except Exception as e:
            print(f"   ⚠️  Fehler beim Lesen von {dateiname}: {e}")
            continue

    # --- Ergebnis speichern ---
    print("\n" + "=" * 60)

    if not alle_buchungen:
        print("⚠️  Keine Buchungen gefunden. Excel-Datei wird nicht erstellt.")
        sys.exit(0)

    print(f"📊 Speichere {len(alle_buchungen)} Buchungen in Excel...")
    speichere_excel(alle_buchungen, ausgabe_datei)
    print(f"\n✅ Fertig! Datei gespeichert: {ausgabe_datei}")
    print("=" * 60)


if __name__ == "__main__":
    main()
