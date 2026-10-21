"""Unit-Tests für die Parser-Funktionen – ohne echte PDFs (nur synthetische Texte).

Ausführen:
    python -m unittest discover -s tests -v

Hinweis: Das Modul "pdf2xlsx-Parser.py" enthält einen Bindestrich und kann
deshalb nicht per import-Statement geladen werden; es wird über importlib
nach Dateipfad geladen.
"""

import importlib.util
import tempfile
import unittest
from datetime import date
from pathlib import Path

import openpyxl

MODUL_PFAD = Path(__file__).resolve().parent.parent / "pdf2xlsx-Parser.py"


def lade_modul():
    spec = importlib.util.spec_from_file_location("pdf2xlsx_parser", MODUL_PFAD)
    modul = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modul)
    return modul


modul = lade_modul()


class BuchungszeilenTests(unittest.TestCase):
    """Einzelne Buchungszeile: Datum, Verwendungszweck, Betrag."""

    def test_einfache_buchung(self):
        zeile = "05.03.2023 Dauerauftrag Strom 300,00"
        self.assertEqual(
            modul.verarbeite_zeile(zeile),
            (date(2023, 3, 5), "Dauerauftrag Strom", 300.0),
        )

    def test_negativer_betrag_mit_tausendertrennzeichen(self):
        zeile = "12.01.2023 Kartenzahlung Supermarkt -1.234,56"
        ergebnis = modul.verarbeite_zeile(zeile)
        self.assertEqual(ergebnis[0], date(2023, 1, 12))
        self.assertEqual(ergebnis[1], "Kartenzahlung Supermarkt")
        self.assertEqual(ergebnis[2], -1234.56)

    def test_positiver_betrag_mit_vorzeichen(self):
        zeile = "02.02.2023 Dauerauftrag +1.234,56"
        self.assertEqual(modul.verarbeite_zeile(zeile)[2], 1234.56)

    def test_fussnote_direkt_am_betrag(self):
        zeile = "05.04.2023 Dauerauftrag 300,00¹"
        self.assertEqual(modul.verarbeite_zeile(zeile)[2], 300.0)

    def test_fussnote_mit_leerzeichen(self):
        zeile = "05.04.2023 Dauerauftrag 300,00 1"
        self.assertEqual(modul.verarbeite_zeile(zeile)[2], 300.0)

    def test_kein_match_ohne_datum(self):
        self.assertIsNone(modul.verarbeite_zeile("Dauerauftrag 300,00"))

    def test_kein_match_ohne_betrag(self):
        # Valuta-Zeilen bzw. Zweck-Fortführungen haben kein "Betrag am Ende"
        self.assertIsNone(modul.verarbeite_zeile("05.04.2023 Beleg ohne Betrag"))


class SeitenTests(unittest.TestCase):
    """Seitenverarbeitung: Überschrift finden, zeilengenaue Buchungserkennung.

    Semantik (Layout der Auszüge): Nur die erste Blockzeile hat den Betrag in
    der rechten Spalte. Alle übrigen Blockzeilen (Valuta-Datum, Zweck-Fortführung)
    werden ignoriert, ohne als Fehler zu gelten – pdfplumber liefert zwischen
    den Blöcken keine Leerzeilen.
    """

    STANDARD_KOPF = "Buchung     Buchung / Verwendungszweck     Betrag (EUR)"

    def test_standardseite_mit_valuta(self):
        text = "\n".join(
            [
                "Kunde Beispiel",
                "Kontoinformationen Beispiel",
                self.STANDARD_KOPF,
                "Valuta",
                "01.03.2023 Dauerauftrag 300,00",
                "",
                "03.03.2023 Kartenzahlung -35,99",
                "",
            ]
        )
        buchungen = modul.extrahiere_seite(text, seiten_nummer=1)
        self.assertEqual(len(buchungen), 2)
        self.assertEqual(buchungen[0], (date(2023, 3, 1), "Dauerauftrag", 300.0))
        self.assertEqual(buchungen[1], (date(2023, 3, 3), "Kartenzahlung", -35.99))

    def test_seite_ohne_valuta_zeile(self):
        # Regression: Ohne "Valuta"-Zeile darf die erste Buchung nicht verloren gehen
        text = "\n".join(
            [
                self.STANDARD_KOPF,
                "01.03.2023 Dauerauftrag 300,00",
                "",
                "03.03.2023 Übertragung -35,99",
            ]
        )
        buchungen = modul.extrahiere_seite(text, seiten_nummer=1)
        self.assertEqual(len(buchungen), 2)

    def test_seite_ohne_kopfzeile(self):
        self.assertEqual(modul.extrahiere_seite("Saldo 12.02.2023 1.503,55", 1), [])

    def test_fortfuehrungszeilen_werden_ignoriert(self):
        # Mehrzeilige Buchung: nur die erste Zeile (Betrag am Ende) zählt.
        # Fortführungszeilen ohne Datum am Anfang sind kein Fehler.
        text = "\n".join(
            [
                self.STANDARD_KOPF,
                "Valuta",
                "01.03.2023 Sepa-Ueberweisung 300,00",
                "Mandatsreferenz 4711",
                "Verwendungszweck Zeile 2",
                "",
                "03.03.2023 Sonstiges -35,99",
                "",
            ]
        )
        buchungen = modul.extrahiere_seite(text, seiten_nummer=1)
        self.assertEqual(buchungen[0], (date(2023, 3, 1), "Sepa-Ueberweisung", 300.0))
        self.assertEqual(len(buchungen), 2)

    def test_zeile_ohne_datum_ist_keine_buchung(self):
        # Kein Datum am Zeilenanfang → keine Buchung
        text = "\n".join(
            [
                self.STANDARD_KOPF,
                "Valuta",
                "01.03.2023 Dauerauftrag 300,00",
                "",
                "Ungültige Zeile ohne Datum",
                "",
                "03.03.2023 Übertragung -35,99",
                "",
            ]
        )
        buchungen = modul.extrahiere_seite(text, seiten_nummer=1)
        self.assertEqual(len(buchungen), 2)

    def test_nur_datum_ist_valuta_und_wird_ignoriert(self):
        # Reine "Valuta"-Zeile (nur Datum, ohne Text/Betrag): im Auszugs-Layout
        # steht das Valuta-Datum in der linken Spalte unter dem Buchungsdatum.
        text = "\n".join(
            [
                self.STANDARD_KOPF,
                "01.03.2023 Dauerauftrag 300,00",
                "02.03.2023",
                "",
                "03.03.2023 Sonstiges -35,99",
                "",
            ]
        )
        buchungen = modul.extrahiere_seite(text, seiten_nummer=1)
        self.assertEqual(len(buchungen), 2)

    def test_datum_ohne_betrag_wird_ignoriert(self):
        # Datum am Zeilenanfang, aber kein Betrag am Zeilenende: Das ist die
        # Fortführungszeile mit Valuta-Datum (kein Fehler, keine Buchung).
        text = "\n".join(
            [
                self.STANDARD_KOPF,
                "Valuta",
                "01.03.2023 Sepa-Überweisung 300,00",
                "02.03.2023 Fortfuehrung/Laufende-Anmerkung:00:00",
                "",
                "03.03.2023 Sonstiges -35,99",
                "",
            ]
        )
        buchungen = modul.extrahiere_seite(text, seiten_nummer=1)
        self.assertEqual(len(buchungen), 2)

    def test_rechnerischer_saldo_ohne_datum_wird_nicht_gezaehlt(self):
        # Die rechnerische Zeile unter der letzten Buchung beginnt nicht mit
        # einem Datum und erhält keinen Betrag am Zeilenende im Muster …
        text = "\n".join(
            [
                self.STANDARD_KOPF,
                "Valuta",
                "01.03.2023 Dauerauftrag 300,00",
                "05.03.2023 Sonstiges -35,99",
                "Rechnerischer Saldo der Buchungen",
                "",
            ]
        )
        buchungen = modul.extrahiere_seite(text, seiten_nummer=1)
        self.assertEqual(len(buchungen), 2)

    def test_block_ohne_letzte_leerzeile(self):
        text = "\n".join([self.STANDARD_KOPF, "Valuta", "01.03.2023 Dauerauftrag 300,00"])
        buchungen = modul.extrahiere_seite(text, seiten_nummer=1)
        self.assertEqual(len(buchungen), 1)


class DateisucheTests(unittest.TestCase):
    """--prefix-Suche im Ordner."""

    def test_oder_mit_stern(self):
        with tempfile.TemporaryDirectory() as tmp:
            ordner = Path(tmp)
            (ordner / "Girokonto_111_Kontoauszug_20230131.pdf").touch()
            (ordner / "Girokonto_222_Kontoauszug_20230228.pdf").touch()
            (ordner / "Girokonto_333_Uebersicht.pdf").touch()
            (ordner / "Firmenkonto_111_Kontoauszug_20230131.pdf").touch()
            (ordner / "Girokonto_111_Kontoauszug_20230131.xlsx").touch()

            ergebnis = modul.finde_pdf_dateien(str(ordner), "Girokonto_*_Kontoauszug_*.pdf")
            self.assertEqual(
                ergebnis,
                [
                    "Girokonto_111_Kontoauszug_20230131.pdf",
                    "Girokonto_222_Kontoauszug_20230228.pdf",
                ],
            )

    def test_ohne_stern_bedeutet_prefix(self):
        with tempfile.TemporaryDirectory() as tmp:
            ordner = Path(tmp)
            (ordner / "Girokonto_111_Kontoauszug_20230131.pdf").touch()
            (ordner / "Anderes_Kontoauszug.pdf").touch()

            ergebnis = modul.finde_pdf_dateien(str(ordner), "Girokonto_")
            self.assertEqual(ergebnis, ["Girokonto_111_Kontoauszug_20230131.pdf"])


class ExcelTests(unittest.TestCase):
    """Excel-Ausgabe: echte Typen und Zahlenformate."""

    def test_excel_typen_und_formate(self):
        daten = [
            (date(2023, 3, 1), "Dauerauftrag", 300.0),
            (date(2023, 3, 3), "Kartenzahlung", -35.99),
        ]
        with tempfile.TemporaryDirectory() as tmp:
            ausgabe = Path(tmp) / "ergebnis.xlsx"
            modul.speichere_excel(daten, str(ausgabe))

            arbeitsbuch = openpyxl.load_workbook(ausgabe)
            tabelle = arbeitsbuch["Buchungen"]
            self.assertEqual(tabelle.max_row, 3)
            self.assertIsInstance(tabelle["A2"].value, date)
            self.assertEqual(tabelle["A2"].number_format, "DD.MM.YYYY")
            self.assertEqual(tabelle["C3"].value, -35.99)
            self.assertEqual(tabelle["C3"].number_format, "0.00")
            self.assertTrue(tabelle["A1"].font.bold)


if __name__ == "__main__":
    unittest.main()
