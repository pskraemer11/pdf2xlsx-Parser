## README Bank Statement Parser pdf2xlsx - May 2026

Python script for extracting transaction history from PDF bank statements.

This Python script was developed to convert multi-page account statements into Excel files, which can then be imported into Excel or OpenOffice Calc for further processing.

No Adobe installation is required.

## Prerequisites
- Python 3.10+
- `pdfplumber` and `openpyxl`

## Installing Dependencies
```bash
pip install pdfplumber openpyxl
```

## How the parser works

A single transaction in the PDF spans several lines ("block"), but only the first line is relevant, because only it carries the amount in the right column. 
A line is extracted as a transaction when **both** conditions hold (logical AND):

1. the line starts with a date (`DD.MM.YYYY`), **and**
2. the line ends with a decimal amount with two decimal places (optionally with thousands separators, a leading `+`/`−` and footnote characters).

Continuation lines of a transaction (further purpose text) and lines without an amount (e.g. the calculated balance) are ignored.

The recognition parameters are configuration constants at the top of the script (`HEADER_MARKER`, `BUCHUNG_REGEX`), so minor layout changes can be adapted there. 
A fully structured extraction using the table geometry of the PDF is a possible future improvement.

## Standard Usage
```bash
python pdf2xlsx-Parser.py /path/to/pdf-folder

# With a custom output name
python pdf2xlsx-Parser.py /path/to/pdf-folder --output my_result.xlsx

# With a custom search pattern
python pdf2xlsx-Parser.py /path/to/pdf-folder --prefix "Bank_Statement_*"
```

## Testing
The parser functions are covered by unit tests that run without real PDF files (synthetic text only). No extra test framework is required (stdlib `unittest`):

```bash
python -m unittest discover -s tests -v
```
