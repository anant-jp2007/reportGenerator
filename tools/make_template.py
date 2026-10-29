"""Build templates/test_report.xlsx from the sample test report workbook.

Takes one size sheet (all 8 sheets share the same layout), removes the other sheets,
empties every cell that the app fills per report, fixes the row 21 typo and sets
one-page A4 printing. Formatting (fonts, borders, merges, superscripts) is kept.

    python tools/make_template.py [sample.xlsx] [output.xlsx]
"""
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
SRC = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "data/source/test_report_sample.xlsx"
OUT = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "templates/test_report.xlsx"
SOURCE_SHEET = "315"

# Cells the app writes for every report (header, spec limits from the chart, test values).
VARIABLE_CELLS = [
    "D7", "D8", "D9", "D10", "D11",          # report no., dates, batch
    "J7", "J8", "J9", "J10",                 # grade, PN, DN, SDR
    "J17", "J18",                            # stripe width, depth
    "E20", "G20",                            # pipe end: spec max, value
    "G21", "G23",                            # length, coiling: left empty
    "E24", "E25", "J24", "J25",              # outside diameter
    "E26", "E27", "J26", "J27",              # wall thickness
    "E28", "G28",                            # ovality
    "G30", "G31", "G33", "G35", "G36", "G37",  # reversion, CBC, MFR, density, elongation, yield
]


def main():
    wb = openpyxl.load_workbook(SRC, rich_text=True)  # rich_text keeps the superscript 2 / 3
    for ws in list(wb.worksheets):
        if ws.title != SOURCE_SHEET:
            wb.remove(ws)
    ws = wb[SOURCE_SHEET]
    ws.title = "Test Report"

    for ref in VARIABLE_CELLS:
        ws[ref].value = None
    ws["D21"].value = "As Declared by Manufacturer"
    ws["G31"].number_format = "0.00"  # CBC was "General" in the sample; keep 2 decimals like 2.50
    # PN / DN / SDR were "0", which would print SDR 13.6 as 14 and PN 12.5 as 13
    for ref in ("J8", "J9", "J10"):
        ws[ref].number_format = "General"
    ws["D11"].number_format = "@"  # batch is text; the sample had a date format here

    ws.print_area = "A1:L40"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "portrait"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"template written: {OUT}")


if __name__ == "__main__":
    main()
