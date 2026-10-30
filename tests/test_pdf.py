"""PDF output: every certificate value is printed, on one A4 page, in the template's fonts.

    python tests/test_pdf.py [production.xls]

Exact positions were matched against Excel's own PDF export (median difference 0.2 pt over
all 174 text pieces); this test guards the contents and page setup so it stays that way.
"""
import sys
import tempfile
from pathlib import Path

import pymupdf

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from app.production import read_production  # noqa: E402
from app.render_pdf import write_pdf, write_pdf_combined  # noqa: E402
from app.report import build_reports  # noqa: E402

REAL = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "data/source/production_sample.xls"
TMP = Path(tempfile.mkdtemp(prefix="pdftest_"))
failures = 0


def check(name, cond, detail=""):
    global failures
    failures += not cond
    print(("PASS " if cond else "FAIL ") + name + (f"  [{detail}]" if detail else ""))


def plain(t):
    return t.replace("^{", "").replace("}", "")


rows = read_production(REAL).rows
picked = [r for r in rows if r.ok][::40] + [r for r in rows if r.ok and (not r.batch or r.sdr is None)][:3]
reports = build_reports(picked)

missing, pages, sizes, fonts = [], [], set(), set()
for rep in reports:
    doc = pymupdf.open(write_pdf(rep, TMP / f"{rep.file_name}.pdf"))
    pages.append(len(doc))
    page = doc[0]
    sizes.add((round(page.rect.width), round(page.rect.height)))
    words = " ".join(page.get_text().split())
    for ref, cell in rep.cells.items():
        t = " ".join(plain(cell.text).split())
        if t and t not in words:
            # wrapped texts break across lines; every word must still be there
            if not all(w in words for w in t.split()):
                missing.append((rep.file_name, ref, t))
    fonts |= {f[3] for f in page.get_fonts()}
    if page.get_images():
        missing.append((rep.file_name, "Balaji has no pictures (signed by hand)", "has an image"))

check(f"{len(reports)} PDFs: every value of the certificate is printed", not missing, missing[:5])
check("one page each", set(pages) == {1}, pages)
check("A4 (595 x 842 pt)", sizes == {(595, 842)}, sizes)
check("template fonts embedded (Times New Roman, Cambria)",
      any("Times" in f for f in fonts) and any("Cambria" in f for f in fonts), fonts)

doc = pymupdf.open(write_pdf_combined(reports, TMP / "All.pdf"))
check("combined PDF: one page per certificate, in order",
      len(doc) == len(reports) and all(f"{r.cells['D8'].text}" in doc[i].get_text() for i, r in enumerate(reports)))

rep = next(r for r in reports if r.cells["J10"].text == "13.6")
text = pymupdf.open(TMP / f"{rep.file_name}.pdf")[0].get_text()
check("SDR 13.6 printed as 13.6", "13.6" in text)
check("superscripts drawn small and raised (Kgs/Cm², kg/m³)",
      any(s["text"].strip() in ("2", "3") and s["size"] < 9
          for b in pymupdf.open(TMP / f"{rep.file_name}.pdf")[0].get_text("dict")["blocks"]
          for ln in b.get("lines", []) for s in ln["spans"]))

print(f"\n(files in {TMP})")
print("ALL PASS" if not failures else f"{failures} FAILED")
sys.exit(1 if failures else 0)
