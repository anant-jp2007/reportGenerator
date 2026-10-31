"""Preview and changing values on the check screen: the files show exactly what the preview showed.

    python tests/test_preview.py

Uses the sample registers from tools/make_sample_registers.py (5 entries that can be certified).
"""
import html as htmllib
import io
import re
import sys
import tempfile
import zipfile
from datetime import date
from pathlib import Path

import docx
import openpyxl
import pymupdf

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tools"))
from app import web  # noqa: E402
from app.production import find_new_rows  # noqa: E402
from app.report import apply_values, build_report, editable_cells  # noqa: E402
from app.settings import load_settings, with_company  # noqa: E402
from make_sample_registers import make  # noqa: E402

TMP = Path(tempfile.mkdtemp(prefix="preview_"))
web.OUTPUT, web.UPLOADS, web.LEDGER = TMP / "output", TMP / "uploads", TMP / "ledger"
client = web.app.test_client()
failures = 0


def check(name, cond, detail=""):
    global failures
    failures += not cond
    print(("PASS " if cond else "FAIL ") + name + (f"  [{detail}]" if detail else ""))


def inputs(fragment: str) -> dict[str, tuple[str, str]]:
    """cell -> (kind, value) of the input boxes in a preview."""
    return {m.group(1): (m.group(2), htmllib.unescape(m.group(3)))
            for m in re.finditer(r'data-cell="(\w+)" data-kind="(\w+)" value="([^"]*)"', fragment)}


before, after = make(TMP / "registers")
settings = with_company(load_settings(), "balaji")

# ---------- which values can be changed ----------
kinds = editable_cells(settings)
hc = settings["header_cells"]
check("the 3 dates and the batch can be changed",
      [kinds[hc[k]] for k in ("date_of_manufacturing", "testing_start", "testing_end", "batch")] == ["date", "date", "date", "batch"])
tests = sorted(r for r, k in kinds.items() if k == "test")
check("every test value obtained can be changed (15 cells)", tests == sorted(
    ["J17", "J18", "G20", "G23", "J24", "J25", "J26", "J27", "G28", "G30", "G31", "G33", "G35", "G36", "G37"]), tests)
check("headers, fixed texts and chart limits can't", not {"A1", "A5", "E24", "G14", hc["grade"], hc["pn"]} & kinds.keys())

# ---------- applying typed values ----------
row = find_new_rows(before, after).new[0]
rep = build_report(row, None, settings)
apply_values(rep, {"J24": "63.05", "G23": "1229", "D8": "01-05-2025", "G30": "N/A", "G31": "", "A1": "HACKED",
                   "D11": "BATCH 7"}, settings)
c = rep.cells
check("decimal number: kept as a number, printed as typed", c["J24"].value == 63.05 and c["J24"].text == "63.05")
check("whole number stays whole", c["G23"].value == 1229 and isinstance(c["G23"].value, int))
check("date read as a date, printed dd-mm-yyyy", c["D8"].value == date(2025, 5, 1) and c["D8"].text == "01-05-2025")
check("text in a test cell printed as typed", c["G30"].value == "N/A" and c["G30"].text == "N/A")
check("emptied value prints blank", c["G31"].value is None and c["G31"].text == "")
check("a cell that can't be changed is ignored", c["A1"].text == "BALAJI POLYPLAST", c["A1"].text)
check("batch printed as typed", c["D11"].text == "BATCH 7")
rep2 = build_report(row, None, settings)
same = dict(rep2.cells)
apply_values(rep2, {r: same[r].text for r in kinds}, settings)
check("values sent back unchanged keep their exact cell (number, date)", all(rep2.cells[r] == same[r] for r in kinds))

# ---------- the preview ----------
html = htmllib.unescape(client.post("/check", data={
    "company": "balaji", "old": (io.BytesIO(before.read_bytes()), before.name),
    "new": (io.BytesIO(after.read_bytes()), after.name)}, content_type="multipart/form-data").get_data(as_text=True))
token = html.split('name="token" value="')[1].split('"')[0]
ids = [s.split('"')[0] for s in html.split('name="row" value="')[1:]]
check("check screen: a View button for every entry", html.count('class="view"') == len(ids) == 5)
check("check screen: the preview window is there, closed", '<dialog class="viewer" id="viewer"' in html)

r = client.get("/preview", query_string={"token": token, "row": ids[0], "no": "BPP63/P1"})
frag = r.get_data(as_text=True)
boxes = inputs(frag)
check("preview: a certificate with the company header", r.status_code == 200 and 'class="cert"' in frag
      and "BALAJI POLYPLAST" in frag)
check("preview: an input box for each value that can be changed, plus the Report No.",
      set(boxes) == set(kinds) | {hc["report_no"]} and boxes[hc["report_no"]] == ("number", "BPP63/P1"), sorted(boxes))
check("preview: the register's values in the boxes", boxes[hc["date_of_manufacturing"]][1] == "30-04-2025"
      and boxes[hc["batch"]][1] == "2025 04 30 001 1", boxes[hc["date_of_manufacturing"]])
check("preview: lines broken as in the PDF", "Colour of pipe shall be black<br>with blue identification strips." in frag)
check("preview: header and chart limits are plain text", 'data-cell="A1"' not in frag and 'data-cell="E24"' not in frag)
again = inputs(client.get("/preview", query_string={"token": token, "row": ids[0]}).get_data(as_text=True))
check("preview again: new random values (the page keeps the ones shown)",
      any(again[t][1] != boxes[t][1] for t in tests))
r = client.get("/preview", query_string={"token": token, "row": "PE63-2025-999-1"})
check("preview of an entry not in the register: clear message", r.status_code == 400)
r = client.get("/preview", query_string={"token": "nothere", "row": ids[0]})
check("preview after the files expired: asks to choose again", r.status_code == 410)

# ---------- download: exactly what the preview showed, changes included ----------
shown = {cell: v for cell, (kind, v) in boxes.items() if kind != "number"}
changed = {**shown, "J24": "63.05", "D8": "01-05-2025", "G36": "580"}
form = {"token": token, "row": ids[:2], "files": "each", "format_choice": "1", "format": ["excel", "pdf", "word"],
        "report_choice": "1", "report": ["1", "2"], f"no-{ids[0]}": "BPP63/P1"}
form.update({f"v-{ids[0]}-{cell}": v for cell, v in changed.items()})
r = client.post("/download", data=form)
z = zipfile.ZipFile(io.BytesIO(r.data))
xl = sorted(n for n in z.namelist() if n.startswith("PE 63/") and n.endswith(".xlsx"))[0]
ws = openpyxl.load_workbook(io.BytesIO(z.read(xl))).active
printed = {}
for cell in shown:
    v = ws[cell].value
    printed[cell] = (v.strftime("%d-%m-%Y") if hasattr(v, "strftime") else
                     "" if v is None else f"{v:.{len(ws[cell].number_format.split('.')[1])}f}"
                     if isinstance(v, float) and "." in ws[cell].number_format else str(v))
check("Excel: every value exactly as in the preview, changes included",
      printed == changed, {k: (printed[k], changed[k]) for k in changed if printed.get(k) != changed[k]})
check("Excel: a changed number is a number, a changed date a date", ws["J24"].value == 63.05
      and ws["J24"].number_format == "0.00" and ws["D8"].value.date() == date(2025, 5, 1))
check("Excel: Report No. typed in the preview", ws[hc["report_no"]].value == "BPP63/P1")
words = " ".join(pymupdf.open(stream=z.read(xl.replace(".xlsx", ".pdf")), filetype="pdf")[0].get_text().split())
check("PDF: changed values printed", all(t in words for t in ("63.05", "01-05-2025", "580", "BPP63/P1")))
wtext = " ".join(cell.text for t in docx.Document(io.BytesIO(z.read(xl.replace(".xlsx", ".docx")))).tables
                 for row_ in t.rows for cell in row_.cells)
check("Word: changed values printed", all(t in wtext for t in ("63.05", "01-05-2025", "580")))
r2 = openpyxl.load_workbook(io.BytesIO(z.read("Report 2.xlsx"))).active
rows2 = {src: int(rw) for rw, src in settings["report2"]["rows"].items() if not rw.startswith("_")}
check("Report 2: the same changed values", r2[f"D{rows2['J24']}"].value == 63.05 and r2[f"D{rows2['G36']}"].value == 580,
      (r2[f"D{rows2['J24']}"].value, r2[f"D{rows2['G36']}"].value))
other = [n for n in z.namelist() if n.endswith(".xlsx") and n != xl and "Report 2" not in n]
check("an entry never previewed is made as before (random values)", len(other) == 1
      and openpyxl.load_workbook(io.BytesIO(z.read(other[0]))).active["J24"].value is not None)

# a company with a logo and a stamp: both drawn in the preview
html = htmllib.unescape(client.post("/check", data={
    "company": "abhay", "old": (io.BytesIO(before.read_bytes()), before.name),
    "new": (io.BytesIO(after.read_bytes()), after.name)}, content_type="multipart/form-data").get_data(as_text=True))
token = html.split('name="token" value="')[1].split('"')[0]
rid = html.split('name="row" value="')[1].split('"')[0]
frag = client.get("/preview", query_string={"token": token, "row": rid}).get_data(as_text=True)
check("Abhay preview: its header, logo and stamp", "ABHAY POLYPLAST" in frag and frag.count('class="cert-pic"') == 2)

print(f"\n(files in {TMP})")
print("ALL PASS" if not failures else f"{failures} FAILED")
sys.exit(1 if failures else 0)
