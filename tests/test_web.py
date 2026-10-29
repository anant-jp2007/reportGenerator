"""Website: choose files -> check entries -> download, through Flask's test client.

    python tests/test_web.py [production.xls]
"""
import html as htmllib
import io
import sys
import tempfile
import zipfile
from pathlib import Path

import openpyxl
import xlrd

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from app import web  # noqa: E402

REAL = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "data/source/production_sample.xls"
TMP = Path(tempfile.mkdtemp(prefix="webtest_"))
web.OUTPUT = TMP / "output"
web.UPLOADS = TMP / "uploads"
client = web.app.test_client()


def text(resp):
    return htmllib.unescape(resp.get_data(as_text=True))

failures = 0


def check(name, cond, detail=""):
    global failures
    failures += not cond
    print(("PASS " if cond else "FAIL ") + name + (f"  [{detail}]" if detail else ""))


def old_copy(blank: dict) -> bytes:
    """The real register with some rows reset to empty (only Sr. No. and Report No. left)."""
    src = xlrd.open_workbook(str(REAL))
    wb = openpyxl.Workbook(write_only=True)
    for sh in src.sheets():
        ws = wb.create_sheet(sh.name)
        for r in range(sh.nrows):
            vals = []
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                v = xlrd.xldate.xldate_as_datetime(cell.value, src.datemode) if cell.ctype == 3 else (cell.value or None)
                if (r + 1) in blank.get(sh.name, []) and c not in (0, 3):
                    v = None
                vals.append(v)
            ws.append(vals)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def upload(old: bytes, new: bytes, old_name="old.xlsx", new_name="PRODUCTION.xls"):
    return client.post("/check", data={"old": (io.BytesIO(old), old_name), "new": (io.BytesIO(new), new_name)},
                       content_type="multipart/form-data")


real = REAL.read_bytes()

r = client.get("/")
html = text(r)
check("start page opens with both file boxes", r.status_code == 200 and "Old register" in html and "New register" in html)

r = client.post("/check", data={"new": (io.BytesIO(real), "PRODUCTION.xls")}, content_type="multipart/form-data")
check("missing old register: clear message", r.status_code == 400 and "Choose the old register" in text(r))

r = upload(b"hello", real, old_name="notes.txt")
check("not an Excel file: clear message", r.status_code == 400 and "isn't an Excel file" in text(r))

r = upload(b"not really excel", real, old_name="broken.xlsx")
check("broken Excel file: clear message", r.status_code == 400 and "couldn't be opened" in text(r))

# PE 63 Sr 171, PE 100 Sr 15 (110 x 6, no batch) and Sr 18, 19 are new; PE 63 row 343 (40 X 10) is new but can't be made
old = old_copy({"Production 63": [368, 343], "Production 100": [49, 52, 53]})
r = upload(old, real)
html = text(r)
check("review lists 4 new entries", r.status_code == 200 and "4 new entries to certify" in html, html[html.find("<h1"):][:80])
check("grouped by grade, PE 63 before PE 100", 0 < html.find(">PE 63<") < html.find(">PE 100<"))
check("report numbers pre-filled 1-4", all(f'value="{n}"' in html for n in (1, 2, 3, 4)))
check("warnings shown (SDR taken from PN, empty batch)", "SDR not written" in html and "Batch number is empty" in html)
check("skipped entry shown with reason", "No certificate for 1 entry" in html and "PE 63 has no PN 10" in html)
token = html.split('name="token" value="')[1].split('"')[0]
ids = [s.split('"')[0] for s in html.split('name="row" value="')[1:]]

form = {"token": token, "row": ids, "files": "each"}
form.update({f"no-{ids[0]}": "BPP63/EX001", f"no-{ids[1]}": "7"})
r = client.post("/download", data=form)
check("download returns a ZIP", r.status_code == 200 and r.mimetype == "application/zip", r.status_code)
z = zipfile.ZipFile(io.BytesIO(r.data))
names = sorted(z.namelist())
check("one Excel file per certificate, named from Report No. column",
      names == sorted(["MP63-171.xlsx", "MP100-015.xlsx", "MP100-018.xlsx", "MP100-019.xlsx"]), names)
d7 = openpyxl.load_workbook(io.BytesIO(z.read("MP63-171.xlsx"))).active["D7"].value
check("typed Report No. 'BPP63/EX001' printed on the certificate", d7 == "BPP63/EX001", d7)
d7 = openpyxl.load_workbook(io.BytesIO(z.read("MP100-015.xlsx"))).active["D7"].value
check("typed Report No. 7 kept as a number", d7 == 7, d7)
check("a copy is saved in the output folder", r.headers["X-Certificates"] == "4" and
      len(list((web.OUTPUT).glob("*/*.xlsx"))) == 4, r.headers.get("X-Saved-To"))

r = client.post("/download", data={"token": token, "row": ids[2:], "files": "one"})
z = zipfile.ZipFile(io.BytesIO(r.data))
check("only ticked entries, all in one file", len(z.namelist()) == 1 and
      len(openpyxl.load_workbook(io.BytesIO(z.read(z.namelist()[0]))).sheetnames) == 2, z.namelist())

r = client.post("/download", data={"token": token, "files": "each"})
check("nothing ticked: clear message", r.status_code == 400 and "Tick at least one entry" in text(r))
r = client.post("/download", data={"token": "abc123", "row": ids, "files": "each"})
check("expired files: asks to choose again", r.status_code == 410 and "Choose the registers again" in text(r))

r = upload(real, real, old_name="PRODUCTION old.xls")
check("same file twice: 'No new entries' with a hint about swapped files",
      "No new entries" in text(r) and "swapped" in text(r))

print(f"\n(files in {TMP})")
print("ALL PASS" if not failures else f"{failures} FAILED")
sys.exit(1 if failures else 0)
