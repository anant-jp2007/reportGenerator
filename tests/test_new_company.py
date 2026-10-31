"""Adding a company from one of its certificates: the header read back, the page still on one A4,
and the website's screens (the new company in the company menu).

    python tests/test_new_company.py

Every existing company's own certificate is used as the sample, so the header read from it must be
the one in settings.json. config/settings.json is changed by the save test and put back afterwards.
"""
import html as htmllib
import io
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from app import new_company as nc, web  # noqa: E402
from app.layout import load_layout  # noqa: E402
from app.render_excel import write_excel  # noqa: E402
from app.report import build_report  # noqa: E402
from app.settings import SETTINGS_FILE, companies, load_settings, with_company  # noqa: E402

TMP = Path(tempfile.mkdtemp(prefix="newcompany_"))
web.OUTPUT, web.UPLOADS, web.LEDGER = TMP / "output", TMP / "uploads", TMP / "ledger"
client = web.app.test_client()
failures = 0


def check(name, cond, detail=""):
    global failures
    failures += not cond
    print(("PASS " if cond else "FAIL ") + name + (f"  [{detail}]" if detail else ""))


def same(a, b):
    return " ".join(str(a).split()) == " ".join(str(b).split())


def to_xls(xlsx: Path) -> Path | None:
    """The same certificate saved by Excel as .xls (None when Excel isn't on this PC)."""
    out = xlsx.with_suffix(".xls")
    ps = (f"$x = New-Object -ComObject Excel.Application; $x.DisplayAlerts = $false; "
          f"$w = $x.Workbooks.Open('{xlsx}'); $w.SaveAs('{out}', 56); $w.Close($false); $x.Quit()")
    try:
        subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True, timeout=120)
    except Exception:
        return None
    return out if out.is_file() else None


s = load_settings()
row = nc.sample_row()
room = 842 - 15.5905

# ---------- every company read back from its own certificate ----------
samples = {}
for cid, name in companies(s):
    sc = with_company(s, cid)
    profile = sc["company"]["profiles"][cid]
    cert = write_excel(build_report(row, None, sc), TMP / f"{cid}.xlsx")
    samples[cid] = cert
    found = nc.read_sample(cert)
    draft = nc.make_draft(name, found, has_stamp=bool(found.stamp))
    first = "C" if any(i["cell"] == "A1" and "logo" in i["file"] for i in profile.get("images", [])) else "A"
    wanted = [profile.get(f"{first}{r}", profile.get(f"A{r}")) for r in (1, 2, 3)] + [profile.get("A4"), profile.get("A5")]
    rows_ok = [w is None or same(w, t) for w, t in zip(wanted, draft.texts)]
    check(f"{name}: the five header rows read back", all(rows_ok),
          "; ".join(f"row {i + 1}: {draft.texts[i]!r} vs {wanted[i]!r}" for i, ok in enumerate(rows_ok) if not ok))
    has_logo = any("logo" in i["file"] for i in profile.get("images", []))
    has_stamp = any("stamp" in i["file"] for i in profile.get("images", []))
    check(f"{name}: logo {'found' if has_logo else 'none'}, stamp {'found' if has_stamp else 'none'}",
          bool(found.logo) == has_logo and bool(found.stamp) == has_stamp)
    check(f"{name}: the phone sign is kept (not a Wingdings letter)", ("✆" in draft.texts[3]) == ("✆" in str(wanted[3])))

    ratio = None
    if found.logo:
        from PIL import Image
        w, h = Image.open(io.BytesIO(found.logo)).size
        ratio = w / h
    prof = nc.make_profile(draft, "x" + cid, ratio, bool(found.stamp))
    lay = load_layout(nc.build_template(prof, TMP / f"x{cid}.xlsx"))
    check(f"{name}: the new template fits A4 at 100% ({lay.height:.1f} of {room - 1:.1f} pt)", lay.height <= room - 1)
    table_same = lay.rows[5:37] == load_layout(ROOT / profile["template"]).rows[5:37] and len(lay.cols) == 12
    check(f"{name}: the table rows are unchanged", table_same)

# ---------- an .xls certificate (as the company's own may be) ----------
xls = to_xls(samples["avtar"])
if xls:
    found = nc.read_sample(xls)
    draft = nc.make_draft("Avtar", found)
    check("an .xls certificate: the same header as the .xlsx",
          draft.texts == nc.make_draft("Avtar", nc.read_sample(samples["avtar"])).texts, str(draft.texts))
else:
    print("(skipped: Excel isn't on this PC to save an .xls sample)")

# ---------- things that aren't a certificate ----------
bad = TMP / "notes.xlsx"
import openpyxl  # noqa: E402
wb = openpyxl.Workbook()
wb.active["A1"] = "Shopping list"
wb.save(bad)
try:
    nc.read_sample(bad)
    check("a workbook without 'Test Report No.' is refused", False)
except nc.CompanyError as e:
    check("a workbook without 'Test Report No.' is refused", "Report No" in str(e), str(e))
(TMP / "fake.xlsx").write_bytes(b"not a workbook")
try:
    nc.read_sample(TMP / "fake.xlsx")
    check("a damaged file is refused in plain words", False)
except nc.CompanyError as e:
    check("a damaged file is refused in plain words", "couldn't be opened" in str(e), str(e))

# a header far too big for the page: refused, nothing saved
huge = nc.Draft("Huge Pipes", ["HUGE " * 60] * 5, [30] * 5, [True] * 5)
before = SETTINGS_FILE.read_bytes()
try:
    nc.save_company(huge, None, None)
    check("a header that can't fit is refused", False)
except nc.CompanyError as e:
    check("a header that can't fit is refused", "room for" in str(e), str(e))
check("... and settings.json is untouched", SETTINGS_FILE.read_bytes() == before)
check("... and no template is left behind", not (ROOT / "templates" / "test_report_huge.xlsx").exists())

# ---------- the website ----------
page = client.get("/").get_data(as_text=True)
check("start page: an 'Add a company' link", 'href="/company"' in page)
page = client.get("/company")
check("add-company page opens", page.status_code == 200 and 'name="sample"' in page.get_data(as_text=True))


def post(name, file, **extra):
    data = {"name": name, "sample": (io.BytesIO(file.read_bytes()), file.name), **extra}
    return client.post("/company", data=data, content_type="multipart/form-data")


def text(r):
    return htmllib.unescape(r.get_data(as_text=True))


r = post("", samples["avtar"])
check("no name: asked for", r.status_code == 400 and "Type the company" in r.get_data(as_text=True))
existing = dict(companies(s))["avtar"]
r = post(existing, samples["avtar"])
check("a name already in the menu is refused", r.status_code == 400 and "already a company" in r.get_data(as_text=True))
r = client.post("/company", data={"name": "Sample Pipes", "sample": (io.BytesIO(b"x"), "list.pdf")},
                content_type="multipart/form-data")
check("a file that isn't Excel is refused", r.status_code == 400 and "isn't an Excel file" in text(r))
r = post("Sample Pipes", samples["avtar"], logo=(io.BytesIO(b"no picture"), "logo.png"))
check("a logo that isn't a picture is refused", r.status_code == 400 and "couldn't be read as a picture" in text(r))

r = post("Sample Pipes", samples["avtar"])
html = r.get_data(as_text=True)
check("sample read: the check page with a sample certificate", r.status_code == 200 and 'class="cert"' in html
      and "AVTAR POLYPLAST" in html and 'class="cert-pic"' in html, str(r.status_code))
token = re.search(r'name="token" value="(\w+)"', html).group(1)
form = {"token": token, "name": "Sample Pipes", "text1": "SAMPLE PIPES LTD", "size1": "22", "bold1": "on"}
for i in range(2, 6):
    form[f"text{i}"] = re.search(rf'name="text{i}"[^>]*>(.*?)</textarea>', html, re.S).group(1)
    form[f"size{i}"] = re.search(rf'name="size{i}" value="([\d.]+)"', html).group(1)
form["text2"] = "Plot 1, Sample Estate, Rajkot"
r = client.post("/company/update", data=form)
html = r.get_data(as_text=True)
check("a changed row shows on the sample certificate", "SAMPLE PIPES LTD" in html and "Sample Estate" in html
      and "AVTAR POLYPLAST" not in html.split('class="cert"')[1])
r = client.post("/company/update", data={**form, "remove_logo": "on"})
check("the logo can be removed", 'class="cert-pic"' not in r.get_data(as_text=True))

settings_before = SETTINGS_FILE.read_bytes()
previous = SETTINGS_FILE.with_name("settings.previous.json")
previous_before = previous.read_bytes() if previous.exists() else None
made = []
try:
    r = client.post("/company/save", data={"token": token})
    check("adding: back to the first screen", r.status_code == 302 and "added=sample" in r.headers["Location"],
          r.headers.get("Location", str(r.status_code)))
    s2 = load_settings()
    prof = s2["company"]["profiles"].get("sample", {})
    made = [ROOT / prof.get("template", "templates/test_report_sample.xlsx")]
    check("the new company is in settings.json", prof.get("name") == "Sample Pipes" and prof.get("C1", prof.get("A1")) == "SAMPLE PIPES LTD")
    check("settings.json still passes its checks", not __import__("app.settings", fromlist=["validate"]).validate(s2))
    page = client.get(r.headers["Location"]).get_data(as_text=True)
    check("the company menu lists it, chosen, with a notice", '<option value="sample" selected>Sample Pipes</option>' in page
          and "Sample Pipes was added" in page)
    lay = load_layout(ROOT / prof["template"])
    check(f"its template fits A4 at 100% ({lay.height:.1f} pt)", lay.height <= room - 1)
    cert = build_report(row, None, with_company(s2, "sample"))
    check("certificates for it carry its header", cert.cells["A1"].text == "SAMPLE PIPES LTD"
          or cert.cells.get("C1", cert.cells["A1"]).text == "SAMPLE PIPES LTD")
    check("the upload folder is cleared after adding", not (web.UPLOADS / token).exists())
    r = post("sample pipes", samples["avtar"])
    check("the same name can't be added twice", r.status_code == 400)
finally:
    SETTINGS_FILE.write_bytes(settings_before)
    if previous_before is None:
        previous.unlink(missing_ok=True)
    else:
        previous.write_bytes(previous_before)
    for p in made + list((ROOT / "assets").glob("sample_*.png")):
        p.unlink(missing_ok=True)
check("config/settings.json put back", SETTINGS_FILE.read_bytes() == settings_before)

print(f"\n(files in {TMP})")
shutil.rmtree(TMP, ignore_errors=True)
print("ALL PASS" if not failures else f"{failures} FAILED")
sys.exit(1 if failures else 0)
