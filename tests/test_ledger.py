"""The record of certified entries (ledger): one register is enough once a company has a record.

    python tests/test_ledger.py

Uses the sample registers from tools/make_sample_registers.py: "before" has 4 entries, "after" 10
(6 new, of which 5 can be certified and 1 can't: 40 X 10 on PE 63).
"""
import html as htmllib
import io
import json
import sys
import tempfile
import zipfile
from datetime import date
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tools"))
from app import web  # noqa: E402
from app.ledger import ledger_path, load, readable_grades  # noqa: E402
from app.production import read_production  # noqa: E402
from make_sample_registers import BEFORE, NEW, build, make  # noqa: E402

TMP = Path(tempfile.mkdtemp(prefix="ledger_"))
web.OUTPUT, web.UPLOADS, web.LEDGER = TMP / "output", TMP / "uploads", TMP / "ledger"
client = web.app.test_client()
failures = 0


def check(name, cond, detail=""):
    global failures
    failures += not cond
    print(("PASS " if cond else "FAIL ") + name + (f"  [{detail}]" if detail else ""))


def text(resp):
    return htmllib.unescape(resp.get_data(as_text=True))


def post_check(company, new: Path, old: Path | None = None):
    data = {"company": company, "new": (io.BytesIO(new.read_bytes()), new.name)}
    if old:
        data["old"] = (io.BytesIO(old.read_bytes()), old.name)
    return client.post("/check", data=data, content_type="multipart/form-data")


def token_ids(html):
    token = html.split('name="token" value="')[1].split('"')[0] if 'name="token"' in html else None
    return token, [s.split('"')[0] for s in html.split('name="row" value="')[1:]]


def download(token, ids, **extra):
    return client.post("/download", data={"token": token, "row": ids, "files": "each", **extra})


def record(company):
    return load(ledger_path(web.LEDGER, company))


def without_sheet(src: Path, sheet: str, name: str) -> Path:
    wb = openpyxl.load_workbook(src)
    del wb[sheet]
    out = TMP / "registers" / name
    wb.save(out)
    return out


def start_page():
    html = text(client.get("/"))
    return json.loads(html.split("var RECORDS = ")[1].split(";\n")[0]), html


before, after = make(TMP / "registers")

# ---------- a company without a record needs the old register once ----------
check("no record yet: empty, not started", not record("balaji").started and record("balaji").entries == [])
r = post_check("balaji", after)
check("one register without a record: asks for the old register too", r.status_code == 400
      and "Choose the old register too" in text(r) and "no record yet" in text(r), text(r)[:0])
records, html = start_page()
check("start page: Balaji has no record, the old register box is shown and required",
      records["balaji"]["started"] is False and 'name="old" accept=".xls,.xlsx" required' in html)

# ---------- two registers: compared as before, then the record is started ----------
html = text(post_check("balaji", after, before))
token, ids = token_ids(html)
check("two registers: 5 new entries, as before", "5 new entries to certify" in html and len(ids) == 5, len(ids))
check("ids differ for every entry", len(set(ids)) == len(ids))
r = download(token, ids[:3], **{f"no-{ids[0]}": "BPP63/X1"})
check("download of 3 entries: recorded", r.status_code == 200 and r.headers["X-Recorded"] == "ok", r.headers.get("X-Recorded"))
led = record("balaji")
check("record: old register's 4 entries + 3 certified", len(led.entries) == 7 and led.certified == 3,
      (len(led.entries), led.certified))
check("record knows the 3 grades of the old register", led.grades == {"PE 63", "PE 80", "PE 100"}, led.grades)
check("record keeps the Report No. used, typed or from the register",
      sorted(e["report_no"] for e in led.entries if e["source"] == "certified") == ["BPP63/X1", "MP80/002", "MP80/003"],
      sorted(e["report_no"] for e in led.entries if e["source"] == "certified"))
check("record saved as ledger/balaji.json", (web.LEDGER / "balaji.json").exists())

r = download(token, ids[:3])
check("same entries downloaded again (another format): nothing added twice",
      r.status_code == 200 and len(record("balaji").entries) == 7)
check("previous version of the record kept", (web.LEDGER / "balaji.previous.json").exists())

records, html = start_page()
check("start page: Balaji now has a record (3 certified, today)",
      records["balaji"]["started"] and records["balaji"]["certified"] == 3
      and records["balaji"]["last"] == date.today().strftime("%d-%m-%Y"), records["balaji"])
check("start page: with a record the old register box is hidden and not required",
      'id="old-file" hidden' in html and 'name="old" accept=".xls,.xlsx">' in html)

# ---------- one register: checked against the record ----------
html = text(post_check("balaji", after))
token1, ids1 = token_ids(html)
check("one register: only the 2 entries not yet certified", "2 new entries to certify" in html and ids1 == ids[3:],
      (ids1, ids[3:]))
check("review says it was checked against the record", "against the record of certified entries" in html)
check("the entry that can't be certified is still listed with its reason", "No certificate for 1 entry" in html)

r = download(token1, ids1[:1])
check("download 1 of the 2: recorded", r.headers["X-Recorded"] == "ok" and record("balaji").certified == 4)
r = download(token1, ids1)
check("same check screen, both entries: still offered (the record as it was when checked), only 1 added",
      r.status_code == 200 and record("balaji").certified == 5 and len(record("balaji").entries) == 9,
      (record("balaji").certified, len(record("balaji").entries)))
r = download(token, ids)
check("an older check screen downloading everything again: nothing added twice", len(record("balaji").entries) == 9)

html = text(post_check("balaji", after))
check("everything certified: 'No new entries', and no hint about swapped files",
      "No new entries" in html and "already got a certificate" in html and "swapped" not in html)

# edited entries stay done, a new one is found
edited = [(g, sr, d, s, ("CHANGED" if (g, sr) == ("PE 80", 2) else b), m, rem) for g, sr, d, s, b, m, rem in BEFORE + NEW]
more = edited + [("PE 100", 3, date(2026, 4, 9), "63 X 5 ( SDR 26 )", "2026 04 09 001 1", 200, "")]
build(more, TMP / "registers" / "PRODUCTION later.xlsx")
html = text(post_check("balaji", TMP / "registers" / "PRODUCTION later.xlsx"))
_, ids2 = token_ids(html)
check("later register: a certified entry with an edited batch stays done, the new PE 100 Sr 3 is found",
      len(ids2) == 1 and ids2[0].startswith("PE100-2026-3-"), ids2)

# two registers with a record: compared as before, the record is only added to
html = text(post_check("balaji", after, before))
check("old register chosen although there's a record: compared as before (5 new)", "5 new entries to certify" in html)

# ---------- the same Sr. No. twice ----------
dup = BEFORE + NEW + [("PE 63", 3, date(2025, 4, 30), "63 X 4 ( SDR 21 )", "2025 04 30 001 2", 600, "")]
build(dup, TMP / "registers" / "PRODUCTION dup.xlsx")
html = text(post_check("newton", TMP / "registers" / "PRODUCTION dup.xlsx", before))
token, ids = token_ids(html)
sr3 = [i for i in ids if i.startswith("PE63-2025-3-")]
check("two new rows with Sr. 3: two separate entries", len(sr3) == 2 and len(set(ids)) == len(ids) == 6, ids)
r = download(token, sr3)
z = zipfile.ZipFile(io.BytesIO(r.data))
check("both Sr. 3 rows get their own certificate", len(z.namelist()) == 2, z.namelist())
check("both recorded", sum(e["sr"] == 3 and e["grade"] == "PE 63" and e["source"] == "certified"
                           for e in record("newton").entries) == 2)
html = text(post_check("newton", TMP / "registers" / "PRODUCTION dup.xlsx"))
_, ids = token_ids(html)
check("next check: neither Sr. 3 row is new again", not any(i.startswith("PE63-2025-3-") for i in ids) and len(ids) == 4, ids)

# ---------- Report 2 alone also records ----------
html = text(post_check("abhay", after, before))
token, ids = token_ids(html)
r = download(token, ids[:1], report_choice="1", report="2")
check("Report 2 alone: recorded too", r.mimetype.endswith("spreadsheetml.sheet") and r.headers["X-Recorded"] == "ok"
      and record("abhay").certified == 1)

# ---------- grades the record doesn't know, sheets that went missing ----------
before_no80 = without_sheet(before, "Production 80", "before no 80.xlsx")
after_no80 = without_sheet(after, "Production 80", "after no 80.xlsx")
html = text(post_check("avtar", after_no80, before_no80))
token, ids = token_ids(html)
download(token, ids[:1])
check("record from registers without PE 80 doesn't know PE 80", record("avtar").grades == {"PE 63", "PE 100"})
html = text(post_check("avtar", after))
_, ids = token_ids(html)
check("register with a PE 80 sheet: PE 80 skipped with a clear message, other grades checked",
      "PE 80 isn't in the record of certified entries yet" in html and not any(i.startswith("PE80") for i in ids)
      and len(ids) == 2, ids)
after_no100 = without_sheet(after, "Production 100", "after no 100.xlsx")
html = text(post_check("avtar", after_no100))
check("register missing a sheet the record has: said plainly", "no sheet for PE 100" in html)

# an old register whose sheet can't be read must not count that grade as recorded
broken = openpyxl.load_workbook(before)
broken["Production 80"]["C2"] = "Something else"
broken.save(TMP / "registers" / "before broken 80.xlsx")
read = read_production(TMP / "registers" / "before broken 80.xlsx")
check("unreadable sheet: its grade isn't counted as recorded", "PE 80" in read.grades
      and readable_grades(read) == {"PE 63", "PE 100"}, readable_grades(read))

# ---------- a damaged record ----------
(web.LEDGER / "balaji.json").write_text("{ not json", encoding="utf-8")
records, html = start_page()
check("damaged record: start page still opens, asks for the old register", records["balaji"]["broken"]
      and not records["balaji"]["started"])
r = post_check("balaji", after)
check("damaged record, one register: clear message, nothing treated as new",
      r.status_code == 400 and "can't be read" in text(r) and "Choose the old register too" in text(r))
html = text(post_check("balaji", after, before))
token, ids = token_ids(html)
r = download(token, ids[:2])
check("two registers: certificates made, damaged record kept aside, a new one started",
      r.status_code == 200 and r.headers["X-Recorded"] == "ok" and list(web.LEDGER.glob("balaji.damaged *.json"))
      and record("balaji").certified == 2 and len(record("balaji").entries) == 6)

print(f"\n(files in {TMP})")
print("ALL PASS" if not failures else f"{failures} FAILED")
sys.exit(1 if failures else 0)
