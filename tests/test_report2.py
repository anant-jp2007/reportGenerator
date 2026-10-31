"""Report 2 (testing record): entries side by side, 4 per sheet, same values as Report 1; continuing an
incomplete Report 2.

    python tests/test_report2.py
"""
import copy
import sys
import tempfile
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from app.production import Row  # noqa: E402
from app.chart import lookup  # noqa: E402
from app.report import build_reports  # noqa: E402
from app.report2 import Report2Error, read_incomplete, write_report2  # noqa: E402
from app.settings import load_settings, validate, with_company  # noqa: E402

from datetime import date  # noqa: E402

TMP = Path(tempfile.mkdtemp(prefix="report2test_"))
S = with_company(load_settings(), "balaji")
failures = 0


def check(name, cond, detail=""):
    global failures
    failures += not cond
    print(("PASS " if cond else "FAIL ") + name + (f"  [{detail}]" if detail else ""))


def row(i, grade="PE 100", od=63, pn=10, sdr=13.6):
    r = Row(grade, 10 + i, i, date(2026, 5, 1 + i), f"{od} X {pn:g} ( SDR {sdr:g} )", f"MP{grade[3:]}/{i:03d}",
            f"2026 05 {1 + i:02d} 001 1", od, pn, sdr)
    r.spec = lookup(grade, od, pn, sdr)
    return r


rows = [row(i) for i in range(1, 8)] + [row(8, "PE 63", 75, 6, 13.6), row(9, "PE 80", 40, 10, 11),
                                         row(10, "PE 100", 800, 10, 13.6)]
reports = build_reports(rows, S)
path, sheets = write_report2(reports, TMP / "r2.xlsx", settings=S)
wb = openpyxl.load_workbook(path)
check("10 entries: 3 sheets of 4, 4 and 2, in one workbook",
      [len(sh.reports) for sh in sheets] == [4, 4, 2] and wb.sheetnames == ["Report 2 - 1", "Report 2 - 2", "Report 2 - 3"],
      wb.sheetnames)

rows_map = S["report2"]["rows"]
ws = wb["Report 2 - 1"]
same = []
for col, rep in zip("DEFG", reports[:4]):
    for r, source in rows_map.items():
        if source == "grade_number":
            continue
        want, got = rep.cells[source].value, ws[f"{col}{r}"].value
        if hasattr(got, "date") and not isinstance(want, str):
            got = got.date()
        if isinstance(want, str):
            want = want.strip()
        if got != want:
            same.append((col, r, source, got, want))
check("every value is the one Report 1 prints for that entry", not same, same[:4])
check("in grade order; designation of the first (PE 63, 75 X 6): 63 / 6 / 75 / 13.6",
      [ws[f"D{r}"].value for r in (2, 3, 4, 5)] == [63, 6, 75, 13.6], [ws[f"D{r}"].value for r in (2, 3, 4, 5)])
check("rows Report 1 doesn't have (18-24, signature) stay empty",
      all(ws[f"{c}{r}"].value is None for c in "DEFG" for r in list(range(22, 29)) + [38]))
check("numbers keep Report 1's decimals (OD 63.10 -> 0.00), dates dd-mm-yyyy",
      ws["D13"].number_format == "0.00" and ws["D6"].number_format == "dd\\-mm\\-yyyy" and ws["D32"].number_format == "0.000")
check("Report Number row = Report No. (column D)", [ws[f"{c}37"].value for c in "DEFG"] == [r.number for r in reports[:4]])
last = wb["Report 2 - 3"]
check("last sheet: 2 entries, columns F and G empty",
      last["E37"].value == "MP100/010" and all(last[f"{c}{r}"].value is None for c in "FG" for r in range(2, 39)))
check("800 mm: ovality left empty, as on Report 1", last["E17"].value is None and reports[9].cells["G28"].text == "")
check("page: A4 landscape, one page", ws.page_setup.orientation == "landscape" and ws.page_setup.fitToHeight == 1
      and ws.sheet_properties.pageSetUpPr.fitToPage)
tmpl = openpyxl.load_workbook(ROOT / S["report2"]["template"]).active
check("labels and borders as in the template",
      all(ws.cell(r, 2).value == tmpl.cell(r, 2).value and repr(ws.cell(r, 4).border) == repr(tmpl.cell(r, 4).border) for r in range(1, 39)))

# continuing: the incomplete last sheet keeps its 2 columns and gets 2 more, then a new sheet
more = build_reports([row(i, "PE 80", 40, 10, 11) for i in range(20, 23)], S)
start = read_incomplete(path, S)
check("incomplete Report 2: last sheet found with 2 filled columns", start.title == "Report 2 - 3" and len(start.kept) == 2)
path2, sheets2 = write_report2(more, TMP / "r2_more.xlsx", path, S)
wb2 = openpyxl.load_workbook(path2)
ws2 = wb2["Report 2 - 3"]
check("its sheet comes first, then 'Report 2 - 4' for the rest", wb2.sheetnames == ["Report 2 - 3", "Report 2 - 4"], wb2.sheetnames)
check("old columns unchanged (values and formats), new entries in F and G",
      all(ws2[f"{c}{r}"].value == last[f"{c}{r}"].value and ws2[f"{c}{r}"].number_format == last[f"{c}{r}"].number_format
          for c in "DE" for r in range(2, 39))
      and [ws2[f"{c}37"].value for c in "FG"] == ["MP80/020", "MP80/021"] and wb2["Report 2 - 4"]["D37"].value == "MP80/022")

full, _ = write_report2(reports[:4], TMP / "full.xlsx", settings=S)
try:
    read_incomplete(full, S)
    check("a full last sheet is refused", False)
except Report2Error as e:
    check("a full last sheet is refused with a clear reason", "already full" in str(e), e)
other = openpyxl.Workbook(); other.active["A1"] = "something else"; other.save(TMP / "other.xlsx")
try:
    read_incomplete(TMP / "other.xlsx", S)
    check("a file that isn't a Report 2 is refused", False)
except Report2Error as e:
    check("a file that isn't a Report 2 is refused with a clear reason", "laid out like Report 2" in str(e), e)
(TMP / "junk.xlsx").write_bytes(b"not a spreadsheet")
try:
    read_incomplete(TMP / "junk.xlsx", S)
    check("a damaged file is refused", False)
except Report2Error as e:
    check("a damaged file is refused with a clear reason", "couldn't be opened" in str(e), e)

# settings checks
bad = copy.deepcopy(S); bad["report2"]["rows"]["13"] = "K99"
check("a Report 2 row pointing at a cell Report 1 doesn't fill is caught",
      any("report2 row 13" in p for p in validate(bad)), validate(bad))
bad = copy.deepcopy(S); bad["report2"]["template"] = "templates/nope.xlsx"
check("missing Report 2 template is caught", any("report2: template" in p for p in validate(bad)))

print(f"\n(files in {TMP})")
print("ALL PASS" if not failures else f"{failures} FAILED")
sys.exit(1 if failures else 0)
