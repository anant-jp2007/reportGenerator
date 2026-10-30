"""Report engine: generate 1000+ reports from the real register and check every rule.

    python tests/test_report.py [production.xls]
"""
import copy
import sys
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from app.chart import lookup  # noqa: E402
from app.production import Row, read_production  # noqa: E402
from app.report import build_report, build_reports  # noqa: E402
from app.settings import company_profile, load_settings  # noqa: E402

REAL = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "data/source/production_sample.xls"
S = load_settings()
RV = S["random_values"]
failures = 0
problems = defaultdict(list)   # rule -> first few failing examples


def check(name, cond, detail=""):
    global failures
    failures += not cond
    print(("PASS " if cond else "FAIL ") + name + (f"  [{detail}]" if detail else ""))


def rule(name, cond, example):
    problems[name]  # register the rule so it's reported even when it never fails
    if not cond and len(problems[name]) < 3:
        problems[name].append(example)


def decimals_ok(cell, d):
    t = cell.text
    return (t.split(".")[1].__len__() == d) if d else ("." not in t)


def on_step(x, step):
    return abs(round(x / step) * step - x) < 1e-9


rows = read_production(REAL).rows
good = [r for r in rows if r.ok]
reports = build_reports(rows) + build_reports(rows)     # two rounds: 1036 reports
check(f"{len(reports)} reports generated from {len(good)} usable rows (row with problems left out)",
      len(reports) == 2 * len(good) and all(r.row.ok for r in reports))

company = company_profile(S)
mfr = defaultdict(list)
for rep in reports:
    c, sp, row = rep.cells, rep.row.spec, rep.row
    tag = f"report {rep.number} {row.where()}"

    rule("company header and every fixed text printed as in settings",
         all(c[k].text == str(v) for k, v in {**company, **S["fixed_text"]}.items()), tag)
    made = row.date
    rule("dates: manufacturing = row date, start = same day, end = start + 3 days",
         (c["D8"].value, c["D9"].value, c["D10"].value) == (made, made, made + timedelta(days=3))
         and c["D8"].text == made.strftime("%d-%m-%Y"), tag)
    rule("batch, grade, PN, DN, SDR from the row",
         (c["D11"].text, c["J7"].text, c["J8"].text, c["J9"].text, c["J10"].text)
         == (row.batch, row.grade, f"{sp.pn:g}", f"{sp.od:g}", f"{sp.sdr:g}"), tag)
    rule("spec limits E20/E24-E28 from the chart, 2 decimals",
         [c[k].text for k in ("E20", "E24", "E25", "E26", "E27", "E28")]
         == [f"{v:.2f}" for v in (sp.pipe_end, sp.od_min, sp.od_max, sp.wall_min, sp.wall_max, sp.ovality_max)], tag)

    for name in ("longitudinal_reversion", "carbon_black_content",
                 "density", "elongation_at_break", "yield_strength"):
        conf, cell = RV[name], c[RV[name]["cell"]]
        rule(f"{name}: inside {conf['min']}-{conf['max']}, {conf['decimals']} decimals",
             conf["min"] <= cell.value <= conf["max"] and decimals_ok(cell, conf["decimals"]), (tag, cell.text))
    rule("elongation is a multiple of 5", c["G36"].value % 5 == 0, (tag, c["G36"].text))
    rule("longitudinal reversion ends in 0 or 5", c["G30"].text[-1] in "05", (tag, c["G30"].text))
    rule("density between 948 and 952", 948 <= c["G35"].value <= 952, (tag, c["G35"].text))

    w = c["J17"].value
    rule("stripe width: 3.25-4.25 up to DN 110, 4.25-5.50 above, 2 decimals",
         (3.25 <= w <= 4.25 if sp.od <= 110 else 4.25 <= w <= 5.50) and decimals_ok(c["J17"], 2), (tag, sp.od, w))
    d = c["J18"].value
    rule("stripe depth: below 0.20 when chart wall min <= 10 mm, below 0.45 above, 2 decimals",
         (0.10 <= d < 0.20 if sp.wall_min <= 10 else 0.25 <= d < 0.45) and decimals_ok(c["J18"], 2),
         (tag, sp.wall_min, d))

    pe = c["G20"]
    rule("square of pipe end: half the chart max to half + 0.50 (2 -> 1.00-1.50, 3 -> 1.50-2.00), ends in 0 or 5",
         sp.pipe_end / 2 <= pe.value <= sp.pipe_end / 2 + 0.5 and pe.value < sp.pipe_end
         and pe.text[-1] in "05" and decimals_ok(pe, 2), (tag, pe.text, sp.pipe_end))

    a, b = c["J24"].value, c["J25"].value
    gap = RV["outside_diameter"]["min_gap_fraction"] * (sp.od_max - sp.od_min) - 1e-9
    rule("OD obtained: strictly inside chart min-max, maximum >= minimum + 20% of the range, 0.1 steps, 2 decimals",
         sp.od_min < a < b < sp.od_max and b - a >= gap and on_step(a, 0.1) and on_step(b, 0.1)
         and decimals_ok(c["J24"], 2), (tag, a, b, sp.od_min, sp.od_max))
    a, b = c["J26"].value, c["J27"].value
    gap = RV["wall_thickness"]["min_gap_fraction"] * (sp.wall_max - sp.wall_min) - 1e-9
    rule("wall obtained: strictly inside chart min-max, maximum >= minimum + 25% of the range (never 3.24 / 3.25)",
         sp.wall_min < a < b < sp.wall_max and b - a >= gap and decimals_ok(c["J26"], 2),
         (tag, a, b, sp.wall_min, sp.wall_max))
    ov = c["G28"].value
    rule("ovality: between 20% and 50% of chart max, ends in 0 or 5",
         0.20 * sp.ovality_max - 1e-9 <= ov <= 0.50 * sp.ovality_max + 1e-9 and c["G28"].text[-1] in "05",
         (tag, c["G28"].text, sp.ovality_max))

    mlo, mhi = RV["melt_flow_rate"]["by_grade"][row.grade]
    rule("MFR inside its grade range, 3 decimals",
         mlo <= c["G33"].value <= mhi and decimals_ok(c["G33"], 3), (tag, c["G33"].text))
    mfr[row.grade].append(c["G33"].value)

    coil = c["G23"]
    rule("length of pipes '---'; coiling 18 x DN + 50 to + 100 up to DN 110, '---' above; units kept",
         c["G21"].text == "---" and c["J21"].text == "meters" and c["J23"].text == "mm"
         and (18 * sp.od + 50 <= coil.value <= 18 * sp.od + 100 and decimals_ok(coil, 0)
              if sp.od <= 110 else coil.text == "---"), (tag, sp.od, coil.text))
    rule("file name = Report No. column with / replaced by -",
         rep.file_name.split("-")[0:2] == row.report_no.replace("/", "-").split("-")[0:2], (tag, rep.file_name))

for name, bad in problems.items():
    check(name, not bad, bad)

check("MFR always PE 100 < PE 80 < PE 63 across all reports",
      max(mfr["PE 100"]) < min(mfr["PE 80"]) and max(mfr["PE 80"]) < min(mfr["PE 63"]),
      {g: (min(v), max(v)) for g, v in mfr.items()})

first = build_reports(rows)
nums = [r.number for r in first]
grades = [r.row.grade for r in first]
check("Report No. (D7) = the register's column D, PE 63 first, then PE 80, then PE 100",
      all(r.cells["D7"].text == r.row.report_no for r in first) and nums == [r.row.report_no for r in first]
      and grades == sorted(grades, key=["PE 63", "PE 80", "PE 100"].index), nums[:3])
check("CM/L number with leading zeros: 0003712656", "CM/L - 0003712656" in first[0].cells["A5"].text)
check("Balaji: no pictures, 'Tested by :-' left blank to sign by hand",
      first[0].pictures == [] and first[0].cells["G39"].text == "Tested by :-")

RANDOM_CELLS = {"J17", "J18", "G20", "G23", "J24", "J25", "J26", "J27", "G28", "G30", "G31", "G33", "G35", "G36", "G37"}

# a second company: its own header, template and pictures; the table is filled exactly the same way
from app.settings import with_company  # noqa: E402

ab = build_report(good[0], None, with_company(S, "abhay"))
check("Abhay: own header (name in C1, phone sign in A4, title without CM/L), own template",
      ab.cells["C1"].text == "ABHAY POLYPLAST" and "✆" in ab.cells["A4"].text and "CM/L" not in ab.cells["A5"].text
      and ab.template.name == "test_report_abhay.xlsx")
check("Abhay: no 'Tested by :-', logo and stamp pictures",
      ab.cells["G39"].text == "" and [p.file.name for p in ab.pictures] == ["abhay_logo.png", "abhay_stamp.png"])
check("Abhay: Balaji's header cells not printed", "A1" not in ab.cells and "A2" not in ab.cells)
same = [k for k in first[0].cells if int(k[1:]) >= 6 and k != "G39" and k not in RANDOM_CELLS
        and first[0].cells[k] != build_report(first[0].row, None, with_company(S, "abhay")).cells.get(k)]
check("Abhay: every non-random table cell same as Balaji's", not same, same[:5])
names = [r.file_name.lower() for r in first]
check("file names unique (repeats like MP63/007 in 2024 and 2025 get -2)", len(names) == len(set(names)),
      [r.file_name for r in first if r.file_name.endswith("-2")][:3])

# fresh random values every time
row = good[0]
r1, r2 = build_report(row, 1), build_report(row, 1)
value_cells = ["J17", "J18", "G20", "J24", "J25", "J26", "J27", "G28", "G30", "G31", "G33", "G35", "G36", "G37"]
diff = [k for k in value_cells if r1.cells[k].text != r2.cells[k].text]
check("same row generated twice gets different test values", len(diff) >= 9, f"{len(diff)} of 14 differ")

# calendar: +3 days across month, year and leap-day boundaries
def with_date(d):
    r = copy.copy(good[0])
    r.date = d
    return build_report(r, 1).cells["D10"].text


cases = {date(2025, 1, 30): "02-02-2025", date(2025, 12, 30): "02-01-2026", date(2024, 2, 27): "01-03-2024",
         date(2025, 2, 27): "02-03-2025", date(2025, 4, 30): "03-05-2025"}
got = {d.isoformat(): with_date(d) for d in cases}
check("testing end +3 days handles month end, year end, leap year", list(got.values()) == list(cases.values()), got)

# PN/SDR with decimals print as written
r = copy.copy(good[0])
r.grade, r.spec, r.warnings = "PE 80", lookup("PE 80", 63, 12.5), []
c = build_report(r, 1).cells
check("PN 12.5 / SDR 9 printed as 12.5 and 9", (c["J8"].text, c["J10"].text) == ("12.5", "9"))
r.spec = lookup("PE 63", 32, 6)
r.grade = "PE 63"
check("SDR 13.6 printed as 13.6 (not 14)", build_report(r, 1).cells["J10"].text == "13.6")

# sizes with no ovality in the chart (710 mm and above)
r.grade, r.spec = "PE 100", lookup("PE 100", 710, 5)
rep = build_report(r, 1)
check("710 mm: no ovality in chart -> both ovality cells blank, with a warning",
      rep.cells["E28"].text == "" and rep.cells["G28"].text == "" and any("ovality" in w for w in rep.warnings))

# every size and SDR in the chart is wide enough for strictly-inside values (no fallback to the limit)
import random  # noqa: E402
from app.chart import chart  # noqa: E402
from app.report import pick_pair  # noqa: E402

rng, tight = random.Random(1), []
for od, size in chart()["sizes"].items():
    for sdr, w in size["wall"].items():
        for _ in range(20):
            od_gap, wall_gap = RV["outside_diameter"]["min_gap_fraction"], RV["wall_thickness"]["min_gap_fraction"]
            a, b = pick_pair(rng, size["od_min"], size["od_max"], 2, 0.1, True, od_gap)
            x, y = pick_pair(rng, w["min"], w["max"], 2, 0.01, True, wall_gap)
            if not (size["od_min"] < a < b < size["od_max"] and w["min"] < x < y < w["max"]
                    and b - a >= od_gap * (size["od_max"] - size["od_min"]) - 1e-9
                    and y - x >= wall_gap * (w["max"] - w["min"]) - 1e-9):
                tight.append((od, sdr, a, b, x, y))
check("all chart sizes x SDRs: OD and wall strictly inside, with the minimum gap, never a fallback",
      not tight, tight[:3])

# every chart size: pipe end and ovality ranges hold values ending in 0 or 5
from app.report import pick, pipe_end_range  # noqa: E402

empty = []
for od, size in chart()["sizes"].items():
    for what, mx, (lo, hi) in (("pipe end", size.get("pipe_end"), (0, 0)), ("ovality", size.get("ovality_max"), (0.2, 0.5))):
        if mx is None:
            continue
        lo, hi = pipe_end_range(RV["pipe_end"], mx) if what == "pipe end" else (mx * lo, mx * hi)
        try:
            pick(rng, lo, hi, 2, 0.05)
        except ValueError:
            empty.append((od, what, mx))
check("all chart sizes: pipe end and ovality ranges have values ending in 0 or 5", not empty, empty)

# big sizes (the real register stops at 110): wider stripes, deeper stripes for walls over 10 mm, no coiling
r = copy.copy(good[0])
r.grade, r.warnings = "PE 100", []
for od, pn, width, depth in ((110, 6, (3.25, 4.25), (0.10, 0.19)), (125, 6, (4.25, 5.50), (0.10, 0.19)),
                             (110, 12.5, (3.25, 4.25), (0.10, 0.19)),   # SDR 11, wall 10.00-11.10: still up to 10
                             (315, 16, (4.25, 5.50), (0.25, 0.44))):    # wall 28.6: above 10
    r.spec = lookup("PE 100", od, pn)
    bad = []
    for _ in range(200):
        c = build_report(r, 1).cells
        if not (width[0] <= c["J17"].value <= width[1] and depth[0] <= c["J18"].value <= depth[1]
                and (c["G23"].text == "---") == (od > 110)):
            bad.append((c["J17"].text, c["J18"].text, c["G23"].text))
    check(f"DN {od} PN {pn} (wall min {r.spec.wall_min:.2f}): width {width}, depth {depth}, "
          f"coiling {'---' if od > 110 else 'number'}", not bad, bad[:3])

# no date in the register: certificate made, all three dates blank, note kept
r = copy.copy(good[0])
r.date, r.warnings = None, ["date is empty; the certificate's dates are left blank"]
c = build_report(r, 1).cells
check("no date: D8, D9, D10 printed blank, everything else filled",
      (c["D8"].text, c["D9"].text, c["D10"].text) == ("", "", "") and c["J9"].text and c["J24"].text)

# warnings from the row travel with the report
nb = next(x for x in good if not x.batch)
check("empty batch: report still made, batch blank, warning kept",
      build_report(nb, 1).cells["D11"].text == "" and "batch number is empty" in build_report(nb, 1).warnings)

print("\nALL PASS" if not failures else f"\n{failures} FAILED")
sys.exit(1 if failures else 0)
