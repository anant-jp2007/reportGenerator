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

    for name in ("stripe_width", "stripe_depth", "longitudinal_reversion", "carbon_black_content",
                 "density", "elongation_at_break", "yield_strength"):
        conf, cell = RV[name], c[RV[name]["cell"]]
        rule(f"{name}: inside {conf['min']}-{conf['max']}, {conf['decimals']} decimals",
             conf["min"] <= cell.value <= conf["max"] and decimals_ok(cell, conf["decimals"]), (tag, cell.text))
    rule("elongation is a multiple of 5", c["G36"].value % 5 == 0, (tag, c["G36"].text))

    pe = c["G20"]
    lo, hi = RV["pipe_end"]["range_by_chart_max"][f"{sp.pipe_end:g}"]
    rule("square of pipe end: inside its range for the chart max, ends in 0 or 5",
         lo <= pe.value <= hi and pe.value < sp.pipe_end and pe.text[-1] in "05" and decimals_ok(pe, 2),
         (tag, pe.text, sp.pipe_end))

    a, b = c["J24"].value, c["J25"].value
    rule("OD obtained: strictly inside chart min-max (never on a limit), minimum < maximum, 0.1 steps, 2 decimals",
         sp.od_min < a < b < sp.od_max and on_step(a, 0.1) and on_step(b, 0.1)
         and decimals_ok(c["J24"], 2), (tag, a, b, sp.od_min, sp.od_max))
    a, b = c["J26"].value, c["J27"].value
    rule("wall obtained: strictly inside chart min-max (never on a limit), minimum < maximum, 2 decimals",
         sp.wall_min < a < b < sp.wall_max and decimals_ok(c["J26"], 2), (tag, a, b, sp.wall_min, sp.wall_max))
    ov = c["G28"].value
    rule("ovality: between 25% and 75% of chart max (away from 0 and the max)",
         0.25 * sp.ovality_max - 1e-9 <= ov <= 0.75 * sp.ovality_max + 1e-9, (tag, ov, sp.ovality_max))

    mlo, mhi = RV["melt_flow_rate"]["by_grade"][row.grade]
    rule("MFR inside its grade range, 3 decimals",
         mlo <= c["G33"].value <= mhi and decimals_ok(c["G33"], 3), (tag, c["G33"].text))
    mfr[row.grade].append(c["G33"].value)

    rule("length and coiling values empty, units kept",
         c["G21"].text == "" and c["G23"].text == "" and c["J21"].text == "meters" and c["J23"].text == "mm", tag)
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
check("numbered 1, 2, 3... with PE 63 first, then PE 80, then PE 100",
      nums == list(range(1, len(first) + 1)) and grades == sorted(grades, key=["PE 63", "PE 80", "PE 100"].index))
names = [r.file_name.lower() for r in first]
check("file names unique (repeats like MP63/007 in 2024 and 2025 get -2)", len(names) == len(set(names)),
      [r.file_name for r in first if r.file_name.endswith("-2")][:3])

# fresh random values every time
row = good[0]
r1, r2 = build_report(row, 1), build_report(row, 1)
value_cells = ["J17", "J18", "G20", "J24", "J25", "J26", "J27", "G28", "G30", "G31", "G33", "G35", "G36", "G37"]
diff = [k for k in value_cells if r1.cells[k].text != r2.cells[k].text]
check("same row generated twice gets different test values", len(diff) >= 10, f"{len(diff)} of 14 differ")

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
            a, b = pick_pair(rng, size["od_min"], size["od_max"], 2, 0.1, True)
            x, y = pick_pair(rng, w["min"], w["max"], 2, 0.01, True)
            if not (size["od_min"] < a < b < size["od_max"] and w["min"] < x < y < w["max"]):
                tight.append((od, sdr, a, b, x, y))
check("all chart sizes x SDRs: OD and wall land strictly inside, minimum < maximum", not tight, tight[:3])

# warnings from the row travel with the report
nb = next(x for x in good if not x.batch)
check("empty batch: report still made, batch blank, warning kept",
      build_report(nb, 1).cells["D11"].text == "" and "batch number is empty" in build_report(nb, 1).warnings)

print("\nALL PASS" if not failures else f"\n{failures} FAILED")
sys.exit(1 if failures else 0)
