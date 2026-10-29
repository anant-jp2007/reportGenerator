"""Print what reports would contain, without making any files.

    python tools/preview_report.py PRODUCTION.xls                  # newest row of each grade
    python tools/preview_report.py PRODUCTION.xls "PE 100" 53      # one row: grade + Excel row
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from app.production import read_production  # noqa: E402
from app.report import build_reports  # noqa: E402

LINES = [
    ("Test Report No.", "D7"), ("Date of Manufacturing", "D8"), ("Date of Testing Start", "D9"),
    ("Date of Testing End", "D10"), ("Batch Number", "D11"), ("M. Grade", "J7"), ("PN", "J8"),
    ("SIZE ( DN )", "J9"), ("SDR", "J10"),
    None,
    ("Stripes width (min.)", "J17"), ("Stripes depth (max.)", "J18"),
    ("Square of pipe end", "E20", "G20"),
    ("Outside dia. minimum", "E24", "J24"), ("Outside dia. maximum", "E25", "J25"),
    ("Wall minimum", "E26", "J26"), ("Wall maximum", "E27", "J27"),
    ("Ovality", "E28", "G28"),
    ("Longitudinal reversion", "G30"), ("Carbon black content", "G31"), ("Melt flow rate", "G33"),
    ("Density", "G35"), ("Elongation at break", "G36"), ("Yield strength", "G37"),
]


def show(rep):
    print(f"\n=== {rep.file_name}  ({rep.row.where()})")
    for line in LINES:
        if line is None:
            print(f"  {'':24} {'spec limit':>10}   {'value obtained':>14}")
            continue
        label, *refs = line
        if len(refs) == 2:
            print(f"  {label:24} {rep.cells[refs[0]].text:>10}   {rep.cells[refs[1]].text:>14}")
        else:
            print(f"  {label:24} {'':>10}   {rep.cells[refs[0]].text:>14}")
    for w in rep.warnings:
        print(f"  ! {w}")


def main():
    path = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "data/source/production_sample.xls"
    rows = read_production(path).rows
    if len(sys.argv) > 3:
        rows = [r for r in rows if r.grade == sys.argv[2] and r.excel_row == int(sys.argv[3])]
    else:
        rows = [max((r for r in rows if r.grade == g and r.ok), key=lambda r: r.excel_row)
                for g in ("PE 63", "PE 80", "PE 100")]
    for rep in build_reports(rows):
        show(rep)


if __name__ == "__main__":
    main()
