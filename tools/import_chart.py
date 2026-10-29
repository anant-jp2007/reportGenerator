"""Convert the IS 4984 weight chart (.xls) into data/is4984_chart.json for the app.

Run again whenever the chart is amended:

    python tools/import_chart.py [weight_chart.xls] [output.json]

Reads the first sheet ("IS - 4984 - 2016"), which has every SDR side by side:
    B Nominal OD | C,D Mean OD min/max | E Max ovality | F Pipe end
    then 5 columns per SDR (41, 33, 26, 21, 17, 13.6, 11, 9, 7.4, 6):
    wall min | wall max | (average) | (inside dia) | weight kg/m
Rows 4-6 map each SDR to its PN for PE 63 / PE 80 / PE 100 ("----" = not made).
The "Page - 01/02" sheets are print copies of the same data; they are cross-checked.
"""
import json
import re
import sys
from datetime import date
from pathlib import Path

import xlrd

ROOT = Path(__file__).resolve().parent.parent
SRC = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "data/source/weight_chart.xls"
OUT = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "data/is4984_chart.json"

FIRST_SDR_COL = 6          # column G
COLS_PER_SDR = 5
GRADE_ROWS = {"PE 63": 3, "PE 80": 4, "PE 100": 5}   # 0-based rows 4, 5, 6
SDR_ROW = 2                                          # row 3
FIRST_DATA_ROW = 9                                   # row 10


def num_key(x: float) -> str:
    """16.0 -> '16', 13.6 -> '13.6' (stable keys for OD / SDR / PN)."""
    return f"{x:g}"


def number(v):
    return float(v) if isinstance(v, (int, float)) and v != "" else None


def parse_label(text, prefix):
    m = re.fullmatch(rf"\s*{prefix}\s*([\d.]+)\s*", str(text))
    return float(m.group(1)) if m else None


def read_main(sh):
    sdrs = []
    c = FIRST_SDR_COL
    while c < sh.ncols:
        sdr = parse_label(sh.cell_value(SDR_ROW, c), "SDR")
        if sdr is None:
            break
        sdrs.append((num_key(sdr), c))
        c += COLS_PER_SDR

    pn = {}
    for grade, r in GRADE_ROWS.items():
        pn[grade] = {}
        for sdr, col in sdrs:
            value = parse_label(sh.cell_value(r, col), "PN")
            if value is not None:
                pn[grade][sdr] = value

    sizes = {}
    for r in range(FIRST_DATA_ROW, sh.nrows):
        od = number(sh.cell_value(r, 1))
        if od is None:
            continue
        walls = {}
        for sdr, col in sdrs:
            wmin, wmax, weight = (number(sh.cell_value(r, col + i)) for i in (0, 1, 4))
            if wmin is not None and wmax is not None:
                walls[sdr] = {"min": wmin, "max": wmax,
                              "weight": round(weight, 6) if weight is not None else None}
        sizes[num_key(od)] = {
            "od_min": number(sh.cell_value(r, 2)),
            "od_max": number(sh.cell_value(r, 3)),
            "ovality_max": number(sh.cell_value(r, 4)),
            "pipe_end": number(sh.cell_value(r, 5)),
            "wall": walls,
        }
    return [s for s, _ in sdrs], pn, sizes


def cross_check(wb, sizes):
    """The print sheets hold the same numbers split in two halves; report any difference."""
    problems = []
    for sh in wb.sheets()[1:]:
        hdr = next((r for r in range(min(12, sh.nrows))
                    if any(parse_label(v, "SDR") for v in sh.row_values(r))), None)
        if hdr is None:
            continue
        blocks = [(num_key(parse_label(v, "SDR")), c) for c, v in enumerate(sh.row_values(hdr))
                  if parse_label(v, "SDR")]
        for r in range(hdr + 1, sh.nrows):
            od = number(sh.cell_value(r, 1))
            if od is None or num_key(od) not in sizes:
                continue
            size = sizes[num_key(od)]
            for key, col in (("od_min", 2), ("od_max", 3), ("ovality_max", 4), ("pipe_end", 5)):
                v = number(sh.cell_value(r, col))
                if v != size[key]:
                    problems.append(f"{sh.name}: OD {num_key(od)} {key} {v} != {size[key]}")
            for sdr, col in blocks:
                wmin, wmax = number(sh.cell_value(r, col)), number(sh.cell_value(r, col + 1))
                main = size["wall"].get(sdr)
                if (main is None) != (wmin is None) or (main and (main["min"], main["max"]) != (wmin, wmax)):
                    problems.append(f"{sh.name}: OD {num_key(od)} SDR {sdr} wall {wmin}-{wmax} != {main}")
    return problems


def main():
    wb = xlrd.open_workbook(str(SRC))
    sh = wb.sheet_by_index(0)
    sdrs, pn, sizes = read_main(sh)
    problems = cross_check(wb, sizes)
    chart = {
        "title": str(sh.cell_value(1, 1)).strip(),
        "source_file": SRC.name,
        "imported_on": date.today().isoformat(),
        "sdr_list": sdrs,
        "pn_by_grade_and_sdr": pn,
        "sizes": sizes,
    }
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(chart, indent=1), encoding="utf-8")
    print(f"{len(sizes)} sizes, SDRs {', '.join(sdrs)} -> {OUT}")
    for g, m in pn.items():
        print(f"  {g}: " + ", ".join(f"SDR {s}=PN {p:g}" for s, p in m.items()))
    print("print sheets match main sheet" if not problems else "DIFFERENCES:\n  " + "\n  ".join(problems))


if __name__ == "__main__":
    main()
