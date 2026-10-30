"""Turn production rows into the full contents of test certificates.

    reports = build_reports(diff.new)         # in grade order; Report No. = column D (MP63/171)
    r = reports[0]
    r.cells["J24"].text    -> "315.50"   (what is printed)
    r.cells["J24"].value   -> 315.5      (number, for Excel)

Every cell of the report area is produced here (company header, fixed texts, header
values, chart limits, random test values), so Excel, PDF and Word print the same thing.
Random values are drawn fresh on every call (user's rule: never repeat a report's values).
"""
from __future__ import annotations

import math
import random
import re
from dataclasses import dataclass, field
from datetime import date, timedelta
from pathlib import Path

from app.production import Row
from app.settings import ROOT, company, company_profile, load_settings

INVALID_FILE_CHARS = re.compile(r'[\\/:*?"<>|]')
# signs written in another font: printed character -> (font, character in that font)
SYMBOLS = {"✆": ("Wingdings", ")")}
# "^{2}" (superscript) or a symbol, for the writers to split texts into runs
MARKUP = re.compile(r"\^\{([^}]*)\}|(" + "|".join(SYMBOLS) + ")")


@dataclass(frozen=True)
class Cell:
    value: object       # str / int / float / date / None
    text: str           # exactly what is printed; "^{2}" marks a superscript, "✆" the phone sign


@dataclass(frozen=True)
class Picture:
    """A picture on the page (logo, stamp): top-left at `cell` + (dx, dy) points, `height` points tall."""
    file: Path
    cell: str
    dx: float
    dy: float
    height: float


@dataclass
class Report:
    number: int | str
    row: Row
    file_name: str                       # without extension
    cells: dict[str, Cell]
    template: Path                       # the company's template (same table, its own header)
    pictures: list[Picture] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


# ---------- formatting ----------

def fmt_number(x: float, decimals: int) -> str:
    return f"{x:.{decimals}f}"


def fmt_plain(x: float) -> str:
    """PN / DN / SDR: 5, 12.5, 13.6, 315 (no trailing zeros)."""
    return f"{x:g}"


def plain_value(x: float):
    return int(x) if float(x).is_integer() else float(x)


def fmt_date(d: date, pattern: str) -> str:
    return (pattern.replace("dd", f"{d.day:02d}").replace("mm", f"{d.month:02d}")
            .replace("yyyy", f"{d.year:04d}"))


# ---------- random values ----------

def pick(rng: random.Random, lo: float, hi: float, decimals: int, step: float | None = None) -> float:
    """Random value in [lo, hi] on a grid of `step` (default: smallest printed unit)."""
    step = step or 10 ** -decimals
    first, last = math.ceil(round(lo / step, 9)), math.floor(round(hi / step, 9))
    if last < first:
        raise ValueError(f"no value with step {step} between {lo} and {hi}")
    return round(rng.randint(first, last) * step, decimals)


def pick_pair(rng, lo, hi, decimals, step, strictly_inside=False, min_gap_fraction=0.0) -> tuple[float, float]:
    """Obtained minimum and maximum: both inside [lo, hi], minimum below maximum.

    strictly_inside: never exactly on lo or hi (falls back to touching only if the
    range is too narrow to hold two values inside it).
    min_gap_fraction: the maximum is at least this share of (hi - lo) above the minimum,
    so the two never look almost equal (3.24 / 3.25). If the range can't hold that gap,
    the widest possible pair is used.
    """
    step = step or 10 ** -decimals
    first, last = math.ceil(round(lo / step, 9)), math.floor(round(hi / step, 9))
    if strictly_inside and (last - 1) - (first + 1) >= 1:
        first, last = first + 1, last - 1
    gap = max(1, math.ceil(round(min_gap_fraction * (hi - lo) / step, 9)))
    if last - first >= gap:
        a = rng.randint(first, last - gap)
        b = rng.randint(a + gap, last)
    else:
        a, b = first, last
    return round(a * step, decimals), round(b * step, decimals)


def pipe_end_range(conf: dict, chart_max: float) -> tuple[float, float]:
    """Half the chart maximum up to half + width_above_half (max 2 -> 1.00-1.50, 7 -> 3.50-4.00)."""
    return chart_max / 2, chart_max / 2 + conf["width_above_half"]


def by_limit(conf: dict, size: float, limit_key: str) -> tuple[float, float]:
    """The range for sizes up to conf[limit_key], or the one for bigger sizes."""
    return tuple(conf["range_up_to_limit" if size <= conf[limit_key] else "range_above_limit"])


# ---------- building one report ----------

def build_report(row: Row, number: int | str | None = None, settings: dict | None = None,
                 rng: random.Random | None = None) -> Report:
    """number: Report No. printed in D7; by default the row's column D (REPORT NO., e.g. MP63/171)."""
    if not row.ok:
        raise ValueError(f"{row.where()}: {'; '.join(row.problems)}")
    s = settings or load_settings()
    rng = rng or random.SystemRandom()
    spec = row.spec
    cells: dict[str, Cell] = {}
    warnings = list(row.warnings)

    def put(ref, value, text=None):
        cells[ref] = Cell(value, text if text is not None else ("" if value is None else str(value)))

    def put_num(ref, x, decimals):
        if x is None:
            put(ref, None, "")
        else:
            x = round(x, decimals)
            put(ref, int(x) if decimals == 0 else x, fmt_number(x, decimals))

    # company header and fixed texts
    # fixed texts, then the company's header (which may replace a fixed text, e.g. G39 left empty)
    for ref, text in s["fixed_text"].items():
        put(ref, text, str(text))
    for ref, text in company_profile(s).items():
        put(ref, text or None, str(text))

    # header values from the production row
    rep, hc = s["report"], s["header_cells"]
    made = row.date             # None: the register had no readable date, so all three print blank
    start = made + timedelta(days=rep["testing_start_days_after_manufacturing"]) if made else None
    end = start + timedelta(days=rep["testing_end_days_after_start"]) if made else None
    # the report no. may be typed by the user on the review screen ("BPP100/EX008"); digits stay a number
    number_text = str(row.report_no if number is None else number).strip()
    put(hc["report_no"], int(number_text) if number_text.isdigit() else number_text, number_text)
    for key, d in (("date_of_manufacturing", made), ("testing_start", start), ("testing_end", end)):
        put(hc[key], d, fmt_date(d, rep["date_format"]) if d else "")
    put(hc["batch"], row.batch)
    put(hc["grade"], row.grade)
    put(hc["pn"], plain_value(spec.pn), fmt_plain(spec.pn))
    put(hc["dn"], plain_value(spec.od), fmt_plain(spec.od))
    put(hc["sdr"], plain_value(spec.sdr), fmt_plain(spec.sdr))

    # specification limits from the chart
    cc = s["chart_cells"]
    for ref, name in cc.items():
        if ref != "decimals":
            put_num(ref, getattr(spec, name), cc["decimals"])

    # test values obtained
    rv = s["random_values"]
    for name in ("longitudinal_reversion", "carbon_black_content",
                 "density", "elongation_at_break", "yield_strength"):
        c = rv[name]
        put_num(c["cell"], pick(rng, c["min"], c["max"], c["decimals"], c.get("step")), c["decimals"])

    for ref, text in s["fixed_values"].items():
        put(ref, text)

    # stripes: width by size (DN), depth by the chart's wall minimum
    for name, size, limit in (("stripe_width", spec.od, "dn_limit"), ("stripe_depth", spec.wall_min, "wall_limit")):
        c = rv[name]
        lo, hi = by_limit(c, size, limit)
        put_num(c["cell"], pick(rng, lo, hi, c["decimals"], c.get("step")), c["decimals"])

    # coiling: a little above 18 x DN for coilable sizes, a dash for bigger ones
    c = rv["coiling"]
    if spec.od <= c["dn_limit"]:
        lo, hi = (18 * spec.od + extra for extra in c["above_18dn"])
        put_num(c["cell"], pick(rng, lo, hi, c["decimals"], c.get("step")), c["decimals"])
    else:
        put(c["cell"], c["text_above_limit"])

    c = rv["pipe_end"]
    if spec.pipe_end is None:
        put_num(c["cell"], None, 0)
        warnings.append("chart has no pipe end tolerance for this size; left blank")
    else:
        lo, hi = pipe_end_range(c, spec.pipe_end)
        put_num(c["cell"], pick(rng, lo, hi, c["decimals"], c.get("step")), c["decimals"])

    for name, lo, hi in (("outside_diameter", spec.od_min, spec.od_max),
                         ("wall_thickness", spec.wall_min, spec.wall_max)):
        c = rv[name]
        a, b = pick_pair(rng, lo, hi, c["decimals"], c.get("step"), c.get("strictly_inside", False),
                         c.get("min_gap_fraction", 0.0))
        put_num(c["cells"]["minimum"], a, c["decimals"])
        put_num(c["cells"]["maximum"], b, c["decimals"])

    c = rv["ovality"]
    if spec.ovality_max is None:
        put_num(c["cell"], None, 0)
        warnings.append("chart has no maximum ovality for this size; ovality left blank")
    else:
        lo, hi = c["fraction_of_chart_max"]
        put_num(c["cell"], pick(rng, spec.ovality_max * lo, spec.ovality_max * hi,
                                c["decimals"], c.get("step")), c["decimals"])

    c = rv["melt_flow_rate"]
    lo, hi = c["by_grade"][row.grade]
    put_num(c["cell"], pick(rng, lo, hi, c["decimals"], c.get("step")), c["decimals"])

    if row.report_no:
        file_name = INVALID_FILE_CHARS.sub("-", row.report_no).strip()
    else:
        file_name = f"{row.grade.replace(' ', '')}-row{row.excel_row}"
    co = company(s)
    pictures = [Picture(ROOT / p["file"], p["cell"], p["dx"], p["dy"], p["height"]) for p in co.get("images", [])]
    return Report(number_text, row, file_name, cells, ROOT / co["template"], pictures, warnings)


def build_reports(rows: list[Row], settings: dict | None = None,
                  rng: random.Random | None = None) -> list[Report]:
    """Reports in grade order (PE 63, PE 80, PE 100, top to bottom), Report No. from column D.

    Rows with problems are left out (the caller lists them for the user). File names are
    made unique by adding -2, -3... so no report overwrites another.
    """
    s = settings or load_settings()
    order = s["report"]["grade_order"]
    ok = sorted((r for r in rows if r.ok), key=lambda r: (order.index(r.grade), r.excel_row))
    reports, used = [], {}
    for row in ok:
        rep = build_report(row, None, s, rng)
        base = rep.file_name
        used[base.lower()] = used.get(base.lower(), 0) + 1
        if used[base.lower()] > 1:
            rep.file_name = f"{base}-{used[base.lower()]}"
        reports.append(rep)
    return reports
