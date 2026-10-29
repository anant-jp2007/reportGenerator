"""IS 4984 weight chart lookups (data/is4984_chart.json, built by tools/import_chart.py).

    spec = lookup("PE 100", od=315, pn=5)          # SDR worked out from grade + PN
    spec = lookup("PE 63", od=32, pn=6, sdr=13.6)  # SDR from the size text, checked against PN
    spec.od_min, spec.od_max, spec.ovality_max, spec.pipe_end, spec.wall_min, spec.wall_max
"""
from __future__ import annotations

import json
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path

CHART_FILE = Path(__file__).resolve().parent.parent / "data" / "is4984_chart.json"


class ChartError(ValueError):
    """The size can't be resolved from the chart; the message is shown to the user."""


@dataclass(frozen=True)
class Spec:
    grade: str
    od: float
    pn: float
    sdr: float
    od_min: float
    od_max: float
    ovality_max: float | None
    pipe_end: float | None
    wall_min: float
    wall_max: float
    weight: float | None   # kg per meter


def key(x: float) -> str:
    return f"{float(x):g}"


@lru_cache(maxsize=1)
def chart() -> dict:
    return json.loads(CHART_FILE.read_text(encoding="utf-8"))


def pn_for(grade: str, sdr: float) -> float | None:
    return chart()["pn_by_grade_and_sdr"].get(grade, {}).get(key(sdr))


def sdr_for(grade: str, pn: float) -> float | None:
    for sdr, p in chart()["pn_by_grade_and_sdr"].get(grade, {}).items():
        if p == float(pn):
            return float(sdr)
    return None


def lookup(grade: str, od: float, pn: float, sdr: float | None = None) -> Spec:
    if grade not in chart()["pn_by_grade_and_sdr"]:
        raise ChartError(f"grade '{grade}' not in chart")
    expected_sdr = sdr_for(grade, pn)
    if sdr is None:
        if expected_sdr is None:
            raise ChartError(f"{grade} has no PN {key(pn)} in IS 4984 chart")
        sdr = expected_sdr
    elif expected_sdr is None or expected_sdr != float(sdr):
        right_pn = pn_for(grade, sdr)
        hint = f" (SDR {key(sdr)} is PN {key(right_pn)} for {grade})" if right_pn else ""
        raise ChartError(f"PN {key(pn)} and SDR {key(sdr)} don't match for {grade}{hint}")

    size = chart()["sizes"].get(key(od))
    if size is None:
        raise ChartError(f"OD {key(od)} mm not in IS 4984 chart")
    wall = size["wall"].get(key(sdr))
    if wall is None:
        raise ChartError(f"{key(od)} mm in SDR {key(sdr)} is not made as per IS 4984 chart (shown as --)")
    return Spec(grade, float(od), float(pn), float(sdr),
                size["od_min"], size["od_max"], size["ovality_max"], size["pipe_end"],
                wall["min"], wall["max"], wall["weight"])
