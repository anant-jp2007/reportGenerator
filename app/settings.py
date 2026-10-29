"""Load and check config/settings.json.

    from app.settings import load_settings
    s = load_settings()          # raises SettingsError listing every problem found
"""
from __future__ import annotations

import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SETTINGS_FILE = ROOT / "config" / "settings.json"
CELL_RE = re.compile(r"^[A-L]([1-9]|[1-3]\d|40)$")   # report area A1:L40


class SettingsError(ValueError):
    pass


def _strip_help(obj):
    """Drop "_help"/"_readme" notes so code only sees real settings."""
    if isinstance(obj, dict):
        return {k: _strip_help(v) for k, v in obj.items() if not k.startswith("_")}
    if isinstance(obj, list):
        return [_strip_help(v) for v in obj]
    return obj


def validate(s: dict) -> list[str]:
    problems = []
    used_cells: dict[str, str] = {}

    def use(cell, what):
        if not isinstance(cell, str) or not CELL_RE.match(cell):
            problems.append(f"{what}: '{cell}' is not a cell inside A1:L40")
        elif cell in used_cells:
            problems.append(f"{what}: cell {cell} is already used by {used_cells[cell]}")
        else:
            used_cells[cell] = what

    company = s.get("company", {})
    profile = company.get("profiles", {}).get(company.get("selected"))
    if profile is None:
        problems.append(f"company: selected profile '{company.get('selected')}' not found")
    else:
        for cell in ("A1", "A2", "A3", "A4", "A5"):
            if not str(profile.get(cell, "")).strip():
                problems.append(f"company '{company['selected']}': {cell} is empty")
            use(cell, f"company {cell}")

    rep = s.get("report", {})
    if set(rep.get("grade_order", [])) != {"PE 63", "PE 80", "PE 100"}:
        problems.append("report.grade_order must list PE 63, PE 80 and PE 100")
    for k in ("testing_start_days_after_manufacturing", "testing_end_days_after_start", "number_starts_at"):
        if not isinstance(rep.get(k), int) or rep[k] < 0:
            problems.append(f"report.{k} must be a whole number >= 0")

    for name, cell in s.get("header_cells", {}).items():
        use(cell, f"header {name}")
    for cell, name in s.get("chart_cells", {}).items():
        if cell != "decimals":
            use(cell, f"chart {name}")

    def check_range(what, lo, hi):
        if not (isinstance(lo, (int, float)) and isinstance(hi, (int, float)) and lo <= hi):
            problems.append(f"{what}: range {lo}-{hi} is not valid (min must be <= max)")

    for name, v in s.get("random_values", {}).items():
        if "cell" in v:
            use(v["cell"], name)
        for label, cell in v.get("cells", {}).items():
            use(cell, f"{name} {label}")
        if not isinstance(v.get("decimals"), int):
            problems.append(f"{name}: decimals missing")
        if "min" in v or "max" in v:
            check_range(name, v.get("min"), v.get("max"))
        for mx, (lo, hi) in v.get("range_by_chart_max", {}).items():
            check_range(f"{name} (chart max {mx})", lo, hi)
            if hi > float(mx):
                problems.append(f"{name}: range for chart max {mx} goes above {mx}")
        for key in ("other_max_fraction", "fraction_of_chart_max"):
            if key in v:
                lo, hi = v[key]
                check_range(f"{name} {key}", lo, hi)
                if not 0 < lo <= hi < 1:
                    problems.append(f"{name}: {key} must be between 0 and 1")
        if "by_grade" in v:
            for g, (lo, hi) in v["by_grade"].items():
                check_range(f"{name} {g}", lo, hi)
            r = v["by_grade"]
            if all(g in r for g in ("PE 100", "PE 80", "PE 63")) and not (
                    r["PE 100"][1] < r["PE 80"][0] and r["PE 80"][1] < r["PE 63"][0]):
                problems.append(f"{name}: ranges overlap, so PE 100 < PE 80 < PE 63 is not guaranteed")

    for cell in s.get("empty_cells", {}).get("cells", []):
        use(cell, "empty value")
    for cell in s.get("fixed_text", {}):
        use(cell, "fixed text")

    sig = s.get("signature", {})
    if not (ROOT / sig.get("file", "")).is_file():
        problems.append(f"signature file '{sig.get('file')}' not found")
    return problems


def load_settings(path: Path = SETTINGS_FILE) -> dict:
    try:
        raw = json.loads(Path(path).read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        raise SettingsError(f"{Path(path).name} is not valid JSON: line {e.lineno}, column {e.colno}: {e.msg}")
    s = _strip_help(raw)
    problems = validate(s)
    if problems:
        raise SettingsError("Problems in settings:\n  " + "\n  ".join(problems))
    return s


def company_profile(s: dict) -> dict:
    return s["company"]["profiles"][s["company"]["selected"]]
