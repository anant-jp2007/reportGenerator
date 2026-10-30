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

    rep = s.get("report", {})
    if set(rep.get("grade_order", [])) != {"PE 63", "PE 80", "PE 100"}:
        problems.append("report.grade_order must list PE 63, PE 80 and PE 100")
    for k in ("testing_start_days_after_manufacturing", "testing_end_days_after_start"):
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
        for key in ("range_up_to_limit", "range_above_limit", "above_18dn"):
            if key in v:
                check_range(f"{name} {key}", *v[key])
        if "width_above_half" in v and not 0 < v["width_above_half"]:
            problems.append(f"{name}: width_above_half must be above 0")
        for key in ("fraction_of_chart_max",):
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

    for cell in s.get("fixed_values", {}):
        use(cell, "fixed value")
    for cell in s.get("fixed_text", {}):
        use(cell, "fixed text")

    # companies: own template and header; a header cell may replace a fixed text, never a value cell
    company = s.get("company", {})
    if company.get("selected") not in company.get("profiles", {}):
        problems.append(f"company: selected profile '{company.get('selected')}' not found")
    for cid, profile in company.get("profiles", {}).items():
        what = f"company '{cid}'"
        if not (ROOT / str(profile.get("template", ""))).is_file():
            problems.append(f"{what}: template '{profile.get('template')}' not found")
        cells = {k: v for k, v in profile.items() if CELL_RE.match(k)}
        if not any(str(v).strip() for k, v in cells.items() if int(k[1:]) <= 5):
            problems.append(f"{what}: no header text in rows 1-5")
        for cell in cells:
            if cell in used_cells and used_cells[cell] != "fixed text":
                problems.append(f"{what}: cell {cell} is already used by {used_cells[cell]}")
        for i, img in enumerate(profile.get("images", [])):
            where = f"{what} image {i + 1}"
            if not (ROOT / str(img.get("file", ""))).is_file():
                problems.append(f"{where}: file '{img.get('file')}' not found")
            if not CELL_RE.match(str(img.get("cell", ""))):
                problems.append(f"{where}: '{img.get('cell')}' is not a cell inside A1:L40")
            if not all(isinstance(img.get(k), (int, float)) and img[k] >= 0 for k in ("dx", "dy")) \
                    or not (isinstance(img.get("height"), (int, float)) and img["height"] > 0):
                problems.append(f"{where}: dx and dy must be 0 or more, height above 0 (points)")
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


def company(s: dict, company_id: str | None = None) -> dict:
    """A company's whole profile (name, template, images, header cells); the selected one by default."""
    return s["company"]["profiles"][company_id or s["company"]["selected"]]


def company_profile(s: dict, company_id: str | None = None) -> dict:
    """The cells a company fills (its header, and any fixed text it replaces)."""
    return {k: v for k, v in company(s, company_id).items() if CELL_RE.match(k)}


def companies(s: dict) -> list[tuple[str, str]]:
    """(id, name shown in the company menu), default company first."""
    profiles = s["company"]["profiles"]
    order = [s["company"]["selected"]] + [k for k in profiles if k != s["company"]["selected"]]
    return [(k, profiles[k].get("name") or profiles[k]["A1"]) for k in order]


def with_company(s: dict, company_id: str) -> dict:
    """A copy of the settings with another company selected."""
    if company_id not in s["company"]["profiles"]:
        raise SettingsError(f"company '{company_id}' not found")
    return {**s, "company": {**s["company"], "selected": company_id}}
