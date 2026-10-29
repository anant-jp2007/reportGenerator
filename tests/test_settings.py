"""Settings must load cleanly and describe exactly the template layout.

    python tests/test_settings.py
"""
import copy
import sys
from pathlib import Path

import openpyxl
from openpyxl.cell.rich_text import CellRichText

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from app.settings import company_profile, load_settings, validate  # noqa: E402

failures = 0


def check(name, cond, detail=""):
    global failures
    failures += not cond
    print(("PASS " if cond else "FAIL ") + name + (f"  [{detail}]" if detail else ""))


s = load_settings()
check("settings load with no problems", True)

# Every text cell of the template is described in settings, with the same text.
ws = openpyxl.load_workbook(ROOT / "templates/test_report.xlsx", rich_text=True).active


def plain(v):
    if isinstance(v, CellRichText):
        return "".join(b if isinstance(b, str) else
                       ("^{%s}" % b.text if b.font.vertAlign == "superscript" else b.text) for b in v)
    return v


template_texts = {c.coordinate: plain(c.value) for row in ws.iter_rows(max_row=40) for c in row if c.value is not None}
settings_texts = {**s["fixed_text"], **company_profile(s)}
missing = sorted(set(template_texts) - set(settings_texts))
extra = sorted(set(settings_texts) - set(template_texts))
diff = [c for c in template_texts if c in settings_texts and template_texts[c] != settings_texts[c]]
check("every template text is in settings", not missing, missing)
check("settings has no text the template lacks", not extra, extra)
check("texts are identical", not diff, [(c, template_texts[c], settings_texts[c]) for c in diff])

# Cells the app fills must be empty in the template.
filled = [s["header_cells"][k] for k in s["header_cells"]]
filled += [c for c in s["chart_cells"] if c != "decimals"]
for v in s["random_values"].values():
    filled += [v["cell"]] if "cell" in v else list(v["cells"].values())
filled += s["empty_cells"]["cells"]
not_empty = [c for c in filled if ws[c].value is not None]
check("all value cells are empty in the template", not not_empty, not_empty)
check("PN/DN/SDR cells print decimals (13.6 not 14)", all(ws[c].number_format == "General" for c in ("J8", "J9", "J10")))


# The checks catch broken edits.
def broken(edit):
    t = copy.deepcopy(s)
    edit(t)
    return validate(t)


check("overlapping MFR ranges are caught",
      any("overlap" in p for p in broken(lambda t: t["random_values"]["melt_flow_rate"]["by_grade"].__setitem__("PE 80", [0.2, 0.5]))))
check("min above max is caught",
      any("not valid" in p for p in broken(lambda t: t["random_values"]["density"].update(min=960, max=950))))
check("pipe end range above chart max is caught",
      any("goes above" in p for p in broken(lambda t: t["random_values"]["pipe_end"]["range_by_chart_max"].__setitem__("2", [1, 2.5]))))
check("two values in one cell is caught",
      any("already used" in p for p in broken(lambda t: t["random_values"]["density"].__setitem__("cell", "G33"))))
check("missing company profile is caught",
      any("not found" in p for p in broken(lambda t: t["company"].__setitem__("selected", "xyz"))))

print("\nALL PASS" if not failures else f"\n{failures} FAILED")
sys.exit(1 if failures else 0)
