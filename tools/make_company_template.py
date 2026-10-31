"""Build a company's certificate template from templates/test_report.xlsx.

    python tools/make_company_template.py abhay

The table (rows 6 onwards) stays exactly the same; only the header changes, as described by the
company's profile in config/settings.json:
- template_layout.merges: the header's merged cells (rows 1-5 are unmerged first);
- template_layout.fonts: {cell: [size, bold]} for the header cells (Times New Roman, centred, wrapped);
- template_layout.row_heights: {row: height in points}, e.g. a taller header and, to keep the page at
  100% on A4, shorter empty rows under the table;
- the header texts (and any fixed text the company replaces, '' = empty).
Logo and stamp are not stored in the template: every certificate places them (profile "images").
The website's "Add a company" builds templates the same way (app/new_company.py).
"""
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from app.new_company import BASE, build_template  # noqa: E402
from app.settings import _strip_help  # noqa: E402


def main(company_id: str) -> None:
    # read the settings without checking them: the template being built doesn't exist yet
    s = _strip_help(json.loads((ROOT / "config" / "settings.json").read_text(encoding="utf-8")))
    profile = s["company"]["profiles"][company_id]
    out = ROOT / profile["template"]
    if out.resolve() == BASE.resolve():
        sys.exit("This company uses the base template; build that one with tools/make_template.py.")
    print(f"template written: {build_template(profile, out)}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    main(sys.argv[1])
