"""Read the production register and find rows that are new since an older copy.

Each grade has its own sheet ("Production 63", "Production 80", "Production 100").
Columns are found by their HEADING, not their position, so extra columns (Brand Name,
Meters, Remark, ...) or a moved column don't matter. Only five columns are used:
Sr. No., Date, Size, Report No., Batch Number.

Sr. No. restarts every April, so a row is identified by (grade, financial year, Sr. No.).
A row counts as "filled" only when it has a Sr. No., a date and a size; the sheets
pre-type report numbers on empty rows, and those are ignored.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from app.chart import ChartError, Spec, lookup

SHEETS = {"Production 63": "PE 63", "Production 80": "PE 80", "Production 100": "PE 100"}
GRADE_ORDER = ["PE 63", "PE 80", "PE 100"]
HEADER_SEARCH_ROWS = 30

SIZE_RE = re.compile(
    r"^\s*(\d+(?:\.\d+)?)\s*[xX×]\s*(\d+(?:\.\d+)?)\s*(?:\(\s*SDR\s*(\d+(?:\.\d+)?)\s*\))?\s*$"
)

# Heading matching, on cleaned text (lowercase, letters and digits only).
# Tier 1: the exact heading; tier 2: known variations; tier 3: contains the keyword.
# The best tier wins; two headings in the best tier is ambiguous and stops the sheet.
COLUMNS = {
    "sr_no": {
        "label": "Sr. No.",
        "exact": {"srno"},
        "variants": {"sno", "slno", "srnumber", "serialno", "serialnumber", "srn", "sr"},
        "keyword": None,
    },
    "date": {
        "label": "Date",
        "exact": {"date"},
        "variants": {"productiondate", "dateofproduction", "mfgdate", "manufacturingdate",
                     "dateofmanufacturing", "dateofmanufacture", "proddate"},
        "keyword": "date",
    },
    "size": {
        "label": "Size",
        "exact": {"size"},
        "variants": {"pipesize", "sizeofpipe", "sizemm", "pipesizemm"},
        "keyword": "size",
    },
    "report_no": {
        "label": "Report No.",
        "exact": {"reportno"},
        "variants": {"reportnumber", "reportnum", "testreportno", "testreportnumber", "rptno"},
        "keyword": "report",
    },
    "batch": {
        "label": "Batch Number",
        "exact": {"batchnumber"},
        "variants": {"batchno", "batchnum", "batch"},
        "keyword": "batch",
    },
}


def clean_heading(v) -> str:
    return re.sub(r"[^a-z0-9]", "", str(v or "").lower())


class SheetError(ValueError):
    """A sheet can't be read safely; the message is shown to the user."""


@dataclass
class Row:
    grade: str
    excel_row: int          # 1-based row number in the sheet, for messages
    sr_no: int
    date: date
    size_text: str
    report_no: str          # column "Report No.", used only for the file name
    batch: str
    od: float | None = None
    pn: float | None = None
    sdr: float | None = None        # as written in the size text (None if missing)
    spec: Spec | None = None        # chart limits; None when the row can't be resolved
    problems: list[str] = field(default_factory=list)   # stop this row's report
    warnings: list[str] = field(default_factory=list)   # report still generated

    @property
    def key(self) -> tuple:
        return (self.grade, financial_year(self.date), self.sr_no)

    @property
    def ok(self) -> bool:
        return not self.problems

    def where(self) -> str:
        return f"Production {self.grade[3:]} row {self.excel_row} (Sr. {self.sr_no}, {self.size_text!r})"


def financial_year(d: date) -> int:
    return d.year if d.month >= 4 else d.year - 1


def parse_size(text: str) -> tuple[float | None, float | None, float | None]:
    m = SIZE_RE.match(text or "")
    if not m:
        return None, None, None
    od, pn, sdr = (float(g) if g else None for g in m.groups())
    return od, pn, sdr


def _num(v):
    if v in (None, ""):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _to_date(v) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y", "%d-%m-%y", "%d/%m/%y"):
            try:
                return datetime.strptime(v.strip(), fmt).date()
            except ValueError:
                pass
    return None


def _sheets(path: Path):
    """Yield (sheet_name, rows) for each grade sheet present; rows are lists of cell values."""
    if path.suffix.lower() == ".xls":
        import xlrd
        wb = xlrd.open_workbook(str(path))
        for name in SHEETS:
            if name not in wb.sheet_names():
                continue
            sh = wb.sheet_by_name(name)
            rows = []
            for r in range(sh.nrows):
                vals = []
                for c in range(sh.ncols):
                    cell = sh.cell(r, c)
                    v = cell.value
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        v = xlrd.xldate.xldate_as_datetime(v, wb.datemode)
                    elif cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                        v = None
                    vals.append(v)
                rows.append(vals)
            yield name, rows
    else:
        import openpyxl
        wb = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
        for name in SHEETS:
            if name in wb.sheetnames:
                yield name, [list(r) for r in wb[name].iter_rows(values_only=True)]


def find_columns(name: str, rows: list[list]) -> tuple[int, dict[str, int]]:
    """Return (header row index, {field: column index}). Raises SheetError if unsure."""
    for r, values in enumerate(rows[:HEADER_SEARCH_ROWS]):
        cleaned = [clean_heading(v) for v in values]
        found, problems, unclear = {}, [], set()
        for fld, spec in COLUMNS.items():
            tiers = (
                [c for c, h in enumerate(cleaned) if h in spec["exact"]],
                [c for c, h in enumerate(cleaned) if h in spec["variants"]],
                [c for c, h in enumerate(cleaned) if spec["keyword"] and h and spec["keyword"] in h],
            )
            best = next((t for t in tiers if t), [])
            if len(best) == 1:
                found[fld] = best[0]
            elif len(best) > 1:
                unclear.add(fld)
                heads = ", ".join(f"'{_text(values[c])}'" for c in best)
                problems.append(f"more than one column could be '{spec['label']}' ({heads}); rename one")
        if not {"sr_no", "date", "size"} <= found.keys() | unclear:
            continue  # not the header row
        if problems:
            raise SheetError(f"{name}: " + "; ".join(problems))
        missing = [COLUMNS[f]["label"] for f in COLUMNS if f not in found]
        if missing:
            raise SheetError(f"{name}: " + ", ".join(f"'{m}'" for m in missing) + " column not found")
        return r, found
    raise SheetError(f"{name}: heading row with Sr. No., Date and Size not found "
                     f"in the first {HEADER_SEARCH_ROWS} rows")


def _resolve(row: Row) -> None:
    """Work out PN/SDR and chart limits; record problems (skip) and warnings (generate anyway)."""
    row.od, row.pn, row.sdr = parse_size(row.size_text)
    if row.od is None:
        row.problems.append(f"size '{row.size_text}' not understood (expected like '63 X 10 ( SDR 13.6 )')")
        return
    try:
        row.spec = lookup(row.grade, row.od, row.pn, row.sdr)
    except ChartError as e:
        row.problems.append(str(e))
        return
    if row.sdr is None:
        row.warnings.append(f"SDR not written; taken as SDR {row.spec.sdr:g} from {row.grade} PN {row.pn:g}")
    if not row.batch:
        row.warnings.append("batch number is empty")
    if not row.report_no:
        row.warnings.append("Report No. is empty; file named by grade and row instead")


@dataclass
class ReadResult:
    rows: list[Row]
    errors: dict[str, str]          # grade -> why the sheet couldn't be read
    notes: list[str]                # rows skipped while reading (e.g. no date)


def read_production(path: str | Path) -> ReadResult:
    path = Path(path)
    rows: list[Row] = []
    errors: dict[str, str] = {}
    notes: list[str] = []
    present = set()
    for name, values in _sheets(path):
        present.add(name)
        grade = SHEETS[name]
        try:
            header, col = find_columns(name, values)
        except SheetError as e:
            errors[grade] = str(e)
            continue

        def get(vals, fld):
            c = col[fld]
            return vals[c] if c < len(vals) else None

        for i, vals in enumerate(values[header + 1:], start=header + 2):
            sr = _num(get(vals, "sr_no"))
            size = _text(get(vals, "size"))
            if sr is None or not size:
                continue  # month title, TOTAL row, or pre-numbered empty row
            d = _to_date(get(vals, "date"))
            if d is None:
                notes.append(f"{name} row {i}: Sr. {int(sr)} has a size but no readable date; skipped")
                continue
            row = Row(grade, i, int(sr), d, size, _text(get(vals, "report_no")), _text(get(vals, "batch")))
            _resolve(row)
            rows.append(row)
    for name, grade in SHEETS.items():
        if name not in present:
            errors[grade] = f"sheet '{name}' not found"
    return ReadResult(rows, errors, notes)


@dataclass
class Diff:
    new: list[Row]                  # in report order: PE 63, PE 80, PE 100, top to bottom
    errors: list[str]               # grades that couldn't be compared, with the reason
    notes: list[str]


def find_new_rows(old_path, new_path) -> Diff:
    """Rows filled in the new file that weren't filled in the old one.

    Rows that exist in both files are never reported again, even if edited (user's decision).
    A grade whose sheet can't be read in either file is skipped entirely, with the reason;
    otherwise every row of it would look new.
    """
    old, new = read_production(old_path), read_production(new_path)
    errors = [f"old file: {e}" for e in old.errors.values()] + [f"new file: {e}" for e in new.errors.values()]
    blocked = old.errors.keys() | new.errors.keys()
    old_keys = {r.key for r in old.rows}
    fresh = [r for r in new.rows if r.grade not in blocked and r.key not in old_keys]
    fresh.sort(key=lambda r: (GRADE_ORDER.index(r.grade), r.excel_row))
    return Diff(fresh, errors, [f"new file: {n}" for n in new.notes])
