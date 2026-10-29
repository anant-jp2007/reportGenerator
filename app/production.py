"""Read the production register and find rows that are new since an older copy.

Each grade has its own sheet ("Production 63", "Production 80", "Production 100").
Columns: A Sr. No. | B Date | C Size | D Report No. | E Batch | F Meters | G Total | H MT | I Remark.

Sr. No. restarts every April, so a row is identified by (grade, financial year, Sr. No.).
A row counts as "filled" only when it has a Sr. No., a date and a size; the sheets
pre-type report numbers on empty rows, and those are ignored.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

SHEETS = {"Production 63": "PE 63", "Production 80": "PE 80", "Production 100": "PE 100"}
GRADE_ORDER = ["PE 63", "PE 80", "PE 100"]

SIZE_RE = re.compile(
    r"^\s*(\d+(?:\.\d+)?)\s*[xX]\s*(\d+(?:\.\d+)?)\s*(?:\(\s*SDR\s*(\d+(?:\.\d+)?)\s*\))?\s*$"
)


@dataclass
class Row:
    grade: str
    excel_row: int          # 1-based row number in the sheet, for messages
    sr_no: int
    date: date
    size_text: str
    report_no: str          # column D, used only for the file name
    batch: str
    meters: float | None
    od: float | None = None
    pn: float | None = None
    sdr: float | None = None
    problems: list[str] = field(default_factory=list)

    @property
    def key(self) -> tuple:
        return (self.grade, financial_year(self.date), self.sr_no)

    def content(self) -> tuple:
        return (self.date, self.size_text.strip().upper(), self.batch.strip(), self.meters)


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


def _to_date(v, datemode=0) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)) and v > 0:
        import xlrd
        return xlrd.xldate.xldate_as_datetime(v, datemode).date()
    if isinstance(v, str):
        for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y"):
            try:
                return datetime.strptime(v.strip(), fmt).date()
            except ValueError:
                pass
    return None


def _sheet_rows(path: Path):
    """Yield (sheet_name, [row values A..I]) for each grade sheet, xls or xlsx."""
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
                for c in range(min(9, sh.ncols)):
                    cell = sh.cell(r, c)
                    v = cell.value
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        v = _to_date(v, wb.datemode)
                    vals.append(v)
                rows.append(vals + [None] * (9 - len(vals)))
            yield name, rows
    else:
        import openpyxl
        wb = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
        for name in SHEETS:
            if name not in wb.sheetnames:
                continue
            rows = [list(r[:9]) + [None] * (9 - len(r[:9]))
                    for r in wb[name].iter_rows(values_only=True)]
            yield name, rows


def read_production(path: str | Path) -> tuple[list[Row], list[str]]:
    """Return (filled rows, warnings). Missing sheets are reported as warnings."""
    path = Path(path)
    rows: list[Row] = []
    warnings: list[str] = []
    seen = set()
    for name, values in _sheet_rows(path):
        seen.add(name)
        grade = SHEETS[name]
        for i, (a, b, c, d, e, f, *_rest) in enumerate(values, start=1):
            sr = _num(a)
            size = _text(c)
            if sr is None or not size:
                continue  # header, month title, TOTAL row, or pre-numbered empty row
            d_ = _to_date(b)
            if d_ is None:
                warnings.append(f"{name} row {i}: Sr. No. {int(sr)} has a size but no readable date, skipped")
                continue
            row = Row(grade, i, int(sr), d_, size, _text(d), _text(e), _num(f))
            row.od, row.pn, row.sdr = parse_size(size)
            if row.od is None:
                row.problems.append(f"size '{size}' not understood")
            elif row.sdr is None:
                row.problems.append("SDR missing in size")
            if not row.batch:
                row.problems.append("batch number missing")
            rows.append(row)
    for name in SHEETS:
        if name not in seen:
            warnings.append(f"sheet '{name}' not found")
    return rows, warnings


@dataclass
class Diff:
    new: list[Row]
    modified: list[tuple[Row, Row]]   # (old, new)
    warnings: list[str]


def find_new_rows(old_path, new_path) -> Diff:
    old_rows, old_warn = read_production(old_path)
    new_rows, new_warn = read_production(new_path)
    old_by_key = {r.key: r for r in old_rows}
    new, modified = [], []
    for r in new_rows:
        prev = old_by_key.get(r.key)
        if prev is None:
            new.append(r)
        elif prev.content() != r.content():
            modified.append((prev, r))
    new.sort(key=lambda r: (GRADE_ORDER.index(r.grade), r.excel_row))
    warnings = [f"old file: {w}" for w in old_warn] + [f"new file: {w}" for w in new_warn]
    return Diff(new, modified, warnings)
