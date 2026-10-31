"""Add a company from one of its own certificates (a sample), without editing the settings by hand.

    sample = read_sample("their report.xls")        # header lines (text, size, bold), logo, stamp
    draft = make_draft("Shree Pipes", sample)        # 5 header rows, as on our certificates
    profile = make_profile(draft, "shree", logo_ratio, ...)   # names, merges, fonts, row heights
    build_template(profile, out)                     # its template: our table, its header
    company_id = save_company(draft, ...)            # files written, settings.json updated

Every company's certificate is our table (rows 6-42, the same for all) under its own header in rows
1-5. The sample's header is everything above its "Test Report No." row: its lines are put in our five
header rows (name, two rows of details, contact line, title), in Times New Roman like the others, with
row heights worked out from the text. A logo goes top left (the text then starts in column C), a stamp
bottom right. Everything is shown on a preview before it's saved.

The page must stay one A4 page at 100%, so the header has about 116 pt: if the text needs more, the
smaller lines get a smaller font, then the name; if it still doesn't fit, the company can't be added
here (the message says so).
"""
from __future__ import annotations

import io
import json
import math
import re
import shutil
from copy import copy
from dataclasses import asdict, dataclass, field
from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.cell.rich_text import CellRichText
from openpyxl.styles import Alignment, Border, Font, Side
from PIL import Image

from app.layout import ROW_SCALE, load_layout
from app.render_pdf import CELL_PAD, _segments, _width, _wrap, font_for
from app.report import SYMBOLS
from app.settings import CELL_RE, ROOT, SETTINGS_FILE, _strip_help, load_settings, validate

BASE = ROOT / "templates" / "test_report.xlsx"
FONT = "Times New Roman"
ROOM = 842 - 15.5905 - 1          # A4 minus the top margin, 1 pt to spare (as tests/test_settings.py checks)
BOTTOM_ROWS = {"38": 6, "39": 15, "40": 13.5, "41": 13.5, "42": 13.5}   # as for the other companies
STAMP = {"cell": "J38", "dx": 22, "dy": 3, "height": 54}
LOGO_DX, LOGO_DY, LOGO_MAX_H, LOGO_MIN_W = 6, 4, 70, 56
CONTACT = re.compile(r"@|www\.|\bph(one)?\b|\bmo(b|bile)?\b|tel|☎|✆|\+91|\b\d{5}\s?\d{5}\b", re.I)
TITLE = re.compile(r"test\s*(certificate|report)|is\s*:?\s*4984", re.I)


class CompanyError(ValueError):
    """The sample can't be used; the message is shown to the user."""


@dataclass
class Line:
    text: str
    size: float
    bold: bool


@dataclass
class Sample:
    lines: list[Line]                  # the header, one line per sheet row (cells of a row joined)
    logo: bytes | None = None          # PNG
    stamp: bytes | None = None         # PNG
    notes: list[str] = field(default_factory=list)


@dataclass
class Draft:
    """What the add-company screen shows and lets the user change: the five header rows."""
    name: str
    texts: list[str]                   # rows 1-5
    sizes: list[float]
    bold: list[bool]
    tested_by: bool = True             # keep "Tested by :-" at the bottom (a stamp usually replaces it)

    def to_json(self) -> str:
        return json.dumps(asdict(self), ensure_ascii=False)

    @classmethod
    def from_json(cls, text: str) -> "Draft":
        return cls(**json.loads(text))


# ---------- reading the sample ----------

def _norm(text: str) -> str:
    return re.sub(r"[^a-z]", "", str(text).lower())


def _png(data: bytes) -> bytes:
    with Image.open(io.BytesIO(data)) as im:
        buf = io.BytesIO()
        im.convert("RGBA").save(buf, "PNG")
    return buf.getvalue()


def _header_end(rows: list[list[str]]) -> int:
    """Index of the "Test Report No." row: the header is above it."""
    for i, cells in enumerate(rows[:40]):
        if any("reportno" in _norm(t) for t in cells):
            return i
    raise CompanyError("the sample has no \"Test Report No.\" row in its first 40 rows, so its header "
                       "can't be found. Choose one of the company's test certificates")


def read_sample(path: str | Path) -> Sample:
    path = Path(path)
    try:
        return _read_xls(path) if path.suffix.lower() == ".xls" else _read_xlsx(path)
    except CompanyError:
        raise
    except Exception as e:                      # not an Excel file, damaged...
        raise CompanyError(f"the sample couldn't be opened as an Excel file ({e})") from e


SYMBOL_OF = {(font.lower(), char): sym for sym, (font, char) in SYMBOLS.items()}   # ("wingdings", ")") -> "✆"


def _xlsx_text(value) -> str:
    """A cell's text, with the phone sign (Wingdings ")") and superscripts written as on our certificates."""
    if not isinstance(value, CellRichText):
        return "" if value is None else str(value).strip()
    out = []
    for block in value:
        if isinstance(block, str):
            out.append(block)
            continue
        font = (block.font.rFont or "").lower() if block.font else ""
        text = "".join(SYMBOL_OF.get((font, ch), ch) for ch in block.text)
        out.append(f"^{{{text}}}" if block.font is not None and block.font.vertAlign == "superscript" else text)
    return "".join(out).strip()


def _read_xlsx(path: Path) -> Sample:
    wb = openpyxl.load_workbook(path, rich_text=True)
    for ws in wb.worksheets:
        rows = [[_xlsx_text(c.value) for c in row] for row in ws.iter_rows(max_row=40)]
        try:
            end = _header_end(rows)
        except CompanyError:
            continue
        lines = []
        for r in range(1, end + 1):
            if ws.row_dimensions[r].hidden:
                continue
            cells = [c for c in ws[r] if _xlsx_text(c.value)]
            if cells:
                lines.append(Line("  ".join(_xlsx_text(c.value) for c in cells),
                                  max(float(c.font.sz or 11) for c in cells), any(bool(c.font.b) for c in cells)))
        sample = Sample(lines)
        for img in ws._images:
            row = getattr(getattr(img.anchor, "_from", None), "row", None)
            if row is None:
                continue
            if row < end and sample.logo is None:
                sample.logo = _png(img._data())
            elif row >= 30 and sample.stamp is None:
                sample.stamp = _png(img._data())
        return sample
    raise CompanyError("the sample has no \"Test Report No.\" row in its first 40 rows, so its header "
                       "can't be found. Choose one of the company's test certificates")


def _read_xls(path: Path) -> Sample:
    import xlrd
    book = xlrd.open_workbook(str(path), formatting_info=True)
    for sh in book.sheets():
        rows = [[str(sh.cell_value(r, c)) for c in range(sh.ncols)] for r in range(min(sh.nrows, 40))]
        try:
            end = _header_end(rows)
        except CompanyError:
            continue
        lines = []
        for r in range(end):
            info = sh.rowinfo_map.get(r)
            if info is not None and info.hidden:
                continue
            cells = [c for c in range(sh.ncols) if str(sh.cell_value(r, c)).strip()]
            if not cells:
                continue
            fonts = [book.font_list[book.xf_list[sh.cell_xf_index(r, c)].font_index] for c in cells]
            text = "  ".join(_xls_text(book, sh, r, c) for c in cells)
            lines.append(Line(text, max(f.height / 20 for f in fonts), any(f.bold or f.weight >= 700 for f in fonts)))
        return Sample(lines, notes=["Pictures can't be read from an .xls file: if the company has a logo or stamp, "
                                    "choose them below (or save the sample as .xlsx in Excel and choose that)."])
    raise CompanyError("the sample has no \"Test Report No.\" row in its first 40 rows, so its header "
                       "can't be found. Choose one of the company's test certificates")


def _xls_text(book, sh, r: int, c: int) -> str:
    """Like _xlsx_text for an .xls cell: its text runs tell which characters are in Wingdings."""
    v = sh.cell_value(r, c)
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    text = str(v)
    runs = sh.rich_text_runlist_map.get((r, c))
    if runs:
        bounds = [off for off, _ in runs] + [len(text)]
        parts = [text[:bounds[0]]]
        for (off, fi), end in zip(runs, bounds[1:]):
            font = book.font_list[fi]
            piece = "".join(SYMBOL_OF.get((font.name.lower(), ch), ch) for ch in text[off:end])
            parts.append(f"^{{{piece}}}" if font.escapement == 1 else piece)
        text = "".join(parts)
    return text.strip()


# ---------- the five header rows ----------

def make_draft(name: str, sample: Sample, has_stamp: bool = False) -> Draft:
    """Put the sample's header lines into our five rows: name, details, details, contact line, title."""
    lines = [ln for ln in sample.lines if ln.text.strip()]
    if not lines:
        raise CompanyError("the sample's header (above \"Test Report No.\") is empty")
    titles = [ln for ln in lines if TITLE.search(ln.text)]
    title = titles[-1] if titles else lines[-1]
    rest = [ln for ln in lines if ln is not title]
    head = max(rest, key=lambda ln: ln.size) if rest else Line(name.upper(), 24, True)
    rest = [ln for ln in rest if ln is not head]
    contact = [ln for ln in rest if CONTACT.search(ln.text)]
    other = [ln for ln in rest if ln not in contact]
    if not contact and len(other) > 2:
        contact = [other.pop()]
    row2 = other[:1]
    row3 = other[1:]

    def join(group, sep):
        return (sep.join(ln.text for ln in group),
                max((ln.size for ln in group), default=10), any(ln.bold for ln in group))

    rows = [(head.text, head.size, head.bold), join(row2, "\n"), join(row3, "\n"), join(contact, "  "),
            (title.text, title.size, True)]
    sizes = [_clamp(rows[0][1], 14, 26)] + [_clamp(r[1], 8, 14) for r in rows[1:4]] + [_clamp(rows[4][1], 10, 14)]
    return Draft(name.strip(), [r[0] for r in rows], sizes, [True, rows[1][2], rows[2][2], rows[3][2], True],
                 tested_by=not has_stamp)


def _clamp(x, lo, hi):
    return float(min(hi, max(lo, round(x * 2) / 2)))


# ---------- the profile: cells, merges, fonts, row heights, pictures ----------

def _row_height(text: str, size: float, bold: bool, width: float, minimum: float) -> float:
    """Template height (points) for the text, wrapped like the PDF; rounded up to 0.75 like Excel's rows."""
    if not text.strip():
        return 4.5
    n = len(_wrap(text, font_for(FONT, bold, False), size, width - 2 * CELL_PAD))
    need = n * size * 1.2 + 1.5
    return max(minimum, math.ceil(need / 0.75) * 0.75)


def make_profile(draft: Draft, company_id: str, logo_ratio: float | None, stamp: bool) -> dict:
    """The settings profile for the draft (pictures as assets/<id>_logo.png / _stamp.png)."""
    cols = load_layout(BASE).cols
    left = "C" if logo_ratio else "A"
    cells = [f"{left}1", f"{left}2", f"{left}3", "A4", "A5"]
    widths = [sum(cols[2:]) if logo_ratio else sum(cols)] * 3 + [sum(cols)] * 2
    sizes = list(draft.sizes)
    budget = ROOM / ROW_SCALE - sum(load_layout(BASE).rows[5:37]) / ROW_SCALE - sum(BOTTOM_ROWS.values())

    def heights():
        hs = [_row_height(t, s, b, w, m) for t, s, b, w, m in
              zip(draft.texts, sizes, draft.bold, widths, (15, 12, 12, 12, 18.75))]
        if logo_ratio and sum(hs[:3]) < LOGO_MAX_H * 0.85 + 2 * LOGO_DY:   # room for the logo beside rows 1-3
            extra = LOGO_MAX_H * 0.85 + 2 * LOGO_DY - sum(hs[:3])
            hs[1] += math.ceil(extra / 2 / 0.75) * 0.75
            hs[2] += math.ceil(extra / 2 / 0.75) * 0.75
        return hs

    hs = heights()
    while sum(hs) > budget:                      # too tall for one A4 page: smaller details, then name
        i = next((i for i in (1, 2, 3) if sizes[i] > 8 and draft.texts[i].strip()), None)
        if i is None and sizes[0] > 16:
            i = 0
        if i is None:
            raise CompanyError(f"the header needs {sum(hs):.0f} pt but a certificate has room for {budget:.0f} pt "
                               "on one A4 page. Shorten the rows on the left (or leave a row empty)")
        sizes[i] -= 0.5
        hs = heights()

    profile = {"name": draft.name, "template": f"templates/test_report_{company_id}.xlsx", "images": []}
    if logo_ratio:
        # the logo may reach into column C up to where the widest centred line of rows 1-3 begins
        widest = max((_width(_segments(ln), font_for(FONT, b, False), s)
                      for t, s, b in zip(draft.texts[:3], sizes, draft.bold)
                      for ln in _wrap(t, font_for(FONT, b, False), s, widths[0] - 2 * CELL_PAD)), default=0)
        room = sum(cols[:2]) + (widths[0] - widest) / 2 - LOGO_DX - 6
        h = min(LOGO_MAX_H, sum(hs[:3]) - 2 * LOGO_DY, max(LOGO_MIN_W, room) / logo_ratio)
        profile["images"].append({"file": f"assets/{company_id}_logo.png", "cell": "A1", "dx": LOGO_DX, "dy": LOGO_DY,
                                  "height": round(h, 1)})
    if stamp:
        profile["images"].append({"file": f"assets/{company_id}_stamp.png", **STAMP})
    for ref, text in zip(cells, draft.texts):
        profile[ref] = text
    if not draft.tested_by:
        profile["G39"] = ""
    profile["template_layout"] = {
        "merges": [f"{ref}:L{ref[1:]}" for ref in cells],
        "fonts": {ref: [s, b] for ref, s, b in zip(cells, sizes, draft.bold)},
        "row_heights": {**{str(r): h for r, h in enumerate(hs, start=1)}, **BOTTOM_ROWS},
    }
    return profile


# ---------- the template ----------

def build_template(profile: dict, out: Path) -> Path:
    """The company's template: templates/test_report.xlsx with rows 1-5 replaced by its header.

    The table (rows 6 onwards) stays exactly the same; template_layout gives the header's merged cells,
    fonts ({cell: [size, bold]}, Times New Roman, centred, wrapped) and row heights (also the short empty
    rows under the table that keep the page at 100% on A4); the profile's cells give the texts.
    Logo and stamp are not stored in the template: every certificate places them (profile "images").
    """
    from app.render_excel import _rich
    lay = profile["template_layout"]
    wb = openpyxl.load_workbook(BASE, rich_text=True)   # rich_text keeps the superscript 2 / 3
    ws = wb.active
    for rng in [r for r in ws.merged_cells.ranges if r.max_row <= 5]:
        ws.unmerge_cells(str(rng))
    for row in ws.iter_rows(min_row=1, max_row=5):
        for c in row:
            c.value = None
    for rng in lay["merges"]:
        ws.merge_cells(rng)
    for ref, (size, bold) in lay["fonts"].items():
        cell = ws[ref]
        cell.font = Font(name=FONT, size=size, bold=bold)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    # header box as in the base: outline round rows 1-5, a line above the title (row 5); nothing inside
    # (unmerging leaves the old merged cells' inner edges behind)
    thin, none = Side(style="thin"), Side()
    for r in range(1, 6):
        for c in range(1, 13):
            ws.cell(r, c).border = Border(left=thin if c == 1 else none, right=thin if c == 12 else none,
                                          top=thin if r in (1, 5) else none, bottom=thin if r == 5 else none)
    for r, h in lay["row_heights"].items():
        ws.row_dimensions[int(r)].height = h
    for ref, text in profile.items():
        if CELL_RE.match(ref):
            ws[ref].value = _rich(text, copy(ws[ref].font)) if text else None
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    load_layout.cache_clear()          # a rebuilt template at the same path must be read again
    return out


# ---------- saving ----------

def new_id(name: str, taken) -> str:
    """'Shree Ram Pipes' -> 'shree' (then shree2, shree3... if taken)."""
    words = re.findall(r"[a-z]+", name.lower())
    base = words[0] if words else "company"
    cid, n = base, 2
    while cid in taken:
        cid, n = f"{base}{n}", n + 1
    return cid


def _insert_profile(text: str, company_id: str, profile: dict) -> str:
    """settings.json with the profile added at the end of company.profiles, the rest of the file
    (its layout and notes) unchanged."""
    start = text.index('"profiles"')
    i = text.index("{", start)
    depth, in_str, esc = 0, False, False
    for j in range(i, len(text)):
        ch = text[j]
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
        elif ch == '"':
            in_str = True
        elif ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                close = j
                break
    body = json.dumps(profile, ensure_ascii=False, indent=2).replace("\n", "\n      ")
    head = text[:close].rstrip()
    return f'{head},\n      "{company_id}": {body}\n    {text[close:]}'


def save_company(draft: Draft, logo: bytes | None, stamp: bytes | None, settings_file: Path = SETTINGS_FILE) -> str:
    """Write the pictures, the template and the profile; returns the new company's id.

    settings.json is checked after the change; if anything is wrong the old file is put back and
    the new files are removed, so a failed save leaves nothing behind."""
    text = Path(settings_file).read_text(encoding="utf-8")
    profiles = json.loads(text)["company"]["profiles"]
    cid = new_id(draft.name, profiles)
    if any(p.get("name", "").strip().lower() == draft.name.strip().lower() for p in profiles.values()):
        raise CompanyError(f"there's already a company called {draft.name}")
    ratio = Image.open(io.BytesIO(logo)).size if logo else None
    profile = make_profile(draft, cid, ratio[0] / ratio[1] if ratio else None, bool(stamp))
    written = []
    try:
        for data, kind in ((logo, "logo"), (stamp, "stamp")):
            if data:
                p = ROOT / "assets" / f"{cid}_{kind}.png"
                p.parent.mkdir(parents=True, exist_ok=True)
                p.write_bytes(data)
                written.append(p)
        written.append(build_template(profile, ROOT / profile["template"]))
        new_text = _insert_profile(text, cid, profile)
        problems = validate(_strip_help(json.loads(new_text)))
        if problems:
            raise CompanyError("the new company doesn't pass the settings checks: " + "; ".join(problems))
        shutil.copyfile(settings_file, Path(settings_file).with_name("settings.previous.json"))
        Path(settings_file).write_text(new_text, encoding="utf-8")
    except Exception:
        for p in written:
            p.unlink(missing_ok=True)
        raise
    return cid


def preview_settings(draft: Draft, folder: Path, logo: bytes | None, stamp: bytes | None) -> dict:
    """Settings with the draft as the selected company, its template and pictures in `folder`
    (the add-company screen's own folder), for drawing a sample certificate."""
    s = load_settings()
    ratio = Image.open(io.BytesIO(logo)).size if logo else None
    profile = make_profile(draft, "preview", ratio[0] / ratio[1] if ratio else None, bool(stamp))
    profile["template"] = str(build_template(profile, folder / "template.xlsx"))
    for img in profile["images"]:
        img["file"] = str(folder / Path(img["file"]).name.replace("preview_", ""))
    return {**s, "company": {**s["company"], "selected": "preview",
                             "profiles": {**s["company"]["profiles"], "preview": profile}}}


def sample_row():
    """A made-up entry for the preview certificate (PE 100, 63 mm, PN 10)."""
    from app.production import Row, _resolve
    row = Row("PE 100", 0, 1, date.today(), "63 X 10 ( SDR 17 )", "SAMPLE/001", "SAMPLE BATCH")
    _resolve(row)
    return row
