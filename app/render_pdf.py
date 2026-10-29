"""Write reports as PDF, drawn to match Excel's print of templates/test_report.xlsx.

    write_pdf(report, "MP100-019.pdf")            # one certificate
    write_pdf_combined(reports, "All.pdf")        # one page per certificate

Uses the template's layout (app/layout.py) and the Windows fonts the template uses, so the PDF
looks like the Excel file printed. Doesn't need Excel installed.
"""
from __future__ import annotations

import re
from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

from app.layout import TABLE_SHIFT_X, THIN_LINE, Layout, load_layout, merged_area_of
from app.report import Report

ROOT = Path(__file__).resolve().parent.parent
FONT_DIR = Path("C:/Windows/Fonts")
# font name in the template -> (regular, bold, italic, bold italic) files in Windows\Fonts
FONT_FILES = {
    "Times New Roman": ("times.ttf", "timesbd.ttf", "timesi.ttf", "timesbi.ttf"),
    "Cambria": ("cambria.ttc", "cambriab.ttf", "cambriai.ttf", "cambriaz.ttf"),
    "Calibri": ("calibri.ttf", "calibrib.ttf", "calibrii.ttf", "calibriz.ttf"),
    "Arial": ("arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf"),
}
BUILTIN = ("Times-Roman", "Times-Bold", "Times-Italic", "Times-BoldItalic")

CELL_PAD = 2.25             # Excel keeps about 3 px between the text and the cell edge
LINE_SPACING = 1.265        # line height in font sizes, as measured in Excel's PDF (13.92 pt for 11 pt)
BASELINE = 0.975            # baseline below the top of its line, in font sizes (fitted to Excel's PDF)
SUPER_SIZE = 0.67           # superscript size and rise, as measured in Excel's PDF
SUPER_RISE = 0.41
SIGNATURE_HEIGHT = 48.0     # points (64 px in the Excel file)
SIGNATURE_LIFT = 4.5        # points above the top of the "Tested by" row (6 px)
SUPERSCRIPT = re.compile(r"\^\{([^}]*)\}")

_registered: dict[tuple, str] = {}


def font_for(name: str, bold: bool, italic: bool) -> str:
    """Register the template font with reportlab (once) and return its reportlab name."""
    key = (name, bold, italic)
    if key in _registered:
        return _registered[key]
    idx = (2 if italic else 0) + (1 if bold else 0)
    files = FONT_FILES.get(name)
    rl_name = BUILTIN[idx]
    if files and (FONT_DIR / files[idx]).is_file():
        rl_name = f"{name}-{idx}"
        f = FONT_DIR / files[idx]
        pdfmetrics.registerFont(TTFont(rl_name, str(f), subfontIndex=0) if f.suffix == ".ttc"
                                else TTFont(rl_name, str(f)))
    elif files is None and name != "Times New Roman":
        return font_for("Times New Roman", bold, italic)
    _registered[key] = rl_name
    return rl_name


def _segments(text: str) -> list[tuple[str, bool]]:
    """'Kgs/Cm^{2}' -> [('Kgs/Cm', False), ('2', True)]"""
    out, pos = [], 0
    for m in SUPERSCRIPT.finditer(text):
        if m.start() > pos:
            out.append((text[pos:m.start()], False))
        out.append((m.group(1), True))
        pos = m.end()
    if pos < len(text):
        out.append((text[pos:], False))
    return out


def _width(segs, font, size) -> float:
    return sum(pdfmetrics.stringWidth(t, font, size * (SUPER_SIZE if sup else 1)) for t, sup in segs)


def _wrap(text: str, font: str, size: float, width: float) -> list[str]:
    """Break text into lines like Excel: at line breaks, and between words when too wide.
    Trailing spaces don't count towards the width (Excel lets them hang past the edge)."""
    lines = []
    for para in text.split("\n"):
        words, line = para.split(" "), ""
        for w in words:
            trial = w if not line else f"{line} {w}"
            if line.strip() and w and _width(_segments(trial.rstrip()), font, size) > width:
                lines.append(line.rstrip())
                line = w
            else:
                line = trial
        lines.append(line.rstrip())
    # Excel ignores spaces at the start of a line that follows a line break
    return lines[:1] + [ln.lstrip() for ln in lines[1:]]


def _draw_page(c: canvas.Canvas, report: Report, lay: Layout) -> None:
    page_w, page_h = A4
    left_m, right_m, top_m, _ = lay.page_margins
    x0 = left_m + ((page_w - left_m - right_m - lay.width) / 2 if lay.centered else 0) + TABLE_SHIFT_X
    y0 = page_h - top_m

    def X(x):
        return x0 + x

    def Y(y):
        return y0 - y

    c.setLineWidth(THIN_LINE)
    c.setStrokeColorRGB(0, 0, 0)
    for x1, y1, x2, y2 in lay.borders:
        c.line(X(x1), Y(y1), X(x2), Y(y2))

    for ref, cell in report.cells.items():
        if not cell.text or ref not in lay.boxes:
            continue
        box, st = lay.boxes[ref], lay.styles[ref]
        font = font_for(st.font, st.bold, st.italic)
        size = st.size
        inner = box.w - 2 * CELL_PAD
        lines = _wrap(cell.text, font, size, inner) if st.wrap else [cell.text.replace("\n", " ")]

        leading = size * LINE_SPACING
        block = leading * len(lines)
        if st.v_align == "top":
            top = box.y + 1
        elif st.v_align == "center":
            top = box.y + (box.h - block) / 2
        else:
            top = box.y + box.h - block - 1

        align = st.h_align
        if align in (None, "general"):
            align = "right" if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool) else "left"

        for i, line in enumerate(lines):
            segs = _segments(line)
            w = _width(segs, font, size)
            if align == "center" or align == "centerContinuous":
                x = box.x + (box.w - w) / 2
            elif align == "right":
                x = box.x + box.w - CELL_PAD - w
            else:
                x = box.x + CELL_PAD
            baseline = top + i * leading + size * BASELINE
            for text, sup in segs:
                s = size * (SUPER_SIZE if sup else 1)
                c.setFont(font, s)
                c.drawString(X(x), Y(baseline - (size * SUPER_RISE if sup else 0)), text)
                x += pdfmetrics.stringWidth(text, font, s)

    sig = ROOT / report.signature_file
    if sig.is_file():
        min_col, min_row, max_col, _ = merged_area_of(report.signature_next_to)
        x = sum(lay.cols[:max_col])
        y = sum(lay.rows[:min_row - 1]) - SIGNATURE_LIFT
        img = ImageReader(str(sig))
        iw, ih = img.getSize()
        w = SIGNATURE_HEIGHT * iw / ih
        c.drawImage(img, X(x), Y(y + SIGNATURE_HEIGHT), w, SIGNATURE_HEIGHT, mask="auto")


def _canvas(path: Path, title: str) -> canvas.Canvas:
    path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(path), pagesize=A4)
    c.setTitle(title)
    c.setAuthor("Test certificates")
    return c


def write_pdf(report: Report, path: str | Path) -> Path:
    path = Path(path)
    c = _canvas(path, f"Test certificate {report.file_name}")
    _draw_page(c, report, load_layout())
    c.showPage()
    c.save()
    return path


def write_pdf_combined(reports: list[Report], path: str | Path) -> Path:
    path = Path(path)
    c = _canvas(path, "Test certificates")
    lay = load_layout()
    for report in reports:
        _draw_page(c, report, lay)
        c.showPage()
    c.save()
    return path
