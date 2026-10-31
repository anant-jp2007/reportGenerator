"""A certificate drawn in the browser, from the same layout the PDF and Word writers use.

    certificate_html(report, settings)   -> HTML for the check screen's preview window

Positions and sizes are in points, as in app/layout.py (the template's grid, measured against
Excel's print), so the preview shows the page as it prints. The values that can be changed
(report.editable_cells, plus the Report No.) are input boxes in their own cells.
"""
from __future__ import annotations

import base64
import io
from html import escape

from markupsafe import Markup
from PIL import Image

from app.layout import load_layout
from app.render_pdf import CELL_PAD, LINE_SPACING, _wrap, font_for, picture_position
from app.report import MARKUP, Report, editable_cells

FONT_FALLBACK = "'Times New Roman', Times, serif"


def _rich(text: str) -> str:
    """Printed text as HTML: '^{2}' as a superscript; symbols kept as their own character."""
    out, pos = [], 0
    for m in MARKUP.finditer(text):
        out.append(escape(text[pos:m.start()]))
        out.append(f"<sup>{escape(m.group(1))}</sup>" if m.group(1) is not None else escape(m.group(2)))
        pos = m.end()
    out.append(escape(text[pos:]))
    return "".join(out)


def _picture(path) -> tuple[str, float]:
    """(data URL, width / height) of a logo or stamp."""
    with Image.open(path) as im:
        ratio = im.width / im.height
        buf = io.BytesIO()
        im.save(buf, "PNG")
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode(), ratio


def _labels(settings: dict) -> dict[str, str]:
    """cell -> name of the value, for the input boxes (Date of manufacturing, Outside diameter minimum...)."""
    out = {ref: key.replace("_", " ").capitalize() for key, ref in settings["header_cells"].items() if key != "_help"}
    out[settings["header_cells"]["report_no"]] = "Report No."
    for name, c in settings["random_values"].items():
        if isinstance(c, dict):
            label = name.replace("_", " ").capitalize()
            for ref, which in ((c["cell"], ""),) if "cell" in c else ((r, f" {k}") for k, r in c["cells"].items()):
                out[ref] = label + which
    return out


def certificate_html(report: Report, settings: dict, editable: bool = True) -> Markup:
    """editable=False: every value as plain text (the add-company screen shows a sample certificate)."""
    lay = load_layout(report.template)
    kinds = dict(editable_cells(settings)) if editable else {}
    if editable:
        kinds[settings["header_cells"]["report_no"]] = "number"
    labels = _labels(settings)
    parts = [f'<div class="cert" style="width:{lay.width:.2f}pt;height:{lay.height:.2f}pt">',
             f'<svg class="cert-lines" viewBox="0 0 {lay.width:.2f} {lay.height:.2f}" aria-hidden="true">']
    parts += [f'<line x1="{x1:.2f}" y1="{y1:.2f}" x2="{x2:.2f}" y2="{y2:.2f}"/>' for x1, y1, x2, y2 in lay.borders]
    parts.append("</svg>")

    for ref in sorted(set(report.cells) | set(kinds), key=lambda r: (int(r.lstrip("ABCDEFGHIJKLMNOPQRSTUVWXYZ")), r)):
        if ref not in lay.boxes:
            continue
        cell = report.cells.get(ref)
        text = cell.text if cell else ""
        kind = kinds.get(ref)
        if not text and not kind:
            continue
        box, st = lay.boxes[ref], lay.styles[ref]
        align = st.h_align
        if align in (None, "general"):
            numeric = cell is not None and isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool)
            align = "right" if numeric else "left"
        align = {"centerContinuous": "center", "fill": "left", "justify": "left", "distributed": "center"}.get(align, align)
        style = (f"left:{box.x:.2f}pt;top:{box.y:.2f}pt;width:{box.w:.2f}pt;height:{box.h:.2f}pt;"
                 f"font-family:'{escape(st.font)}',{FONT_FALLBACK};font-size:{st.size:g}pt;"
                 f"line-height:{st.size * LINE_SPACING:.2f}pt;padding:0 {CELL_PAD}pt;text-align:{align};"
                 f"justify-content:{ {'center': 'center', 'right': 'flex-end'}.get(align, 'flex-start') };"
                 f"font-weight:{700 if st.bold else 400};font-style:{'italic' if st.italic else 'normal'}")
        v = {"top": "start", "center": "center"}.get(st.v_align, "end")
        cls = f"c v-{v}"
        if kind:
            parts.append(f'<label class="{cls} edit" style="{style}"><input class="cert-in" data-cell="{ref}" '
                         f'data-kind="{kind}" value="{escape(text)}" aria-label="{escape(labels.get(ref, ref))}" autocomplete="off" '
                         f'spellcheck="false"></label>')
        else:
            # lines broken exactly where the PDF breaks them (same font measurements)
            font = font_for(st.font, st.bold, st.italic)
            lines = _wrap(text, font, st.size, box.w - 2 * CELL_PAD) if st.wrap else [text.replace("\n", " ")]
            parts.append(f'<div class="{cls}" style="{style}"><span>{"<br>".join(_rich(ln) for ln in lines)}</span></div>')

    for pic in report.pictures:
        x, y = picture_position(lay, pic)
        url, ratio = _picture(pic.file)
        parts.append(f'<img class="cert-pic" src="{url}" alt="" style="left:{x:.2f}pt;top:{y:.2f}pt;'
                     f'height:{pic.height:.2f}pt;width:{pic.height * ratio:.2f}pt">')
    parts.append("</div>")
    return Markup("".join(parts))
