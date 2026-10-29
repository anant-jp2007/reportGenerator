# Comparing our PDF / Word certificates with Excel's own print

Development tools only (need Excel and Word installed). Used to tune app/render_pdf.py and app/render_word.py.

- `xl2pdf.ps1 <in.xlsx> <out.pdf>`: Excel exports the workbook to PDF (prints column widths / row heights too).
- `doc2pdf.ps1 <in.docx> <out.pdf>`: Word exports the document to PDF.
- `compare.py <folder> [dpi]`: makes one certificate as Excel (exported by Excel) and as our PDF in <folder>,
  writes overlay.png and excel.pdf / mine.pdf.
- `textcmp.py <excel.pdf> <other.pdf>`: matches every text piece, prints median/max position differences in pt.

Results so far: PDF median 0.24 pt across / 0.19 pt down (max 1.3 pt); Word median 0.24 / 0.36 (max 3.8 pt,
the "Maximum " in D20, same in both, an Excel quirk in that cell).
