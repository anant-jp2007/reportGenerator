"""Make one report as Excel (exported to PDF by Excel) and as our PDF; overlay both pages."""
import sys, subprocess, random
sys.path.insert(0, r"E:\PROjectssss")
import pymupdf as fitz
from PIL import Image, ImageChops
from app.production import read_production
from app.report import build_report
from app.render_excel import write_excel
from app.render_pdf import write_pdf
SP = sys.argv[1]
rows = read_production(r"E:\PROjectssss\data\source\production_sample.xls").rows
row = next(r for r in rows if r.grade == "PE 100" and r.excel_row == 53)
rep = build_report(row, 3, rng=random.Random(7))
write_excel(rep, f"{SP}/cmp.xlsx"); write_pdf(rep, f"{SP}/mine.pdf")
subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", f"{SP}/xl2pdf.ps1",
                f"{SP}\cmp.xlsx".replace("/", "\\"), f"{SP}\excel.pdf".replace("/", "\\")], capture_output=True)
dpi = int(sys.argv[2]) if len(sys.argv) > 2 else 100
a = fitz.open(f"{SP}/excel.pdf")[0].get_pixmap(dpi=dpi); b = fitz.open(f"{SP}/mine.pdf")[0].get_pixmap(dpi=dpi)
A = Image.frombytes("RGB", (a.width, a.height), a.samples).convert("L")
B = Image.frombytes("RGB", (b.width, b.height), b.samples).convert("L")
# overlay: Excel in red, ours in blue, where both are dark -> dark purple
ov = Image.merge("RGB", (B, Image.new("L", A.size, 255).point(lambda v: v) if False else ImageChops.lighter(A, B), A))
ov.save(f"{SP}/overlay.png")
diff = ImageChops.difference(A, B).point(lambda v: 255 if v > 80 else 0)
print("pages", len(fitz.open(f"{SP}/mine.pdf")), "| differing pixels:", sum(1 for v in diff.getdata() if v), "of", A.size[0]*A.size[1])
