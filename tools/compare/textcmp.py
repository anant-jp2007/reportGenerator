"""Match each text piece in Excel's PDF with ours; report position differences in points."""
import sys, statistics, pymupdf as fitz
def spans(path):
    out = []
    for b in fitz.open(path)[0].get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                t = s["text"].strip()
                if t: out.append((t, s["bbox"][0], s["origin"][1], s["size"], s["font"]))
    return out
ex, mine = spans(sys.argv[1]), spans(sys.argv[2])
used, rows = set(), []
for t, x, y, size, font in ex:
    cands = [(abs(mx - x) + abs(my - y), i) for i, (mt, mx, my, ms, mf) in enumerate(mine) if mt == t and i not in used]
    if not cands:
        rows.append((t, None)); continue
    _, i = min(cands); used.add(i)
    rows.append((t, (round(mine[i][1] - x, 2), round(mine[i][2] - y, 2), round(mine[i][3] - size, 2))))
missing = [t for t, d in rows if d is None]
d = [r[1] for r in rows if r[1]]
print(f"matched {len(d)} of {len(ex)} text pieces; not matched: {missing[:12]}")
print("dx  median %.2f  max %.2f" % (statistics.median(abs(a) for a, b, c in d), max(abs(a) for a, b, c in d)))
print("dy  median %.2f  max %.2f" % (statistics.median(abs(b) for a, b, c in d), max(abs(b) for a, b, c in d)))
print("size differences:", sorted({c for a, b, c in d if abs(c) > 0.05}))
worst = sorted(((abs(a) + abs(b), t, a, b) for t, (a, b, c) in [(t, v) for t, v in rows if v]), reverse=True)[:15]
for s, t, a, b in worst: print(f"  {t[:38]!r:42} dx {a:+.2f} dy {b:+.2f}")
