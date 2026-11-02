"""Per-page text line audit from the rendered PDF: overlaps, off-page, orphan lines."""
import argparse
import re, subprocess, sys, html

parser = argparse.ArgumentParser(description="Audit text lines in rendered PDF pages")
parser.add_argument("pdf", help="Path to PDF file")
parser.add_argument("pages", nargs="*", type=int, help="Page numbers to check (default: all)")
args = parser.parse_args()

pdf = args.pdf
pages = args.pages or None
x = subprocess.run(["pdftotext", "-bbox-layout", pdf, "-"], capture_output=True, text=True).stdout
out = []
for pi, pg in enumerate(re.findall(r"<page.*?</page>", x, re.S), 1):
    if pages and pi not in pages: continue
    L = []
    for bi, blk in enumerate(re.findall(r"<block.*?</block>", pg, re.S)):
        for ln in re.findall(r'<line xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">(.*?)</line>', blk, re.S):
            words = "".join(html.unescape(w) for w in re.findall(r">([^<]*)</word>", ln[4]))
            L.append((bi, *map(float, ln[:4]), words))
    for i, a in enumerate(L):
        if a[4] > 540 * 0.985 or a[3] > 960: out.append(f"p{pi:02d} EDGE {tuple(round(v) for v in a[1:5])} {a[5][:30]}")
        if len(a[5]) == 1 and ord(a[5]) > 0x2e80 and sum(1 for b in L if b[0] == a[0]) > 1:
            out.append(f"p{pi:02d} ORPHAN '{a[5]}'")
        for b in L[i + 1:]:
            ix = min(a[3], b[3]) - max(a[1], b[1]); iy = min(a[4], b[4]) - max(a[2], b[2])
            if ix > 1 and iy > 2:
                out.append(f"p{pi:02d} OVER {a[5][:20]!r} ~ {b[5][:20]!r} ({ix:.0f}x{iy:.0f})")
print("\n".join(out) or "clean")
