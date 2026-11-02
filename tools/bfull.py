"""Full G2 deck to a temp dir: python3 /tmp/g2/bfull.py <outdir>"""
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, "/home/ansty/3D_learing/SEA-loop/reports/中期答辩_20261104/build")
import theme  # noqa: E402

theme.G2 = True
import build  # noqa: E402

out = Path(sys.argv[1])
out.mkdir(parents=True, exist_ok=True)
build.LIGHT.update({"TXT": "17252B", "SUB": "303C42", "DIM": "303C42", "ACC": "214C40"})
build.use_light_palette()
prs = theme.new_deck()
for f in build.main_slides():
    f(prs)
build.slides_end.backups(prs)
build.keep_left_margin(prs)
prs.save(out / "deck.pptx")
hidden = sum(1 for s in prs.slides if s.element.get("show") == "0")
print(f"{len(prs.slides)} slides, {hidden} hidden")
subprocess.run(["soffice", "--headless", f"-env:UserInstallation=file://{out}/lo", "--convert-to", "pdf",
                "--outdir", str(out), str(out / "deck.pptx")], check=True, stdout=subprocess.DEVNULL)
for f in out.glob("p-*.png"):
    f.unlink()
subprocess.run(["pdftoppm", "-r", "144", "-png", str(out / "deck.pdf"), str(out / "p")], check=True)
