"""Full G2 deck to a temp dir: python3 bfull.py --build-dir <path> --output <outdir>"""
import argparse
import subprocess
import sys
from pathlib import Path

parser = argparse.ArgumentParser(description="Build full G2 deck to a temp directory")
parser.add_argument("--build-dir", required=True, help="Path to build directory containing theme.py and build.py")
parser.add_argument("--output", required=True, help="Output directory for generated files")
args = parser.parse_args()

sys.path.insert(0, args.build_dir)
import theme  # noqa: E402
theme.G2 = True
import build  # noqa: E402

out = Path(args.output)
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
