import argparse
from pathlib import Path

parser = argparse.ArgumentParser(description="Patch EISART util_ecm.py for newer NumPy versions.")
parser.add_argument(
    "path",
    nargs="?",
    default=str(Path(__file__).resolve().parent / "util_ecm.py"),
    help="Path to util_ecm.py. Defaults to the bundled copy.",
)
args = parser.parse_args()

p = Path(args.path)
text = p.read_text(encoding="utf-8", errors="ignore")

old = "ranked_peak_indices = np.array(ranked_peak_indices, dtype=np.int)"
new = "ranked_peak_indices = np.array(ranked_peak_indices, dtype=int)"

if old in text:
    text = text.replace(old, new)
    p.write_text(text, encoding="utf-8")
    print("已修补 util_ecm.py")
else:
    print("没找到目标语句，请手动搜索 np.int")
