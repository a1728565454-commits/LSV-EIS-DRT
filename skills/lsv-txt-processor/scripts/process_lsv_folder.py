from __future__ import annotations

import argparse
import importlib.util
import sys
from pathlib import Path


DEFAULT_OUTPUT = Path.cwd() / "processed"


def bundled_project_script() -> Path:
    repo_script = (
        Path(__file__).resolve().parents[3]
        / "support"
        / "analysis_scripts"
        / "build_lsv_template_batch.py"
    )
    if not repo_script.exists():
        raise RuntimeError(
            "Bundled support script not found. Run this skill from the repository "
            "or set up the repository support files next to the skills directory."
        )
    return repo_script


def load_module(path: Path):
    spec = importlib.util.spec_from_file_location("lsv_batch", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load project LSV builder: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules["lsv_batch"] = module
    spec.loader.exec_module(module)
    return module


def output_path_for(module, folder: Path, output_root: Path) -> Path:
    try:
        relative = folder.resolve().relative_to(module.ROOT.resolve())
        return output_root / relative.parent / f"{folder.name}-LSV.xlsx"
    except ValueError:
        return output_root / f"{folder.name}-LSV.xlsx"


def main() -> None:
    parser = argparse.ArgumentParser(description="Process one selected folder of TXT LSV/IPV files.")
    parser.add_argument("--input-dir", required=True, help="Selected folder containing exported TXT files.")
    parser.add_argument("--output-root", default=str(DEFAULT_OUTPUT), help="Processed-data output root.")
    args = parser.parse_args()

    folder = Path(args.input_dir).resolve()
    output_root = Path(args.output_root).resolve()
    if not folder.exists() or not folder.is_dir():
        raise SystemExit(f"Input folder does not exist or is not a directory: {folder}")

    module = load_module(bundled_project_script())
    files = [p for p in folder.glob("*.txt") if module.re.search(r"(?i)(lsv|ipv)", p.name)]
    if not files:
        raise SystemExit(f"No LSV/IPV TXT files found in: {folder}")

    module.OUTPUT_DIR = output_root
    module.lsv_dirs = lambda: [folder]
    module.output_path = lambda f: output_path_for(module, f, output_root)

    output, count = module.build_one(folder)
    print(f"built\t{count}\t{output}")


if __name__ == "__main__":
    main()
