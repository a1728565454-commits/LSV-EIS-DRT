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
        / "build_eis_drt_fit_batch.py"
    )
    if not repo_script.exists():
        raise RuntimeError(
            "Bundled support script not found. Run this skill from the repository "
            "or set up the repository support files next to the skills directory."
        )
    return repo_script


def load_module(path: Path):
    spec = importlib.util.spec_from_file_location("eis_drt_batch", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load project EIS+DRT builder: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules["eis_drt_batch"] = module
    spec.loader.exec_module(module)
    return module


def accepted_txt(path: Path, selected_folder: Path, allow_sta: bool = False) -> bool:
    lowered = str(path).lower()
    folder_lowered = str(selected_folder).lower()
    if path.suffix.lower() != ".txt":
        return False
    if not allow_sta and "sta" in lowered:
        return False
    if "eis" in path.name.lower():
        return True
    if "阻抗" in str(selected_folder):
        return True
    return False


def output_path_for(module, folder: Path, output_root: Path) -> Path:
    output_name = module.TEMPLATE.name.replace("文件夹名称", folder.name)
    try:
        relative = folder.resolve().relative_to(module.ROOT.resolve())
        return output_root / relative.parent / output_name
    except ValueError:
        return output_root / output_name


def main() -> None:
    parser = argparse.ArgumentParser(description="Process one selected folder into EIS+DRT+fit workbook.")
    parser.add_argument("--input-dir", required=True, help="Selected folder containing exported TXT EIS/impedance files.")
    parser.add_argument("--output-root", default=str(DEFAULT_OUTPUT), help="Processed-data output root.")
    parser.add_argument("--allow-sta", action="store_true", help="Allow STA paths. Default is to exclude STA.")
    args = parser.parse_args()

    folder = Path(args.input_dir).resolve()
    output_root = Path(args.output_root).resolve()
    if not folder.exists() or not folder.is_dir():
        raise SystemExit(f"Input folder does not exist or is not a directory: {folder}")

    module = load_module(bundled_project_script())
    files = [p for p in folder.glob("*.txt") if accepted_txt(p, folder, args.allow_sta)]
    if not files:
        raise SystemExit(f"No accepted EIS/impedance TXT files found in: {folder}")

    module.OUTPUT_DIR = output_root
    module.is_excluded = lambda path: (not args.allow_sta and "sta" in str(path).lower())
    module.is_eis_source = lambda path: accepted_txt(path, folder, args.allow_sta)
    module.eis_dirs = lambda: [folder]
    module.folder_output_path = lambda f: output_path_for(module, f, output_root)

    runner, kernel, settings_dir = module.setup_drt_runner()
    import shutil
    import tempfile

    temp_dir = Path(tempfile.mkdtemp(prefix="selected_eis_drt_"))
    try:
        drt_root = module.run_drt_for_folder(runner, kernel, settings_dir, folder, temp_dir)
        output = module.build_workbook(folder, drt_root)
        print(f"built\t{len(files)}\t{output}")
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == "__main__":
    main()
