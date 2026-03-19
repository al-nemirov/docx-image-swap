#!/usr/bin/env python3
"""
DOCX Image Swap — CLI Runner
=============================
Extract images from DOCX → Manual swap → Reinject → Save result.

Usage:
    python run.py                      (uses first .docx in source/)
    python run.py document.docx        (specific file)
"""

import sys
import json
import shutil
from pathlib import Path

ROOT = Path(__file__).parent
CONFIG_FILE = ROOT / "config.json"
SOURCE_DIR = ROOT / "source"
WORK_DIR = ROOT / "temp"
IMAGES_DIR = ROOT / "images"
OUTPUT_DIR = ROOT / "output"


def log(msg: str):
    print(f"  {msg}")


def load_config() -> dict:
    with open(CONFIG_FILE, encoding="utf-8") as f:
        return json.load(f)


def find_input_file(arg: str = None) -> Path:
    """Find input DOCX: from argument or first file in source/."""
    if arg:
        p = Path(arg)
        if p.exists():
            return p
        p = SOURCE_DIR / arg
        if p.exists():
            return p
        print(f"File not found: {arg}")
        sys.exit(1)

    SOURCE_DIR.mkdir(exist_ok=True)
    docx_files = list(SOURCE_DIR.glob("*.docx"))
    if not docx_files:
        print(f"No .docx files found in {SOURCE_DIR}/")
        print("Place a .docx file there or pass a path as argument.")
        sys.exit(1)
    return docx_files[0]


def main():
    config = load_config()
    steps = config.get("steps", [])

    # Find input file
    arg = sys.argv[1] if len(sys.argv) > 1 else None
    input_file = find_input_file(arg)
    print(f"\nDOCX Image Swap")
    print(f"Input: {input_file.name}\n")

    # Prepare work directory
    WORK_DIR.mkdir(exist_ok=True)
    IMAGES_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)

    # Copy source to work dir
    work_docx = WORK_DIR / "working.docx"
    shutil.copy2(input_file, work_docx)

    # Also copy to source info for step_04
    source_info = {"original_name": input_file.stem, "original_path": str(input_file)}
    with open(WORK_DIR / "source_info.json", "w", encoding="utf-8") as f:
        json.dump(source_info, f, ensure_ascii=False)

    # Import step modules
    from modules.step_01_extract_images import run as step_01
    from modules.step_02_manual_swap import run as step_02
    from modules.step_03_insert_images import run as step_03
    from modules.step_04_save_result import run as step_04

    step_funcs = [step_01, step_02, step_03, step_04]

    for i, (step_cfg, step_fn) in enumerate(zip(steps, step_funcs)):
        name = step_cfg.get("name", f"Step {i + 1}")
        if not step_cfg.get("enabled", True):
            print(f"[SKIP] {name}")
            continue

        print(f"[{i + 1}/4] {name}")

        # Merge step config
        merged = {**step_cfg}
        if "config" in step_cfg:
            merged.update(step_cfg["config"])

        # Manual step — pause for user
        if step_cfg.get("type") == "manual":
            step_fn(WORK_DIR, merged, log)
            print(f"\n  Images are in: {IMAGES_DIR.resolve()}")
            print(f"  Replace the files (keep filenames) and press Enter.\n")
            input("  Press Enter to continue...")
            print()
            continue

        ok = step_fn(WORK_DIR, merged, log)
        if not ok:
            print(f"\n[ERROR] {name} failed. Aborting.")
            sys.exit(1)
        print()

    print(f"Done! Result saved to: {OUTPUT_DIR.resolve()}\n")


if __name__ == "__main__":
    main()
