#!/usr/bin/env python3
"""
DOCX Image Swap — CLI Runner
=============================
Extract images from DOCX -> Manual swap -> Reinject -> Save result.

Usage:
    python run.py                      (uses first .docx in source/)
    python run.py document.docx        (specific file)
    python run.py --yes                (skip manual confirmation prompts)
"""

import sys
import json
import shutil
import argparse
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

ROOT: Path = Path(__file__).parent
CONFIG_FILE: Path = ROOT / "config.json"
SOURCE_DIR: Path = ROOT / "source"
WORK_DIR: Path = ROOT / "temp"
IMAGES_DIR: Path = ROOT / "images"
OUTPUT_DIR: Path = ROOT / "output"

# Step module function signature
StepFunc = Callable[..., bool]


def log(msg: str) -> None:
    """Print an indented log message to stdout. / Вывести сообщение в консоль."""
    print(f"  {msg}")


def load_config() -> Dict[str, Any]:
    """Load pipeline configuration from config.json.
    Загрузить конфигурацию пайплайна из config.json.
    """
    if not CONFIG_FILE.exists():
        print("[ОШИБКА] Файл конфигурации не найден / Config file not found: config.json")
        sys.exit(1)
    try:
        with open(CONFIG_FILE, encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError as e:
        print(f"[ОШИБКА] Некорректный JSON в config.json / Invalid JSON in config.json: {e}")
        sys.exit(1)


def find_input_file(arg: Optional[str] = None) -> Path:
    """Find input DOCX: from CLI argument or first file in source/.
    Найти входной DOCX: из аргумента командной строки или первый файл в source/.
    """
    if arg:
        p = Path(arg)
        if p.exists():
            return p
        p = SOURCE_DIR / arg
        if p.exists():
            return p
        print(f"[ОШИБКА] Файл не найден / File not found: {arg}")
        sys.exit(1)

    SOURCE_DIR.mkdir(exist_ok=True)
    docx_files: List[Path] = sorted(SOURCE_DIR.glob("*.docx"))
    if not docx_files:
        print(f"[ОШИБКА] Нет .docx файлов в {SOURCE_DIR}/")
        print(f"         No .docx files found in {SOURCE_DIR}/")
        print("  Поместите .docx файл в эту папку или укажите путь аргументом.")
        print("  Place a .docx file there or pass a path as argument.")
        sys.exit(1)
    return docx_files[0]


def import_step_modules() -> List[StepFunc]:
    """Import and return all step runner functions.
    Импортировать и вернуть все функции шагов пайплайна.
    """
    try:
        from modules.step_01_extract_images import run as step_01
        from modules.step_02_manual_swap import run as step_02
        from modules.step_03_insert_images import run as step_03
        from modules.step_04_save_result import run as step_04
    except ImportError as e:
        print(f"[ОШИБКА] Не удалось импортировать модули / Failed to import modules: {e}")
        sys.exit(1)

    return [step_01, step_02, step_03, step_04]


def prepare_work_dir(input_file: Path) -> Path:
    """Create working directories and prepare the working copy.
    Создать рабочие директории и подготовить рабочую копию документа.
    """
    WORK_DIR.mkdir(exist_ok=True)
    IMAGES_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)

    work_docx: Path = WORK_DIR / "working.docx"
    shutil.copy2(input_file, work_docx)

    source_info: Dict[str, str] = {
        "original_name": input_file.stem,
        "original_path": str(input_file),
    }
    with open(WORK_DIR / "source_info.json", "w", encoding="utf-8") as f:
        json.dump(source_info, f, ensure_ascii=False)

    return work_docx


def run_step(
    index: int,
    step_cfg: Dict[str, Any],
    step_fn: StepFunc,
    total: int,
    no_confirm: bool = False,
) -> bool:
    """Execute a single pipeline step. Returns True on success.
    Выполнить один шаг пайплайна. Возвращает True при успехе.
    """
    name: str = step_cfg.get("name", f"Step {index + 1}")

    if not step_cfg.get("enabled", True):
        print(f"[ПРОПУСК/SKIP] {name}")
        return True

    print(f"[{index + 1}/{total}] {name}")

    # Merge step-level config
    merged: Dict[str, Any] = {**step_cfg}
    if "config" in step_cfg:
        merged.update(step_cfg["config"])

    # Manual step — pause for user input
    if step_cfg.get("type") == "manual":
        try:
            step_fn(ROOT, merged, log)
        except Exception as e:
            print(f"\n[ОШИБКА/ERROR] {name}: {e}")
            return False
        print(f"\n  Изображения находятся в: {IMAGES_DIR.resolve()}")
        print(f"  Images are in: {IMAGES_DIR.resolve()}")
        if no_confirm:
            print(f"  --yes: автопродолжение / auto-continuing...\n")
        else:
            print(f"  Замените файлы (сохраните имена) и нажмите Enter.")
            print(f"  Replace the files (keep filenames) and press Enter.\n")
            input("  Нажмите Enter / Press Enter to continue...")
        print()
        return True

    # Regular step
    try:
        ok: bool = step_fn(ROOT, merged, log)
    except Exception as e:
        print(f"\n[ОШИБКА/ERROR] Исключение в шаге '{name}' / Exception in step '{name}': {e}")
        return False

    if not ok:
        print(f"\n[ОШИБКА/ERROR] Шаг '{name}' не выполнен / Step '{name}' failed. Прерывание / Aborting.")
        return False

    print()
    return True


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments.
    Разбор аргументов командной строки.
    """
    parser = argparse.ArgumentParser(
        description="DOCX Image Swap — extract, replace and reinject images."
    )
    parser.add_argument(
        "file", nargs="?", default=None,
        help="Path to .docx file (optional, defaults to first file in source/)"
    )
    parser.add_argument(
        "--yes", "--no-confirm", dest="no_confirm", action="store_true",
        help="Skip manual confirmation prompts (auto-continue)"
    )
    return parser.parse_args()


def main() -> None:
    """Main entry point: load config, find input, run all steps.
    Главная точка входа: загрузка конфигурации, поиск входного файла, запуск всех шагов.
    """
    args: argparse.Namespace = parse_args()

    config: Dict[str, Any] = load_config()
    steps: List[Dict[str, Any]] = config.get("steps", [])

    if not steps:
        print("[ОШИБКА] В config.json нет шагов / No steps defined in config.json")
        sys.exit(1)

    # Find input file
    input_file: Path = find_input_file(args.file)

    print(f"\n{'=' * 50}")
    print(f"  DOCX Image Swap")
    print(f"  Файл / Input: {input_file.name}")
    print(f"{'=' * 50}\n")

    # Prepare work directory
    prepare_work_dir(input_file)

    # Import step modules
    step_funcs: List[StepFunc] = import_step_modules()

    # Validate: config steps must match available handler functions
    if len(steps) != len(step_funcs):
        print(
            f"[ОШИБКА] Несовпадение: {len(steps)} шагов в config.json, "
            f"но {len(step_funcs)} обработчиков / "
            f"Mismatch: {len(steps)} steps in config.json, "
            f"but {len(step_funcs)} handler functions available."
        )
        sys.exit(1)

    total: int = len(steps)

    # Run pipeline
    for i, (step_cfg, step_fn) in enumerate(zip(steps, step_funcs)):
        success: bool = run_step(i, step_cfg, step_fn, total, no_confirm=args.no_confirm)
        if not success:
            print(f"\n[ОШИБКА] Пайплайн прерван на шаге {i + 1} / Pipeline aborted at step {i + 1}.")
            sys.exit(1)

    print(f"{'=' * 50}")
    print(f"  Готово! Результат сохранён в: {OUTPUT_DIR.resolve()}")
    print(f"  Done!  Result saved to:       {OUTPUT_DIR.resolve()}")
    print(f"{'=' * 50}\n")


if __name__ == "__main__":
    main()
