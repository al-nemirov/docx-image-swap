"""
DOCX Image Swap — step_04_save_result
======================================
Save result to output folder.

Copies working.docx and (optionally) images/.
"""

from pathlib import Path
import shutil
import json
import time
from datetime import datetime
from typing import Any, Callable, Dict, List

try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


def run(work_path: Path, step_config: Dict[str, Any], log_func: Callable[[str], None], ai_client: Any = None, **kwargs: Any) -> bool:
    """Save final result to output folder."""
    start_time = time.time()

    try:
        # --- PARAMETERS ---
        config = step_config.get("config", step_config)
        copy_metadata = config.get("copy_metadata", True)
        copy_images = config.get("copy_images", False)
        filename_template = config.get(
            "filename_template",
            "{name}_img_swap_{timestamp}.docx"
        )

        log_func("")
        log_func("=" * 70)
        log_func("   SAVING RESULT")
        log_func("=" * 70)

        # --- PATHS ---
        temp_dir = work_path / "temp"
        output_dir = work_path / "output"
        images_dir = work_path / "images"
        working_docx = temp_dir / "working.docx"
        source_info_file = temp_dir / "source_info.json"

        # --- CHECKS ---
        if not working_docx.exists():
            log_func("[ERROR] working.docx not found in temp/ directory")
            return False

        file_size_mb = working_docx.stat().st_size / (1024 * 1024)
        log_func(f"Source file: {file_size_mb:.2f} MB")

        # --- OUTPUT DIR ---
        output_dir.mkdir(parents=True, exist_ok=True)

        # --- FILENAME ---
        doc_name = "Document"

        if source_info_file.exists():
            try:
                with open(source_info_file, "r", encoding="utf-8") as f:
                    info = json.load(f)
                    doc_name = info.get("original_name", doc_name).strip()
            except Exception:
                pass

        safe_name = "".join(
            c if c.isalnum() or c in (' ', '_', '-') else '_'
            for c in doc_name
        ).strip().replace(' ', '_')[:100]

        now = datetime.now()
        timestamp = now.strftime("%Y%m%d_%H%M")

        output_filename = filename_template.format(
            name=safe_name,
            version="1.0",
            timestamp=timestamp,
            year=now.year,
            month=now.strftime("%m"),
            day=now.strftime("%d"),
            hour=now.strftime("%H"),
            minute=now.strftime("%M")
        ).replace("__", "_").strip("_")

        output_path = output_dir / output_filename

        # Conflict resolution
        if output_path.exists():
            counter = 2
            base_stem = output_path.stem
            base_suffix = output_path.suffix
            while output_path.exists():
                output_path = output_dir / f"{base_stem}_v{counter}{base_suffix}"
                counter += 1

        log_func(f"Output filename: {output_path.name}")

        # --- COPY ---
        log_func("")
        log_func("Copying document...")
        shutil.copy2(working_docx, output_path)

        if not output_path.exists():
            log_func("[ERROR] Output file was not created — check disk space and permissions")
            return False

        final_size_mb = output_path.stat().st_size / (1024 * 1024)
        log_func(f"Copied: {final_size_mb:.2f} MB")

        # Validate
        if HAS_DOCX:
            try:
                doc = Document(output_path)
                log_func(f"   Paragraphs: {len(doc.paragraphs)}")
            except Exception:
                pass

        # --- METADATA ---
        copied_files: List[str] = []
        if copy_metadata:
            for name in ["image_map.json", "source_info.json"]:
                src = temp_dir / name
                if src.exists():
                    try:
                        shutil.copy2(src, output_dir / src.name)
                        copied_files.append(src.name)
                    except Exception:
                        pass
            if copied_files:
                log_func(f"Metadata: {', '.join(copied_files)}")

        # --- IMAGES ---
        if copy_images and images_dir.exists():
            output_images_dir = output_dir / "images"
            output_images_dir.mkdir(exist_ok=True)

            copied_count = 0
            for img in images_dir.glob("*.*"):
                try:
                    shutil.copy2(img, output_images_dir / img.name)
                    copied_count += 1
                except Exception:
                    pass

            if copied_count > 0:
                log_func(f"Images: {copied_count} files")

        # --- SUMMARY ---
        duration = time.time() - start_time

        log_func("")
        log_func("=" * 70)
        log_func("DONE!")
        log_func("=" * 70)
        log_func(f"   File: {output_path.resolve()}")
        log_func(f"   Size: {final_size_mb:.2f} MB")
        log_func(f"   Duration: {duration:.1f} sec")
        log_func("=" * 70)

        return True

    except Exception as e:
        log_func(f"[CRITICAL] Unhandled error while saving result: {e}")
        import traceback
        for line in traceback.format_exc().splitlines():
            if line.strip():
                log_func(f"   {line}")
        return False
