"""
DOCX Image Swap — step_02_manual_swap
======================================
Manual pause for image replacement.

User replaces files in images/ folder (keeping filenames)
and confirms to continue.
"""

from pathlib import Path
import json
from typing import Dict, Any, Callable


def run(work_path: Path, step_config: Dict[str, Any], log_func: Callable[[str], None], ai_client=None, **kwargs) -> bool:
    """
    Manual step: display instructions and wait for confirmation.

    Steps with type="manual" pause the pipeline and wait
    for user confirmation. This module only displays info.
    """
    try:
        images_dir = work_path / "images"
        temp_dir = work_path / "temp"
        image_map_path = temp_dir / "image_map.json"

        log_func("")
        log_func("=" * 70)
        log_func("   IMAGE REPLACEMENT (MANUAL STEP)")
        log_func("=" * 70)
        log_func("")

        # Show current images
        if images_dir.exists():
            image_files = sorted(images_dir.glob("*.jpg")) + sorted(images_dir.glob("*.png"))

            if image_files:
                log_func(f"Folder: {images_dir.resolve()}")
                log_func(f"   Files found: {len(image_files)}")
                log_func("")

                total_size = 0
                for img_file in image_files:
                    size_kb = img_file.stat().st_size / 1024
                    total_size += size_kb
                    log_func(f"   - {img_file.name}  ({size_kb:.1f} KB)")

                log_func(f"   ─────────────────────────────")
                log_func(f"   Total: {total_size:.1f} KB")
            else:
                log_func("Warning: no images in images/ folder")
        else:
            log_func("Warning: images/ folder not found")

        # Show anchor map
        if image_map_path.exists():
            try:
                with open(image_map_path, "r", encoding="utf-8") as f:
                    image_map = json.load(f)
                log_func("")
                log_func(f"Anchor map ({len(image_map)} entries):")
                for anchor, filename in image_map.items():
                    log_func(f"   {anchor} -> {filename}")
            except Exception:
                pass

        log_func("")
        log_func("-" * 70)
        log_func("INSTRUCTIONS:")
        log_func("   1. Open the images/ folder")
        log_func("   2. Replace files with new ones (KEEP THE FILENAMES!)")
        log_func("      img1.jpg -> replace with new img1.jpg")
        log_func("      img2.jpg -> replace with new img2.jpg")
        log_func("      etc.")
        log_func("   3. Press Enter when ready to continue")
        log_func("-" * 70)
        log_func("")

        return True

    except Exception as e:
        log_func(f"Error: {e}")
        return False
