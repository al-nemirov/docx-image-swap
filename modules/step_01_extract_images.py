"""
DOCX Image Swap — step_01_extract_images
=========================================
1. Extract all images (inline + floating) from DOCX
2. Convert to JPG with size/quality optimization
3. Fix stretching (DPI normalization)
4. Replace with anchors {{img_N}}
5. Save correspondence map
"""

from pathlib import Path
import shutil
import json
import io
from typing import Dict, Set

try:
    from docx import Document
    from docx.oxml.ns import qn
    from PIL import Image
    HAS_LIBS = True
except ImportError:
    HAS_LIBS = False


def run(work_path: Path, step_config: dict, log_func, ai_client=None, **kwargs) -> bool:
    """
    Main function: extract, optimize, and replace images with anchors.

    Args:
        work_path: project working directory
        step_config: step parameters from config
        log_func: logging callback
        ai_client: AI client (unused)

    Returns:
        bool: True on success
    """
    if not HAS_LIBS:
        log_func("Missing libraries: python-docx, Pillow")
        log_func("   Install: pip install python-docx pillow")
        return False

    try:
        # --- SETUP PATHS ---
        source_dir = work_path / "source"
        temp_dir = work_path / "temp"
        images_dir = work_path / "images"

        temp_dir.mkdir(parents=True, exist_ok=True)

        # Clean images folder
        if images_dir.exists():
            shutil.rmtree(images_dir)
        images_dir.mkdir(parents=True, exist_ok=True)
        log_func("Images folder cleaned")

        # --- FIND SOURCE FILE ---
        log_func("Looking for source DOCX...")

        working_docx = temp_dir / "working.docx"

        if working_docx.exists():
            source_file = working_docx
            log_func("Found: temp/working.docx")
        else:
            input_name = step_config.get("input_file", "")
            if input_name:
                source_file = source_dir / input_name
            else:
                docx_files = list(source_dir.glob("*.docx"))
                if not docx_files:
                    log_func("No DOCX files found")
                    return False
                source_file = docx_files[0]

            log_func(f"Found: {source_file.name}")
            shutil.copy2(source_file, working_docx)
            source_file = working_docx
            log_func("Copied to temp/working.docx")

        # Save original backup
        original_backup = temp_dir / "original.docx"
        if not original_backup.exists():
            shutil.copy2(source_file, original_backup)
            log_func("Original saved: temp/original.docx")

        # --- LOAD DOCUMENT ---
        doc = Document(source_file)
        log_func("Document loaded")

        # --- OPTIMIZATION PARAMS ---
        jpeg_quality = step_config.get("jpeg_quality", 85)
        max_width = step_config.get("max_width", 2048)
        max_height = step_config.get("max_height", 2048)
        optimize = step_config.get("optimize", True)

        # --- VARIABLES ---
        img_counter = 0
        image_map: Dict[str, str] = {}
        processed_rids: Set[str] = set()
        skipped = 0
        errors = 0

        stats = {
            'optimized': 0,
            'saved_bytes': 0,
            'total_original': 0,
            'total_final': 0
        }

        # --- COLLECT IMAGE RELATIONSHIPS ---
        image_parts = {}
        skipped_rels = 0

        for rel in doc.part.rels.values():
            if "image" in rel.reltype:
                rId = rel.rId
                try:
                    image_parts[rId] = rel.target_part
                except Exception:
                    skipped_rels += 1

        if skipped_rels > 0:
            log_func(f"   Skipped {skipped_rels} external images")

        log_func(f"   Images found: {len(image_parts)}")

        if not image_parts:
            log_func("No images found — creating empty map")
            image_map_path = temp_dir / "image_map.json"
            with open(image_map_path, "w", encoding="utf-8") as f:
                json.dump({}, f, ensure_ascii=False, indent=2)
            doc.save(working_docx)
            return True

        # --- FIND ALL DRAWING ELEMENTS ---
        body = doc.element.body
        drawings = body.xpath('.//w:drawing | .//wp:inline | .//wp:anchor')
        log_func(f"   Drawing elements found: {len(drawings)}")

        # --- PROCESS IMAGES ---
        for drawing in drawings:
            blips = drawing.xpath('.//a:blip[@r:embed]')
            if not blips:
                continue

            blip = blips[0]
            rId = blip.get(qn('r:embed'))

            if not rId:
                continue

            if rId in processed_rids:
                skipped += 1
                continue

            if rId not in image_parts:
                log_func(f"   Warning: rId {rId} not found in relationships")
                errors += 1
                continue

            try:
                image_part = image_parts[rId]
                image_blob = image_part.blob
                original_size = len(image_blob)
                stats['total_original'] += original_size

                img_counter += 1

                # --- OPTIMIZE AND CONVERT ---
                filename = f"img{img_counter}.jpg"
                file_path = images_dir / filename

                with Image.open(io.BytesIO(image_blob)) as img:
                    original_width, original_height = img.size

                    # DPI fix
                    dpi = img.info.get('dpi', (72, 72))
                    if isinstance(dpi, tuple):
                        dpi_x, dpi_y = dpi
                    else:
                        dpi_x = dpi_y = dpi

                    if dpi_x > 96 or dpi_y > 96:
                        scale_factor = 96 / max(dpi_x, dpi_y)
                        new_width = int(original_width * scale_factor)
                        new_height = int(original_height * scale_factor)
                        img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                        log_func(f"   img{img_counter}: DPI fix {original_width}x{original_height} -> {new_width}x{new_height}")
                        stats['optimized'] += 1
                        original_width, original_height = new_width, new_height

                    # Max size limit
                    if original_width > max_width or original_height > max_height:
                        scale = min(max_width / original_width, max_height / original_height)
                        new_width = int(original_width * scale)
                        new_height = int(original_height * scale)
                        img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                        log_func(f"   img{img_counter}: Resize {original_width}x{original_height} -> {new_width}x{new_height}")
                        stats['optimized'] += 1

                    # Color mode conversion
                    if img.mode in ("RGBA", "LA", "PA"):
                        bg = Image.new("RGB", img.size, (255, 255, 255))
                        try:
                            bg.paste(img, mask=img.split()[-1])
                        except Exception:
                            bg.paste(img)
                        img = bg
                    elif img.mode == "P":
                        img = img.convert("RGBA")
                        bg = Image.new("RGB", img.size, (255, 255, 255))
                        try:
                            bg.paste(img, mask=img.split()[-1])
                        except Exception:
                            bg.paste(img)
                        img = bg
                    elif img.mode != "RGB":
                        img = img.convert("RGB")

                    # Save optimized
                    img.save(
                        file_path,
                        format="JPEG",
                        quality=jpeg_quality,
                        optimize=optimize,
                        dpi=(96, 96)
                    )

                # Stats
                final_size = file_path.stat().st_size
                stats['total_final'] += final_size

                if final_size < original_size:
                    saved = original_size - final_size
                    stats['saved_bytes'] += saved
                    savings_percent = (saved / original_size) * 100
                    log_func(f"   {filename}: {original_size/1024:.1f}KB -> {final_size/1024:.1f}KB (-{savings_percent:.1f}%)")
                else:
                    log_func(f"   {filename}: {final_size/1024:.1f}KB")

                # --- CREATE ANCHOR ---
                marker = f"{{{{img_{img_counter}}}}}"
                image_map[marker] = filename

                # --- REPLACE IN DOCUMENT ---
                run_element = drawing.getparent()
                para_element = run_element.getparent()

                if para_element.tag != qn('w:p'):
                    while para_element is not None and para_element.tag != qn('w:p'):
                        para_element = para_element.getparent()

                if para_element is not None and para_element.tag == qn('w:p'):
                    from lxml import etree

                    new_run = etree.SubElement(para_element, qn('w:r'))
                    new_text = etree.SubElement(new_run, qn('w:t'))
                    new_text.text = marker
                    new_text.set(qn('xml:space'), 'preserve')

                    para_element.insert(list(para_element).index(run_element), new_run)
                    para_element.remove(run_element)
                else:
                    log_func(f"   Warning: could not find paragraph for {marker}")

                processed_rids.add(rId)

            except Exception as e:
                log_func(f"   Error processing img{img_counter} (rId {rId}): {e}")
                errors += 1

        # --- SAVE RESULTS ---
        doc.save(working_docx)
        log_func(f"Document saved: {working_docx.name}")

        image_map_path = temp_dir / "image_map.json"
        with open(image_map_path, "w", encoding="utf-8") as f:
            json.dump(image_map, f, ensure_ascii=False, indent=2)
        log_func(f"Map saved: {image_map_path.name}")

        # --- SUMMARY ---
        log_func("")
        log_func(f"Images extracted: {img_counter}")

        if stats['optimized'] > 0:
            log_func(f"   Optimized: {stats['optimized']} images")

        if stats['saved_bytes'] > 0:
            saved_mb = stats['saved_bytes'] / (1024 * 1024)
            total_orig_mb = stats['total_original'] / (1024 * 1024)
            total_final_mb = stats['total_final'] / (1024 * 1024)
            savings_percent = (stats['saved_bytes'] / stats['total_original']) * 100

            log_func(f"   Size before: {total_orig_mb:.2f} MB")
            log_func(f"   Size after: {total_final_mb:.2f} MB")
            log_func(f"   Saved: {saved_mb:.2f} MB ({savings_percent:.1f}%)")

        if skipped:
            log_func(f"   Skipped duplicates: {skipped}")

        if errors:
            log_func(f"   Errors: {errors}")

        if img_counter > 0:
            log_func(f"   Files: img1.jpg ... img{img_counter}.jpg")

        return True

    except Exception as e:
        log_func(f"Critical error: {e}")
        import traceback
        log_func("Traceback:")
        for line in traceback.format_exc().split('\n'):
            if line.strip():
                log_func(f"   {line}")
        return False
