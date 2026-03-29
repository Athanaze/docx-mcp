"""
Visual preview: convert .docx → PDF → PNG pages.

Uses LibreOffice (headless) for docx→PDF, then poppler's pdftoppm for PDF→PNG.
Returns paths to the generated PNG files so Cursor/Claude can visually inspect them.
"""
import os
import subprocess
import shutil
import tempfile
import glob
import struct
import zlib
import re

from word_document_server.paths import resolve_docx, resolve_path


def preview_document(filename, pages=None, dpi=200, output_dir=None):
    """
    Convert a Word document to PNG images for visual inspection.

    Args:
        filename: Path to the .docx file.
        pages: Which pages to render (e.g. "1" or "1-3"). None = all pages.
        dpi: Resolution in dots per inch (default 200).
        output_dir: Directory for output PNGs. Defaults to same dir as the docx.

    Returns:
        JSON-like string with the list of generated PNG paths.
    """
    import json

    try:
        path = resolve_docx(filename)
    except ValueError as e:
        return str(e)
    if not os.path.exists(path):
        return f"Document {path} does not exist"

    lo = shutil.which('libreoffice') or shutil.which('soffice')
    if not lo:
        return "LibreOffice not found. Install it: sudo pacman -S libreoffice-fresh"

    if not shutil.which('pdftoppm'):
        return "pdftoppm not found. Install poppler: sudo pacman -S poppler"

    if output_dir is None:
        output_dir = os.path.dirname(path) or '.'
    else:
        from word_document_server.paths import resolve_directory
        try:
            output_dir = resolve_directory(output_dir)
        except ValueError as e:
            return str(e)

    # Step 1: docx → PDF via LibreOffice
    with tempfile.TemporaryDirectory(prefix="word_mcp_preview_") as tmpdir:
        try:
            result = subprocess.run(
                [lo, '--headless', '--convert-to', 'pdf', '--outdir', tmpdir, path],
                capture_output=True, text=True, timeout=120,
            )
        except subprocess.TimeoutExpired:
            return "PDF conversion timed out"
        if result.returncode != 0:
            return f"LibreOffice conversion failed: {result.stderr}"

        base = os.path.splitext(os.path.basename(path))[0]
        pdf_path = os.path.join(tmpdir, f"{base}.pdf")
        if not os.path.exists(pdf_path):
            return "PDF conversion produced no output"

        # Also save PDF alongside the docx for convenience
        final_pdf = os.path.join(output_dir, f"{base}.pdf")
        shutil.copy2(pdf_path, final_pdf)

        # Step 2: PDF → PNG via pdftoppm
        png_prefix = os.path.join(output_dir, f"{base}_page")
        cmd = ['pdftoppm', '-png', '-r', str(int(dpi))]
        if pages:
            if '-' in str(pages):
                first, last = str(pages).split('-', 1)
                cmd += ['-f', first, '-l', last]
            else:
                cmd += ['-f', str(pages), '-l', str(pages)]
        cmd += [pdf_path, png_prefix]

        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        except subprocess.TimeoutExpired:
            return "PNG conversion timed out"
        if result.returncode != 0:
            return f"pdftoppm failed: {result.stderr}"

    # Collect generated PNGs
    pattern = f"{png_prefix}*.png"
    pngs = sorted(glob.glob(pattern))

    if not pngs:
        return json.dumps({
            "success": False,
            "message": "No PNG files generated",
            "pdf": final_pdf,
        })

    page_images = []
    for idx, p in enumerate(pngs, start=1):
        width = height = None
        try:
            with open(p, "rb") as f:
                header = f.read(24)
            if header[:8] == b"\x89PNG\r\n\x1a\n" and header[12:16] == b"IHDR":
                width, height = struct.unpack(">II", header[16:24])
        except Exception:
            pass
        page_images.append({
            "page_number": idx,
            "path": p,
            "size_bytes": os.path.getsize(p),
            "width_px": width,
            "height_px": height,
        })

    return json.dumps({
        "success": True,
        "pdf": final_pdf,
        "pages": pngs,  # backwards-compatible
        "page_images": page_images,
        "count": len(pngs),
        "dpi": dpi,
    }, indent=2)


def render_document_pages(filename, pages=None, dpi=200, output_dir=None):
    """Alias of preview_document with a layout-review oriented name."""
    return preview_document(filename, pages=pages, dpi=dpi, output_dir=output_dir)


def _parse_page_number(path):
    m = re.search(r"(\d+)(?=\.png$)", os.path.basename(path), re.IGNORECASE)
    return int(m.group(1)) if m else None


def _load_png_rgb(path):
    """Decode non-interlaced 8-bit PNG into (width, height, bpp, raw-bytes)."""
    with open(path, "rb") as f:
        data = f.read()
    if data[:8] != b"\x89PNG\r\n\x1a\n":
        raise ValueError(f"Not a PNG file: {path}")

    idx = 8
    width = height = bit_depth = color_type = interlace = None
    idat = bytearray()
    while idx + 8 <= len(data):
        length = struct.unpack(">I", data[idx:idx + 4])[0]
        ctype = data[idx + 4:idx + 8]
        chunk_data = data[idx + 8:idx + 8 + length]
        idx += 12 + length
        if ctype == b"IHDR":
            width, height, bit_depth, color_type, _comp, _filt, interlace = struct.unpack(
                ">IIBBBBB", chunk_data
            )
        elif ctype == b"IDAT":
            idat.extend(chunk_data)
        elif ctype == b"IEND":
            break

    if bit_depth != 8:
        raise ValueError(f"Unsupported PNG bit depth {bit_depth} in {path}")
    if interlace != 0:
        raise ValueError(f"Interlaced PNG not supported in {path}")
    if color_type == 2:
        bpp = 3
    elif color_type == 6:
        bpp = 4
    elif color_type == 0:
        bpp = 1
    else:
        raise ValueError(f"Unsupported PNG color type {color_type} in {path}")

    raw = zlib.decompress(bytes(idat))
    stride = width * bpp
    out = bytearray(height * stride)
    pos = 0
    prev = bytearray(stride)
    for y in range(height):
        ftype = raw[pos]
        pos += 1
        filt = bytearray(raw[pos:pos + stride])
        pos += stride
        row = bytearray(stride)
        if ftype == 0:
            row[:] = filt
        elif ftype == 1:  # Sub
            for i in range(stride):
                left = row[i - bpp] if i >= bpp else 0
                row[i] = (filt[i] + left) & 0xFF
        elif ftype == 2:  # Up
            for i in range(stride):
                row[i] = (filt[i] + prev[i]) & 0xFF
        elif ftype == 3:  # Average
            for i in range(stride):
                left = row[i - bpp] if i >= bpp else 0
                row[i] = (filt[i] + ((left + prev[i]) // 2)) & 0xFF
        elif ftype == 4:  # Paeth
            for i in range(stride):
                a = row[i - bpp] if i >= bpp else 0
                b = prev[i]
                c = prev[i - bpp] if i >= bpp else 0
                p = a + b - c
                pa = abs(p - a)
                pb = abs(p - b)
                pc = abs(p - c)
                pr = a if pa <= pb and pa <= pc else (b if pb <= pc else c)
                row[i] = (filt[i] + pr) & 0xFF
        else:
            raise ValueError(f"Unsupported PNG filter {ftype} in {path}")
        out[y * stride:(y + 1) * stride] = row
        prev = row
    return width, height, bpp, bytes(out)


def _compare_pngs(before_path, after_path):
    bw, bh, bbpp, braw = _load_png_rgb(before_path)
    aw, ah, abpp, araw = _load_png_rgb(after_path)
    if (bw, bh, bbpp) != (aw, ah, abpp):
        return {
            "changed": True,
            "dimensions_changed": True,
            "before": {"width_px": bw, "height_px": bh, "channels": bbpp},
            "after": {"width_px": aw, "height_px": ah, "channels": abpp},
            "pixel_diff_percent": 100.0,
            "mean_abs_channel_diff": None,
        }
    total = len(braw)
    diff_count = 0
    abs_sum = 0
    for x, y in zip(braw, araw):
        if x != y:
            diff_count += 1
        abs_sum += abs(x - y)
    return {
        "changed": diff_count > 0,
        "dimensions_changed": False,
        "before": {"width_px": bw, "height_px": bh, "channels": bbpp},
        "after": {"width_px": aw, "height_px": ah, "channels": abpp},
        "pixel_diff_percent": round((diff_count * 100.0) / total, 6) if total else 0.0,
        "mean_abs_channel_diff": round(abs_sum / total, 6) if total else 0.0,
    }


def compare_rendered_pages(before_pages, after_pages, change_threshold_percent=0.1):
    """
    Compare before/after rendered PNG pages and report changed pages.

    Args:
        before_pages: list[str] of PNG file paths from earlier render.
        after_pages: list[str] of PNG file paths from current render.
        change_threshold_percent: minimum channel-diff percent to classify as changed.
    """
    import json

    if not before_pages or not after_pages:
        return json.dumps({
            "success": False,
            "message": "before_pages and after_pages must be non-empty lists",
        }, indent=2)
    threshold = float(change_threshold_percent)

    def _build_map(paths):
        out = {}
        for p in paths:
            resolved = resolve_path(p)
            if not os.path.exists(resolved):
                raise FileNotFoundError(resolved)
            n = _parse_page_number(resolved)
            key = str(n) if n is not None else os.path.basename(resolved)
            out[key] = resolved
        return out

    try:
        before_map = _build_map(before_pages)
        after_map = _build_map(after_pages)
    except Exception as e:
        return json.dumps({"success": False, "message": f"Path resolution failed: {e}"}, indent=2)

    all_keys = sorted(set(before_map) | set(after_map), key=lambda k: (len(k), k))
    per_page = []
    changed_pages = []
    removed_pages = []
    added_pages = []
    for key in all_keys:
        b = before_map.get(key)
        a = after_map.get(key)
        if b and not a:
            removed_pages.append(key)
            per_page.append({
                "page_key": key,
                "status": "removed",
                "before_path": b,
                "after_path": None,
                "changed": True,
            })
            continue
        if a and not b:
            added_pages.append(key)
            per_page.append({
                "page_key": key,
                "status": "added",
                "before_path": None,
                "after_path": a,
                "changed": True,
            })
            continue
        try:
            stats = _compare_pngs(b, a)
            changed = stats["dimensions_changed"] or stats["pixel_diff_percent"] >= threshold
            rec = {
                "page_key": key,
                "status": "matched",
                "before_path": b,
                "after_path": a,
                "changed": changed,
                **stats,
            }
            if changed:
                changed_pages.append(key)
            per_page.append(rec)
        except Exception as e:
            per_page.append({
                "page_key": key,
                "status": "error",
                "before_path": b,
                "after_path": a,
                "changed": True,
                "error": str(e),
            })
            changed_pages.append(key)

    return json.dumps({
        "success": True,
        "change_threshold_percent": threshold,
        "before_count": len(before_map),
        "after_count": len(after_map),
        "matched_count": len([r for r in per_page if r["status"] == "matched"]),
        "changed_pages": changed_pages,
        "added_pages": added_pages,
        "removed_pages": removed_pages,
        "changed_count": len(changed_pages) + len(added_pages) + len(removed_pages),
        "per_page": per_page,
    }, indent=2)
