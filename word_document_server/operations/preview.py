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

from word_document_server.paths import resolve_docx


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

    return json.dumps({
        "success": True,
        "pdf": final_pdf,
        "pages": pngs,
        "count": len(pngs),
        "dpi": dpi,
    }, indent=2)
