"""
PDF conversion via LibreOffice (headless).
"""
import os
import subprocess
import shutil

from word_document_server.paths import resolve_docx


def convert_to_pdf(filename, output_filename=None):
    """Convert a Word document to PDF using LibreOffice."""
    try:
        path = resolve_docx(filename)
    except ValueError as e:
        return str(e)

    if not os.path.exists(path):
        return f"Document {path} does not exist"

    lo = shutil.which('libreoffice') or shutil.which('soffice')
    if not lo:
        return ("LibreOffice not found. Install it for PDF conversion: "
                "sudo pacman -S libreoffice-fresh")

    outdir = os.path.dirname(path) or '.'

    try:
        result = subprocess.run(
            [lo, '--headless', '--convert-to', 'pdf', '--outdir', outdir, path],
            capture_output=True, text=True, timeout=120,
        )
    except subprocess.TimeoutExpired:
        return "PDF conversion timed out after 120 seconds"
    except Exception as e:
        return f"PDF conversion failed: {e}"

    if result.returncode != 0:
        return f"PDF conversion failed: {result.stderr}"

    base = os.path.splitext(os.path.basename(path))[0]
    pdf_path = os.path.join(outdir, f"{base}.pdf")

    if output_filename and os.path.exists(pdf_path):
        from word_document_server.paths import resolve_path
        try:
            final = resolve_path(output_filename)
        except ValueError as e:
            return str(e)
        if not final.lower().endswith('.pdf'):
            final = os.path.splitext(final)[0] + '.pdf'
        os.makedirs(os.path.dirname(final) or '.', exist_ok=True)
        shutil.move(pdf_path, final)
        pdf_path = final

    if os.path.exists(pdf_path):
        return f"Converted to PDF: {pdf_path}"
    return "PDF conversion completed but output file not found"
