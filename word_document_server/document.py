"""
Document context decorator for Word Document Server.

Provides the @docx_tool decorator that eliminates boilerplate from every MCP tool:
path resolution, file validation, Document open/save, and error handling.
"""
import os
from functools import wraps
from docx import Document

from word_document_server.paths import resolve_docx


def docx_tool(*, readonly=False, creates=False):
    """
    Decorator that handles path resolution, validation, Document open/save,
    and error handling for all MCP tools.

    The decorated function receives (doc, filename, **kwargs) where doc is
    an open Document object and filename is the resolved absolute path.

    Args:
        readonly: If True, skip doc.save() after the operation.
        creates: If True, create a new blank Document instead of opening an existing one.
    """
    def decorator(fn):
        @wraps(fn)
        def wrapper(filename: str, **kwargs):
            try:
                filename = resolve_docx(filename)
            except ValueError as e:
                return str(e)

            if creates:
                parent = os.path.dirname(filename)
                if parent and not os.path.exists(parent):
                    os.makedirs(parent, exist_ok=True)
                doc = Document()
            else:
                if not os.path.exists(filename):
                    return f"Document {filename} does not exist"
                if not readonly:
                    if not os.access(filename, os.W_OK):
                        return (
                            f"Cannot modify document: {filename} is not writeable. "
                            f"Consider creating a copy first."
                        )
                try:
                    doc = Document(filename)
                except Exception as e:
                    return f"Failed to open document: {e}"

            try:
                result = fn(doc, filename, **kwargs)
                if not readonly and not creates:
                    doc.save(filename)
                elif creates:
                    doc.save(filename)
                return result
            except Exception as e:
                return f"Failed: {e}"
        return wrapper
    return decorator


def raw_docx_tool():
    """
    Decorator for tools that need raw file access (e.g. footnotes via zipfile,
    encryption via msoffcrypto) rather than a python-docx Document object.

    The decorated function receives (filename, **kwargs) where filename is
    the resolved absolute path. The function handles its own file I/O.
    """
    def decorator(fn):
        @wraps(fn)
        def wrapper(filename: str, **kwargs):
            try:
                filename = resolve_docx(filename)
            except ValueError as e:
                return str(e)
            if not os.path.exists(filename):
                return f"Document {filename} does not exist"
            try:
                return fn(filename, **kwargs)
            except Exception as e:
                return f"Failed: {e}"
        return wrapper
    return decorator
