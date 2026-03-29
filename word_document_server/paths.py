"""
Path resolution and DOCUMENT_ROOT sandboxing for Word Document Server.
"""
import os
import shutil
from typing import Optional, Tuple


def get_document_root() -> Optional[str]:
    """
    Get the document root directory from the DOCUMENT_ROOT environment variable.

    When set, all relative file paths passed to MCP tools will be resolved
    relative to this directory instead of the server's current working directory.
    """
    root = os.environ.get('DOCUMENT_ROOT', '').strip()
    if root:
        root = os.path.abspath(os.path.expanduser(root))
        os.makedirs(root, exist_ok=True)
        return root
    return None


def _sandbox_check(resolved: str, root: str, original: str) -> str:
    root_real = os.path.realpath(root)
    if not resolved.startswith(root_real + os.sep) and resolved != root_real:
        raise ValueError(
            f"Access denied: path '{original}' resolves to '{resolved}' "
            f"which is outside DOCUMENT_ROOT '{root_real}'"
        )
    return resolved


def resolve_path(filepath: str) -> str:
    """
    Resolve a file path against DOCUMENT_ROOT with traversal protection.

    When DOCUMENT_ROOT is set, both relative and absolute paths are validated
    to stay within the root. When not set, paths resolve against CWD.
    """
    root = get_document_root()
    if root:
        if os.path.isabs(filepath):
            resolved = os.path.realpath(filepath)
        else:
            resolved = os.path.realpath(os.path.join(root, filepath))
        return _sandbox_check(resolved, root, filepath)
    if os.path.isabs(filepath):
        return filepath
    return os.path.abspath(filepath)


def resolve_directory(dirpath: str) -> str:
    """Resolve a directory path with the same sandboxing as resolve_path."""
    root = get_document_root()
    if root:
        if os.path.isabs(dirpath):
            resolved = os.path.realpath(dirpath)
        else:
            resolved = os.path.realpath(os.path.join(root, dirpath))
        return _sandbox_check(resolved, root, dirpath)
    if os.path.isabs(dirpath):
        return dirpath
    return os.path.abspath(dirpath)


def resolve_docx(filename: str) -> str:
    """Ensure filename has .docx extension and resolve against DOCUMENT_ROOT."""
    if not filename.endswith('.docx'):
        filename = filename + '.docx'
    return resolve_path(filename)


def copy_document(source: str, dest: Optional[str] = None) -> Tuple[bool, str, Optional[str]]:
    """Create a copy of a document."""
    if not os.path.exists(source):
        return False, f"Source document {source} does not exist", None
    if not dest:
        base, ext = os.path.splitext(source)
        dest = f"{base}_copy{ext}"
    try:
        shutil.copy2(source, dest)
        return True, f"Document copied to {dest}", dest
    except Exception as e:
        return False, f"Failed to copy document: {e}", None
