"""
File utility functions for Word Document Server.
"""
import os
from typing import Tuple, Optional
import shutil


def check_file_writeable(filepath: str) -> Tuple[bool, str]:
    """
    Check if a file can be written to.
    
    Args:
        filepath: Path to the file
        
    Returns:
        Tuple of (is_writeable, error_message)
    """
    # If file doesn't exist, check if directory is writeable
    if not os.path.exists(filepath):
        directory = os.path.dirname(filepath)
        # If no directory is specified (empty string), use current directory
        if directory == '':
            directory = '.'
        if not os.path.exists(directory):
            return False, f"Directory {directory} does not exist"
        if not os.access(directory, os.W_OK):
            return False, f"Directory {directory} is not writeable"
        return True, ""
    
    # If file exists, check if it's writeable
    if not os.access(filepath, os.W_OK):
        return False, f"File {filepath} is not writeable (permission denied)"
    
    # Try to open the file for writing to see if it's locked
    try:
        with open(filepath, 'a'):
            pass
        return True, ""
    except IOError as e:
        return False, f"File {filepath} is not writeable: {str(e)}"
    except Exception as e:
        return False, f"Unknown error checking file permissions: {str(e)}"


def create_document_copy(source_path: str, dest_path: Optional[str] = None) -> Tuple[bool, str, Optional[str]]:
    """
    Create a copy of a document.
    
    Args:
        source_path: Path to the source document
        dest_path: Optional path for the new document. If not provided, will use source_path + '_copy.docx'
        
    Returns:
        Tuple of (success, message, new_filepath)
    """
    if not os.path.exists(source_path):
        return False, f"Source document {source_path} does not exist", None
    
    if not dest_path:
        # Generate a new filename if not provided
        base, ext = os.path.splitext(source_path)
        dest_path = f"{base}_copy{ext}"
    
    try:
        # Simple file copy
        shutil.copy2(source_path, dest_path)
        return True, f"Document copied to {dest_path}", dest_path
    except Exception as e:
        return False, f"Failed to copy document: {str(e)}", None


def get_document_root() -> Optional[str]:
    """
    Get the document root directory from the DOCUMENT_ROOT environment variable.
    
    When set, all relative file paths passed to MCP tools will be resolved
    relative to this directory instead of the server's current working directory.
    This is critical for deployments via uvx, Claude Desktop, or any environment
    where the server's CWD is unpredictable.
    
    Returns:
        The absolute path to the document root, or None if not set.
    """
    root = os.environ.get('DOCUMENT_ROOT', '').strip()
    if root:
        root = os.path.abspath(os.path.expanduser(root))
        os.makedirs(root, exist_ok=True)
        return root
    return None


def resolve_document_path(filepath: str) -> str:
    """
    Resolve a file path against DOCUMENT_ROOT if it is a relative path.
    
    - If the path is already absolute, return it as-is.
    - If DOCUMENT_ROOT is set and the path is relative, resolve against DOCUMENT_ROOT.
    - Otherwise, resolve against the current working directory (original behavior).
    
    Args:
        filepath: The file path to resolve
        
    Returns:
        Absolute file path
    """
    if os.path.isabs(filepath):
        return filepath
    
    root = get_document_root()
    if root:
        return os.path.join(root, filepath)
    
    return os.path.abspath(filepath)


def ensure_docx_extension(filename: str) -> str:
    """
    Ensure filename has .docx extension and resolve against DOCUMENT_ROOT.
    
    This is the central path resolution point for all MCP tools.
    When DOCUMENT_ROOT is set, relative paths like "report.docx" will be
    resolved to "$DOCUMENT_ROOT/report.docx" instead of the server's CWD.
    
    Args:
        filename: The filename to check
        
    Returns:
        Absolute filename with .docx extension
    """
    if not filename.endswith('.docx'):
        filename = filename + '.docx'
    return resolve_document_path(filename)

