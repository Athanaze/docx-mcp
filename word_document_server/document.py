"""
Document context decorator for Word Document Server.

Provides the @docx_tool decorator that eliminates boilerplate from every MCP tool:
path resolution, file validation, Document open/save, and error handling.
"""
import os
import json
import threading
import time
from functools import wraps
from docx import Document

from word_document_server.paths import resolve_docx

_PATH_LOCKS_GUARD = threading.Lock()
_PATH_LOCKS: dict[str, threading.RLock] = {}


def _get_path_lock(filename: str) -> threading.RLock:
    with _PATH_LOCKS_GUARD:
        lock = _PATH_LOCKS.get(filename)
        if lock is None:
            lock = threading.RLock()
            _PATH_LOCKS[filename] = lock
        return lock


def _lock_timeout_seconds() -> float:
    try:
        return max(0.1, float(os.getenv("WORD_MCP_LOCK_TIMEOUT_SEC", "30")))
    except Exception:
        return 30.0


def _lock_log_interval_seconds() -> float:
    try:
        return max(0.05, float(os.getenv("WORD_MCP_LOCK_LOG_INTERVAL_SEC", "5")))
    except Exception:
        return 5.0


def _acquire_write_lock(filename: str, lock: threading.RLock) -> tuple[bool, str | None, dict]:
    if lock.acquire(blocking=False):
        return True, None, {"contended": False, "waited_seconds": 0.0, "timeout_seconds": _lock_timeout_seconds()}

    timeout = _lock_timeout_seconds()
    interval = _lock_log_interval_seconds()
    started = time.monotonic()
    while True:
        elapsed = time.monotonic() - started
        remaining = timeout - elapsed
        if remaining <= 0:
            return False, (
                f"Timed out waiting for writer lock for {filename} "
                f"after {timeout:.1f}s"
            ), {
                "contended": True,
                "waited_seconds": round(elapsed, 3),
                "timeout_seconds": timeout,
            }
        wait_for = min(interval, remaining)
        if lock.acquire(timeout=wait_for):
            waited = time.monotonic() - started
            return True, None, {
                "contended": True,
                "waited_seconds": round(waited, 3),
                "timeout_seconds": timeout,
            }


def _inject_lock_diagnostics(result, filename: str, lock_diag: dict):
    """Attach lock-wait diagnostics into normal tool responses (HTTP payload)."""
    if not isinstance(result, str):
        return result
    if not lock_diag or not lock_diag.get("contended"):
        return result
    lock_payload = {
        "path": filename,
        "contended": True,
        "waited_seconds": lock_diag.get("waited_seconds", 0.0),
        "timeout_seconds": lock_diag.get("timeout_seconds"),
    }
    try:
        obj = json.loads(result)
        if isinstance(obj, dict):
            obj["_lock"] = lock_payload
            return json.dumps(obj, indent=2)
    except Exception:
        pass
    return (
        result
        + "\n"
        + f"[lock] waited {lock_payload['waited_seconds']:.3f}s for writer lock on {filename}"
    )


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

            needs_write_lock = creates or not readonly
            lock = _get_path_lock(filename)
            acquired = False
            lock_diag = {}
            if needs_write_lock:
                acquired, err, lock_diag = _acquire_write_lock(filename, lock)
                if not acquired:
                    return err
            try:
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
                    return _inject_lock_diagnostics(result, filename, lock_diag)
                except Exception as e:
                    return f"Failed: {e}"
            finally:
                if needs_write_lock and acquired:
                    lock.release()
        return wrapper
    return decorator


def raw_docx_tool(*, readonly=False):
    """
    Decorator for tools that need raw file access (e.g. footnotes via zipfile)
    rather than a python-docx Document object.

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
            needs_write_lock = not readonly
            lock = _get_path_lock(filename)
            acquired = False
            lock_diag = {}
            if needs_write_lock:
                acquired, err, lock_diag = _acquire_write_lock(filename, lock)
                if not acquired:
                    return err
            try:
                result = fn(filename, **kwargs)
                return _inject_lock_diagnostics(result, filename, lock_diag)
            except Exception as e:
                return f"Failed: {e}"
            finally:
                if needs_write_lock and acquired:
                    lock.release()
        return wrapper
    return decorator
