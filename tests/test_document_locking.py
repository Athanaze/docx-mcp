"""Concurrency safety tests for per-document write serialization."""
import threading
import time

from word_document_server.document import docx_tool, _get_path_lock


def test_docx_tool_serializes_writes_on_same_file(monkeypatch, tmp_path):
    class FakeDocument:
        def __init__(self, *_args, **_kwargs):
            pass

        def save(self, _path):
            return None

    # Patch Document class used by decorator so we don't parse real .docx.
    monkeypatch.setattr("word_document_server.document.Document", FakeDocument)

    shared = {"active": 0, "max_active": 0}
    guard = threading.Lock()

    @docx_tool(creates=True)
    def slow_write(doc, filename, marker):
        _ = doc, filename, marker
        with guard:
            shared["active"] += 1
            if shared["active"] > shared["max_active"]:
                shared["max_active"] = shared["active"]
        time.sleep(0.15)
        with guard:
            shared["active"] -= 1
        return "ok"

    target = str(tmp_path / "same.docx")
    t1 = threading.Thread(target=lambda: slow_write(filename=target, marker="a"))
    t2 = threading.Thread(target=lambda: slow_write(filename=target, marker="b"))

    start = time.perf_counter()
    t1.start()
    t2.start()
    t1.join()
    t2.join()
    elapsed = time.perf_counter() - start

    assert shared["max_active"] == 1
    # Two serialized 150ms sections should exceed ~280ms with some scheduler slack.
    assert elapsed >= 0.28


def test_docx_tool_lock_timeout_returns_error(monkeypatch, tmp_path):
    class FakeDocument:
        def __init__(self, *_args, **_kwargs):
            pass

        def save(self, _path):
            return None

    monkeypatch.setattr("word_document_server.document.Document", FakeDocument)
    monkeypatch.setenv("WORD_MCP_LOCK_TIMEOUT_SEC", "0.2")
    monkeypatch.setenv("WORD_MCP_LOCK_LOG_INTERVAL_SEC", "0.05")

    @docx_tool(creates=True)
    def fast_write(doc, filename):
        _ = doc, filename
        return "ok"

    target = str(tmp_path / "timeout.docx")
    lock = _get_path_lock(target)
    hold = threading.Event()

    def holder():
        lock.acquire()
        hold.set()
        time.sleep(0.5)
        lock.release()

    t = threading.Thread(target=holder)
    t.start()
    hold.wait(timeout=1)
    result = fast_write(filename=target)
    t.join()

    assert "Timed out waiting for writer lock" in result


def test_docx_tool_lock_wait_injects_http_diagnostics(monkeypatch, tmp_path):
    class FakeDocument:
        def __init__(self, *_args, **_kwargs):
            pass

        def save(self, _path):
            return None

    monkeypatch.setattr("word_document_server.document.Document", FakeDocument)
    monkeypatch.setenv("WORD_MCP_LOCK_TIMEOUT_SEC", "2")
    monkeypatch.setenv("WORD_MCP_LOCK_LOG_INTERVAL_SEC", "0.05")

    @docx_tool(creates=True)
    def fast_write(doc, filename):
        _ = doc, filename
        return "ok"

    target = str(tmp_path / "diag.docx")
    lock = _get_path_lock(target)
    hold = threading.Event()

    def holder():
        lock.acquire()
        hold.set()
        time.sleep(0.2)
        lock.release()

    t = threading.Thread(target=holder)
    t.start()
    hold.wait(timeout=1)
    result = fast_write(filename=target)
    t.join()

    assert result.startswith("ok")
    assert "[lock] waited " in result
    assert "writer lock on" in result


def test_docx_tool_injects_lock_diagnostics_into_json(monkeypatch, tmp_path):
    import json

    class FakeDocument:
        def __init__(self, *_args, **_kwargs):
            pass

        def save(self, _path):
            return None

    monkeypatch.setattr("word_document_server.document.Document", FakeDocument)
    monkeypatch.setenv("WORD_MCP_LOCK_TIMEOUT_SEC", "2")
    monkeypatch.setenv("WORD_MCP_LOCK_LOG_INTERVAL_SEC", "0.05")

    @docx_tool(creates=True)
    def json_write(doc, filename):
        _ = doc, filename
        return json.dumps({"ok": True})

    target = str(tmp_path / "jsondiag.docx")
    lock = _get_path_lock(target)
    hold = threading.Event()

    def holder():
        lock.acquire()
        hold.set()
        time.sleep(0.2)
        lock.release()

    t = threading.Thread(target=holder)
    t.start()
    hold.wait(timeout=1)
    raw = json_write(filename=target)
    t.join()

    data = json.loads(raw)
    assert data["ok"] is True
    assert data["_lock"]["contended"] is True
    assert data["_lock"]["waited_seconds"] > 0
