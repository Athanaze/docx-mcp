"""
Tests for path resolution and DOCUMENT_ROOT sandboxing.
"""
import os
import pytest
import tempfile

from word_document_server.paths import (
    resolve_path, resolve_directory, resolve_docx,
    get_document_root,
)


class TestResolveWithoutRoot:
    def test_relative_path(self):
        result = resolve_path("test.docx")
        assert os.path.isabs(result)
        assert result.endswith("test.docx")

    def test_absolute_path(self):
        result = resolve_path("/tmp/test.docx")
        assert result == "/tmp/test.docx"


class TestResolveDocx:
    def test_adds_extension(self):
        result = resolve_docx("report")
        assert result.endswith(".docx")

    def test_keeps_extension(self):
        result = resolve_docx("report.docx")
        assert result.endswith(".docx")
        assert not result.endswith(".docx.docx")


class TestDocumentRootSandbox:
    def test_relative_within_root(self, tmp_dir, monkeypatch):
        monkeypatch.setenv('DOCUMENT_ROOT', tmp_dir)
        result = resolve_path("subdir/file.txt")
        assert result.startswith(os.path.realpath(tmp_dir))

    def test_absolute_within_root(self, tmp_dir, monkeypatch):
        monkeypatch.setenv('DOCUMENT_ROOT', tmp_dir)
        inner = os.path.join(tmp_dir, "file.txt")
        result = resolve_path(inner)
        assert result == os.path.realpath(inner)

    def test_absolute_outside_root_rejected(self, tmp_dir, monkeypatch):
        monkeypatch.setenv('DOCUMENT_ROOT', tmp_dir)
        with pytest.raises(ValueError, match="Access denied"):
            resolve_path("/etc/passwd")

    def test_traversal_rejected(self, tmp_dir, monkeypatch):
        monkeypatch.setenv('DOCUMENT_ROOT', tmp_dir)
        with pytest.raises(ValueError, match="Access denied"):
            resolve_path("../../../etc/passwd")

    def test_directory_within_root(self, tmp_dir, monkeypatch):
        monkeypatch.setenv('DOCUMENT_ROOT', tmp_dir)
        result = resolve_directory("subdir")
        assert result.startswith(os.path.realpath(tmp_dir))

    def test_directory_outside_root_rejected(self, tmp_dir, monkeypatch):
        monkeypatch.setenv('DOCUMENT_ROOT', tmp_dir)
        with pytest.raises(ValueError, match="Access denied"):
            resolve_directory("/tmp")


class TestDocumentRoot:
    def test_not_set(self, monkeypatch):
        monkeypatch.delenv('DOCUMENT_ROOT', raising=False)
        assert get_document_root() is None

    def test_empty_string(self, monkeypatch):
        monkeypatch.setenv('DOCUMENT_ROOT', '   ')
        assert get_document_root() is None

    def test_creates_directory(self, monkeypatch, tmp_dir):
        new_dir = os.path.join(tmp_dir, "new_root")
        monkeypatch.setenv('DOCUMENT_ROOT', new_dir)
        root = get_document_root()
        assert root is not None
        assert os.path.isdir(root)
