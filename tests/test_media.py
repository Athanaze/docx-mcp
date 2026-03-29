"""Tests for embedded-image inspection tools."""
import json
import struct
import zlib
from pathlib import Path

from word_document_server.operations.content import add_picture
from word_document_server.operations.media import list_document_images


def _write_minimal_png(path: Path) -> None:
    def chunk(tag: bytes, data: bytes) -> bytes:
        body = tag + data
        crc = zlib.crc32(body) & 0xFFFFFFFF
        return struct.pack(">I", len(data)) + body + struct.pack(">I", crc)

    sig = b"\x89PNG\r\n\x1a\n"
    ihdr = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    raw = b"\x00" + b"\xff\x00\x00"
    data = sig + chunk(b"IHDR", ihdr) + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")
    path.write_bytes(data)


class TestListDocumentImages:
    def test_no_images(self, blank_docx):
        data = json.loads(list_document_images(filename=blank_docx))
        assert data["count"] == 0
        assert data["total_size_bytes"] == 0
        assert data["images"] == []

    def test_embedded_image_inventory(self, blank_docx, tmp_path):
        img = tmp_path / "tiny.png"
        _write_minimal_png(img)
        add_picture(filename=blank_docx, image_path=str(img))

        data = json.loads(list_document_images(filename=blank_docx))
        assert data["count"] == 1
        assert data["total_size_bytes"] > 0
        image = data["images"][0]
        assert image["extension"] == "png"
        assert image["size_bytes"] > 0
        assert image["usage_count"] >= 1
        assert any(p.endswith("document.xml") for p in image["used_in"])
