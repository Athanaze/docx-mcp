"""
Optional regression: real-world DOCX with irregular table rows.

Requires `workspace/realworld/challenge_contract.docx` (Calibre DOCX demo from
filesamples.com). Download e.g.:

  curl -fsSL -o workspace/realworld/challenge_contract.docx \\
    https://filesamples.com/samples/document/docx/sample1.docx

If the file is absent, tests are skipped so CI stays hermetic.
"""
from pathlib import Path

import pytest

from word_document_server.operations.content import get_document_info, get_document_text

CHALLENGE = (
    Path(__file__).resolve().parent.parent
    / "workspace"
    / "realworld"
    / "challenge_contract.docx"
)


@pytest.mark.skipif(not CHALLENGE.is_file(), reason="challenge_contract.docx not in workspace/realworld/")
def test_get_document_text_does_not_fail_on_calibre_demo():
    out = get_document_text(filename=str(CHALLENGE), include_indices=False)
    assert not out.startswith("Failed:"), out
    assert len(out) > 500


@pytest.mark.skipif(not CHALLENGE.is_file(), reason="challenge_contract.docx not in workspace/realworld/")
def test_get_document_info_on_calibre_demo():
    import json

    raw = get_document_info(filename=str(CHALLENGE))
    data = json.loads(raw)
    assert data.get("table_count", 0) >= 1
    assert data.get("block_count", 0) >= 10
    assert data.get("word_count", 0) > data.get("word_count_paragraph_walk", 0)
