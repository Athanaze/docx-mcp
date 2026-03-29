"""
Block item abstraction built on python-docx's iter_inner_content().

In OOXML, the document body is a flat sequence of block-level content:
w:p (paragraphs, headings, list items) and w:tbl (tables). python-docx
exposes this via BlockItemContainer.iter_inner_content(). This module
adds indexed access, type-safe resolution, and text search on top.
"""
import re
import unicodedata
from collections import namedtuple

from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn


BlockItem = namedtuple("BlockItem", ["index", "type", "obj"])


def _classify_paragraph(para):
    """Classify a Paragraph as 'heading', 'list_item', or 'paragraph'."""
    if para.style and para.style.name.lower().startswith("heading"):
        return "heading"
    pPr = para._element.find(qn('w:pPr'))
    if pPr is not None and pPr.find(qn('w:numPr')) is not None:
        return "list_item"
    return "paragraph"


def get_block_items(doc):
    """Return all block items in document order as a list of BlockItem."""
    items = []
    for i, item in enumerate(doc.iter_inner_content()):
        if isinstance(item, Table):
            items.append(BlockItem(i, "table", item))
        else:
            items.append(BlockItem(i, _classify_paragraph(item), item))
    return items


def resolve_block(doc, block_index):
    """Return the BlockItem at the given index, or raise ValueError."""
    block_index = int(block_index)
    items = get_block_items(doc)
    if block_index < 0 or block_index >= len(items):
        raise ValueError(
            f"Invalid block index {block_index}. "
            f"Document has {len(items)} block items."
        )
    return items[block_index]


def resolve_paragraph_block(doc, block_index):
    """Return the Paragraph at block_index, or raise if it's a table."""
    bi = resolve_block(doc, block_index)
    if bi.type == "table":
        raise ValueError(
            f"Block {block_index} is a table, not a paragraph."
        )
    return bi


def resolve_table_block(doc, block_index):
    """Return the Table at block_index, or raise if it's not a table."""
    bi = resolve_block(doc, block_index)
    if bi.type != "table":
        raise ValueError(
            f"Block {block_index} is a {bi.type}, not a table."
        )
    return bi


# ---------------------------------------------------------------------------
# Text normalization (shared with content.py via import)
# ---------------------------------------------------------------------------

_ZERO_WIDTH_RE = re.compile(
    '[\u200b\u200c\u200d\u2060\ufeff\u00ad]'
)


def normalize_text(text):
    """Normalize text for reliable matching: NFC, strip zero-width, collapse whitespace."""
    text = unicodedata.normalize('NFC', text)
    text = _ZERO_WIDTH_RE.sub('', text)
    text = text.replace('\u00a0', ' ')
    return ' '.join(text.split()).strip()


def _block_text(bi):
    """Extract searchable text from a block item."""
    if bi.type == "table":
        parts = []
        for row in bi.obj.rows:
            for cell in row.cells:
                parts.append(cell.text)
        return " ".join(parts)
    return bi.obj.text


def find_block(doc, target_text=None, block_index=None, skip_toc=True):
    """Find a block item by index or normalized text search.

    Returns BlockItem or None. Text search does exact match first,
    then fuzzy (substring) match. Empty blocks are skipped in fuzzy pass.
    """
    if block_index is not None:
        try:
            return resolve_block(doc, block_index)
        except ValueError:
            return None

    if not target_text:
        return None

    target_norm = normalize_text(target_text)
    if not target_norm:
        return None

    items = get_block_items(doc)

    for bi in items:
        if skip_toc and bi.type != "table":
            if bi.obj.style and bi.obj.style.name.upper().startswith("TOC"):
                continue
        text_norm = normalize_text(_block_text(bi))
        if text_norm == target_norm:
            return bi

    for bi in items:
        if skip_toc and bi.type != "table":
            if bi.obj.style and bi.obj.style.name.upper().startswith("TOC"):
                continue
        text_norm = normalize_text(_block_text(bi))
        if not text_norm:
            continue
        if target_norm in text_norm or text_norm in target_norm:
            return bi

    return None
