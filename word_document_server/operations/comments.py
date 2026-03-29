"""
Comment operations using native python-docx 1.2.0 API.

Replaces 376 lines of hand-rolled XML parsing with ~30 lines.
"""
import json
from word_document_server.document import docx_tool
from word_document_server.operations.blocks import resolve_paragraph_block


@docx_tool(readonly=True)
def get_comments(doc, filename, author=None):
    """Get all comments, optionally filtered by author."""
    try:
        comments = doc.comments
    except AttributeError:
        return json.dumps({
            "error": "python-docx version does not support comments API. Requires >= 1.2.0",
            "comments": []
        })

    result = []
    for c in comments:
        entry = {
            "id": c.comment_id,
            "author": c.author or "",
            "text": c.text or "",
        }
        if hasattr(c, 'timestamp') and c.timestamp:
            entry["date"] = str(c.timestamp)
        if hasattr(c, 'initials'):
            entry["initials"] = c.initials or ""
        result.append(entry)

    if author:
        result = [c for c in result if c["author"].lower() == author.lower()]

    return json.dumps({"comments": result, "count": len(result)}, indent=2)


@docx_tool()
def add_comment(doc, filename, block_index, text, author="", initials=""):
    """Add a comment to a specific paragraph block."""
    try:
        bi = resolve_paragraph_block(doc, block_index)
    except ValueError as e:
        return str(e)

    para = bi.obj
    runs = para.runs if para.runs else None

    try:
        if runs:
            doc.add_comment(runs=runs, text=text, author=author, initials=initials)
        else:
            run = para.add_run("")
            doc.add_comment(runs=[run], text=text, author=author, initials=initials)
    except AttributeError:
        return "python-docx version does not support add_comment(). Requires >= 1.2.0"

    return f"Comment added to block {block_index} in {filename}"
