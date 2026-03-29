# Contributing

## Setup

Python **3.11+** is required.

Using [uv](https://github.com/astral-sh/uv) (recommended):

```bash
uv sync --extra dev --extra asgi
```

Or with pip:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -e ".[dev,asgi]"
```

## Tests

Unit tests plus **FastMCP in-process** checks (`tests/test_mcp_integration.py` — `list_tools` / `call_tool` against the real server surface):

```bash
uv run pytest tests/ -q
```

## Layout

- [`word_document_server/server.py`](word_document_server/server.py) — MCP tool registration (`register_tools()`).
- [`word_document_server/operations/`](word_document_server/operations/) — document logic; tools use `@docx_tool` / `@raw_docx_tool` in [`document.py`](word_document_server/document.py).
- [`word_document_server/asgi.py`](word_document_server/asgi.py) — ASGI entry for Uvicorn (HTTP MCP).

When adding a tool, implement the operation in `operations/`, register it in `server.py`, and add tests under `tests/`. Prefer a short **MCP integration** test if the tool is user-facing.

**Optional parameters:** Use `str | None = None` (and similar) on MCP wrappers so JSON `null` from clients validates the same as omitting the key. For string fields that default to `""`, normalize `None` → `""` in the wrapper before calling `operations/`.

## Roadmap (agent-oriented tools)

Implemented in-tree: structured **`get_blocks`**, **`list_document_styles`**, **`set_paragraph_style`**.

Still valuable for future work:

- Headers / footers and page numbers
- Section breaks
- Hyperlinks (add/list)
- Batch edit API with strict validation

## Style

Match existing patterns: focused changes, no drive-by refactors, keep docstrings accurate for MCP consumers.
