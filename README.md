# Office Word MCP Server

MCP (Model Context Protocol) server for **.docx** files: create, read, edit, tables, lists, footnotes, comments, PDF export, and PNG preview (via LibreOffice + poppler where installed).

Implementation: **[python-docx](https://python-docx.readthedocs.io/)** + **[FastMCP](https://github.com/jlowin/fastmcp)**. Block-level content is addressed with a single **`block_index`** in document order (paragraphs, headings, list items, tables). See [`word_document_server/operations/`](word_document_server/operations/) and tool registration in [`word_document_server/server.py`](word_document_server/server.py).

## Running the server (HTTP only)

Use Streamable HTTP (default) on the same machine or a trusted network:

```bash
export MCP_TRANSPORT=streamable-http
export MCP_HOST=127.0.0.1
export MCP_PORT=8000
export DOCUMENT_ROOT="$(pwd)/workspace"   # optional sandbox for paths
uv run word_mcp_server
```

Endpoint: `http://<host>:<port>/mcp` (default path; override with `MCP_PATH`). Optional legacy SSE: `MCP_TRANSPORT=sse` and `MCP_SSE_PATH`.

| Variable | Default | Meaning |
|----------|---------|--------|
| `MCP_TRANSPORT` | `streamable-http` | `streamable-http` or `sse` |
| `MCP_HOST` / `PORT` / `MCP_PORT` | `0.0.0.0` / `8000` | HTTP bind host/port |
| `MCP_PATH` | `/mcp` | Streamable HTTP path |
| `DOCUMENT_ROOT` | _(unset)_ | If set, `.docx` paths are confined under this directory |

Local `127.0.0.1` HTTP does not add TLS; that is up to your environment if you expose the service. See [SECURITY.md](SECURITY.md).

## Requirements

- **Python 3.11+**
- **PDF / PNG preview:** LibreOffice (`soffice` / `libreoffice`) and `pdftoppm` (poppler)

## Install

```bash
git clone https://github.com/Athanaze/docx-mcp.git
cd docx-mcp
uv sync --extra dev
```

Or: `pip install -e ".[dev]"` in a virtualenv.

## What the tools cover (summary)

- **Lifecycle:** create/copy/merge documents, list files, metadata, raw XML dump
- **Reading:** `get_document_text`, `get_document_outline`, `get_blocks` (structured JSON), `find_text` (body + table cells)
- **Writing:** headings, paragraphs, lists, tables, images, page breaks, TOC field, `insert_content` / `delete_block` / `move_block`, search-and-replace
- **Formatting:** runs and table cells, custom styles, table layout (widths, merge, shading, alignment, auto-fit)
- **Footnotes:** add/delete/validate, footnote text style
- **Comments:** list and add (by `block_index`)
- **Export:** `convert_to_pdf`, `preview_document` (PNG pages for visual check)

Exact names and parameters match what your MCP client lists from the server—there is no separate Python API doc in this file.

## Troubleshooting (current behaviour)

These are **not** bugs to “fix” in code unless we extend features; they reflect how Word files and the stack behave:

1. **Styles** — Templates with normal Word styles behave most predictably. If a style is missing, many tools still apply **direct** formatting or fall back to available styles.
2. **Read-only files** — Saves fail with a clear error; use **`copy_document`** then edit the copy.
3. **Images** — `image_path` is resolved like other files (`DOCUMENT_ROOT` applies when set). Use paths the server process can read; formats depend on **python-docx** / Pillow.
4. **Colours** — **`#RRGGBB` or `RRGGBB`** and named colours (`red`, `blue`, …) are accepted (see [`helpers.parse_color`](word_document_server/operations/helpers.py)).
5. **Tables** — Rows/columns are **0-based**. **`block_index`** identifies the whole table in the document. Auto-fit vs fixed widths follows OOXML; changing layout can interact with manual widths.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md).

## License

MIT — see [LICENSE](LICENSE).

## Acknowledgements

[MCP](https://modelcontextprotocol.io/), [python-docx](https://python-docx.readthedocs.io/), [FastMCP](https://github.com/jlowin/fastmcp).
