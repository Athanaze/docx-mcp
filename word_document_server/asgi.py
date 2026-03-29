"""
ASGI application for HTTP deployment with Uvicorn, Hypercorn, or Gunicorn.

Uses Streamable HTTP transport (recommended for remote MCP clients).

Example::

    uvicorn word_document_server.asgi:app --host 0.0.0.0 --port 8000

Environment variables (same as ``run_server``):

- ``MCP_PATH`` — URL path for the MCP endpoint (default ``/mcp``).
- ``DOCUMENT_ROOT`` — optional sandbox directory for all document paths.
"""
from word_document_server.server import get_transport_config, mcp, register_tools

register_tools()

_cfg = get_transport_config()
_path = _cfg.get("path") or "/mcp"

app = mcp.http_app(path=_path, transport="streamable-http")
