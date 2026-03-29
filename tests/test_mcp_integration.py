"""
FastMCP integration: list_tools / call_tool through the real MCP surface.

Catches missing registrations, schema issues, and drift vs. the operations layer.
"""
import json
import os
import tempfile

import pytest
from fastmcp.client import Client

from word_document_server.server import mcp, register_tools

# Critical tools that must stay registered with stable names for agents / clients.
EXPECTED_TOOL_NAMES = frozenset({
    "create_document",
    "get_document_info",
    "get_document_text",
    "get_document_outline",
    "get_blocks",
    "list_document_styles",
    "set_paragraph_style",
    "find_text",
    "insert_content",
    "delete_block",
    "move_block",
    "add_footnote",
    "merge_cells_horizontal",
    "set_table_width",
    "preview_document",
    "render_document_pages",
    "compare_rendered_pages",
    "list_document_images",
})


@pytest.fixture(scope="module")
def _registered():
    register_tools()


@pytest.mark.asyncio
async def test_list_tools_includes_expected_names(_registered):
    async with Client(mcp) as client:
        tools = await client.list_tools()
    names = {t.name for t in tools}
    missing = EXPECTED_TOOL_NAMES - names
    assert not missing, f"Missing MCP tools: {missing}"
    assert len(names) >= 50


@pytest.mark.asyncio
async def test_call_tool_create_and_read_text(_registered):
    tmp = tempfile.mkdtemp(prefix="mcp_int_")
    path = os.path.join(tmp, "doc.docx")
    async with Client(mcp) as client:
        r = await client.call_tool(
            "create_document",
            {"filename": path, "title": "MCP Test", "author": "pytest"},
        )
        assert not r.is_error
        r2 = await client.call_tool(
            "add_paragraph",
            {"filename": path, "text": "Hello from MCP client."},
        )
        assert not r2.is_error
        r3 = await client.call_tool("get_document_text", {"filename": path})
        assert not r3.is_error
        text = r3.content[0].text
        assert "Hello from MCP client." in text


@pytest.mark.asyncio
async def test_call_tool_get_blocks_json(_registered):
    tmp = tempfile.mkdtemp(prefix="mcp_int_")
    path = os.path.join(tmp, "blocks.docx")
    async with Client(mcp) as client:
        await client.call_tool("create_document", {"filename": path})
        await client.call_tool(
            "add_paragraph",
            {"filename": path, "text": "Alpha", "bold": True},
        )
        r = await client.call_tool(
            "get_blocks",
            {"filename": path, "start_block_index": 0, "include_runs": True},
        )
        assert not r.is_error
        data = json.loads(r.content[0].text)
        assert "blocks" in data
        assert len(data["blocks"]) >= 1
        para_blocks = [b for b in data["blocks"] if b.get("type") != "table"]
        assert any("Alpha" in b.get("text", "") for b in para_blocks)


@pytest.mark.asyncio
async def test_insert_content_json_null_optionals_same_as_omitted(_registered):
    """Models often send explicit JSON null for optional tool args — must validate."""
    tmp = tempfile.mkdtemp(prefix="mcp_null_")
    path = os.path.join(tmp, "null.docx")
    async with Client(mcp) as client:
        await client.call_tool("create_document", {"filename": path})
        await client.call_tool(
            "add_paragraph",
            {"filename": path, "text": "Anchor", "style": None},
        )
        r = await client.call_tool(
            "insert_content",
            {
                "filename": path,
                "content_type": "paragraph",
                "text": "After anchor",
                "target_block_index": 0,
                "position": "after",
                "style": None,
                "target_text": None,
                "items": None,
                "table_data": None,
            },
        )
        assert not r.is_error, r.content
        t = await client.call_tool("get_document_text", {"filename": path})
        assert "After anchor" in t.content[0].text
