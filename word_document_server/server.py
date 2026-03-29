"""
Word Document MCP Server — tool registration.

All tools are synchronous functions (the docx_tool decorator handles everything).
All tools use unified block_index to address block items in document order.
"""
import os
import sys

from dotenv import load_dotenv

load_dotenv()
os.environ.setdefault('FASTMCP_LOG_LEVEL', 'INFO')

from fastmcp import FastMCP
from mcp.types import ToolAnnotations

from word_document_server.operations import content, formatting, comments
from word_document_server.operations import footnotes as fn_ops
from word_document_server.operations import pdf, preview, media

mcp = FastMCP("Word Document Server")

_tools_registered = False


def register_tools():
    """Register all MCP tools (idempotent — safe to call multiple times)."""
    global _tools_registered
    if _tools_registered:
        return

    # ------------------------------------------------------------------
    # Document lifecycle
    # ------------------------------------------------------------------

    @mcp.tool(annotations=ToolAnnotations(title="Create Document", destructiveHint=True))
    def create_document(filename: str, title: str | None = None,
                        author: str | None = None):
        """Create a new Word document with optional metadata.

        Parallelism: mutating tool; writes are serialized per filename.
        """
        return content.create_document(filename=filename, title=title, author=author)

    @mcp.tool(annotations=ToolAnnotations(title="Get Document Info", readOnlyHint=True))
    def get_document_info(filename: str):
        """Get document properties, block count, and word counts.

        word_count follows the same block iteration as get_document_text / find_text
        (includes table cell text). word_count_paragraph_walk is python-docx body
        paragraphs only (may miss table words and disagree with Word's status bar).

        Parallelism: read-only tool; safe to run in parallel.
        """
        return content.get_document_info(filename=filename)

    @mcp.tool(annotations=ToolAnnotations(title="Get Document Text", readOnlyHint=True))
    def get_document_text(
        filename: str,
        include_indices: bool = True,
        max_table_rows: int | None = None,
        max_cells_per_row: int | None = None,
        max_chars_per_cell: int | None = None,
    ):
        """Extract all text from a Word document in document order.

        Every block item gets a unified [N] index — paragraphs, headings,
        list items, and tables all share one numbering scheme. Use this
        block_index with ALL tools (format_text, delete_block, insert_content,
        add_table_row, etc.).

        Set include_indices=False for plain text without metadata.

        For very wide/nested tables, set max_table_rows, max_cells_per_row, and/or
        max_chars_per_cell to keep output readable (None = no limit).

        Parallelism: read-only tool; safe to run in parallel.
        """
        return content.get_document_text(
            filename=filename,
            include_indices=include_indices,
            max_table_rows=max_table_rows,
            max_cells_per_row=max_cells_per_row,
            max_chars_per_cell=max_chars_per_cell,
        )

    @mcp.tool(annotations=ToolAnnotations(title="Get Document Outline", readOnlyHint=True))
    def get_document_outline(filename: str, max_blocks: int | None = None):
        """Get the structure of a Word document — all blocks with types and indices.

        On very large documents, set max_blocks to limit JSON size (see total_block_count,
        truncated).

        Parallelism: read-only tool; safe to run in parallel.
        """
        return content.get_document_outline(filename=filename, max_blocks=max_blocks)

    @mcp.tool(annotations=ToolAnnotations(title="Get Blocks (Structured)", readOnlyHint=True))
    def get_blocks(filename: str, start_block_index: int = 0,
                   end_block_index: int | None = None, include_runs: bool = True):
        """Get blocks as structured JSON (text, style, runs; or table cell matrix).

        Parallelism: read-only tool; safe to run in parallel.
        """
        return content.get_blocks(
            filename=filename, start_block_index=start_block_index,
            end_block_index=end_block_index, include_runs=include_runs)

    @mcp.tool(annotations=ToolAnnotations(title="List Document Styles", readOnlyHint=True))
    def list_document_styles(filename: str):
        """List style names and types available in the document.

        Parallelism: read-only tool; safe to run in parallel.
        """
        return content.list_document_styles(filename=filename)

    @mcp.tool(annotations=ToolAnnotations(title="Set Paragraph Style", destructiveHint=True))
    def set_paragraph_style(filename: str, block_index: int, style_name: str):
        """Apply a document style to the paragraph at block_index.

        Parallelism: mutating tool; writes are serialized per filename.
        """
        return content.set_paragraph_style(
            filename=filename, block_index=block_index, style_name=style_name)

    @mcp.tool(annotations=ToolAnnotations(title="List Documents", readOnlyHint=True))
    def list_documents(directory: str = "."):
        """List all .docx files in the specified directory.

        Parallelism: read-only tool; safe to run in parallel.
        """
        return content.list_documents(directory)

    @mcp.tool(annotations=ToolAnnotations(title="Copy Document", destructiveHint=True))
    def copy_document(source_filename: str, destination_filename: str | None = None):
        """Create a copy of a Word document.

        Parallelism: mutating filesystem operation; avoid parallel writes to same destination.
        """
        return content.copy_document(source_filename, destination_filename)

    @mcp.tool(annotations=ToolAnnotations(title="Merge Documents", destructiveHint=True))
    def merge_documents(target_filename: str, source_filenames: list[str],
                        add_page_breaks: bool = True):
        """Merge multiple Word documents into one.

        Parallelism: mutating tool; writes are serialized per target filename.
        """
        return content.merge_documents(target_filename, source_filenames, add_page_breaks)

    @mcp.tool(annotations=ToolAnnotations(title="Get Document XML", readOnlyHint=True))
    def get_document_xml(filename: str):
        """Get the raw XML of a Word document.

        Parallelism: read-only tool; safe to run in parallel.
        """
        return content.get_document_xml(filename)

    @mcp.tool(annotations=ToolAnnotations(title="Convert to PDF", destructiveHint=True))
    def convert_to_pdf(filename: str, output_filename: str | None = None):
        """Convert a Word document to PDF (requires LibreOffice).

        Parallelism: read-only wrt source doc; safe in parallel for distinct output filenames.
        """
        return pdf.convert_to_pdf(filename, output_filename)

    @mcp.tool(annotations=ToolAnnotations(title="Preview Document", readOnlyHint=True))
    def preview_document(filename: str, pages: str | None = None, dpi: int = 200,
                         output_dir: str | None = None):
        """Convert a Word document to PNG images for visual inspection.

        Converts docx -> PDF (LibreOffice) -> PNG (poppler). Returns paths to
        the generated PNG files.

        Parallelism: read-only wrt source doc; safe in parallel for distinct output dirs.
        """
        return preview.preview_document(filename, pages=pages, dpi=dpi,
                                        output_dir=output_dir)

    @mcp.tool(annotations=ToolAnnotations(title="Render Document Pages", readOnlyHint=True))
    def render_document_pages(filename: str, pages: str | None = None, dpi: int = 200,
                              output_dir: str | None = None):
        """Render document pages as PNG metadata for visual layout review.

        Parallelism: read-only wrt source doc; safe in parallel for distinct output dirs.
        """
        return preview.render_document_pages(filename, pages=pages, dpi=dpi,
                                             output_dir=output_dir)

    @mcp.tool(annotations=ToolAnnotations(title="Compare Rendered Pages", readOnlyHint=True))
    def compare_rendered_pages(
        before_pages: list[str],
        after_pages: list[str],
        change_threshold_percent: float = 0.1,
    ):
        """Compare before/after page PNGs and report changed pages and diff stats.

        Parallelism: read-only tool over PNG files; safe to run in parallel.
        """
        return preview.compare_rendered_pages(
            before_pages=before_pages,
            after_pages=after_pages,
            change_threshold_percent=change_threshold_percent,
        )

    @mcp.tool(annotations=ToolAnnotations(title="List Document Images", readOnlyHint=True))
    def list_document_images(filename: str, include_usage: bool = True):
        """List embedded images with sizes and usage locations inside the document.

        Parallelism: read-only tool; safe to run in parallel.
        """
        return media.list_document_images(filename=filename, include_usage=include_usage)

    # ------------------------------------------------------------------
    # Content addition (appends at end of document)
    # ------------------------------------------------------------------

    @mcp.tool(annotations=ToolAnnotations(title="Add Table of Contents"))
    def add_table_of_contents(filename: str, title: str | None = None,
                              max_level: int | None = None):
        """Insert a Table of Contents field that renders when opened in Word/LibreOffice.

        Parallelism: mutating tool; writes are serialized per filename.
        """
        return content.add_table_of_contents(
            filename=filename,
            title="Table of Contents" if title is None else title,
            max_level=3 if max_level is None else max_level,
        )

    @mcp.tool(annotations=ToolAnnotations(title="Add Heading"))
    def add_heading(filename: str, text: str, level: int = 1,
                    font_name: str | None = None, font_size: int | None = None,
                    bold: bool | None = None, italic: bool | None = None,
                    border_bottom: bool = False):
        """Add a heading to the end of a Word document.

        Parallelism: mutating tool; writes are serialized per filename.
        """
        return content.add_heading(filename=filename, text=text, level=level,
                                   font_name=font_name, font_size=font_size,
                                   bold=bold, italic=italic,
                                   border_bottom=border_bottom)

    @mcp.tool(annotations=ToolAnnotations(title="Add Paragraph"))
    def add_paragraph(filename: str, text: str, style: str | None = None,
                      font_name: str | None = None, font_size: int | None = None,
                      bold: bool | None = None, italic: bool | None = None,
                      color: str | None = None):
        """Add a paragraph to the end of a Word document with optional formatting.

        Parallelism: mutating tool; writes are serialized per filename.
        """
        return content.add_paragraph(filename=filename, text=text, style=style,
                                     font_name=font_name, font_size=font_size,
                                     bold=bold, italic=italic, color=color)

    @mcp.tool(annotations=ToolAnnotations(title="Add Table"))
    def add_table(filename: str, rows: int, cols: int,
                  data: list[list[str]] | None = None):
        """Add a table to the end of a Word document.

        Parallelism: mutating tool; writes are serialized per filename.
        """
        return content.add_table(filename=filename, rows=rows, cols=cols, data=data)

    @mcp.tool(annotations=ToolAnnotations(title="Add Picture"))
    def add_picture(filename: str, image_path: str, width: float | None = None):
        """Add an image to a Word document.

        Parallelism: mutating tool; writes are serialized per filename.
        """
        return content.add_picture(filename=filename, image_path=image_path, width=width)

    @mcp.tool(annotations=ToolAnnotations(title="Add Page Break"))
    def add_page_break(filename: str):
        """Add a page break to the document.

        Parallelism: mutating tool; writes are serialized per filename.
        """
        return content.add_page_break(filename=filename)

    @mcp.tool(annotations=ToolAnnotations(title="Add List"))
    def add_list(filename: str, items: list[str], list_type: str = "bullet",
                 level: int = 0):
        """Add a bulleted or numbered list.
        list_type: 'bullet' or 'number'. level: indentation level (0-based).

        Parallelism: mutating tool; writes are serialized per filename.
        """
        return content.add_list(filename=filename, items=items,
                                list_type=list_type, level=level)

    # ------------------------------------------------------------------
    # Positioned insertion & editing
    # ------------------------------------------------------------------

    @mcp.tool(annotations=ToolAnnotations(title="Insert Content"))
    def insert_content(
            filename: str, content_type: str | None = None,
            text: str | None = None, target_text: str | None = None,
            target_block_index: int | None = None,
            position: str | None = None, style: str | None = None,
            items: list[str] | None = None, list_type: str | None = None,
            level: int | None = None,
            table_rows: int | None = None, table_cols: int | None = None,
            table_data: list[list[str]] | None = None):
        """Insert a heading, paragraph, list, or table before/after a target block.

        Find the target by target_text or target_block_index (from get_document_text).
        content_type: 'heading', 'paragraph', 'list', or 'table'.

        Optional fields may be omitted or JSON null (same meaning).

        Parallelism: mutating tool; writes are serialized per filename.
        """
        return content.insert_content(
            filename=filename,
            content_type=content_type or "paragraph",
            text="" if text is None else text,
            target_text=target_text, target_block_index=target_block_index,
            position=position or "after", style=style, items=items,
            list_type=list_type or "bullet",
            level=1 if level is None else level,
            table_rows=table_rows, table_cols=table_cols, table_data=table_data)

    @mcp.tool(annotations=ToolAnnotations(title="Set Paragraph Text"))
    def set_paragraph_text(filename: str, new_text: str,
                           block_index: int | None = None,
                           target_text: str | None = None,
                           preserve_formatting: bool = True):
        """Replace the text of an existing paragraph block.

        Find by block_index or target_text. Keeps formatting by default.

        Parallelism: mutating tool; writes are serialized per filename.
        """
        return content.set_paragraph_text(
            filename=filename, new_text=new_text,
            block_index=block_index, target_text=target_text,
            preserve_formatting=preserve_formatting)

    # ------------------------------------------------------------------
    # Block operations (work on any block type)
    # ------------------------------------------------------------------

    @mcp.tool(annotations=ToolAnnotations(title="Delete Block", destructiveHint=True))
    def delete_block(filename: str, block_index: int):
        """Delete any block item (paragraph, heading, list item, or table) by its index.

        Parallelism: mutating tool; writes are serialized per filename.
        """
        return content.delete_block(filename=filename, block_index=block_index)

    @mcp.tool(annotations=ToolAnnotations(title="Move Block"))
    def move_block(filename: str, source_index: int, target_index: int,
                   position: str = "after"):
        """Move any block item from one position to before/after another.

        Parallelism: mutating tool; writes are serialized per filename.
        """
        return content.move_block(filename=filename,
                                  source_index=source_index,
                                  target_index=target_index,
                                  position=position)

    # ------------------------------------------------------------------
    # Search & replace
    # ------------------------------------------------------------------

    @mcp.tool(annotations=ToolAnnotations(title="Search and Replace", destructiveHint=True))
    def search_and_replace(filename: str, find_text: str, replace_text: str):
        """Search for text and replace all occurrences (preserves formatting).

        Parallelism: mutating tool; writes are serialized per filename.
        """
        return content.search_and_replace(filename=filename,
                                          find_text=find_text,
                                          replace_text=replace_text)

    @mcp.tool(annotations=ToolAnnotations(title="Find Text", readOnlyHint=True))
    def find_text(filename: str, text_to_find: str, match_case: bool = True,
                  whole_word: bool = False, max_results: int | None = None):
        """Find occurrences of text in a document. Returns block_index for each match.

        count is always the total hits; matches lists at most max_results entries
        (default: all). Use max_results on large docs to keep responses small.

        Parallelism: read-only tool; safe to run in parallel.
        """
        return content.find_text(filename=filename,
                                 text_to_find=text_to_find,
                                 match_case=match_case,
                                 whole_word=whole_word,
                                 max_results=max_results)

    @mcp.tool(annotations=ToolAnnotations(title="Replace Block"))
    def replace_block(filename: str, header_text: str | None = None,
                      start_anchor: str | None = None, end_anchor: str | None = None,
                      new_paragraphs: list[str] | None = None,
                      style: str | None = None):
        """Replace content below a header or between text anchors.

        Parallelism: mutating tool; writes are serialized per filename.
        """
        return content.replace_block(filename=filename, header_text=header_text,
                                     start_anchor=start_anchor,
                                     end_anchor=end_anchor,
                                     new_paragraphs=new_paragraphs,
                                     style=style)

    # ------------------------------------------------------------------
    # Text formatting
    # ------------------------------------------------------------------

    @mcp.tool(annotations=ToolAnnotations(title="Format Text"))
    def format_text(filename: str, block_index: int | None = None,
                    start_pos: int | None = None, end_pos: int | None = None,
                    search_text: str | None = None, match_occurrence: int = 1,
                    bold: bool | None = None, italic: bool | None = None,
                    underline: bool | None = None, color: str | None = None,
                    font_size: int | None = None, font_name: str | None = None):
        """Format text within a paragraph block by text search or character positions.

        Preferred: provide search_text to find and format matching text.
        Optionally add block_index to narrow scope. Use match_occurrence
        to select the Nth match (1-based).

        Parallelism: mutating tool; writes are serialized per filename.
        """
        return formatting.format_text(
            filename=filename, block_index=block_index,
            start_pos=start_pos, end_pos=end_pos,
            search_text=search_text, match_occurrence=match_occurrence,
            bold=bold, italic=italic, underline=underline,
            color=color, font_size=font_size, font_name=font_name)

    @mcp.tool(annotations=ToolAnnotations(title="Create Style"))
    def create_style(filename: str, style_name: str, bold: bool | None = None,
                     italic: bool | None = None, font_size: int | None = None,
                     font_name: str | None = None, color: str | None = None,
                     base_style: str | None = None):
        """Create a custom paragraph style in the document.

        Parallelism: mutating tool; writes are serialized per filename.
        """
        return formatting.create_style(
            filename=filename, style_name=style_name,
            bold=bold, italic=italic, font_size=font_size,
            font_name=font_name, color=color, base_style=base_style)

    # ------------------------------------------------------------------
    # Table formatting (addressed by block_index)
    # ------------------------------------------------------------------

    @mcp.tool(annotations=ToolAnnotations(title="Format Table"))
    def format_table(filename: str, block_index: int,
                     has_header_row: bool | None = None,
                     border_style: str | None = None):
        """Format a table (identified by block_index) with borders and header styling.

        Parallelism: mutating tool; writes are serialized per filename.
        """
        return formatting.format_table(
            filename=filename, block_index=block_index,
            has_header_row=has_header_row, border_style=border_style)

    @mcp.tool(annotations=ToolAnnotations(title="Set Column Widths"))
    def set_column_widths(filename: str, block_index: int, widths: list[float],
                          width_type: str = "points"):
        """Set widths of table columns (points, inches, cm, or percent)."""
        return formatting.set_column_widths(
            filename=filename, block_index=block_index,
            widths=widths, width_type=width_type)

    @mcp.tool(annotations=ToolAnnotations(title="Set Table Width"))
    def set_table_width(filename: str, block_index: int, width: float,
                        width_type: str = "points"):
        """Set the overall width of a table (points, inches, cm, percent, or auto)."""
        return formatting.set_table_width(
            filename=filename, block_index=block_index,
            width=width, width_type=width_type)

    @mcp.tool(annotations=ToolAnnotations(title="Auto Fit Table"))
    def auto_fit_table(filename: str, block_index: int):
        """Set table to auto-fit columns based on content."""
        return formatting.auto_fit_table(
            filename=filename, block_index=block_index)

    @mcp.tool(annotations=ToolAnnotations(title="Merge Cells"))
    def merge_cells(filename: str, block_index: int, start_row: int,
                    start_col: int, end_row: int, end_col: int):
        """Merge cells in a rectangular area of a table."""
        return formatting.merge_cells(
            filename=filename, block_index=block_index,
            start_row=start_row, start_col=start_col,
            end_row=end_row, end_col=end_col)

    @mcp.tool(annotations=ToolAnnotations(title="Merge Cells Horizontal"))
    def merge_cells_horizontal(filename: str, block_index: int, row: int,
                               start_col: int, end_col: int):
        """Merge cells horizontally in one row of a table."""
        return formatting.merge_cells_horizontal(
            filename=filename, block_index=block_index,
            row=row, start_col=start_col, end_col=end_col)

    @mcp.tool(annotations=ToolAnnotations(title="Merge Cells Vertical"))
    def merge_cells_vertical(filename: str, block_index: int, col: int,
                             start_row: int, end_row: int):
        """Merge cells vertically in one column of a table."""
        return formatting.merge_cells_vertical(
            filename=filename, block_index=block_index,
            col=col, start_row=start_row, end_row=end_row)

    @mcp.tool(annotations=ToolAnnotations(title="Set Cell Shading"))
    def set_table_cell_shading(filename: str, block_index: int, row_index: int,
                               col_index: int, fill_color: str,
                               pattern: str = "clear"):
        """Apply shading to a specific table cell."""
        return formatting.set_table_cell_shading(
            filename=filename, block_index=block_index,
            row_index=row_index, col_index=col_index,
            fill_color=fill_color, pattern=pattern)

    @mcp.tool(annotations=ToolAnnotations(title="Alternating Row Colors"))
    def apply_table_alternating_rows(filename: str, block_index: int,
                                     color1: str = "FFFFFF",
                                     color2: str = "F2F2F2"):
        """Apply alternating row colors to a table."""
        return formatting.apply_table_alternating_rows(
            filename=filename, block_index=block_index,
            color1=color1, color2=color2)

    @mcp.tool(annotations=ToolAnnotations(title="Highlight Table Header"))
    def highlight_table_header(filename: str, block_index: int,
                               header_color: str = "4472C4",
                               text_color: str = "FFFFFF"):
        """Apply special highlighting to the table header row."""
        return formatting.highlight_table_header(
            filename=filename, block_index=block_index,
            header_color=header_color, text_color=text_color)

    @mcp.tool(annotations=ToolAnnotations(title="Set Cell Alignment"))
    def set_cell_alignment(filename: str, block_index: int, row_index: int,
                           col_index: int, horizontal: str = "left",
                           vertical: str = "top"):
        """Set text alignment for a table cell."""
        return formatting.set_cell_alignment(
            filename=filename, block_index=block_index,
            row_index=row_index, col_index=col_index,
            horizontal=horizontal, vertical=vertical)

    @mcp.tool(annotations=ToolAnnotations(title="Set Table Alignment All"))
    def set_table_alignment_all(filename: str, block_index: int,
                                horizontal: str = "left",
                                vertical: str = "top"):
        """Set text alignment for ALL cells in a table at once."""
        return formatting.set_table_alignment_all(
            filename=filename, block_index=block_index,
            horizontal=horizontal, vertical=vertical)

    @mcp.tool(annotations=ToolAnnotations(title="Format Cell Text"))
    def format_cell_text(filename: str, block_index: int, row_index: int,
                         col_index: int, text_content: str | None = None,
                         bold: bool | None = None, italic: bool | None = None,
                         underline: bool | None = None, color: str | None = None,
                         font_size: int | None = None, font_name: str | None = None):
        """Format text within a specific table cell."""
        return formatting.format_cell_text(
            filename=filename, block_index=block_index,
            row_index=row_index, col_index=col_index,
            text_content=text_content, bold=bold, italic=italic,
            underline=underline, color=color, font_size=font_size,
            font_name=font_name)

    @mcp.tool(annotations=ToolAnnotations(title="Set Cell Padding"))
    def set_cell_padding(filename: str, block_index: int, row_index: int,
                         col_index: int, top: float | None = None,
                         bottom: float | None = None, left: float | None = None,
                         right: float | None = None, unit: str = "points"):
        """Set padding for a specific table cell."""
        return formatting.set_cell_padding(
            filename=filename, block_index=block_index,
            row_index=row_index, col_index=col_index,
            top=top, bottom=bottom, left=left, right=right, unit=unit)

    @mcp.tool(annotations=ToolAnnotations(title="Add Table Row"))
    def add_table_row(filename: str, block_index: int,
                      row_data: list[str] | None = None,
                      position: str = "end"):
        """Add a row to a table identified by block_index."""
        return content.add_table_row(filename=filename, block_index=block_index,
                                     row_data=row_data, position=position)

    @mcp.tool(annotations=ToolAnnotations(title="Delete Table Row", destructiveHint=True))
    def delete_table_row(filename: str, block_index: int, row_index: int):
        """Delete a specific row from a table."""
        return content.delete_table_row(filename=filename,
                                        block_index=block_index,
                                        row_index=row_index)

    @mcp.tool(annotations=ToolAnnotations(title="Add Table Column"))
    def add_table_column(filename: str, block_index: int,
                         col_data: list[str] | None = None,
                         position: str = "end"):
        """Add a column to a table identified by block_index."""
        return content.add_table_column(filename=filename,
                                        block_index=block_index,
                                        col_data=col_data, position=position)

    @mcp.tool(annotations=ToolAnnotations(title="Delete Table Column", destructiveHint=True))
    def delete_table_column(filename: str, block_index: int, col_index: int):
        """Delete a specific column from a table."""
        return content.delete_table_column(filename=filename,
                                           block_index=block_index,
                                           col_index=col_index)

    # ------------------------------------------------------------------
    # Comments
    # ------------------------------------------------------------------

    @mcp.tool(annotations=ToolAnnotations(title="Get Comments", readOnlyHint=True))
    def get_comments(filename: str, author: str | None = None):
        """Get all comments from a document, optionally filtered by author.

        Parallelism: read-only tool; safe to run in parallel.
        """
        return comments.get_comments(filename=filename, author=author)

    @mcp.tool(annotations=ToolAnnotations(title="Add Comment"))
    def add_comment(filename: str, block_index: int, text: str,
                    author: str | None = None, initials: str | None = None):
        """Add a comment to a specific paragraph block.

        Parallelism: mutating tool; writes are serialized per filename.
        """
        return comments.add_comment(
            filename=filename, block_index=block_index, text=text,
            author="" if author is None else author,
            initials="" if initials is None else initials,
        )

    # ------------------------------------------------------------------
    # Footnotes
    # ------------------------------------------------------------------

    @mcp.tool(annotations=ToolAnnotations(title="Add Footnote"))
    def add_footnote(filename: str, search_text: str | None = None,
                     block_index: int | None = None,
                     footnote_text: str | None = None,
                     position: str | None = None):
        """Add a real OOXML footnote to a document.
        Locate by search_text or block_index. Position: 'after' or 'before'.

        Parallelism: mutating tool; writes are serialized per filename.
        """
        return fn_ops.add_footnote(
            filename, search_text=search_text, block_index=block_index,
            footnote_text="" if footnote_text is None else footnote_text,
            position=position or "after",
        )

    @mcp.tool(annotations=ToolAnnotations(title="Delete Footnote", destructiveHint=True))
    def delete_footnote(filename: str, footnote_id: int | None = None,
                        search_text: str | None = None, clean_orphans: bool = True):
        """Delete a footnote by ID or by searching for text near it.

        Parallelism: mutating tool; writes are serialized per filename.
        """
        return fn_ops.delete_footnote(filename, footnote_id=footnote_id,
                                      search_text=search_text,
                                      clean_orphans=clean_orphans)

    @mcp.tool(annotations=ToolAnnotations(title="Validate Footnotes", readOnlyHint=True))
    def validate_footnotes(filename: str):
        """Validate all footnotes for coherence and compliance.

        Parallelism: read-only tool; safe to run in parallel.
        """
        return fn_ops.validate_footnotes(filename)

    @mcp.tool(annotations=ToolAnnotations(title="Customize Footnote Style"))
    def customize_footnote_style(filename: str, font_name: str | None = None,
                                 font_size: int | None = None):
        """Customize the Footnote Text style (font name and size).

        Parallelism: mutating tool; writes are serialized per filename.
        """
        return fn_ops.customize_footnote_style(filename, font_name=font_name,
                                               font_size=font_size)

    _tools_registered = True


def get_transport_config():
    config = {
        'transport': 'stdio', 'host': '0.0.0.0', 'port': 8000,
        'path': '/mcp', 'sse_path': '/sse',
    }
    transport = os.getenv('MCP_TRANSPORT', 'stdio').lower()
    valid = ('stdio', 'streamable-http', 'sse')
    if transport not in valid:
        transport = 'stdio'
    config['transport'] = transport
    config['host'] = os.getenv('MCP_HOST', config['host'])
    config['port'] = int(os.getenv('PORT', os.getenv('MCP_PORT', config['port'])))
    config['path'] = os.getenv('MCP_PATH', config['path'])
    config['sse_path'] = os.getenv('MCP_SSE_PATH', config['sse_path'])
    return config


def run_server():
    """Run the Word Document MCP Server."""
    config = get_transport_config()
    register_tools()

    transport = config['transport']
    print(f"Starting Word Document MCP Server ({transport} transport)...")

    try:
        if transport == 'stdio':
            mcp.run(transport='stdio')
        elif transport == 'streamable-http':
            mcp.run(transport='streamable-http', host=config['host'],
                     port=config['port'], path=config['path'])
        elif transport == 'sse':
            mcp.run(transport='sse', host=config['host'],
                     port=config['port'], path=config['sse_path'])
    except KeyboardInterrupt:
        print("\nShutting down server...")
    except Exception as e:
        print(f"Error starting server: {e}")
        sys.exit(1)

    return mcp


def main():
    run_server()


if __name__ == "__main__":
    main()
