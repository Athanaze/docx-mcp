"""
Footnote operations using robust OOXML manipulation via lxml + zipfile.

python-docx does not have a high-level API for footnotes, so we manipulate
the underlying XML directly. This module keeps only the production-ready
"robust" implementations and removes the legacy fake footnotes (superscript
text + body paragraph) that don't produce real OOXML footnotes.
"""
import os
import json
import zipfile
import shutil
from lxml import etree

from word_document_server.paths import resolve_docx

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
XML_NS = 'http://www.w3.org/XML/1998/namespace'

RESERVED_FOOTNOTE_IDS = {-1, 0, 1}
MAX_FOOTNOTE_ID = 32767


def _get_safe_footnote_id(footnotes_root):
    nsmap = {'w': W_NS}
    used_ids = set()
    for fn in footnotes_root.xpath('//w:footnote', namespaces=nsmap):
        fn_id = fn.get(f'{{{W_NS}}}id')
        if fn_id:
            try:
                used_ids.add(int(fn_id))
            except ValueError:
                pass
    candidate = 2
    while candidate in used_ids or candidate in RESERVED_FOOTNOTE_IDS:
        candidate += 1
        if candidate > MAX_FOOTNOTE_ID:
            raise ValueError("No available footnote IDs")
    return candidate


def _ensure_content_types(content_types_xml):
    ct_tree = etree.fromstring(content_types_xml)
    nsmap = {'ct': CT_NS}
    existing = ct_tree.xpath(
        "//ct:Override[@PartName='/word/footnotes.xml']", namespaces=nsmap
    )
    if existing:
        return content_types_xml
    override = etree.Element(
        f'{{{CT_NS}}}Override',
        PartName='/word/footnotes.xml',
        ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml'
    )
    ct_tree.append(override)
    return etree.tostring(ct_tree, encoding='UTF-8', xml_declaration=True, standalone="yes")


def _ensure_document_rels(document_rels_xml):
    rels_tree = etree.fromstring(document_rels_xml)
    nsmap = {'r': REL_NS}
    existing = rels_tree.xpath(
        "//r:Relationship[contains(@Type, 'footnotes')]", namespaces=nsmap
    )
    if existing:
        return document_rels_xml
    all_rels = rels_tree.xpath('//r:Relationship', namespaces=nsmap)
    existing_ids = {r.get('Id') for r in all_rels if r.get('Id')}
    rid_num = 1
    while f'rId{rid_num}' in existing_ids:
        rid_num += 1
    rel = etree.Element(
        f'{{{REL_NS}}}Relationship',
        Id=f'rId{rid_num}',
        Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes',
        Target='footnotes.xml'
    )
    rels_tree.append(rel)
    return etree.tostring(rels_tree, encoding='UTF-8', xml_declaration=True, standalone="yes")


def _create_minimal_footnotes_xml():
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="{W_NS}">
    <w:footnote w:type="separator" w:id="-1">
        <w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
            <w:r><w:separator/></w:r></w:p>
    </w:footnote>
    <w:footnote w:type="continuationSeparator" w:id="0">
        <w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
            <w:r><w:continuationSeparator/></w:r></w:p>
    </w:footnote>
</w:footnotes>'''.encode('utf-8')


def _ensure_footnote_styles(styles_root):
    nsmap = {'w': W_NS}
    if not styles_root.xpath('//w:style[@w:styleId="FootnoteReference"]', namespaces=nsmap):
        style = etree.Element(f'{{{W_NS}}}style', attrib={
            f'{{{W_NS}}}type': 'character', f'{{{W_NS}}}styleId': 'FootnoteReference'
        })
        name = etree.SubElement(style, f'{{{W_NS}}}name')
        name.set(f'{{{W_NS}}}val', 'footnote reference')
        base = etree.SubElement(style, f'{{{W_NS}}}basedOn')
        base.set(f'{{{W_NS}}}val', 'DefaultParagraphFont')
        rPr = etree.SubElement(style, f'{{{W_NS}}}rPr')
        va = etree.SubElement(rPr, f'{{{W_NS}}}vertAlign')
        va.set(f'{{{W_NS}}}val', 'superscript')
        styles_root.append(style)

    if not styles_root.xpath('//w:style[@w:styleId="FootnoteText"]', namespaces=nsmap):
        style = etree.Element(f'{{{W_NS}}}style', attrib={
            f'{{{W_NS}}}type': 'paragraph', f'{{{W_NS}}}styleId': 'FootnoteText'
        })
        name = etree.SubElement(style, f'{{{W_NS}}}name')
        name.set(f'{{{W_NS}}}val', 'footnote text')
        base = etree.SubElement(style, f'{{{W_NS}}}basedOn')
        base.set(f'{{{W_NS}}}val', 'Normal')
        pPr = etree.SubElement(style, f'{{{W_NS}}}pPr')
        sz = etree.SubElement(pPr, f'{{{W_NS}}}sz')
        sz.set(f'{{{W_NS}}}val', '20')
        styles_root.append(style)


def _read_docx_parts(filename):
    """Read all relevant parts from the docx zip."""
    parts = {}
    with zipfile.ZipFile(filename, 'r') as zin:
        parts['document'] = zin.read('word/document.xml')
        parts['content_types'] = zin.read('[Content_Types].xml')
        parts['document_rels'] = zin.read('word/_rels/document.xml.rels')
        parts['footnotes'] = (
            zin.read('word/footnotes.xml')
            if 'word/footnotes.xml' in zin.namelist()
            else _create_minimal_footnotes_xml()
        )
        parts['styles'] = (
            zin.read('word/styles.xml')
            if 'word/styles.xml' in zin.namelist()
            else f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="{W_NS}"/>'.encode()
        )
    return parts


def _write_docx(source_filename, target_filename, modified_parts):
    """Write modified parts back to a docx zip."""
    temp = target_filename + '.tmp'
    with zipfile.ZipFile(temp, 'w', zipfile.ZIP_DEFLATED) as zout:
        with zipfile.ZipFile(source_filename, 'r') as zin:
            for item in zin.infolist():
                if item.filename not in modified_parts:
                    zout.writestr(item, zin.read(item.filename))
        for name, data in modified_parts.items():
            zout.writestr(name, data)
    os.replace(temp, target_filename)


# ---------------------------------------------------------------------------
# MCP Tool implementations
# ---------------------------------------------------------------------------

def add_footnote(filename, search_text=None, block_index=None,
                 footnote_text="", position="after"):
    """
    Add a real OOXML footnote to a document.

    Args:
        filename: Path to the Word document.
        search_text: Text to search for (to place the footnote near).
        block_index: Block index of a paragraph (alternative to search_text).
        footnote_text: The footnote content.
        position: "after" or "before" the target location.
    """
    try:
        path = resolve_docx(filename)
    except ValueError as e:
        return str(e)
    if not os.path.exists(path):
        return f"Document {path} does not exist"
    if not search_text and block_index is None:
        return "Must provide either search_text or block_index"

    try:
        parts = _read_docx_parts(path)
        doc_root = etree.fromstring(parts['document'])
        footnotes_root = etree.fromstring(parts['footnotes'])
        styles_root = etree.fromstring(parts['styles'])
        nsmap = {'w': W_NS}

        # Find target paragraph and the specific run/position for the footnote ref
        insert_run_split = None  # (run_element, char_offset) for mid-run insertion

        if search_text:
            target_para = None
            for para in doc_root.xpath('//w:p', namespaces=nsmap):
                text_runs = para.xpath('.//w:r[w:t]', namespaces=nsmap)
                full_text = ''.join(
                    ''.join(r.xpath('.//w:t/text()', namespaces=nsmap))
                    for r in text_runs
                )
                if search_text in full_text:
                    target_para = para
                    # Find the exact run and offset where search_text ends
                    char_count = 0
                    target_end = full_text.index(search_text) + len(search_text)
                    if position == "before":
                        target_end = full_text.index(search_text)
                    for r in text_runs:
                        r_text = ''.join(r.xpath('.//w:t/text()', namespaces=nsmap))
                        if char_count + len(r_text) >= target_end:
                            offset_in_run = target_end - char_count
                            insert_run_split = (r, offset_in_run)
                            break
                        char_count += len(r_text)
                    break
            if target_para is None:
                return f"Text '{search_text}' not found in document"
        else:
            bi = int(block_index)
            body = doc_root.find('w:body', namespaces=nsmap)
            block_items = [
                ch for ch in body
                if ch.tag in (f'{{{W_NS}}}p', f'{{{W_NS}}}tbl')
            ]
            if bi < 0 or bi >= len(block_items):
                return f"Block index {bi} out of range (document has {len(block_items)} blocks)"
            target_el = block_items[bi]
            if target_el.tag != f'{{{W_NS}}}p':
                return f"Block {bi} is a table, not a paragraph"
            target_para = target_el

        # Validate not in header/footer
        parent = target_para.getparent()
        while parent is not None:
            if parent.tag in [f'{{{W_NS}}}hdr', f'{{{W_NS}}}ftr']:
                return "Cannot add footnote in header/footer"
            parent = parent.getparent()

        footnote_id = _get_safe_footnote_id(footnotes_root)

        # Determine insertion position
        if insert_run_split:
            run_el, offset = insert_run_split
            t_el = run_el.find(f'{{{W_NS}}}t')
            if t_el is not None and offset < len(t_el.text or ''):
                # Split the run: text before offset stays, text after goes to new run
                original_text = t_el.text or ''
                t_el.text = original_text[:offset]
                t_el.set(f'{{{XML_NS}}}space', 'preserve')
                # Create after-run with remaining text
                after_run = etree.Element(f'{{{W_NS}}}r')
                rPr = run_el.find(f'{{{W_NS}}}rPr')
                if rPr is not None:
                    import copy
                    after_run.insert(0, copy.deepcopy(rPr))
                after_t = etree.SubElement(after_run, f'{{{W_NS}}}t')
                after_t.text = original_text[offset:]
                after_t.set(f'{{{XML_NS}}}space', 'preserve')
                run_el.addnext(after_run)
            insert_pos = list(target_para).index(run_el) + 1
        elif position == "before":
            runs = target_para.xpath('.//w:r[w:t]', namespaces=nsmap)
            insert_pos = list(target_para).index(runs[0]) if runs else 0
        else:
            runs = target_para.xpath('.//w:r', namespaces=nsmap)
            insert_pos = (list(target_para).index(runs[-1]) + 1) if runs else len(target_para)

        # Create footnote reference run in document
        ref_run = etree.Element(f'{{{W_NS}}}r')
        rPr = etree.SubElement(ref_run, f'{{{W_NS}}}rPr')
        rStyle = etree.SubElement(rPr, f'{{{W_NS}}}rStyle')
        rStyle.set(f'{{{W_NS}}}val', 'FootnoteReference')
        fn_ref = etree.SubElement(ref_run, f'{{{W_NS}}}footnoteReference')
        fn_ref.set(f'{{{W_NS}}}id', str(footnote_id))
        target_para.insert(insert_pos, ref_run)

        # Create footnote content in footnotes.xml
        new_fn = etree.Element(f'{{{W_NS}}}footnote',
                               attrib={f'{{{W_NS}}}id': str(footnote_id)})
        fn_para = etree.SubElement(new_fn, f'{{{W_NS}}}p')
        pPr = etree.SubElement(fn_para, f'{{{W_NS}}}pPr')
        pStyle = etree.SubElement(pPr, f'{{{W_NS}}}pStyle')
        pStyle.set(f'{{{W_NS}}}val', 'FootnoteText')

        marker_run = etree.SubElement(fn_para, f'{{{W_NS}}}r')
        marker_rPr = etree.SubElement(marker_run, f'{{{W_NS}}}rPr')
        marker_rStyle = etree.SubElement(marker_rPr, f'{{{W_NS}}}rStyle')
        marker_rStyle.set(f'{{{W_NS}}}val', 'FootnoteReference')
        etree.SubElement(marker_run, f'{{{W_NS}}}footnoteRef')

        space_run = etree.SubElement(fn_para, f'{{{W_NS}}}r')
        space_t = etree.SubElement(space_run, f'{{{W_NS}}}t')
        space_t.set(f'{{{XML_NS}}}space', 'preserve')
        space_t.text = ' '

        text_run = etree.SubElement(fn_para, f'{{{W_NS}}}r')
        text_el = etree.SubElement(text_run, f'{{{W_NS}}}t')
        text_el.text = footnote_text

        footnotes_root.append(new_fn)
        _ensure_footnote_styles(styles_root)

        modified = {
            'word/document.xml': etree.tostring(doc_root, encoding='UTF-8',
                                                xml_declaration=True, standalone="yes"),
            'word/footnotes.xml': etree.tostring(footnotes_root, encoding='UTF-8',
                                                 xml_declaration=True, standalone="yes"),
            'word/styles.xml': etree.tostring(styles_root, encoding='UTF-8',
                                              xml_declaration=True, standalone="yes"),
            '[Content_Types].xml': _ensure_content_types(parts['content_types']),
            'word/_rels/document.xml.rels': _ensure_document_rels(parts['document_rels']),
        }
        _write_docx(path, path, modified)

        return json.dumps({
            "success": True,
            "message": f"Footnote (ID: {footnote_id}) added to {path}",
            "footnote_id": footnote_id,
            "position": position,
        })

    except Exception as e:
        return f"Failed to add footnote: {e}"


def delete_footnote(filename, footnote_id=None, search_text=None, clean_orphans=True):
    """Delete a footnote by ID or by searching for text near it."""
    try:
        path = resolve_docx(filename)
    except ValueError as e:
        return str(e)
    if not os.path.exists(path):
        return f"Document {path} does not exist"
    if not footnote_id and not search_text:
        return "Must provide either footnote_id or search_text"

    try:
        nsmap = {'w': W_NS}
        with zipfile.ZipFile(path, 'r') as zin:
            doc_xml = zin.read('word/document.xml')
            if 'word/footnotes.xml' not in zin.namelist():
                return "No footnotes in document"
            fn_xml = zin.read('word/footnotes.xml')

        doc_root = etree.fromstring(doc_xml)
        fn_root = etree.fromstring(fn_xml)

        if search_text:
            fid = None
            for para in doc_root.xpath('//w:p', namespaces=nsmap):
                text = ''.join(para.xpath('.//w:t/text()', namespaces=nsmap))
                if search_text in text:
                    refs = para.xpath('.//w:footnoteReference', namespaces=nsmap)
                    if refs:
                        fid = int(refs[0].get(f'{{{W_NS}}}id'))
                        break
            if fid is None:
                return f"No footnote found near text '{search_text}'"
            footnote_id = fid

        footnote_id = int(footnote_id)

        refs_removed = 0
        for ref in doc_root.xpath(f'//w:footnoteReference[@w:id="{footnote_id}"]', namespaces=nsmap):
            run = ref.getparent()
            if run is not None and run.tag == f'{{{W_NS}}}r':
                parent = run.getparent()
                if parent is not None:
                    parent.remove(run)
                    refs_removed += 1

        if refs_removed == 0:
            return f"Footnote {footnote_id} not found in document"

        for fn in fn_root.xpath(f'//w:footnote[@w:id="{footnote_id}"]', namespaces=nsmap):
            fn_root.remove(fn)

        if clean_orphans:
            referenced = {
                ref.get(f'{{{W_NS}}}id')
                for ref in doc_root.xpath('//w:footnoteReference', namespaces=nsmap)
                if ref.get(f'{{{W_NS}}}id')
            }
            for fn in list(fn_root.xpath('//w:footnote', namespaces=nsmap)):
                fid = fn.get(f'{{{W_NS}}}id')
                if fid and fid not in referenced and fid not in ('-1', '0'):
                    fn_root.remove(fn)

        modified = {
            'word/document.xml': etree.tostring(doc_root, encoding='UTF-8',
                                                xml_declaration=True, standalone="yes"),
            'word/footnotes.xml': etree.tostring(fn_root, encoding='UTF-8',
                                                 xml_declaration=True, standalone="yes"),
        }
        _write_docx(path, path, modified)

        return json.dumps({
            "success": True,
            "message": f"Footnote {footnote_id} deleted from {path}",
            "footnote_id": footnote_id,
        })

    except Exception as e:
        return f"Failed to delete footnote: {e}"


def validate_footnotes(filename):
    """Validate all footnotes for coherence and compliance."""
    try:
        path = resolve_docx(filename)
    except ValueError as e:
        return str(e)
    if not os.path.exists(path):
        return f"Document {path} does not exist"

    report = {
        'total_references': 0, 'total_content': 0,
        'orphaned_content': [], 'missing_references': [],
        'invalid_locations': [], 'missing_styles': [],
        'coherence_issues': [],
    }

    try:
        nsmap = {'w': W_NS}
        with zipfile.ZipFile(path, 'r') as zf:
            doc_root = etree.fromstring(zf.read('word/document.xml'))

            ref_ids = set()
            for ref in doc_root.xpath('//w:footnoteReference', namespaces=nsmap):
                rid = ref.get(f'{{{W_NS}}}id')
                if rid:
                    ref_ids.add(rid)
                    report['total_references'] += 1
                    p = ref.getparent()
                    while p is not None:
                        if p.tag in [f'{{{W_NS}}}hdr', f'{{{W_NS}}}ftr']:
                            report['invalid_locations'].append(rid)
                            break
                        p = p.getparent()

            if 'word/footnotes.xml' in zf.namelist():
                fn_root = etree.fromstring(zf.read('word/footnotes.xml'))
                content_ids = set()
                for fn in fn_root.xpath('//w:footnote', namespaces=nsmap):
                    fid = fn.get(f'{{{W_NS}}}id')
                    if fid:
                        content_ids.add(fid)
                        if fid not in ('-1', '0'):
                            report['total_content'] += 1
                report['orphaned_content'] = sorted(content_ids - ref_ids - {'-1', '0'})
                report['missing_references'] = sorted(ref_ids - content_ids)
            elif report['total_references'] > 0:
                report['coherence_issues'].append('References exist but no footnotes.xml')

            if 'word/styles.xml' in zf.namelist():
                styles = etree.fromstring(zf.read('word/styles.xml'))
                if not styles.xpath('//w:style[@w:styleId="FootnoteReference"]', namespaces=nsmap):
                    report['missing_styles'].append('FootnoteReference')
                if not styles.xpath('//w:style[@w:styleId="FootnoteText"]', namespaces=nsmap):
                    report['missing_styles'].append('FootnoteText')

        is_valid = not any([
            report['orphaned_content'], report['missing_references'],
            report['invalid_locations'], report['coherence_issues'],
        ])
        report['valid'] = is_valid
        return json.dumps(report, indent=2)

    except Exception as e:
        return f"Failed to validate footnotes: {e}"


def customize_footnote_style(filename, font_name=None, font_size=None):
    """Customize the Footnote Text style (font name and size)."""
    from docx import Document
    try:
        path = resolve_docx(filename)
    except ValueError as e:
        return str(e)
    if not os.path.exists(path):
        return f"Document {path} does not exist"

    try:
        doc = Document(path)

        try:
            fn_style = doc.styles['Footnote Text']
        except KeyError:
            fn_style = doc.styles.add_style('Footnote Text', 1)

        if font_name:
            fn_style.font.name = font_name
        if font_size:
            from docx.shared import Pt
            fn_style.font.size = Pt(int(font_size))

        doc.save(path)
        return f"Footnote style updated in {path}"

    except Exception as e:
        return f"Failed to customize footnote style: {e}"
