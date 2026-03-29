"""
Proper OOXML numbering definitions for bulleted and numbered lists.

Solves the critical bug where the old code set w:numPr on paragraphs with
hardcoded num_id values but never created the actual numbering definitions
in numbering.xml. Lists would only render on documents whose templates
already defined those IDs.
"""
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


BULLET_CHARS = ['\u2022', 'o', '\u25AA']  # disc, circle, square
NUMBER_FMTS = ['decimal', 'lowerLetter', 'lowerRoman']


def _get_or_create_numbering_part(doc):
    """Get or create the numbering part of the document."""
    try:
        numbering_part = doc.part.numbering_part
    except Exception:
        numbering_part = None

    if numbering_part is None:
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        from docx.parts.numbering import NumberingPart
        from docx.opc.part import PartFactory
        from docx.opc.packuri import PackURI
        import lxml.etree as etree

        numbering_xml = (
            '<w:numbering xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"'
            ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
            ' xmlns:o="urn:schemas-microsoft-com:office:office"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
            ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
            ' xmlns:v="urn:schemas-microsoft-com:vml"'
            ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
            ' xmlns:w10="urn:schemas-microsoft-com:office:word"'
            ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
            ' xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"'
            ' xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"'
            ' xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"'
            ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
            '/>'
        )
        numbering_elm = etree.fromstring(numbering_xml.encode('utf-8'))
        numbering_part = NumberingPart(
            PackURI('/word/numbering.xml'),
            'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
            numbering_elm,
            doc.part.package
        )
        doc.part.relate_to(numbering_part, RT.NUMBERING)

    return numbering_part


def _max_abstract_num_id(numbering_elm):
    """Find the highest abstractNumId currently in the numbering element."""
    max_id = -1
    for abstract in numbering_elm.findall(qn('w:abstractNum')):
        aid = int(abstract.get(qn('w:abstractNumId')))
        if aid > max_id:
            max_id = aid
    return max_id


def _max_num_id(numbering_elm):
    """Find the highest numId currently in the numbering element."""
    max_id = 0
    for num in numbering_elm.findall(qn('w:num')):
        nid = int(num.get(qn('w:numId')))
        if nid > max_id:
            max_id = nid
    return max_id


def _create_abstract_num(abstract_id, is_bullet=True):
    """Create a w:abstractNum element with 3 levels of bullets or numbers."""
    abstract_num = OxmlElement('w:abstractNum')
    abstract_num.set(qn('w:abstractNumId'), str(abstract_id))

    multi_level = OxmlElement('w:multiLevelType')
    multi_level.set(qn('w:val'), 'hybridMultilevel')
    abstract_num.append(multi_level)

    for level in range(3):
        lvl = OxmlElement('w:lvl')
        lvl.set(qn('w:ilvl'), str(level))

        start = OxmlElement('w:start')
        start.set(qn('w:val'), '1')
        lvl.append(start)

        num_fmt = OxmlElement('w:numFmt')
        if is_bullet:
            num_fmt.set(qn('w:val'), 'bullet')
        else:
            num_fmt.set(qn('w:val'), NUMBER_FMTS[level % len(NUMBER_FMTS)])
        lvl.append(num_fmt)

        lvl_text = OxmlElement('w:lvlText')
        if is_bullet:
            lvl_text.set(qn('w:val'), BULLET_CHARS[level % len(BULLET_CHARS)])
        else:
            lvl_text.set(qn('w:val'), f'%{level + 1}.')
        lvl.append(lvl_text)

        lvl_jc = OxmlElement('w:lvlJc')
        lvl_jc.set(qn('w:val'), 'left')
        lvl.append(lvl_jc)

        pPr = OxmlElement('w:pPr')
        ind = OxmlElement('w:ind')
        left = 720 * (level + 1)
        hanging = 360
        ind.set(qn('w:left'), str(left))
        ind.set(qn('w:hanging'), str(hanging))
        pPr.append(ind)
        lvl.append(pPr)

        if is_bullet:
            rPr = OxmlElement('w:rPr')
            rFonts = OxmlElement('w:rFonts')
            if level == 0:
                rFonts.set(qn('w:ascii'), 'Symbol')
                rFonts.set(qn('w:hAnsi'), 'Symbol')
            elif level == 1:
                rFonts.set(qn('w:ascii'), 'Courier New')
                rFonts.set(qn('w:hAnsi'), 'Courier New')
            else:
                rFonts.set(qn('w:ascii'), 'Wingdings')
                rFonts.set(qn('w:hAnsi'), 'Wingdings')
            rPr.append(rFonts)
            lvl.append(rPr)

        abstract_num.append(lvl)

    return abstract_num


def _create_num(num_id, abstract_num_id):
    """Create a w:num element that references an abstractNum."""
    num = OxmlElement('w:num')
    num.set(qn('w:numId'), str(num_id))
    abstract_ref = OxmlElement('w:abstractNumId')
    abstract_ref.set(qn('w:val'), str(abstract_num_id))
    num.append(abstract_ref)
    return num


def _get_abstract_format(numbering_elm, abstract_id):
    """Determine if an abstractNum is 'bullet' or 'decimal' by checking level 0."""
    for abstract in numbering_elm.findall(qn('w:abstractNum')):
        aid = int(abstract.get(qn('w:abstractNumId')))
        if aid == abstract_id:
            lvl0 = abstract.find(qn('w:lvl'))
            if lvl0 is not None:
                fmt = lvl0.find(qn('w:numFmt'))
                if fmt is not None:
                    return fmt.get(qn('w:val'))
    return None


def ensure_list_definitions(doc):
    """
    Ensure that the document has proper numbering definitions for both
    bullets and numbered lists.

    Scans existing definitions by their actual format type (not by ID),
    so this is safe across save/reload cycles where python-docx may
    renumber the abstract definitions.

    Returns (bullet_num_id, number_num_id) tuple.
    """
    numbering_part = _get_or_create_numbering_part(doc)
    numbering_elm = numbering_part.element

    # Build a map: numId -> abstractNumId
    num_to_abstract = {}
    for num in numbering_elm.findall(qn('w:num')):
        nid = int(num.get(qn('w:numId')))
        abstract_ref = num.find(qn('w:abstractNumId'))
        if abstract_ref is not None:
            num_to_abstract[nid] = int(abstract_ref.get(qn('w:val')))

    # Find existing bullet and number definitions by checking actual format
    bullet_num_id = None
    number_num_id = None

    for nid, abstract_id in num_to_abstract.items():
        fmt = _get_abstract_format(numbering_elm, abstract_id)
        if fmt == 'bullet' and bullet_num_id is None:
            bullet_num_id = nid
        elif fmt == 'decimal' and number_num_id is None:
            number_num_id = nid
        if bullet_num_id is not None and number_num_id is not None:
            break

    next_abstract = _max_abstract_num_id(numbering_elm) + 1
    next_num = _max_num_id(numbering_elm) + 1

    if bullet_num_id is None:
        abstract = _create_abstract_num(next_abstract, is_bullet=True)
        numbering_elm.append(abstract)
        num = _create_num(next_num, next_abstract)
        numbering_elm.append(num)
        bullet_num_id = next_num
        next_abstract += 1
        next_num += 1

    if number_num_id is None:
        abstract = _create_abstract_num(next_abstract, is_bullet=False)
        numbering_elm.append(abstract)
        num = _create_num(next_num, next_abstract)
        numbering_elm.append(num)
        number_num_id = next_num

    return bullet_num_id, number_num_id


def create_restart_num_id(doc, list_type="number"):
    """Create a fresh numId that restarts numbering from 1.

    In OOXML, paragraphs sharing the same numId are one continuous list.
    To start a new list, we create a new w:num pointing to the same
    abstractNumId but with a fresh numId. This makes Word restart at 1.
    """
    bullet_id, number_id = ensure_list_definitions(doc)
    base_id = bullet_id if list_type == "bullet" else number_id

    numbering_part = _get_or_create_numbering_part(doc)
    numbering_elm = numbering_part.element

    # Find which abstractNumId the base num references
    abstract_id = None
    for num in numbering_elm.findall(qn('w:num')):
        if int(num.get(qn('w:numId'))) == base_id:
            ref = num.find(qn('w:abstractNumId'))
            if ref is not None:
                abstract_id = int(ref.get(qn('w:val')))
            break

    if abstract_id is None:
        return base_id

    new_num_id = _max_num_id(numbering_elm) + 1
    num = _create_num(new_num_id, abstract_id)

    # Add numIdMacAtCleanup / lvlOverride to restart at 1
    override = OxmlElement('w:lvlOverride')
    override.set(qn('w:ilvl'), '0')
    start_override = OxmlElement('w:startOverride')
    start_override.set(qn('w:val'), '1')
    override.append(start_override)
    num.append(override)

    numbering_elm.append(num)
    return new_num_id


def set_paragraph_list(paragraph, num_id, level=0):
    """
    Apply list numbering to a paragraph.

    Args:
        paragraph: python-docx Paragraph object
        num_id: Numbering definition ID from ensure_list_definitions()
        level: Indentation level (0-based)
    """
    pPr = paragraph._element.get_or_add_pPr()

    existing = pPr.find(qn('w:numPr'))
    if existing is not None:
        pPr.remove(existing)

    numPr = OxmlElement('w:numPr')

    ilvl = OxmlElement('w:ilvl')
    ilvl.set(qn('w:val'), str(level))
    numPr.append(ilvl)

    numId = OxmlElement('w:numId')
    numId.set(qn('w:val'), str(num_id))
    numPr.append(numId)

    pPr.append(numPr)
    return paragraph
