"""Shared helpers used by multiple operations modules."""
from docx.shared import Pt, RGBColor


_COLOR_MAP = {
    'red': RGBColor(255, 0, 0),
    'blue': RGBColor(0, 0, 255),
    'green': RGBColor(0, 128, 0),
    'yellow': RGBColor(255, 255, 0),
    'black': RGBColor(0, 0, 0),
    'gray': RGBColor(128, 128, 128),
    'grey': RGBColor(128, 128, 128),
    'white': RGBColor(255, 255, 255),
    'purple': RGBColor(128, 0, 128),
    'orange': RGBColor(255, 165, 0),
}


def parse_color(color_str):
    """Parse a color string (name or hex) into an RGBColor."""
    if not color_str:
        return None
    lower = color_str.lower().strip()
    if lower in _COLOR_MAP:
        return _COLOR_MAP[lower]
    cleaned = lower.lstrip('#')
    if len(cleaned) == 6:
        try:
            return RGBColor(int(cleaned[0:2], 16), int(cleaned[2:4], 16), int(cleaned[4:6], 16))
        except ValueError:
            pass
    return None


def apply_run_format(run, bold=None, italic=None, underline=None,
                     color=None, font_size=None, font_name=None):
    """Apply formatting attributes to a run."""
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    if underline is not None:
        run.underline = underline
    rgb = parse_color(color) if isinstance(color, str) else color
    if rgb is not None:
        run.font.color.rgb = rgb
    if font_size is not None:
        run.font.size = Pt(int(font_size))
    if font_name is not None:
        run.font.name = font_name


def resolve_list_style(doc):
    """Return the best available list paragraph style name from the document."""
    for candidate in ('List Paragraph', 'ListParagraph', 'Normal'):
        try:
            _ = doc.styles[candidate]
            return candidate
        except KeyError:
            continue
    return 'Normal'
