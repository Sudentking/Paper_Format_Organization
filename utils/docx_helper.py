import re
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


FONT_SIZE_MAP = {
    '初号': 42, '小初': 36,
    '一号': 26, '小一': 24,
    '二号': 22, '小二': 18,
    '三号': 16, '小三': 15,
    '四号': 14, '小四': 12,
    '五号': 10.5, '小五': 9,
    '六号': 7.5, '小六': 6.5,
    '七号': 5.5, '八号': 5,
}

ALIGNMENT_MAP = {
    'left': WD_ALIGN_PARAGRAPH.LEFT,
    'center': WD_ALIGN_PARAGRAPH.CENTER,
    'right': WD_ALIGN_PARAGRAPH.RIGHT,
    'justify': WD_ALIGN_PARAGRAPH.JUSTIFY,
}

VERTICAL_ALIGNMENT_MAP = {
    'top': WD_ALIGN_VERTICAL.TOP,
    'center': WD_ALIGN_VERTICAL.CENTER,
    'bottom': WD_ALIGN_VERTICAL.BOTTOM,
}


def is_english_or_number(char):
    return bool(re.match(r'[a-zA-Z0-9]', char))


def split_text_by_language(text):
    if not text:
        return []

    segments = []
    current = ""
    current_is_en = is_english_or_number(text[0])

    for ch in text:
        ch_is_en = is_english_or_number(ch)
        if ch_is_en == current_is_en:
            current += ch
        else:
            if current:
                segments.append((current, current_is_en))
            current = ch
            current_is_en = ch_is_en

    if current:
        segments.append((current, current_is_en))

    return segments


def set_run_font(run, font_name, font_size, bold=False):
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)


def set_run_font_with_separation(run, cn_font, en_font, font_size, bold=False, text=""):
    segments = split_text_by_language(text) if text else [(run.text, False)]

    run.font.size = Pt(font_size)
    run.font.bold = bold

    has_en = any(is_en for _, is_en in segments)
    has_cn = any(not is_en for _, is_en in segments)

    run.font.name = cn_font
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), cn_font)
    if has_en:
        rFonts.set(qn('w:ascii'), en_font)
        rFonts.set(qn('w:hAnsi'), en_font)


def format_paragraph_text(paragraph, cn_font, en_font, font_size, bold=False):
    text = paragraph.text
    if not text.strip():
        return

    paragraph.clear()
    segments = split_text_by_language(text)

    for seg_text, is_en in segments:
        run = paragraph.add_run(seg_text)
        font_name = en_font if is_en else cn_font
        set_run_font(run, font_name, font_size, bold)


def set_paragraph_format(paragraph, alignment='left', line_spacing=1.0,
                         space_before=0, space_after=0, first_line_indent=0):
    paragraph.alignment = ALIGNMENT_MAP.get(alignment, WD_ALIGN_PARAGRAPH.LEFT)

    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    paragraph.paragraph_format.line_spacing = line_spacing

    paragraph.paragraph_format.space_before = Pt(space_before)
    paragraph.paragraph_format.space_after = Pt(space_after)

    if first_line_indent > 0:
        paragraph.paragraph_format.first_line_indent = Pt(first_line_indent * 12)


def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    existing = tcPr.find(qn('w:tcBorders'))
    if existing is not None:
        tcPr.remove(existing)

    tcBorders = OxmlElement('w:tcBorders')

    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        if edge in kwargs:
            edge_data = kwargs[edge]
            edge_el = OxmlElement(f'w:{edge}')
            edge_el.set(qn('w:val'), edge_data.get('val', 'single'))
            edge_el.set(qn('w:sz'), str(edge_data.get('sz', 4)))
            edge_el.set(qn('w:space'), '0')
            edge_el.set(qn('w:color'), edge_data.get('color', '000000'))
            tcBorders.append(edge_el)

    tcPr.append(tcBorders)


def set_cell_vertical_alignment(cell, alignment='center'):
    cell.vertical_alignment = VERTICAL_ALIGNMENT_MAP.get(alignment, WD_ALIGN_VERTICAL.CENTER)


def match_pattern(text, pattern):
    if not pattern or not text:
        return False
    return bool(re.match(pattern, text.strip()))


HEADING_STYLE_MAP = {
    'Heading 1': 'level1', 'heading 1': 'level1',
    'Heading 2': 'level2', 'heading 2': 'level2',
    'Heading 3': 'level3', 'heading 3': 'level3',
    '标题 1': 'level1', '标题1': 'level1',
    '标题 2': 'level2', '标题2': 'level2',
    '标题 3': 'level3', '标题3': 'level3',
    'Title': 'level1',
    'Subtitle': 'level2',
}


def detect_heading_level(paragraph, patterns=None):
    style_name = paragraph.style.name if paragraph.style else ''
    if style_name in HEADING_STYLE_MAP:
        return HEADING_STYLE_MAP[style_name]

    for key, level in [('Heading 1', 'level1'), ('Heading 2', 'level2'), ('Heading 3', 'level3')]:
        if key.lower() in style_name.lower():
            return level

    try:
        pPr = paragraph._element.find(qn('w:pPr'))
        if pPr is not None:
            outline = pPr.find(qn('w:outlineLvl'))
            if outline is not None:
                val = int(outline.get(qn('w:val')))
                if val == 0:
                    return 'level1'
                elif val == 1:
                    return 'level2'
                elif val == 2:
                    return 'level3'
    except Exception:
        pass

    if patterns:
        text = paragraph.text.strip()
        for level in ('level3', 'level2', 'level1'):
            pattern = patterns.get(level)
            if match_pattern(text, pattern):
                return level

    return None
