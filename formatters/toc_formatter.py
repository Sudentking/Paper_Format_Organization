from .base_formatter import BaseFormatter
from utils.docx_helper import (
    set_paragraph_format, match_pattern, set_run_font,
)
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree
import re

TAB_PAGE_PATTERN = re.compile(r'^(.+?)\t(\S+)$')
DOTS_PAGE_PATTERN = re.compile(r'^(.+?)(\s*\.{2,}\s*\S+)$')
CHAPTER_NUM_PATTERN = re.compile(r'^(\d+(?:\.\d+)*)\s+')
TOC_STYLE_PATTERN = re.compile(r'^TOC\s*\d*$', re.IGNORECASE)
TRAILING_NUM_PATTERN = re.compile(r'^(.*\D)(\d+)$')
TOC_SDT_TAG = 'Table of Contents'


def is_toc_entry(text):
    text = text.strip()
    if '\t' in text and TAB_PAGE_PATTERN.match(text):
        return True
    if DOTS_PAGE_PATTERN.search(text):
        return True
    return False


def is_toc_style(style_name):
    if not style_name:
        return False
    return bool(TOC_STYLE_PATTERN.match(style_name))


def split_toc_parts(text):
    text = text.strip()
    m = TAB_PAGE_PATTERN.match(text)
    if m:
        return m.group(1), '\t' + m.group(2)
    m = DOTS_PAGE_PATTERN.match(text)
    if m:
        return m.group(1), m.group(2)
    return text, ''


def _get_sdt_toc_elements(body):
    results = []
    for child in body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag != 'sdt':
            continue
        sdtPr = child.find(qn('w:sdtPr'))
        if sdtPr is None:
            continue
        docPartObj = sdtPr.find(qn('w:docPartObj'))
        if docPartObj is None:
            continue
        gallery = docPartObj.find(qn('w:docPartGallery'))
        if gallery is not None and gallery.get(qn('w:val')) == TOC_SDT_TAG:
            sdtContent = child.find(qn('w:sdtContent'))
            if sdtContent is not None:
                results.append(sdtContent)
    return results


class TOCFormatter(BaseFormatter):
    def __init__(self, doc, config):
        super().__init__(doc, config)
        self.toc_config = config.get('toc', {})
        self.title_config = self.toc_config.get('title', {})
        self.chapter_config = self.toc_config.get('chapter', {})
        self.section_config = self.toc_config.get('section', {})
        self.subsection_config = self.toc_config.get('subsection', {})
        self.special_entries = self.toc_config.get('special_entries', [
            '摘要', 'Abstract', '结论', '参考文献', '致谢'
        ])
        self.heading_patterns = self.recognition_config.get('heading_patterns', {})

    def format(self):
        sdt_tocs = _get_sdt_toc_elements(self.doc.element.body)
        if sdt_tocs:
            for sdtContent in sdt_tocs:
                self._format_sdt_toc(sdtContent)
        else:
            self._format_paragraph_toc()

    def _format_sdt_toc(self, sdtContent):
        p_elements = sdtContent.findall(qn('w:p'))
        title_cfg = self.title_config
        line_spacing = self.toc_config.get('line_spacing', 1.25)

        for p_elem in p_elements:
            text = self._get_p_text(p_elem)
            style = self._get_p_style(p_elem)

            if not text.strip():
                continue

            if re.match(r'^目\s*录$', text.strip()):
                self._format_sdt_title(p_elem, text, title_cfg, line_spacing)
                continue

            level = self._toc_style_to_level(style)
            if level is not None:
                self._format_sdt_entry(p_elem, text, level, line_spacing)

    def _format_sdt_title(self, p_elem, text, cfg, line_spacing):
        self._clear_runs(p_elem)
        self._set_p_alignment(p_elem, cfg.get('alignment', 'center'))
        self._set_p_spacing(p_elem, cfg.get('space_before', 12), cfg.get('space_after', 12))
        self._set_p_line_spacing(p_elem, line_spacing)

        cn_font = cfg.get('font_name_cn', '黑体')
        en_font = cfg.get('font_name_en', 'Times New Roman')
        size = cfg.get('font_size', 16)
        bold = cfg.get('bold', True)

        self._add_sdt_runs(p_elem, text, cn_font, en_font, size, bold)

    def _format_sdt_entry(self, p_elem, text, level, line_spacing):
        if level == 1:
            cfg = self.chapter_config
        elif level == 2:
            cfg = self.section_config
        else:
            cfg = self.subsection_config

        indent_chars = cfg.get('indent', 0)
        indent_val = indent_chars * 240

        self._set_p_alignment(p_elem, 'left')
        self._set_p_line_spacing(p_elem, line_spacing)

        pPr = p_elem.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr')
            p_elem.insert(0, pPr)

        ind = pPr.find(qn('w:ind'))
        if indent_val > 0:
            if ind is None:
                ind = OxmlElement('w:ind')
                pPr.append(ind)
            ind.set(qn('w:left'), str(indent_val))
        else:
            if ind is not None:
                ind.set(qn('w:left'), '0')

        m = TRAILING_NUM_PATTERN.match(text.strip())
        if m:
            title_part = m.group(1)
            page_part = m.group(2)
        else:
            title_part = text.strip()
            page_part = ''

        cn_font = cfg.get('font_name_cn', '宋体')
        en_font = cfg.get('font_name_en', 'Times New Roman')
        font_size = cfg.get('font_size', 12)
        is_bold = cfg.get('bold', False)

        self._clear_runs(p_elem)
        self._add_sdt_runs(p_elem, title_part, cn_font, en_font, font_size, is_bold)
        if page_part:
            self._add_sdt_runs(p_elem, page_part, cn_font, en_font, font_size, False)

    def _format_paragraph_toc(self):
        self._format_toc_title_para()
        self._format_toc_entries_para()

    def _format_toc_title_para(self):
        title_cfg = self.title_config
        for paragraph in self.doc.paragraphs:
            if self._is_toc_title(paragraph):
                set_paragraph_format(
                    paragraph,
                    alignment=title_cfg.get('alignment', 'center'),
                    space_before=title_cfg.get('space_before', 12),
                    space_after=title_cfg.get('space_after', 12),
                    line_spacing=self.toc_config.get('line_spacing', 1.25),
                )
                self._set_para_title_runs(paragraph, title_cfg)
                break

    def _set_para_title_runs(self, paragraph, cfg):
        text = paragraph.text
        if not text.strip():
            return
        paragraph.clear()
        run = paragraph.add_run(text)
        set_run_font(
            run, cfg.get('font_name_cn', '黑体'),
            cfg.get('font_size', 16), cfg.get('bold', True),
        )
        self._set_run_en_font(run, cfg.get('font_name_en', 'Times New Roman'))

    def _format_toc_entries_para(self):
        in_toc = False
        for paragraph in self.doc.paragraphs:
            text = paragraph.text.strip()
            style_name = paragraph.style.name if paragraph.style else ''

            if self._is_toc_title(paragraph):
                in_toc = True
                continue

            if not in_toc:
                if is_toc_style(style_name):
                    in_toc = True
                else:
                    continue

            if not text:
                continue

            if self._is_toc_end(paragraph, style_name):
                break

            self._format_para_entry(paragraph)

    def _is_toc_title(self, paragraph):
        text = paragraph.text.strip()
        style_name = paragraph.style.name if paragraph.style else ''
        if style_name in ('TOC Heading', 'toc heading'):
            return True
        if re.match(r'^目\s*录$', text):
            return True
        return False

    def _is_toc_end(self, paragraph, style_name):
        text = paragraph.text.strip()
        if not text:
            return False
        if is_toc_entry(text):
            return False
        if is_toc_style(style_name):
            return False
        if style_name in ('TOC Heading', 'toc heading'):
            return False
        text_only = split_toc_parts(text)[0].strip()
        if CHAPTER_NUM_PATTERN.match(text_only):
            if '\t' not in text and not re.search(r'\.{2,}', text):
                return True
        if match_pattern(text, self.recognition_config.get('reference_title_pattern', '')):
            return True
        return False

    def _detect_entry_type(self, title_part):
        title_part = title_part.strip()
        for entry_name in self.special_entries:
            if title_part == entry_name:
                return 'special', entry_name
            if title_part.startswith(entry_name):
                after = title_part[len(entry_name):]
                if not after or after[0] in ' \t.':
                    return 'special', entry_name
        m = CHAPTER_NUM_PATTERN.match(title_part)
        if m:
            num_part = m.group(1)
            depth = num_part.count('.')
            if depth == 0:
                return 'chapter', title_part[m.end():]
            elif depth == 1:
                return 'section', title_part[m.end():]
            else:
                return 'subsection', title_part[m.end():]
        return 'unknown', title_part

    def _format_para_entry(self, paragraph):
        text = paragraph.text
        if not text.strip():
            return
        title_part, trailing_part = split_toc_parts(text)
        entry_type, _ = self._detect_entry_type(title_part)

        if entry_type == 'chapter':
            cfg = self.chapter_config
            indent_chars = cfg.get('indent', 0)
        elif entry_type == 'section':
            cfg = self.section_config
            indent_chars = cfg.get('indent', 2)
        elif entry_type == 'subsection':
            cfg = self.subsection_config
            indent_chars = cfg.get('indent', 4)
        else:
            cfg = self.section_config
            indent_chars = cfg.get('indent', 0)

        indent_pt = indent_chars * 12 if indent_chars else 0
        set_paragraph_format(paragraph, alignment='left', line_spacing=self.toc_config.get('line_spacing', 1.25))
        if indent_pt > 0:
            paragraph.paragraph_format.left_indent = Pt(indent_pt)
        else:
            paragraph.paragraph_format.left_indent = None

        cn_font = cfg.get('font_name_cn', '宋体')
        en_font = cfg.get('font_name_en', 'Times New Roman')
        font_size = cfg.get('font_size', 12)
        is_bold = True if entry_type == 'chapter' else cfg.get('bold', False)

        paragraph.clear()
        self._add_para_runs(paragraph, title_part, cn_font, en_font, font_size, is_bold)
        if trailing_part:
            self._add_para_runs(paragraph, trailing_part, cn_font, en_font, font_size, False)

    def _add_para_runs(self, paragraph, text, cn_font, en_font, font_size, bold):
        if not text:
            return
        for seg_text, is_en in self._split_by_language(text):
            run = paragraph.add_run(seg_text)
            set_run_font(run, en_font if is_en else cn_font, font_size, bold)

    def _add_sdt_runs(self, p_elem, text, cn_font, en_font, font_size, bold):
        if not text:
            return
        for seg_text, is_en in self._split_by_language(text):
            r_elem = OxmlElement('w:r')
            rPr = OxmlElement('w:rPr')

            rFonts = OxmlElement('w:rFonts')
            font = en_font if is_en else cn_font
            rFonts.set(qn('w:ascii'), font)
            rFonts.set(qn('w:hAnsi'), font)
            rFonts.set(qn('w:eastAsia'), cn_font)
            rPr.append(rFonts)

            if bold:
                b_elem = OxmlElement('w:b')
                rPr.append(b_elem)

            sz = OxmlElement('w:sz')
            sz.set(qn('w:val'), str(int(font_size * 2)))
            rPr.append(sz)
            szCs = OxmlElement('w:szCs')
            szCs.set(qn('w:val'), str(int(font_size * 2)))
            rPr.append(szCs)

            r_elem.append(rPr)

            t_elem = OxmlElement('w:t')
            t_elem.set(qn('xml:space'), 'preserve')
            t_elem.text = seg_text
            r_elem.append(t_elem)

            p_elem.append(r_elem)

    def _clear_runs(self, p_elem):
        for r in p_elem.findall(qn('w:r')):
            p_elem.remove(r)
        for hlink in p_elem.findall(qn('w:hyperlink')):
            for r in hlink.findall(qn('w:r')):
                hlink.remove(r)
            p_elem.remove(hlink)

    def _set_p_alignment(self, p_elem, alignment):
        ALIGN_MAP = {'left': 'left', 'center': 'center', 'right': 'right', 'justify': 'both'}
        jc_val = ALIGN_MAP.get(alignment, 'left')
        pPr = p_elem.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr')
            p_elem.insert(0, pPr)
        jc = pPr.find(qn('w:jc'))
        if jc is None:
            jc = OxmlElement('w:jc')
            pPr.append(jc)
        jc.set(qn('w:val'), jc_val)

    def _set_p_spacing(self, p_elem, before, after):
        pPr = p_elem.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr')
            p_elem.insert(0, pPr)
        spacing = pPr.find(qn('w:spacing'))
        if spacing is None:
            spacing = OxmlElement('w:spacing')
            pPr.append(spacing)
        if before is not None:
            spacing.set(qn('w:before'), str(int(before * 20)))
        if after is not None:
            spacing.set(qn('w:after'), str(int(after * 20)))

    def _set_p_line_spacing(self, p_elem, multiplier):
        pPr = p_elem.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr')
            p_elem.insert(0, pPr)
        spacing = pPr.find(qn('w:spacing'))
        if spacing is None:
            spacing = OxmlElement('w:spacing')
            pPr.append(spacing)
        spacing.set(qn('w:line'), str(int(multiplier * 240)))
        spacing.set(qn('w:lineRule'), 'auto')

    @staticmethod
    def _get_p_text(p_elem):
        runs = p_elem.findall('.//' + qn('w:r'))
        parts = []
        for r in runs:
            t = r.find(qn('w:t'))
            if t is not None and t.text:
                parts.append(t.text)
        return ''.join(parts)

    @staticmethod
    def _get_p_style(p_elem):
        pPr = p_elem.find(qn('w:pPr'))
        if pPr is not None:
            pStyle = pPr.find(qn('w:pStyle'))
            if pStyle is not None:
                return pStyle.get(qn('w:val'))
        return ''

    @staticmethod
    def _toc_style_to_level(style):
        if not style:
            return None
        m = re.match(r'^TOC\s*(\d+)$', style, re.IGNORECASE)
        if m:
            return int(m.group(1))
        return None

    def _split_by_language(self, text):
        if not text:
            return []
        segments = []
        current = ''
        current_is_en = self._is_en_or_symbol(text[0])
        for ch in text:
            ch_is_en = self._is_en_or_symbol(ch)
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

    @staticmethod
    def _is_en_or_symbol(ch):
        return bool(re.match(r'[a-zA-Z0-9\s\.\-–—:;,!?\'\"()\[\]{}/@#$%^&*+=<>~`|\\]', ch))

    @staticmethod
    def _set_run_en_font(run, en_font):
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.insert(0, rFonts)
        rFonts.set(qn('w:ascii'), en_font)
        rFonts.set(qn('w:hAnsi'), en_font)

    def update_toc(self):
        try:
            import win32com.client
            import pythoncom
            pythoncom.CoInitialize()
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(self.doc.part.package.filename)
            for field in doc.Fields:
                if field.Type == 13:
                    field.Update()
            doc.Save()
            doc.Close()
            word.Quit()
            pythoncom.CoUninitialize()
            return True
        except Exception:
            return False
