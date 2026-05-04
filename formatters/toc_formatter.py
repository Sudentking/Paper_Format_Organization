from .base_formatter import BaseFormatter
from utils.docx_helper import (
    set_paragraph_format, match_pattern, set_run_font,
)
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re


CHAPTER_PATTERN = re.compile(
    r'^(\d+(?:\.\d+)*)(\s+)(.+?)(\s*\.{2,}\s*\S+)$'
)

SPECIAL_ENTRY_PATTERN = re.compile(
    r'^(摘要|Abstract|结论|参考文献|致谢)(\s*\.{2,}\s*\S+)$'
)

TOC_LINE_PATTERN = re.compile(
    r'^(.+?)(\s*\.{2,}\s*\S+)$'
)


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
        self._format_toc_title()
        self._format_toc_entries()

    def _format_toc_title(self):
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
                self._set_toc_title_runs(paragraph, title_cfg)
                break

    def _set_toc_title_runs(self, paragraph, cfg):
        text = paragraph.text
        if not text.strip():
            return

        paragraph.clear()
        run = paragraph.add_run(text)
        set_run_font(
            run,
            cfg.get('font_name_cn', '黑体'),
            cfg.get('font_size', 16),
            cfg.get('bold', True),
        )
        self._set_run_en_font(run, cfg.get('font_name_en', 'Times New Roman'))

    def _format_toc_entries(self):
        in_toc = False

        for paragraph in self.doc.paragraphs:
            style_name = paragraph.style.name if paragraph.style else ''

            if self._is_toc_title(paragraph):
                in_toc = True
                continue

            if not in_toc:
                continue

            if self._is_toc_end(paragraph, style_name):
                break

            if not paragraph.text.strip():
                continue

            self._format_entry(paragraph)

    def _is_toc_title(self, paragraph):
        text = paragraph.text.strip()
        style_name = paragraph.style.name if paragraph.style else ''

        if 'TOC' in style_name or 'toc' in style_name.lower():
            return True

        if re.match(r'^目\s*录$', text):
            return True

        return False

    def _is_toc_end(self, paragraph, style_name):
        text = paragraph.text.strip()

        if match_pattern(text, self.heading_patterns.get('level1', '')):
            return True

        if text and 'TOC' not in style_name and 'toc' not in style_name.lower():
            if match_pattern(text, self.recognition_config.get('reference_title_pattern', '')):
                return True

        return False

    def _detect_entry_type(self, text):
        text = text.strip()

        for entry_name in self.special_entries:
            if text.startswith(entry_name):
                after = text[len(entry_name):]
                if not after or after[0] in ' \t.':
                    return 'special', entry_name

        m = CHAPTER_PATTERN.match(text)
        if m:
            depth = m.group(1).count('.')
            if depth == 0:
                return 'chapter', m.group(3)
            elif depth == 1:
                return 'section', m.group(3)
            else:
                return 'subsection', m.group(3)

        return 'unknown', text

    def _format_entry(self, paragraph):
        text = paragraph.text
        if not text.strip():
            return

        entry_type, _ = self._detect_entry_type(text)

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

        set_paragraph_format(
            paragraph,
            alignment='left',
            line_spacing=self.toc_config.get('line_spacing', 1.25),
        )

        if indent_pt > 0:
            paragraph.paragraph_format.left_indent = Pt(indent_pt)
        else:
            paragraph.paragraph_format.left_indent = None

        self._format_entry_runs(paragraph, text, cfg, entry_type)

    def _format_entry_runs(self, paragraph, text, cfg, entry_type):
        cn_font = cfg.get('font_name_cn', '宋体')
        en_font = cfg.get('font_name_en', 'Times New Roman')
        font_size = cfg.get('font_size', 12)
        is_bold = cfg.get('bold', False) if entry_type != 'chapter' else True

        m = TOC_LINE_PATTERN.match(text)
        if m:
            title_part = m.group(1)
            trailing_part = m.group(2)
        else:
            title_part = text
            trailing_part = ''

        paragraph.clear()

        self._add_styled_runs(paragraph, title_part, cn_font, en_font, font_size, is_bold)

        if trailing_part:
            self._add_styled_runs(paragraph, trailing_part, cn_font, en_font, font_size, False)

    def _add_styled_runs(self, paragraph, text, cn_font, en_font, font_size, bold):
        if not text:
            return

        segments = self._split_by_language(text)
        for seg_text, is_en in segments:
            run = paragraph.add_run(seg_text)
            font_name = en_font if is_en else cn_font
            set_run_font(run, font_name, font_size, bold)

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
