from .base_formatter import BaseFormatter
from utils.docx_helper import (
    set_paragraph_format, format_paragraph_text, match_pattern,
    set_run_font, split_text_by_language,
)
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re


class TOCFormatter(BaseFormatter):
    def __init__(self, doc, config):
        super().__init__(doc, config)
        self.toc_config = config.get('toc', {})
        self.title_config = self.toc_config.get('title', {})
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
                )
                format_paragraph_text(
                    paragraph,
                    cn_font=title_cfg.get('font_name_cn', '黑体'),
                    en_font=title_cfg.get('font_name_en', 'Times New Roman'),
                    font_size=title_cfg.get('font_size', 16),
                    bold=title_cfg.get('bold', True),
                )
                break

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

            level = self._detect_toc_level(paragraph, style_name)
            if level:
                self._format_entry(paragraph, level)

    def _is_toc_title(self, paragraph):
        text = paragraph.text.strip()
        style_name = paragraph.style.name if paragraph.style else ''

        if 'TOC' in style_name or 'toc' in style_name.lower():
            return True

        if re.match(r'^目录$', text):
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

    def _detect_toc_level(self, paragraph, style_name):
        text = paragraph.text.strip()

        if re.match(r'^[0-9]+\.[0-9]+\.[0-9]+', text):
            return 'level3'
        elif re.match(r'^[0-9]+\.[0-9]+', text):
            return 'level2'
        elif re.match(r'^[0-9]+', text):
            return 'level1'

        if '1' in style_name:
            return 'level1'
        elif '2' in style_name:
            return 'level2'
        elif '3' in style_name:
            return 'level3'

        pPr = paragraph._element.find(qn('w:pPr'))
        if pPr is not None:
            numPr = pPr.find(qn('w:numPr'))
            if numPr is not None:
                ilvl = numPr.find(qn('w:ilvl'))
                if ilvl is not None:
                    val = ilvl.get(qn('w:val'))
                    if val == '0':
                        return 'level1'
                    elif val == '1':
                        return 'level2'
                    elif val == '2':
                        return 'level3'

        indent = paragraph.paragraph_format.left_indent
        if indent is not None:
            indent_pt = indent.pt
            if indent_pt >= 36:
                return 'level3'
            elif indent_pt >= 18:
                return 'level2'

        return 'level1'

    def _format_entry(self, paragraph, level):
        cfg = self.toc_config.get(level, {})

        indent_chars = cfg.get('indent', 0)
        indent_pt = indent_chars * 12

        paragraph.paragraph_format.left_indent = indent_pt if indent_pt > 0 else None

        set_paragraph_format(
            paragraph,
            alignment='left',
            line_spacing=self.toc_config.get('line_spacing', 1.5),
        )

        format_paragraph_text(
            paragraph,
            cn_font=cfg.get('font_name_cn', '宋体'),
            en_font=cfg.get('font_name_en', 'Times New Roman'),
            font_size=cfg.get('font_size', 12),
            bold=cfg.get('bold', False),
        )

    def update_toc(self):
        """通过COM接口更新目录域代码（需要win32com，仅Windows可用）"""
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
