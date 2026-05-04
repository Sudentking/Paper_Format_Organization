import re
from .base_formatter import BaseFormatter
from .toc_formatter import is_toc_entry, is_toc_style
from utils.docx_helper import (
    set_paragraph_format, format_paragraph_text, detect_heading_level,
)
from docx.oxml.ns import qn

SPECIAL_HEADING_PATTERNS = [
    re.compile(r'^摘\s*要$'),
    re.compile(r'^Abstract$', re.IGNORECASE),
    re.compile(r'^关键词[：:]?$'),
    re.compile(r'^Keywords[：:]?$', re.IGNORECASE),
]


class HeadingFormatter(BaseFormatter):
    def __init__(self, doc, config):
        super().__init__(doc, config)
        self.heading_config = config.get('headings', {})
        self.patterns = self.recognition_config.get('heading_patterns', {})

    def format(self):
        for paragraph in self.doc.paragraphs:
            style_name = paragraph.style.name if paragraph.style else ''
            if is_toc_entry(paragraph.text) or is_toc_style(style_name):
                continue

            special_level = self._detect_special_heading(paragraph)
            if special_level:
                self._format_heading(paragraph, special_level)
                continue

            level = detect_heading_level(paragraph, self.patterns)
            if level:
                self._format_heading(paragraph, level)

    def _detect_special_heading(self, paragraph):
        text = paragraph.text.strip()
        if not text:
            return None

        pPr = paragraph._element.find(qn('w:pPr'))
        if pPr is not None:
            ol = pPr.find(qn('w:outlineLvl'))
            if ol is not None:
                return None

        for pattern in SPECIAL_HEADING_PATTERNS:
            if pattern.match(text):
                return 'level1'

        return None

    def _format_heading(self, paragraph, level):
        cfg = self.heading_config.get(level, {})

        set_paragraph_format(
            paragraph,
            alignment=cfg.get('alignment', 'left'),
            line_spacing=cfg.get('line_spacing', 1.0),
            space_before=cfg.get('space_before', 0),
            space_after=cfg.get('space_after', 0),
        )

        format_paragraph_text(
            paragraph,
            cn_font=cfg.get('font_name_cn', '宋体'),
            en_font=cfg.get('font_name_en', 'Times New Roman'),
            font_size=cfg.get('font_size', 12),
            bold=cfg.get('bold', False),
        )
