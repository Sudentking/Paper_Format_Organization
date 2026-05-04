from .base_formatter import BaseFormatter
from .toc_formatter import is_toc_entry, is_toc_style
from utils.docx_helper import (
    set_paragraph_format, format_paragraph_text, detect_heading_level, match_pattern,
)


class ParagraphFormatter(BaseFormatter):
    def __init__(self, doc, config):
        super().__init__(doc, config)
        self.paragraph_config = config.get('paragraph', {})
        self.heading_patterns = self.recognition_config.get('heading_patterns', {})
        self.table_pattern = self.recognition_config.get('table_caption_pattern', '')
        self.figure_pattern = self.recognition_config.get('figure_caption_pattern', '')
        self.reference_pattern = self.recognition_config.get('reference_title_pattern', '')

    def format(self):
        for paragraph in self.doc.paragraphs:
            if self._should_skip(paragraph):
                continue
            self._format_paragraph(paragraph)

    def _should_skip(self, paragraph):
        text = paragraph.text
        if not text.strip():
            return True

        style_name = paragraph.style.name if paragraph.style else ''
        if is_toc_entry(text) or is_toc_style(style_name):
            return True

        if detect_heading_level(paragraph, self.heading_patterns):
            return True

        if match_pattern(text, self.table_pattern):
            return True
        if match_pattern(text, self.figure_pattern):
            return True
        if match_pattern(text, self.reference_pattern):
            return True

        return False

    def _format_paragraph(self, paragraph):
        cfg = self.paragraph_config

        set_paragraph_format(
            paragraph,
            alignment=cfg.get('alignment', 'justify'),
            line_spacing=cfg.get('line_spacing', 1.5),
            space_before=cfg.get('space_before', 0),
            space_after=cfg.get('space_after', 0),
            first_line_indent=cfg.get('first_line_indent', 0),
        )

        format_paragraph_text(
            paragraph,
            cn_font=cfg.get('font_name_cn', '宋体'),
            en_font=cfg.get('font_name_en', 'Times New Roman'),
            font_size=cfg.get('font_size', 12),
            bold=cfg.get('bold', False),
        )
