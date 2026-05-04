from .base_formatter import BaseFormatter
from utils.docx_helper import (
    set_paragraph_format, format_paragraph_text, detect_heading_level,
)


class HeadingFormatter(BaseFormatter):
    def __init__(self, doc, config):
        super().__init__(doc, config)
        self.heading_config = config.get('headings', {})
        self.patterns = self.recognition_config.get('heading_patterns', {})

    def format(self):
        for paragraph in self.doc.paragraphs:
            level = detect_heading_level(paragraph, self.patterns)
            if level:
                self._format_heading(paragraph, level)

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
