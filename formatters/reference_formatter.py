from .base_formatter import BaseFormatter
from utils.docx_helper import (
    set_paragraph_format, format_paragraph_text, match_pattern,
)


class ReferenceFormatter(BaseFormatter):
    def __init__(self, doc, config):
        super().__init__(doc, config)
        self.reference_config = config.get('reference', {})
        self.title_config = self.reference_config.get('title', {})
        self.content_config = self.reference_config.get('content', {})
        self.reference_pattern = self.recognition_config.get('reference_title_pattern', '')

    def format(self):
        in_reference = False

        for paragraph in self.doc.paragraphs:
            if match_pattern(paragraph.text, self.reference_pattern):
                self._format_title(paragraph)
                in_reference = True
                continue

            if in_reference and paragraph.text.strip():
                self._format_content(paragraph)

    def _format_title(self, paragraph):
        cfg = self.title_config

        set_paragraph_format(
            paragraph,
            alignment=cfg.get('alignment', 'center'),
            space_before=cfg.get('space_before', 12),
            space_after=cfg.get('space_after', 6),
        )

        format_paragraph_text(
            paragraph,
            cn_font=cfg.get('font_name_cn', '黑体'),
            en_font=cfg.get('font_name_en', 'Times New Roman'),
            font_size=cfg.get('font_size', 10.5),
            bold=cfg.get('bold', True),
        )

    def _format_content(self, paragraph):
        cfg = self.content_config

        set_paragraph_format(
            paragraph,
            alignment=cfg.get('alignment', 'left'),
            line_spacing=cfg.get('line_spacing', 1.5),
        )

        format_paragraph_text(
            paragraph,
            cn_font=cfg.get('font_name_cn', '宋体'),
            en_font=cfg.get('font_name_en', 'Times New Roman'),
            font_size=cfg.get('font_size', 10.5),
            bold=cfg.get('bold', False),
        )
