from .base_formatter import BaseFormatter
from utils.docx_helper import (
    set_paragraph_format, format_paragraph_text, match_pattern,
)


class CaptionFormatter(BaseFormatter):
    def __init__(self, doc, config):
        super().__init__(doc, config)
        self.figure_config = config.get('figure', {})
        self.caption_config = self.figure_config.get('caption', {})
        self.figure_pattern = self.recognition_config.get('figure_caption_pattern', '')

    def format(self):
        for paragraph in self.doc.paragraphs:
            if match_pattern(paragraph.text, self.figure_pattern):
                self._format_caption(paragraph)

    def _format_caption(self, paragraph):
        cfg = self.caption_config

        set_paragraph_format(
            paragraph,
            alignment=cfg.get('alignment', 'center'),
            space_before=cfg.get('space_before', 6),
            space_after=cfg.get('space_after', 6),
        )

        format_paragraph_text(
            paragraph,
            cn_font=cfg.get('font_name_cn', '宋体'),
            en_font=cfg.get('font_name_en', 'Times New Roman'),
            font_size=cfg.get('font_size', 12),
            bold=cfg.get('bold', True),
        )
