from typing import Dict, Any
from datetime import datetime

class DocumentComponentsBase:

    @staticmethod
    def render_header(config: Dict[str, Any]) -> Dict[str, Any]:
        if not config.get('enabled', True):
            return {'enabled': False}

        return {
            'text': config.get('text', 'Header'),
            'font_size': config.get('font_size', 14),
            'color': config.get('color', '#000000'),
            'bold': config.get('bold', True),
            'alignment': config.get('alignment', 'center'),

            'logo_path': config.get('logo_path'),
            'logo_width': config.get('logo_width', 1.2),
            'logo_height': config.get('logo_height'),
            'logo_alignment': config.get('logo_alignment', 'left'),
            'logo_space_after': config.get('logo_space_after', 10),
        }


    @staticmethod
    def render_body_content(config: Dict[str, Any]) -> Dict[str, Any]:
        if not config.get('enabled', True):
            return {'enabled': False}

        return {
            'title': config.get('title', 'Body Content'),
            'content_type': config.get('content_type', 'text'),
            'text': config.get('text', ''),  # For text-only PDFs
            'font_size': config.get('font_size', 12),
            'color': config.get('color', '#000000'),
            'bold': config.get('bold', False),
            'italic': config.get('italic', False),
            'alignment': config.get('alignment', 'left'),
            'structured_sections': config.get('structured_sections', []),
            'space_after': config.get('space_after', 12)
        }


    @staticmethod
    def render_footer(config: Dict[str, Any]) -> Dict[str, Any]:
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M')
        if not config.get('enabled', True):
            return {'enabled': False}
        return {
            'text': config['text'].format(timestamp=timestamp),
            'font_size': config.get('font_size', 9),
            'color': config.get('color', '#666666'),
            'alignment': config.get('alignment', 'center')
        }


    @staticmethod
    def render_signature(config: Dict[str, Any]) -> Dict[str, Any]:
        if not config.get('enabled', True):
            return {'enabled': False}
        return {
            'name': config.get('name', 'John Doe'),
            'title': config.get('title', 'Director'),
            'email': config.get('email', ''),
            'phone': config.get('phone', ''),
            'font_size': config.get('font_size', 10)
        }


    @staticmethod
    def get_table_style(config: Dict[str, Any]) -> Dict[str, Any]:
        return {
            'header_bg': config.get('header_bg', '#2C5282'),
            'header_text': config.get('header_text', '#FFFFFF'),
            'row_colors': config.get('row_colors', ['#FFFFFF', '#F7FAFC']),
            'border_color': config.get('border_color', '#CCCCCC'),
            'header_font_size': config.get('header_font_size', 11),
            'row_font_size': config.get('row_font_size', 10)
        }

# PDF
class PDFComponents(DocumentComponentsBase):
    pass


# Excel
class ExcelComponents(DocumentComponentsBase):
    pass


# Word
class WordComponents(DocumentComponentsBase):
    pass