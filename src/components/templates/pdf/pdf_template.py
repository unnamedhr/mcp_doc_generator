from typing import Dict, Any, List
from datetime import datetime

class DocumentComponentsBase:
    @staticmethod
    def render_header(config: Dict[str, Any]) -> Dict[str, Any]:
        if not config.get('enabled', True):
            return {'enabled': False}
        return {
            'text': config.get('text', 'Default Header'),
            'font_size': config.get('font_size', 14),
            'color': config.get('color', '#1F3A93'),
            'bold': config.get('bold', True),
            'alignment': config.get('alignment', 'center')
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

class PDFComponents(DocumentComponentsBase):
    pass


# Excel-specific
class ExcelComponents(DocumentComponentsBase):
    pass


# Word-specific
class WordComponents(DocumentComponentsBase):
    pass