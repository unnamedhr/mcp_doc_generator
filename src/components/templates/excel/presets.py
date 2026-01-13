from typing import Dict, Any

class ExcelPresets:

    @staticmethod
    def minimal_theme() -> Dict[str, Any]:
        return {
            'header': {
                'enabled': True,
                'bg_color': '#E5E5E5',
                'color': '#000000',
                'font_size': 11,
                'font_name': 'Calibri',
                'bold': True
            },
            'data': {
                'enabled': True,
                'alternating_colors': True,
                'alternate_color': '#F8F9FA'
            },
            'table_style': {
                'header_bg': '#E5E5E5',
                'header_text': '#000000',
                'row_colors': ['#FFFFFF', '#F8F9FA'],
                'border_style': 'thin',
                'freeze_header': True
            }
        }

    @staticmethod
    def blue_theme() -> Dict[str, Any]:
        return {
            'header': {
                'enabled': True,
                'bg_color': '#1F3A93',
                'color': '#FFFFFF',
                'font_size': 12,
                'font_name': 'Calibri',
                'bold': True
            },
            'data': {
                'enabled': True,
                'alternating_colors': True,
                'alternate_color': '#F0F7FF'
            },
            'table_style': {
                'header_bg': '#1F3A93',
                'header_text': '#FFFFFF',
                'row_colors': ['#FFFFFF', '#F0F7FF'],
                'border_style': 'thin',
                'freeze_header': True
            }
        }

    @staticmethod
    def green_theme() -> Dict[str, Any]:
        return {
            'header': {
                'enabled': True,
                'bg_color': '#38A169',
                'color': '#FFFFFF',
                'font_size': 12,
                'font_name': 'Calibri',
                'bold': True
            },
            'data': {
                'enabled': True,
                'alternating_colors': True,
                'alternate_color': '#F0FFF4'
            },
            'table_style': {
                'header_bg': '#38A169',
                'header_text': '#FFFFFF',
                'row_colors': ['#FFFFFF', '#F0FFF4'],
                'border_style': 'thin',
                'freeze_header': True
            }
        }

    @staticmethod
    def red_theme() -> Dict[str, Any]:
        return {
            'header': {
                'enabled': True,
                'bg_color': '#C53030',
                'color': '#FFFFFF',
                'font_size': 12,
                'font_name': 'Calibri',
                'bold': True
            },
            'data': {
                'enabled': True,
                'alternating_colors': True,
                'alternate_color': '#FFF5F5'
            },
            'table_style': {
                'header_bg': '#C53030',
                'header_text': '#FFFFFF',
                'row_colors': ['#FFFFFF', '#FFF5F5'],
                'border_style': 'thin',
                'freeze_header': True
            }
        }
