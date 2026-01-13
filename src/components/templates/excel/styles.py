from typing import Dict, Any
from .enums import BorderStyle, NumberFormat, AlignmentType


class ExcelStyles:

    @staticmethod
    def get_border_config(config: Dict[str, Any]) -> Dict[str, Any]:
        """
        Get cell border configuration.
        """
        return {
            'style': config.get('style', BorderStyle.THIN.value),
            'color': config.get('color', '#000000'),
            'top': config.get('top', True),
            'bottom': config.get('bottom', True),
            'left': config.get('left', True),
            'right': config.get('right', True),
            'internal': config.get('internal', True)
        }

    @staticmethod
    def get_cell_format(config: Dict[str, Any]) -> Dict[str, Any]:
        return {
            'number_format': config.get('number_format', NumberFormat.GENERAL.value),
            'bg_color': config.get('bg_color', None),
            'font_color': config.get('font_color', None),
            'bold': config.get('bold', False),
            'italic': config.get('italic', False),
            'underline': config.get('underline', False),
            'alignment': config.get('alignment', AlignmentType.CENTER.value),
            'vertical_alignment': config.get('vertical_alignment', AlignmentType.CENTER.value),
            'font_name': config.get('font_name', 'Calibri'),
            'font_size': config.get('font_size', 10),
            'wrap_text': config.get('wrap_text', False),
            'border': ExcelStyles.get_border_config(config.get('border', {}))  # ✅ Now works
        }

    @staticmethod
    def get_column_config(config: Dict[str, Any]) -> Dict[str, Any]:
        return {
            'width': config.get('width', None),
            'auto_width': config.get('auto_width', True),
            'min_width': config.get('min_width', 5),
            'max_width': config.get('max_width', 50),
            'number_format': config.get('number_format', None),
            'alignment': config.get('alignment', AlignmentType.CENTER.value),
            'bg_color': config.get('bg_color', None),
            'font_color': config.get('font_color', None),
            'font_size': config.get('font_size', None),
            'font_name': config.get('font_name', None),
            'bold': config.get('bold', False),
            'italic': config.get('italic', False),
            'underline': config.get('underline', False),
            'wrap_text': config.get('wrap_text', False),
            'hidden': config.get('hidden', False),
            'frozen': config.get('frozen', False)
        }

    @staticmethod
    def get_table_style(config: Dict[str, Any]) -> Dict[str, Any]:
        return {
            'header_bg': config.get('header_bg', '#1F3A93'),
            'header_text': config.get('header_text', '#FFFFFF'),
            'header_font_size': config.get('header_font_size', 12),
            'header_font_name': config.get('header_font_name', 'Calibri'),
            'header_bold': config.get('header_bold', True),
            'header_italic': config.get('header_italic', False),
            'header_underline': config.get('header_underline', False),
            'row_colors': config.get('row_colors', ['#FFFFFF', '#F7FAFC']),
            'row_font_size': config.get('row_font_size', 10),
            'row_font_name': config.get('row_font_name', 'Calibri'),
            'row_height': config.get('row_height', 18),
            'border_color': config.get('border_color', '#CCCCCC'),
            'border_style': config.get('border_style', BorderStyle.THIN.value),
            'cell_padding': config.get('cell_padding', 8),
            'alternating_rows': config.get('alternating_rows', True),
            'freeze_header': config.get('freeze_header', True),
            'header_alignment': config.get('header_alignment', AlignmentType.CENTER.value),
            'data_alignment': config.get('data_alignment', AlignmentType.CENTER.value)
        }
