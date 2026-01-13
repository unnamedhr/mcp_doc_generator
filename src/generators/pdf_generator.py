from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from datetime import datetime
import json
from pathlib import Path
from typing import Dict, Any, List, Optional

OUTPUT_DIR = Path("./generated_documents")


def parse_components_config(components_json: str) -> Dict[str, Dict[str, Any]]:

    defaults = {
        "header": {"enabled": True, "text": "Professional Report", "font_size": 14, "color": "#1F3A93"},
        "footer": {"enabled": True, "text": "Confidential | {timestamp}", "font_size": 9, "color": "#666666"},
        "table_style": {"header_bg": "#2C5282", "header_text": "#FFFFFF", "row_colors": ["#FFFFFF", "#F7FAFC"]},
        "signature": {"enabled": False, "name": "John Doe", "title": "Director"}
    }

    if not components_json or components_json.strip() == "{}":
        return defaults

    try:
        config_dict = json.loads(components_json)
        merged = {}
        for key in defaults:
            merged[key] = {**defaults[key], **config_dict.get(key, {})}
        return merged
    except json.JSONDecodeError:
        return defaults


def _apply_component_header(components: Dict[str, Dict[str, Any]], config: Dict[str, Any],
                            styles, elements: List) -> None:
    comp_header = components.get('header', {})
    if not comp_header.get('enabled', True):
        return

    header_style = ParagraphStyle(
        'ComponentHeader',
        parent=styles['Heading1'],
        fontSize=comp_header.get('font_size', 14),
        textColor=colors.HexColor(comp_header.get('color', '#1F3A93')),
        spaceAfter=comp_header.get('space_after', 6),
        alignment=comp_header.get('alignment', 1)
    )
    elements.append(Paragraph(comp_header['text'], header_style))
    elements.append(Spacer(1, 0.15 * inch))


def _apply_component_footer(components: Dict[str, Dict[str, Any]], styles, elements: List) -> None:
    comp_footer = components.get('footer', {})
    if not comp_footer.get('enabled', True):
        return

    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M')
    footer_text = comp_footer['text'].format(timestamp=timestamp)

    footer_style = ParagraphStyle(
        'ComponentFooter',
        parent=styles['Normal'],
        fontSize=comp_footer.get('font_size', 9),
        textColor=colors.HexColor(comp_footer.get('color', '#666666')),
        alignment=comp_footer.get('alignment', 1),
        spaceBefore=12
    )
    elements.append(Spacer(1, 0.2 * inch))
    elements.append(Paragraph(footer_text, footer_style))


def _apply_component_signature(components: Dict[str, Dict[str, Any]], styles, elements: List) -> None:
    comp_sig = components.get('signature', {})
    if not comp_sig.get('enabled', True):
        return

    sig_parts = [
        comp_sig.get('name', ''),
        comp_sig.get('title', ''),
        comp_sig.get('email', ''),
        comp_sig.get('phone', '')
    ]
    sig_text = " | ".join([p for p in sig_parts if p])

    sig_style = ParagraphStyle(
        'ComponentSignature',
        parent=styles['Normal'],
        fontSize=comp_sig.get('font_size', 10),
        italic=True,
        spaceAfter=30,
        alignment=2
    )
    elements.append(Spacer(1, 0.3 * inch))
    elements.append(Paragraph(sig_text, sig_style))


def _get_table_style_from_components(components: Dict[str, Dict[str, Any]],
                                     config: Dict[str, Any]) -> TableStyle:
    table_cfg = config.get('table', {})
    comp_table = components.get('table_style', {})

    header_bg = table_cfg.get('header_bg', comp_table.get('header_bg', '#2C5282'))
    header_text = table_cfg.get('header_text_color', comp_table.get('header_text', '#FFFFFF'))
    row_colors_cfg = table_cfg.get('row_colors', comp_table.get('row_colors', ['#FFFFFF', '#F7FAFC']))
    border_color = table_cfg.get('grid_color', comp_table.get('border_color', '#CCCCCC'))

    return TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(header_bg)),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor(header_text)),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), table_cfg.get('header_font_size', 12)),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), table_cfg.get('font_size', 10)),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor(c) for c in row_colors_cfg]),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor(border_color)),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ])


def generate_pdf_report(
        title: str,
        body: str,
        report_data: str,
        styling_config: str = "{}",
        components_config: str = "{}",
        include_header: bool = True
) -> str:
    """
    Generate a PDF with styling and reusable components (header, footer, signature, table_style).
    """
    try:
        # Parse inputs
        data = json.loads(report_data)
        config: Dict[str, Any] = json.loads(styling_config) if styling_config else {}
        components = parse_components_config(components_config)  # SELF-CONTAINED

        # Document setup
        doc_cfg = config.get('document', {})
        margin = doc_cfg.get('margin', 0.75) * inch
        page_sizes = {
            'letter': letter,
            'A4': letter,
            'legal': letter
        }
        page_size = page_sizes.get(doc_cfg.get('page_size', 'letter'), letter)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{title.replace(' ', '_')}_{timestamp}.pdf"
        filepath = OUTPUT_DIR / filename

        doc = SimpleDocTemplate(
            str(filepath), pagesize=page_size,
            rightMargin=margin, leftMargin=margin,
            topMargin=margin, bottomMargin=margin
        )

        styles = getSampleStyleSheet()
        elements: List[Any] = []

        # HEADER
        _apply_component_header(components, config, styles, elements)

        # Content (title, timestamp, sections/tables)
        if include_header:
            title_cfg = config.get('title', {})
            title_style = ParagraphStyle(
                'CustomHeaderTitle', parent=styles['Heading1'],
                fontSize=title_cfg.get('font_size', 24),
                textColor=colors.HexColor(title_cfg.get('color', '#1F3A93')),
                spaceAfter=title_cfg.get('space_after', 6),
                alignment=title_cfg.get('alignment', 1)
            )
            elements.append(Paragraph(body, title_style))
            elements.append(Spacer(1, 0.2 * inch))

        # Report title
        report_title_cfg = config.get('report_title', {})
        report_title_style = ParagraphStyle(
            'CustomReportTitle', parent=styles['Heading2'],
            fontSize=report_title_cfg.get('font_size', 18),
            textColor=colors.HexColor(report_title_cfg.get('color', '#2C5282')),
            spaceAfter=12
        )
        elements.append(Paragraph(title, report_title_style))
        elements.append(Spacer(1, 0.2 * inch))

        # Timestamp
        timestamp_cfg = config.get('timestamp', {})
        timestamp_style = ParagraphStyle(
            'CustomTimestamp', parent=styles['Normal'],
            fontSize=timestamp_cfg.get('font_size', 10),
            italic=timestamp_cfg.get('italic', True)
        )
        timestamp_text = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        elements.append(Paragraph(timestamp_text, timestamp_style))
        elements.append(Spacer(1, 0.3 * inch))

        # Sections & Tables
        if 'sections' in data:
            for section in data['sections']:
                section_cfg = config.get('section', {})
                section_style = ParagraphStyle(
                    'CustomSection', parent=styles['Heading3'],
                    fontSize=section_cfg.get('font_size', 14),
                    textColor=colors.HexColor(section_cfg.get('color', '#2D3748')),
                    spaceAfter=section_cfg.get('space_after', 12)
                )
                elements.append(Paragraph(section.get('title', 'Section'), section_style))

                if 'table_data' in section:
                    table_data = section['table_data']
                    if table_data:
                        num_cols = len(table_data[0])
                        col_widths = [2 * inch] * num_cols
                        table_cfg = config.get('table', {})
                        if 'col_widths' in table_cfg:
                            col_widths = [w * inch for w in table_cfg['col_widths']]

                        table = Table(table_data, colWidths=col_widths)
                        table_style = _get_table_style_from_components(components, config)
                        table.setStyle(table_style)
                        elements.append(table)
                        elements.append(Spacer(1, 0.2 * inch))

        # SIGNATURE
        _apply_component_signature(components, styles, elements)

        # FOOTER
        _apply_component_footer(components, styles, elements)

        doc.build(elements)
        return f"✓ PDF with components: {filepath}"

    except json.JSONDecodeError as e:
        return f"✗ Invalid JSON: {str(e)}"
    except Exception as e:
        return f"✗ PDF Error: {str(e)}"