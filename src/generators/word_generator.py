from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml.shared import OxmlElement
import json
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, List
from dataclasses import dataclass, field

OUTPUT_DIR = Path("./generated_documents")


@dataclass
class WordStyleConfig:
    """Parsed Word document styling configuration."""
    document: Dict[str, Any] = field(default_factory=dict)
    header: Dict[str, Any] = field(default_factory=dict)
    title: Dict[str, Any] = field(default_factory=dict)
    timestamp: Dict[str, Any] = field(default_factory=dict)
    paragraph: Dict[str, Any] = field(default_factory=dict)
    headings: Dict[str, Any] = field(default_factory=dict)
    table: Dict[str, Any] = field(default_factory=dict)
    lists: Dict[str, Any] = field(default_factory=dict)


def parse_style_config(styling_config: str) -> WordStyleConfig:
    """Parse JSON styling config."""
    if not styling_config or styling_config == "{}":
        return WordStyleConfig()

    config = json.loads(styling_config)
    return WordStyleConfig(**config)


def set_cell_shading(cell, color_hex: str):
    """Apply background color to the table cell."""
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), color_hex)
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run._element.get_or_add_tcPr().append(shading_elm)
    cell._tc.tcPr.append(shading_elm)


def generate_word_report(
        title: str,
        subtitle: str,
        content: str,  # JSON sections
        styling_config: str = "{}",  # Full styling JSON
        include_table_of_contents: bool = False,
        page_margins: str = "{}"  # Optional margins JSON
) -> str:
    """
    Generate a fully customizable Word document (.docx).

    content JSON example:
    {
        "sections": [
            {"type": "heading1", "text": "Introduction"},
            {"type": "paragraph", "text": "Content..."},
            {"type": "table", "table_data": [["H1","H2"],["D1","D2"]]},
            {"type": "bullet", "text": "• Item 1"},
            {"type": "numbered", "text": "1. Item 1"}
        ]
    }

    styling_config example:
    {
        "document": {
            "margins": {"top": 2.0, "bottom": 2.0, "left": 2.5, "right": 2.5},
            "default_font": "Calibri",
            "default_font_size": 11
        },
        "header": {
            "font_size": 16,
            "color": [31, 58, 147],
            "bold": true
        },
        "title": {
            "font_size": 24,
            "color": [44, 82, 130],
            "bold": true
        },
        "paragraph": {
            "font_size": 11,
            "line_spacing": 1.15,
            "space_after": 6
        },
        "table": {
            "style": "Light Grid Accent 1",
            "header_shading": "2C5282",
            "header_font_size": 11,
            "row_shading_alternate": "F3F4F6"
        }
    }
    """
    try:
        content_data = json.loads(content)
        style_config = parse_style_config(styling_config)
        margins = json.loads(page_margins) if page_margins else {}

        # Create the document
        doc = Document()

        # Page margins
        doc_cfg = style_config.document
        sections = doc.sections
        for section in sections:
            margin_cfg = doc_cfg.get('margins', margins)
            section.top_margin = Cm(margin_cfg.get('top', 2.0))
            section.bottom_margin = Cm(margin_cfg.get('bottom', 2.0))
            section.left_margin = Cm(margin_cfg.get('left', 2.5))
            section.right_margin = Cm(margin_cfg.get('right', 2.5))

        # 1. COMPANY HEADER
        header_cfg = style_config.header
        header_p = doc.add_paragraph()
        header_run = header_p.add_run(subtitle)
        header_run.font.size = Pt(header_cfg.get('font_size', 16))
        header_run.bold = header_cfg.get('bold', True)
        color = header_cfg.get('color', [31, 58, 147])
        header_run.font.color.rgb = RGBColor(*color)
        header_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # 2. DOCUMENT TITLE
        title_cfg = style_config.title
        title_p = doc.add_paragraph()
        title_run = title_p.add_run(title)
        title_run.font.size = Pt(title_cfg.get('font_size', 24))
        title_run.bold = title_cfg.get('bold', True)
        color = title_cfg.get('color', [44, 82, 130])
        title_run.font.color.rgb = RGBColor(*color)
        title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # 3. GENERATION TIMESTAMP
        timestamp_cfg = style_config.timestamp
        timestamp_p = doc.add_paragraph()
        timestamp_run = timestamp_p.add_run(
            f"Generated: {datetime.now().strftime('%B %d, %Y %H:%M:%S')}"
        )
        timestamp_run.font.size = Pt(timestamp_cfg.get('font_size', 10))
        timestamp_run.italic = timestamp_cfg.get('italic', True)
        timestamp_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        doc.add_paragraph()  # Spacer

        # 4. TABLE OF CONTENTS
        if include_table_of_contents:
            toc_heading = doc.add_heading("TABLE OF CONTENTS", level=1)
            toc_heading.runs[0].font.size = Pt(14)
            toc_heading.runs[0].bold = True
            doc.add_paragraph()
            # Note: Full TOC requires docx limitations workaround or post-processing

        # 5. CONTENT SECTIONS
        para_cfg = style_config.paragraph
        default_font = doc_cfg.get('default_font', 'Calibri')
        default_size = Pt(doc_cfg.get('default_font_size', 11))

        for section in content_data.get('sections', []):
            section_type = section.get('type', 'paragraph')
            text = section.get('text', '')

            if section_type == 'heading1':
                p = doc.add_heading(text, level=1)
                for run in p.runs:
                    run.font.name = default_font
                    run.font.size = Pt(style_config.headings.get('h1_size', 16))

            elif section_type == 'heading2':
                p = doc.add_heading(text, level=2)
                for run in p.runs:
                    run.font.size = Pt(style_config.headings.get('h2_size', 14))

            elif section_type == 'heading3':
                p = doc.add_heading(text, level=3)
                for run in p.runs:
                    run.font.size = Pt(style_config.headings.get('h3_size', 12))

            elif section_type == 'paragraph':
                p = doc.add_paragraph(text)
                for run in p.runs:
                    run.font.name = default_font
                    run.font.size = para_cfg.get('font_size', default_size.pt)
                p.paragraph_format.space_after = Pt(para_cfg.get('space_after', 6))
                p.paragraph_format.line_spacing = para_cfg.get('line_spacing', 1.15)

            elif section_type == 'bullet':
                p = doc.add_paragraph(text, style='List Bullet')
                for run in p.runs:
                    run.font.name = default_font
                    run.font.size = Pt(style_config.lists.get('bullet_size', 11))

            elif section_type == 'numbered':
                p = doc.add_paragraph(text, style='List Number')
                for run in p.runs:
                    run.font.name = default_font
                    run.font.size = Pt(style_config.lists.get('numbered_size', 11))

            elif section_type == 'table':
                if 'table_data' in section:
                    table_data = section['table_data']
                    if not table_data or not table_data[0]:
                        continue

                    num_rows, num_cols = len(table_data), len(table_data[0])
                    table = doc.add_table(rows=num_rows, cols=num_cols)
                    table.style = style_config.table.get('style', 'Light Grid Accent 1')

                    table.autofit = True

                    # Apply table styling
                    table_cfg = style_config.table

                    for row_idx, row_data in enumerate(table_data):
                        for col_idx, cell_value in enumerate(row_data):
                            cell = table.rows[row_idx].cells[col_idx]
                            cell.text = str(cell_value)

                            # Header row special styling
                            if row_idx == 0:
                                for paragraph in cell.paragraphs:
                                    for run in paragraph.runs:
                                        run.font.bold = True
                                        run.font.size = Pt(table_cfg.get('header_font_size', 11))
                                        run.font.color.rgb = RGBColor(255, 255, 255)

                                # Header shading
                                if table_cfg.get('header_shading'):
                                    set_cell_shading(cell, table_cfg['header_shading'])

                            # Alternating row shading
                            else:
                                if row_idx % 2 == 1 and table_cfg.get('row_shading_alternate'):
                                    set_cell_shading(cell, table_cfg['row_shading_alternate'])

            # Section spacing
            if section.get('add_spacing', True):
                doc.add_paragraph()

        # Save document
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{title.replace(' ', '_')}_{timestamp}.docx"
        filepath = OUTPUT_DIR / filename
        doc.save(str(filepath))

        return f"✓ Word document generated successfully: {filepath}"

    except json.JSONDecodeError as e:
        return f"✗ Error: Invalid JSON content/config - {str(e)}"
    except Exception as e:
        return f"✗ Error generating Word document: {str(e)}"
