from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml.shared import OxmlElement
import json
from pathlib import Path
from datetime import datetime
from typing import Dict, Any
from dataclasses import dataclass, field

OUTPUT_DIR = Path("./generated_documents")
OUTPUT_DIR.mkdir(exist_ok=True)


@dataclass
class WordStyleConfig:
    document: Dict[str, Any] = field(default_factory=dict)
    header: Dict[str, Any] = field(default_factory=dict)
    title: Dict[str, Any] = field(default_factory=dict)
    paragraph: Dict[str, Any] = field(default_factory=dict)
    headings: Dict[str, Any] = field(default_factory=dict)
    table: Dict[str, Any] = field(default_factory=dict)
    lists: Dict[str, Any] = field(default_factory=dict)


def parse_style_config(styling_config: str) -> WordStyleConfig:
    if not styling_config or styling_config == "{}":
        return WordStyleConfig()
    return WordStyleConfig(**json.loads(styling_config))


# Style helper functions
def ensure_paragraph_style(doc, name, base="Normal", size=11, bold=False):
    if name in doc.styles:
        return
    style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    style.base_style = doc.styles[base]
    style.font.size = Pt(size)
    style.font.bold = bold


def ensure_character_style(doc, name, base=None, bold=False, italic=False):
    if name in doc.styles:
        return
    style = doc.styles.add_style(name, WD_STYLE_TYPE.CHARACTER)
    if base:
        style.base_style = doc.styles[base]
    style.font.bold = bold
    style.font.italic = italic


def initialize_styles(doc):
    ensure_paragraph_style(doc, "Title", size=18, bold=True)
    ensure_paragraph_style(doc, "Body Text", size=11)
    ensure_paragraph_style(doc, "Caption", size=9)

    heading_sizes = {
        1: 16,
        2: 14,
        3: 12,
        4: 11,
        5: 10,
        6: 9,
    }

    for level, size in heading_sizes.items():
        style = doc.styles[f"Heading {level}"]
        style.font.size = Pt(size)
        style.font.bold = True

        style.paragraph_format.space_before = Pt(12 if level == 1 else 8)
        style.paragraph_format.space_after = Pt(6)

    # Character styles
    ensure_character_style(doc, "Emphasis", italic=True)
    ensure_character_style(doc, "Strong", bold=True)


# Header and footer
def add_header_footer(doc, subtitle: str):
    section = doc.sections[0]
    section.different_first_page_header_footer = True

    header = section.header
    header.is_linked_to_previous = False
    p = header.paragraphs[0]
    p.text = subtitle
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    footer = section.footer
    p = footer.paragraphs[0]
    run = p.add_run()

    fld_begin = OxmlElement('w:fldChar')
    fld_begin.set(qn('w:fldCharType'), 'begin')

    instr = OxmlElement('w:instrText')
    instr.text = "PAGE"

    fld_end = OxmlElement('w:fldChar')
    fld_end.set(qn('w:fldCharType'), 'end')

    run._r.extend([fld_begin, instr, fld_end])
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER


# Table helper functions
def set_cell_shading(cell, color_hex):
    """Apply background color to a table cell."""
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), color_hex)
    cell._tc.get_or_add_tcPr().append(shd)


def repeat_table_header(row):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    header = OxmlElement('w:tblHeader')
    trPr.append(header)


# Internal document generator
def _generate_word_report_internal(
        title: str,
        subtitle: str,
        content: str,
        styling_config: str = "{}",
        page_margins: str = "{}"
) -> str:
    """
    Internal Word document generation implementation.
    """
    try:
        content_data = json.loads(content)
        style_config = parse_style_config(styling_config)
        margins = json.loads(page_margins) if page_margins else {}

        # Create the document
        doc = Document()
        # Apply styling
        initialize_styles(doc)

        # Page margins
        for section in doc.sections:
            cfg = style_config.document.get("margins", margins)
            section.top_margin = Cm(cfg.get("top", 2))
            section.bottom_margin = Cm(cfg.get("bottom", 2))
            section.left_margin = Cm(cfg.get("left", 2.5))
            section.right_margin = Cm(cfg.get("right", 2.5))

        # HEADER AND FOOTER
        add_header_footer(doc, subtitle)

        # DOCUMENT TITLE
        doc.add_paragraph(title, style="Title").alignment = WD_ALIGN_PARAGRAPH.CENTER

        # CONTENT SECTIONS
        for section in content_data.get("sections", []):
            t = section.get("type", "paragraph")

            if t.startswith("heading"):
                level = min(int(t.replace("heading", "")), 6)
                doc.add_paragraph(section["text"], style=f"Heading {level}")

            elif t == "paragraph":
                doc.add_paragraph(section["text"], style="Body Text")

            elif t == "table":
                data = section["table_data"]
                rows, cols = len(data), len(data[0])
                table = doc.add_table(rows=rows, cols=cols)
                table.style = style_config.table.get("style", "Light Grid Accent 1")
                table.autofit = True

                repeat_table_header(table.rows[0])

                for r, row in enumerate(data):
                    for c, value in enumerate(row):
                        cell = table.rows[r].cells[c]
                        cell.text = str(value)

                        if r == 0:
                            set_cell_shading(cell, style_config.table.get("header_shading", "2C5282"))

                if section.get("merge"):
                    for m in section["merge"]:
                        table.cell(*m[:2]).merge(table.cell(*m[2:]))

            elif t == "caption":
                doc.add_paragraph(section["text"], style="Caption")

            elif t == "page_break":
                doc.add_page_break()

            doc.add_paragraph()

        # Save document
        filename = f"{title.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        path = OUTPUT_DIR / filename
        doc.save(path)

        return str(path)

    except Exception as e:
        raise ValueError(f"Word generation failed: {str(e)}")


def generate_word_report(
        template_path: str = None,
        output_path: str = None,
        data: Dict[str, Any] = None,
        **kwargs
) -> str:
    """
    Wrapper for Word document generation.
    """
    if not data:
        raise ValueError("Missing 'data' parameter with Word document configuration")

    try:
        # Extract parameters from data dict
        title = data.get("title", "Document")
        subtitle = data.get("subtitle", "")
        content = data.get("content", "{}")
        styling_config = data.get("styling_config", "{}")
        page_margins = data.get("page_margins", "{}")

        # Convert dict inputs to JSON strings (agent may send nested objects)
        if isinstance(content, dict):
            content = json.dumps(content)
        if isinstance(styling_config, dict):
            styling_config = json.dumps(styling_config)
        if isinstance(page_margins, dict):
            page_margins = json.dumps(page_margins)

        # Call internal generator
        filepath = _generate_word_report_internal(
            title=title,
            subtitle=subtitle,
            content=content,
            styling_config=styling_config,
            page_margins=page_margins
        )

        try:
            content_data = json.loads(content) if isinstance(content, str) else content
            section_count = len(content_data.get("sections", []))
        except:
            section_count = 0

        return f"Word document generated: {filepath} | Sections: {section_count} | Title: {title}"

    except Exception as e:
        raise RuntimeError(f"Word generation failed: {str(e)}")