from __future__ import annotations

import base64
import json
import re
import io
from typing import Union, Literal, Dict, Any, List, Optional, TypedDict
from datetime import datetime
from pathlib import Path

from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4, legal, landscape
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    Image,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

REPO_ROOT = Path(__file__).resolve()
REPO_ROOT = REPO_ROOT.parents[2]
OUTPUT_DIR = REPO_ROOT / "generated_documents"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

AlignmentValue = Union[int, str, None]
ImageAlignOutput = Literal["LEFT", "CENTER", "CENTRE", "RIGHT"]


class PdfResult(TypedDict):
    type: str
    filename: str
    mime_type: str
    base64: str
    size_bytes: int
    path: str
    meta: Dict[str, Any]

def _safe_filename(name: str) -> str:
    s = (name or "document").strip()
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^A-Za-z0-9_\-\.]", "", s)
    return s[:80] if s else "document"


def parse_components_config(components_json: str) -> Dict[str, Dict[str, Any]]:
    defaults = {
        "header": {
            "enabled": True,
            "text": "Generated Document",
            "font_size": 14,
            "color": "#000000",
            "alignment": TA_CENTER,
        },
        "footer": {
            "enabled": True,
            "text": "Generated: {timestamp}",
            "font_size": 9,
            "color": "#666666",
            "alignment": TA_CENTER,
        },
        "table_style": {
            "header_bg": "#2C5282",
            "header_text": "#FFFFFF",
            "row_colors": ["#FFFFFF", "#F7FAFC"],
            "border_color": "#CCCCCC",
        },
        "signature": {
            "enabled": False,
            "font_size": 10,
            "alignment": TA_RIGHT,
        },
    }

    if not components_json:
        return defaults

    if isinstance(components_json, str) and components_json.strip() == "{}":
        return defaults

    try:
        user_cfg = json.loads(components_json) if isinstance(components_json, str) else (components_json or {})
        merged = {}
        for key, base in defaults.items():
            merged[key] = {**base, **user_cfg.get(key, {})}
        return merged
    except json.JSONDecodeError:
        return defaults


def _register_fonts(font_cfg: Dict[str, Any]) -> Dict[str, str]:
    registered: Dict[str, str] = {}
    for role, spec in (font_cfg or {}).items():
        name = (spec or {}).get("name")
        path = (spec or {}).get("ttf_path")
        if not name or not path:
            continue
        try:
            pdfmetrics.registerFont(TTFont(name, path))
            registered[role] = name
        except Exception:
            pass
    return registered


def _resolve_alignment(raw: AlignmentValue) -> int:
    if raw is None:
        return TA_LEFT
    if isinstance(raw, int):
        return raw
    if isinstance(raw, str):
        s = raw.lower().strip()
        return {
            "left": TA_LEFT,
            "center": TA_CENTER,
            "centre": TA_CENTER,
            "right": TA_RIGHT,
            "justify": TA_JUSTIFY,
        }.get(s, TA_LEFT)
    return TA_LEFT


def _parse_color(value: Any, profile: str = "RGB", default="#000000"):
    if not value:
        value = default

    if isinstance(value, dict) and profile.upper() == "CMYK":
        return colors.CMYKColor(
            float(value.get("c", 0)) * 100,
            float(value.get("m", 0)) * 100,
            float(value.get("y", 0)) * 100,
            float(value.get("k", 0)) * 100,
        )

    if isinstance(value, str):
        if value.startswith("#"):
            return colors.HexColor(value)
        if value.lower().startswith("rgb("):
            r, g, b = [int(x.strip()) for x in value[4:-1].split(",")]
            return colors.Color(r / 255, g / 255, b / 255)

    return colors.HexColor(default)


def _get_page_size(doc_cfg: Dict[str, Any]):
    base = {
        "LETTER": letter,
        "A4": A4,
        "LEGAL": legal,
    }.get(str((doc_cfg or {}).get("page_size", "LETTER")).upper(), letter)

    if str((doc_cfg or {}).get("orientation", "portrait")).lower() == "landscape":
        return landscape(base)
    return base


def _resolve_image_align(raw: AlignmentValue) -> ImageAlignOutput:
    if isinstance(raw, str):
        return raw.upper()
    if isinstance(raw, int):
        return {0: "LEFT", 1: "CENTER", 2: "RIGHT"}.get(raw, "LEFT")
    return "LEFT"


def _render_logo_base64(logo_base64: str, elements: List[Any], width_in: float, align):
    if not logo_base64:
        return

    try:
        b64 = logo_base64.strip()

        # Allow data URL formats
        if b64.startswith("data:") and "base64," in b64:
            b64 = b64.split("base64,", 1)[1].strip()

        img_bytes = base64.b64decode(b64, validate=False)
        img_buf = io.BytesIO(img_bytes)

        img = Image(img_buf, width=width_in * inch)
        img.hAlign = _resolve_image_align(align)
        elements.append(img)
        elements.append(Spacer(1, 0.15 * inch))
    except Exception:
        return


def _apply_component_header(components, styles, elements, color_profile, font_roles):
    cfg = components["header"]
    if not cfg.get("enabled", True):
        return

    if cfg.get("logo_base64"):
        _render_logo_base64(
            cfg["logo_base64"],
            elements,
            cfg.get("logo_width", 1.2),
            cfg.get("logo_alignment"),
        )

    style = ParagraphStyle(
        "ComponentHeader",
        parent=styles["Heading1"],
        fontSize=cfg.get("font_size", 14),
        fontName=font_roles.get("heading", "Helvetica-Bold"),
        textColor=_parse_color(cfg.get("color"), color_profile),
        alignment=_resolve_alignment(cfg.get("alignment")),
        spaceAfter=12,
    )

    elements.append(Paragraph(cfg.get("text", ""), style))
    elements.append(Spacer(1, 0.15 * inch))


def _apply_component_footer(components, styles, elements, color_profile, font_roles):
    cfg = components["footer"]
    if not cfg.get("enabled", True):
        return

    text = cfg.get("text", "").format(timestamp=datetime.now().strftime("%Y-%m-%d %H:%M"))

    style = ParagraphStyle(
        "ComponentFooter",
        parent=styles["Normal"],
        fontSize=cfg.get("font_size", 9),
        fontName=font_roles.get("body", "Helvetica"),
        textColor=_parse_color(cfg.get("color"), color_profile),
        alignment=_resolve_alignment(cfg.get("alignment")),
        spaceBefore=12,
    )

    elements.append(Spacer(1, 0.2 * inch))
    elements.append(Paragraph(text, style))


def _apply_component_signature(components, styles, elements, font_roles):
    cfg = components["signature"]
    if not cfg.get("enabled", False):
        return

    text = " | ".join(
        filter(
            None,
            [cfg.get("name"), cfg.get("title"), cfg.get("email"), cfg.get("phone")],
        )
    )

    style = ParagraphStyle(
        "ComponentSignature",
        parent=styles["Normal"],
        fontSize=cfg.get("font_size", 10),
        fontName=font_roles.get("body", "Helvetica"),
        alignment=_resolve_alignment(cfg.get("alignment")),
        italic=True,
        spaceAfter=30,
    )

    elements.append(Spacer(1, 0.3 * inch))
    elements.append(Paragraph(text, style))


def _get_table_style_from_components(components, config, color_profile):
    table_cfg = (config or {}).get("table", {})
    comp = components["table_style"]

    row_colors = [
        _parse_color(c, color_profile)
        for c in table_cfg.get("row_colors", comp["row_colors"])
    ]

    return TableStyle(
        [
            ("BACKGROUND", (0, 0), (-1, 0), _parse_color(comp["header_bg"], color_profile)),
            ("TEXTCOLOR", (0, 0), (-1, 0), _parse_color(comp["header_text"], color_profile)),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), table_cfg.get("header_font_size", 12)),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), row_colors),
            ("ALIGN", (0, 1), (-1, -1), "CENTER"),
            ("GRID", (0, 0), (-1, -1), 1, _parse_color(comp["border_color"], color_profile)),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]
    )


def _apply_body_content(body, config, styles, elements, color_profile, font_roles):
    if not body:
        return

    cfg = (config or {}).get("body", {})
    style = ParagraphStyle(
        "BodyText",
        parent=styles["Normal"],
        fontName=font_roles.get("body", "Helvetica"),
        fontSize=cfg.get("font_size", 11),
        leading=cfg.get("leading", 1.4) * cfg.get("font_size", 11),
        textColor=_parse_color(cfg.get("color"), color_profile),
        alignment=_resolve_alignment(cfg.get("alignment", TA_JUSTIFY)),
        spaceAfter=cfg.get("space_after", 12),
    )

    for para in body.split("\n\n"):
        elements.append(Paragraph(para.replace("\n", "<br/>"), style))


def _generate_pdf_report_internal(
        out_path: Path,
        title: str,
        body: str,
        report_data: str,
        styling_config: str = "{}",
        components_config: str = "{}",
) -> Path:
    """Internal PDF generation implementation."""
    try:
        data = json.loads(report_data) if isinstance(report_data, str) else (report_data or {})
        config = json.loads(styling_config) if isinstance(styling_config, str) and styling_config else (styling_config or {})
        components = parse_components_config(components_config)

        font_roles = _register_fonts((config or {}).get("fonts", {}))

        doc_cfg = (config or {}).get("document", {})
        page_size = _get_page_size(doc_cfg)

        doc = SimpleDocTemplate(
            str(out_path),
            pagesize=page_size,
            topMargin=doc_cfg.get("margin_top", 0.75) * inch,
            bottomMargin=doc_cfg.get("margin_bottom", 0.75) * inch,
            leftMargin=doc_cfg.get("margin_left", 0.75) * inch,
            rightMargin=doc_cfg.get("margin_right", 0.75) * inch,
        )

        styles = getSampleStyleSheet()
        elements: List[Any] = []

        _apply_component_header(components, styles, elements, "RGB", font_roles)

        elements.append(Paragraph(title, styles["Heading2"]))
        elements.append(Spacer(1, 0.2 * inch))

        _apply_body_content(body, config, styles, elements, "RGB", font_roles)

        for section in (data or {}).get("sections", []):
            elements.append(Paragraph(section.get("title", "Section"), styles["Heading3"]))

            if section.get("table_data"):
                table = Table(section["table_data"])
                table.setStyle(_get_table_style_from_components(components, config, "RGB"))
                elements.append(table)
                elements.append(Spacer(1, 0.2 * inch))

        _apply_component_signature(components, styles, elements, font_roles)
        _apply_component_footer(components, styles, elements, "RGB", font_roles)

        doc.build(elements)
        return out_path

    except Exception as e:
        raise ValueError(f"PDF generation failed: {str(e)}")


def generate_pdf(
        template_path: Optional[str] = None,
        output_path: Optional[str] = None,
        data: Optional[Dict[str, Any]] = None,
        **kwargs
) -> PdfResult:
    """
    Wrapper for PDF generation.
    """
    if not data:
        raise ValueError("Missing 'data' parameter with PDF configuration")

    try:
        # Extract parameters from data dict
        title = data.get("title", "Report")
        body = data.get("body", "")
        report_data = data.get("report_data", "{}")
        styling_config = data.get("styling_config", "{}")
        components_config = data.get("components_config", "{}")

        # Convert dict inputs to JSON strings
        if isinstance(report_data, str):
            report_data_str = report_data
        else:
            report_data_str = json.dumps(report_data)

        if isinstance(styling_config, str):
            styling_config_str = styling_config
        else:
            styling_config_str = json.dumps(styling_config)

        if isinstance(components_config, str):
            components_config_str = components_config
        else:
            components_config_str = json.dumps(components_config)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe = _safe_filename(title)
        filename = f"{safe}_{timestamp}.pdf"
        out_path = Path(output_path) if output_path else (OUTPUT_DIR / filename)

        # Generate PDF file
        generated_path = _generate_pdf_report_internal(
            out_path=out_path,
            title=title,
            body=body,
            report_data=report_data_str,
            styling_config=styling_config_str,
            components_config=components_config_str,
        )

        # Read file & encode
        file_bytes = generated_path.read_bytes()

        MAX_BYTES = 8 * 1024 * 1024
        if len(file_bytes) > MAX_BYTES:
            raise RuntimeError(f"Generated PDF is too large ({len(file_bytes)} bytes). Reduce content or tables.")

        return {
            "type": "file_base64",
            "filename": generated_path.name,
            "path": str(generated_path),
            "mime_type": "application/pdf",
            "base64": base64.b64encode(file_bytes).decode("utf-8"),
            "size_bytes": len(file_bytes),
            "meta": {
                "title": title,
                "body_chars": len(body or ""),
                "sections": len(json.loads(report_data_str).get("sections", [])) if report_data_str else 0,
                "timestamp": timestamp,
            },
        }
    except Exception as e:
        raise RuntimeError(f"PDF generation failed: {str(e)}")