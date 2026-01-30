from typing import Any
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import letter, A4, legal, landscape
from constants import AlignmentValue, DEFAULT_ALIGNMENT_MAP


def resolve_alignment(value: AlignmentValue) -> int:
    if isinstance(value, int):
        return value
    if isinstance(value, str):
        return DEFAULT_ALIGNMENT_MAP.get(value.lower(), TA_LEFT)
    return TA_LEFT


def resolve_page_size(doc_cfg: dict):
    base = {
        "LETTER": letter,
        "A4": A4,
        "LEGAL": legal,
    }.get(str(doc_cfg.get("page_size", "LETTER")).upper(), letter)

    return landscape(base) if doc_cfg.get("orientation") == "landscape" else base


def parse_color(value: Any, default="#000000"):
    if not value:
        return colors.HexColor(default)
    if isinstance(value, str):
        return colors.HexColor(value)
    return colors.HexColor(default)