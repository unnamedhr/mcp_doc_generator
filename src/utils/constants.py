from pathlib import Path
from typing import Union, Literal
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY

OUTPUT_DIR = Path("./generated_documents")
OUTPUT_DIR.mkdir(exist_ok=True)

AlignmentValue = Union[int, str, None]
ImageAlign = Literal["LEFT", "CENTER", "RIGHT"]

DEFAULT_ALIGNMENT_MAP = {
    "left": TA_LEFT,
    "center": TA_CENTER,
    "centre": TA_CENTER,
    "right": TA_RIGHT,
    "justify": TA_JUSTIFY,
}
