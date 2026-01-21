from typing import Dict, Any
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


def register_fonts(font_cfg: Dict[str, Any]) -> Dict[str, str]:
    roles = {}
    for role, spec in font_cfg.items():
        try:
            pdfmetrics.registerFont(TTFont(spec["name"], spec["ttf_path"]))
            roles[role] = spec["name"]
        except Exception:
            pass
    return roles
