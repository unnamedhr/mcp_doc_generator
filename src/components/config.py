from typing import Dict, Any, Optional
from dataclasses import dataclass
from datetime import datetime
import json

@dataclass
class DocumentComponents:
    header: Optional[Dict[str, Any]] = None
    footer: Optional[Dict[str, Any]] = None
    table_style: Optional[Dict[str, Any]] = None
    signature: Optional[Dict[str, Any]] = None

def get_default_components() -> DocumentComponents:
    return DocumentComponents(
        header={
            "enabled": True,
            "text": "Document content",
            "font_size": 14,
            "color": "#1F3A93",
            "bold": True
        },
        footer={
            "enabled": True,
            "text": "Confidential | Generated: {timestamp}",
            "font_size": 9,
            "color": "#666666",
            "alignment": "center"
        },
        table_style={
            "header_bg": "#2C5282",
            "header_text": "#FFFFFF",
            "row_colors": ["#FFFFFF", "#F7FAFC"],
            "border_color": "#CCCCCC"
        },
        signature={
            "enabled": False,
            "name": "John Doe",
            "title": "Director",
            "email": "john@company.com",
            "phone": "**********"
        }
    )

def parse_components_config(components_json: str) -> DocumentComponents:
    if not components_json or components_json == "{}":
        return get_default_components()
    return DocumentComponents(**json.loads(components_json))
