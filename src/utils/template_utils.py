import re
import json
from pathlib import Path
from typing import List, Dict, Any, Optional, Callable
import base64
import tempfile
from docxtpl import DocxTemplate

from components.templates.template_manager import TemplateManager


def extract_jinja_placeholders_from_text(text: str) -> List[str]:
    pattern = r'\{\{\s*([a-zA-Z_][a-zA-Z0-9_.]*)\s*\}\}'
    matches = re.findall(pattern, text, re.IGNORECASE)
    return sorted(list(set(matches)))


def validate_template_placeholders(placeholders: List[str]) -> List[str]:
    valid_chars = re.compile(r'^[a-zA-Z_][a-zA-Z0-9_.]*$')
    return [ph for ph in placeholders if valid_chars.match(ph)]


def generate_default_mapping(placeholders: List[str]) -> List[Dict[str, Any]]:
    return [{"placeholder": ph, "path": ph} for ph in placeholders]


def create_temp_template_from_base64(base64_content: str) -> Path:
    binary_data = base64.b64decode(base64_content)
    temp_file = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    temp_file.write(binary_data)
    temp_file.close()
    return Path(temp_file.name)


def cleanup_old_templates(template_manager: TemplateManager, max_age_days: int = 7) -> int:
    old_template_ids = template_manager.get_old_templates(max_age_days)
    deleted = 0

    for template_id in old_template_ids:
        if template_manager.delete_template(template_id):
            deleted += 1

    if deleted > 0:
        print(f"🧹 Cleaned up {deleted} old templates")

    return deleted


def safe_render_template(
        template_path: Path,
        context: Dict[str, Any],
        error_handler: Optional[Callable[[str, str], Any]] = None
) -> Dict[str, Any]:
    missing_vars = set()

    try:
        tpl = DocxTemplate(str(template_path))
        undeclared = tpl.get_undeclared_template_variables()

        for var in undeclared:
            missing_vars.add(var)
            if error_handler:
                error_handler(var, f"Missing in context")

        context.update({var: "" for var in missing_vars})
        return context

    except Exception as e:
        print(f"Template validation failed: {e}")
        return context


def generate_preview_mapping(
        data_sample: Dict[str, Any],
        placeholders: List[str],
        max_suggestions: int = 10
) -> List[Dict[str, Any]]:
    suggestions = []

    for ph in placeholders:
        if ph in data_sample:
            suggestions.append({"placeholder": ph, "path": ph})
            continue

        for key in data_sample.keys():
            if ph.startswith(key.lower()):
                suggestions.append({
                    "placeholder": ph,
                    "path": f"{key}.{ph[len(key):]}"
                })

        suggestions.append({"placeholder": ph, "path": ph})

    return suggestions[:max_suggestions]


def format_mapping_for_display(mappings: List[Dict[str, Any]]) -> str:
    lines = []
    for m in mappings:
        ph = m["placeholder"]
        path = m.get("path", ph)
        transform = m.get("transform", "")
        transform_str = f" [{transform}]" if transform else ""
        lines.append(f"  {ph} ← {path}{transform_str}")
    return "\n".join(lines)


def export_mappings_to_json(mappings: List[Dict[str, Any]], filename: str) -> Path:
    path = Path(filename)
    path.write_text(json.dumps(mappings, indent=2))
    return path


def load_mappings_from_json(filename: str) -> List[Dict[str, Any]]:
    path = Path(filename)
    if not path.exists():
        raise FileNotFoundError(f"Mapping file not found: {filename}")
    return json.loads(path.read_text())