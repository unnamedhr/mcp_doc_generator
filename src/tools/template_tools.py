from pathlib import Path
from typing import Dict, Any, List, Optional
from pydantic import BaseModel, Field
from components.template_config.mapping_resolver import MappingResolver
from generators.template_generator import TemplateDocumentGenerator
from utils.template_utils import (
    create_temp_template_from_base64,
    generate_preview_mapping,
    safe_render_template,
    format_mapping_for_display,
    validate_template_placeholders
)
from docxtpl import DocxTemplate
from openpyxl import load_workbook
import re


class GenerateFromBase64TemplateReq(BaseModel):
    base64_template: str = Field(..., description="Base64-encoded .docx or .xlsx template")
    data: Dict[str, Any] = Field(..., description="Source JSON data")
    mapping: Optional[List[Dict[str, Any]]] = Field(default=None, description="Optional explicit mapping rules")
    output_filename: Optional[str] = Field(default=None, description="Custom output filename")
    output_format: Optional[str] = Field(default=None, description="Optional: docx, xlsx, pdf")
    auto_fallback: bool = Field(default=True, description="Fallback to direct key match if path fails")
    preview_suggestions: bool = Field(default=False, description="Return mapping suggestions, no generation")
    validate_template: bool = Field(default=True, description="Validate placeholders before rendering")
    table_mappings: Optional[List[Dict[str, Any]]] = Field(default=None, description="Table column mappings")
    table_placeholder: Optional[str] = Field(default=None, description="Table placeholder name")


class GenerateFromBase64TemplateOutput(BaseModel):
    document_id: str = Field(..., description="Generated document filename")
    filepath: str = Field(..., description="Full path to generated file")
    used_mapping: Optional[str] = Field(default=None, description="Summary of applied mapping")
    missing_vars: List[str] = Field(default_factory=list, description="Unresolved placeholders")
    placeholders: Optional[List[str]] = Field(default=None, description="Detected placeholders (scalars + tables)")
    table_mappings: Optional[Dict[str, List[Dict[str, str]]]] = Field(default=None, description="Detected table structures")


_PLACEHOLDER_RE = re.compile(r"\{\{\s*([a-zA-Z_][a-zA-Z0-9_.]*)\s*\}\}", re.IGNORECASE)
_TABLE_PLACEHOLDER_RE = re.compile(r"^\s*\{\{\s*/([a-zA-Z_][a-zA-Z0-9_.]*)\s*\}\}\s*$")


def extract_template_structure(template_path: Path) -> tuple[List[str], Dict[str, List[Dict[str, str]]]]:
    suffix = template_path.suffix.lower()
    placeholders = set()
    table_mappings: Dict[str, List[Dict[str, str]]] = {}

    if suffix == ".docx":
        try:
            tpl = DocxTemplate(str(template_path))
            placeholders.update(tpl.get_undeclared_template_variables())
        except Exception:
            pass

    if suffix == ".xlsx":
        try:
            wb = load_workbook(str(template_path), data_only=False)

            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]

                # Detect tables
                table_mappings.update(_detect_excel_tables(ws))

                # Detect scalar placeholders
                for row in ws.iter_rows():
                    for cell in row:
                        if isinstance(cell.value, str):
                            placeholders.update(_PLACEHOLDER_RE.findall(cell.value))

        except Exception as e:
            print(f"XLSX extraction warning: {e}")

    return sorted(list(placeholders)), table_mappings


def _detect_excel_tables(ws) -> Dict[str, List[Dict[str, str]]]:
    table_mappings: Dict[str, List[Dict[str, str]]] = {}

    for row_idx in range(1, ws.max_row + 1):
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if not isinstance(cell.value, str):
                continue

            match = _TABLE_PLACEHOLDER_RE.match(cell.value)
            if not match:
                continue

            table_name = match.group(1)
            header_row_idx = row_idx - 1
            if header_row_idx < 1:
                continue

            # Extract headers
            headers: List[str] = []
            c = col_idx
            while c <= ws.max_column:
                h_cell = ws.cell(row=header_row_idx, column=c)
                h_val = str(h_cell.value).strip() if h_cell.value else ""
                if not h_val:
                    break
                headers.append(h_val)
                c += 1

            if headers:
                table_mappings[table_name] = [
                    {"column": header, "key": re.sub(r"[^a-zA-Z0-9_]", "_", header.lower())}
                    for header in headers
                ]

    return table_mappings


class TemplateTools:
    def __init__(self, template_generator: TemplateDocumentGenerator):
        self.tg = template_generator

    def generate_from_template(self, req: GenerateFromBase64TemplateReq) -> GenerateFromBase64TemplateOutput:
        """ Generate Report from Template."""
        template_path = create_temp_template_from_base64(req.base64_template)
        try:
            raw_placeholders, detected_table_mappings = extract_template_structure(template_path)
            meta_placeholders = validate_template_placeholders(raw_placeholders)

            if req.preview_suggestions:
                suggestions = generate_preview_mapping(req.data, meta_placeholders)
                return GenerateFromBase64TemplateOutput(
                    document_id="preview",
                    filepath="",
                    used_mapping=format_mapping_for_display(suggestions),
                    missing_vars=[],
                    placeholders=meta_placeholders,
                    table_mappings=detected_table_mappings
                )

            missing_vars: List[str] = []
            if req.validate_template and template_path.suffix.lower() == ".docx":
                def error_handler(ph: str, msg: str):
                    if not MappingResolver.is_table_placeholder(ph):  # Skip tables in validation
                        missing_vars.append(ph)
                safe_render_template(template_path, {}, error_handler)

            # Table logic (preserved exactly)
            table_placeholder = req.table_placeholder
            detected_tables = detected_table_mappings
            if table_placeholder is None:
                if len(detected_tables) == 1:
                    table_placeholder = next(iter(detected_tables.keys()))
                else:
                    table_placeholder = None

            effective_table_mappings = req.table_mappings
            if effective_table_mappings is None and table_placeholder and table_placeholder in detected_tables:
                effective_table_mappings = detected_tables[table_placeholder]

            # Separate scalar vs table placeholders
            scalar_placeholders = [ph for ph in meta_placeholders if not MappingResolver.is_table_placeholder(ph)]
            auto_mappings = [{"placeholder": ph, "path": ph} for ph in scalar_placeholders]
            effective_mappings = req.mapping or auto_mappings

            context = MappingResolver.build_context_from_dict(
                data=req.data,
                mapping_dict=effective_mappings,
                table_mappings=effective_table_mappings,
                table_placeholder=table_placeholder,
                auto_fallback=req.auto_fallback
            )

            used_mapping = "Auto-generated from scalar placeholders"
            if req.mapping:
                used_mapping = format_mapping_for_display(req.mapping)
            if effective_table_mappings:
                used_mapping += f"\nTable {table_placeholder}: {len(effective_table_mappings)} columns"

            filepath = self.tg.generate(
                template_path=template_path,
                context=context,
                output_filename=req.output_filename,
                output_format=req.output_format
            )

            return GenerateFromBase64TemplateOutput(
                document_id=Path(filepath).name,
                filepath=filepath,
                used_mapping=used_mapping,
                missing_vars=missing_vars,
                placeholders=meta_placeholders,
                table_mappings=detected_table_mappings
            )
        except Exception as e:
            raise ValueError(f"Generation failed: {e}")
        finally:
            if template_path.exists():
                template_path.unlink()