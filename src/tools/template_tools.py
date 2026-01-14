from pathlib import Path
from typing import Dict, Any, List, Optional
from pydantic import BaseModel, Field

from components.templates.template_manager import TemplateManager
from components.template_config.mapping_resolver import MappingResolver
from generators.template_generator import TemplateDocumentGenerator
from utils.template_utils import (
    validate_template_placeholders,
    generate_default_mapping,
    generate_preview_mapping,
    safe_render_template,
    format_mapping_for_display,
)

class UploadTemplateReq(BaseModel):
    template_path: str = Field(..., description="Path to .docx template file")
    name: str = Field(..., description="Template name")
    description: str = Field(default="", description="Template description")


class UploadTemplateOutput(BaseModel):
    template_id: str = Field(..., description="Unique ID of uploaded template")
    placeholders: List[str] = Field(..., description="List of detected {{placeholder}} names")
    default_mapping: List[Dict[str, Any]] = Field(..., description="Auto-generated default mapping")


class ListTemplatesOutput(BaseModel):
    templates: List[Dict[str, Any]] = Field(..., description="List of all templates with metadata")


class GetPlaceholdersReq(BaseModel):
    template_id: str = Field(..., description="ID of template")


class GetPlaceholdersOutput(BaseModel):
    template_id: str
    placeholders: List[str] = Field(..., description="All {{placeholder}} names in template")


class GenerateFromTemplateReq(BaseModel):
    template_id: str = Field(..., description="ID of template to use")
    data: Dict[str, Any] = Field(..., description="Source JSON data")
    mapping: Optional[List[Dict[str, Any]]] = Field(
        default=None,
        description="""Optional explicit mapping rules:
            [
              {"placeholder": "full_name", "path": "person.name", "transform": "capitalize"},
              {"placeholder": "invoice_amount", "path": "items[*].price", "transform": "currency:EUR"},
            ]
            If omitted, placeholders are matched directly to data keys (e.g., {{foo}} → data['foo']).
        """
    )
    output_filename: Optional[str] = Field(
        default=None,
        description="Custom output filename (default: auto-generated)"
    )
    auto_fallback: bool = Field(
        default=True,
        description="If mapping fails, try placeholder name as JSONPath"
    )
    preview_suggestions: bool = Field(
        default=False,
        description="If true and no mapping, return mapping suggestions (no generation)"
    )
    validate_template: bool = Field(
        default=True,
        description="Validate template placeholders before rendering"
    )
    table_mappings: Optional[List[Dict[str, Any]]] = Field(
        default=None,
        description="Array table mappings: [{'column': 'Product', 'path': 'items[*].name'}]"
    )
    table_placeholder: Optional[str] = Field(
        default="table_data",
        description="Placeholder name for auto-populated table (default: table_data)"
    )


class GenerateFromTemplateOutput(BaseModel):
    document_id: str = Field(..., description="ID/filename of generated document")
    file_path: str = Field(..., description="Full path to generated file")
    used_mapping: Optional[str] = Field(..., description="Readable summary of applied mapping")
    missing_vars: List[str] = Field(default=[], description="Template vars that were missing")


class DeleteTemplateReq(BaseModel):
    template_id: str = Field(..., description="ID of template to delete")


class DeleteTemplateOutput(BaseModel):
    success: bool
    message: str


class TemplateTools:

    def __init__(
            self,
            template_manager: TemplateManager,
            template_generator: TemplateDocumentGenerator
    ):
        self.tm = template_manager
        self.tg = template_generator

    def upload_template(self, req: UploadTemplateReq) -> UploadTemplateOutput:
        try:
            template_path = Path(req.template_path)
            template_id, raw_placeholders = self.tm.upload_template(
                source_path=template_path,
                name=req.name,
                description=req.description
            )

            # Validate and generate defaults
            validated_placeholders = validate_template_placeholders(raw_placeholders)
            default_mapping = generate_default_mapping(validated_placeholders)

            return UploadTemplateOutput(
                template_id=template_id,
                placeholders=validated_placeholders,
                default_mapping=default_mapping
            )
        except Exception as e:
            raise ValueError(f"Upload failed: {e}")

    def list_templates(self) -> ListTemplatesOutput:
        """List all uploaded templates."""
        templates = self.tm.list_templates()
        return ListTemplatesOutput(templates=templates)

    def get_placeholders(self, req: GetPlaceholdersReq) -> GetPlaceholdersOutput:
        """Get validated placeholders for template."""
        meta = self.tm.get_template_metadata(req.template_id)
        if not meta:
            raise ValueError(f"Template {req.template_id} not found")

        validated = validate_template_placeholders(meta.placeholders)
        return GetPlaceholdersOutput(
            template_id=req.template_id,
            placeholders=validated
        )

    def generate_from_template(self, req: GenerateFromTemplateReq) -> GenerateFromTemplateOutput:
        # Get template
        template_path = self.tm.get_template_path(req.template_id)
        if not template_path:
            raise ValueError(f"Template {req.template_id} not found")

        if req.preview_suggestions and not req.mapping:
            meta = self.tm.get_template_metadata(req.template_id)
            suggestions = generate_preview_mapping(req.data, meta.placeholders)
            raise ValueError(f"PREVIEW_SUGGESTIONS:\n{format_mapping_for_display(suggestions)}")

        missing_vars = []
        if req.validate_template:
            def error_handler(ph: str, msg: str):
                missing_vars.append(ph)

            safe_render_template(template_path, {}, error_handler)

        # Build context
        if req.mapping or req.table_mappings:
            context = MappingResolver.build_context_from_dict(
                data=req.data,
                mapping_dict=req.mapping or [],
                table_mappings=req.table_mappings,
                table_placeholder=req.table_placeholder,
                auto_fallback=req.auto_fallback
            )
            used_mapping = format_mapping_for_display(req.mapping or []) + "Table: " + str(len(req.table_mappings or [])) + " columns"
        else:
            context = req.data
            used_mapping = "Direct mapping"

        # Generate document
        try:
            file_path = self.tg.generate(
                template_path=template_path,
                context=context,
                output_filename=req.output_filename
            )

            return GenerateFromTemplateOutput(
                document_id=Path(file_path).name,
                file_path=file_path,
                used_mapping=used_mapping,
                missing_vars=missing_vars
            )
        except Exception as e:
            raise ValueError(f"Generation failed: {e}")

    def delete_template(self, req: DeleteTemplateReq) -> DeleteTemplateOutput:
        """Delete a template."""
        success = self.tm.delete_template(req.template_id)
        return DeleteTemplateOutput(
            success=success,
            message="Deleted" if success else "Not found"
        )