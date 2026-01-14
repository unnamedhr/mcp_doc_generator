from pathlib import Path
from typing import Dict, Any, Optional
from datetime import datetime
from docxtpl import DocxTemplate


class TemplateDocumentGenerator:
    """Generate documents from templates."""

    def __init__(self, output_dir: Path):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def generate(
            self,
            template_path: Path,
            context: Dict[str, Any],
            output_filename: Optional[str] = None,
            output_format: str = "docx"
    ) -> str:
        """
        Generate a document from a template.

        Args:
            template_path: Path to .docx template
            context: Dict {placeholder: resolved_value, ...}
            output_filename: Custom filename (default: auto-generated)
            output_format: "docx" (PDF support via LibreOffice future)
        """
        if not template_path.exists():
            raise ValueError(f"Template not found: {template_path}")

        try:
            tpl = DocxTemplate(str(template_path))
        except Exception as e:
            raise ValueError(f"Failed to load template: {e}")

        # Render
        try:
            tpl.render(context)
        except Exception as e:
            raise ValueError(f"Jinja2 render error: {e}")

        # Generate filename if not provided
        if not output_filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"generated_{timestamp}.docx"

        output_path = self.output_dir / output_filename

        # Save
        try:
            tpl.save(str(output_path))
        except Exception as e:
            raise ValueError(f"Failed to save document: {e}")

        return str(output_path)

    def convert_to_pdf(self, docx_path: Path) -> Optional[str]:
        """
        Convert DOCX to PDF.
        """
        import subprocess

        if not docx_path.exists():
            raise ValueError(f"File not found: {docx_path}")

        try:
            result = subprocess.run(
                [
                    "libreoffice",
                    "--headless",
                    "--convert-to", "pdf",
                    "--outdir", str(self.output_dir),
                    str(docx_path)
                ],
                capture_output=True,
                timeout=30
            )

            if result.returncode == 0:
                pdf_path = docx_path.with_suffix(".pdf")
                return str(pdf_path)
            else:
                return None
        except Exception as e:
            print(f"PDF conversion failed: {e}")
            return None