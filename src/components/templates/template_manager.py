import uuid
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
import shutil
from docxtpl import DocxTemplate


class TemplateMetadata:

    def __init__(
            self,
            template_id: str,
            name: str,
            file_path: str,
            placeholders: List[str],
            created_at: str,
            description: str = ""
    ):
        self.template_id = template_id
        self.name = name
        self.file_path = file_path
        self.placeholders = placeholders
        self.created_at = created_at
        self.description = description

    def to_dict(self) -> Dict[str, Any]:
        return {
            "template_id": self.template_id,
            "name": self.name,
            "file_path": self.file_path,
            "placeholders": self.placeholders,
            "created_at": self.created_at,
            "description": self.description,
        }


class TemplateManager:
    """This class manages the template lifecycle.
     Includes the upload, introspection, storing, and retrieval of templates."""

    def __init__(self, templates_dir: Path):
        """
        Args:
            templates_dir: Path to store .docx template files
        """
        self.templates_dir = Path(templates_dir)
        self.templates_dir.mkdir(parents=True, exist_ok=True)
        self.metadata_file = self.templates_dir / ".templates_metadata.json"
        self._metadata_cache: Dict[str, TemplateMetadata] = {}
        self._load_metadata()

    def _load_metadata(self):
        """Load cached template metadata from disk."""
        if self.metadata_file.exists():
            import json
            try:
                data = json.loads(self.metadata_file.read_text())
                for tid, meta_dict in data.items():
                    self._metadata_cache[tid] = TemplateMetadata(**meta_dict)
            except Exception as e:
                print(f"Failed to load template metadata: {e}")

    def _save_metadata(self):
        """Persist metadata to disk."""
        import json
        data = {tid: meta.to_dict() for tid, meta in self._metadata_cache.items()}
        self.metadata_file.write_text(json.dumps(data, indent=2))

    def upload_template(
            self,
            source_path: Path,
            name: str,
            description: str = ""
    ) -> Tuple[str, List[str]]:
        """
        Upload a .docx template file.
        Args:
            source_path: Path to .docx file
            name: Template name
            description: File description
        Returns:
            (template_id, list_of_placeholders)
        """
        if not source_path.exists():
            raise ValueError(f"Template file not found: {source_path}")

        # Extract placeholders
        placeholders = self.extract_placeholders(source_path)

        # Generate ID and save
        template_id = str(uuid.uuid4())
        dest_path = self.templates_dir / f"{template_id}.docx"

        try:
            shutil.copy2(str(source_path), str(dest_path))
        except Exception as e:
            raise ValueError(f"Failed to copy template: {e}")

        # Cache metadata
        from datetime import datetime
        meta = TemplateMetadata(
            template_id=template_id,
            name=name,
            file_path=str(dest_path),
            placeholders=placeholders,
            created_at=datetime.now().isoformat(),
            description=description
        )
        self._metadata_cache[template_id] = meta
        self._save_metadata()

        return template_id, placeholders

    @staticmethod
    def extract_placeholders(template_path: Path) -> List[str]:
        """
        Extract all {{placeholder}} names from a .docx template.
        """
        try:
            tpl = DocxTemplate(str(template_path))
            undeclared = tpl.get_undeclared_template_variables()
            return sorted(list(undeclared))
        except Exception as e:
            raise ValueError(f"Failed to parse template: {e}")

    def get_template_path(self, template_id: str) -> Optional[Path]:
        """Get file path to stored template."""
        if template_id in self._metadata_cache:
            return Path(self._metadata_cache[template_id].file_path)
        return None

    def get_template_metadata(self, template_id: str) -> Optional[TemplateMetadata]:
        """Retrieve metadata for a template."""
        return self._metadata_cache.get(template_id)

    def list_templates(self) -> List[Dict[str, Any]]:
        """List all uploaded templates."""
        return [meta.to_dict() for meta in self._metadata_cache.values()]

    def delete_template(self, template_id: str) -> bool:
        """Delete a template and its metadata."""
        if template_id not in self._metadata_cache:
            return False

        path = Path(self._metadata_cache[template_id].file_path)
        try:
            if path.exists():
                path.unlink()
            del self._metadata_cache[template_id]
            self._save_metadata()
            return True
        except Exception as e:
            print(f"⚠️  Failed to delete template: {e}")
            return False

    def get_old_templates(self, max_age_days: int = 7) -> List[str]:
        """
        Get the list of template_ids older than max_age_days.
        Args:
            max_age_days: Age threshold
        """
        from datetime import datetime, timedelta
        cutoff = datetime.now() - timedelta(days=max_age_days)
        old_ids = []

        for template_id, meta in self._metadata_cache.items():
            created = datetime.fromisoformat(meta.created_at)
            if created < cutoff:
                old_ids.append(template_id)

        return old_ids

