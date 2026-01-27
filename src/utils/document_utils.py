from pathlib import Path
from typing import Dict, Any, List
from datetime import datetime
import json

OUTPUT_DIR = Path("./generated_documents")


def _get_file_info(file_path: Path) -> Dict[str, Any]:
    """Extract file metadata."""
    stat = file_path.stat()
    return {
        "name": file_path.name,
        "path": str(file_path),
        "size_kb": round(stat.st_size / 1024, 2),
        "size_mb": round(stat.st_size / (1024 * 1024), 2),
        "modified": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
        "extension": file_path.suffix.lower(),
        "type": _get_file_type(file_path.suffix.lower())
    }


def _get_file_type(extension: str) -> str:
    """Map file extension to the document type."""
    type_map = {
        ".pdf": "PDF",
        ".docx": "Word",
        ".xlsx": "Excel",
        ".doc": "Word (Legacy)",
        ".xls": "Excel (Legacy)"
    }
    return type_map.get(extension, "Unknown")


def _format_list_output(files_info: List[Dict[str, Any]], limit: int) -> str:
    """Format file list as human-readable string."""
    if not files_info:
        return "No documents generated yet."

    total_count = len(files_info)
    shown = files_info[:limit]

    result = f"Generated Documents ({total_count} total, showing {len(shown)}):\n\n"

    for i, file in enumerate(shown, 1):
        result += (
            f"{i}. {file['name']}\n"
            f"   Type: {file['type']} | Size: {file['size_kb']} KB | "
            f"Modified: {file['modified']}\n"
        )

    if total_count > limit:
        result += f"\n... and {total_count - limit} more files"

    return result


def list_generated_documents(
        template_path: str = None,
        output_path: str = None,
        data: Dict[str, Any] = None,
        **kwargs
) -> str:
    """
    List all generated documents.
    """
    try:
        # Parse parameters
        if not data:
            data = {}

        limit = data.get("limit", 20)
        file_type = data.get("file_type", "all").lower()
        output_format = data.get("format", "text").lower()

        # Check if the directory exists
        if not OUTPUT_DIR.exists():
            return json.dumps({"files": [], "total": 0}) if output_format == "json" else "No documents generated yet."

        # Get all files
        all_files = list(OUTPUT_DIR.glob("*"))

        # Filter by type if specified
        if file_type != "all":
            extension_map = {
                "pdf": ".pdf",
                "word": ".docx",
                "docx": ".docx",
                "excel": ".xlsx",
                "xlsx": ".xlsx"
            }
            target_ext = extension_map.get(file_type)
            if target_ext:
                all_files = [f for f in all_files if f.suffix.lower() == target_ext]

        # Sort by modification time (newest first)
        all_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)

        if not all_files:
            return json.dumps({"files": [], "total": 0}) if output_format == "json" else "No documents found."

        # Extract file info
        files_info = [_get_file_info(f) for f in all_files]

        if output_format == "json":
            return json.dumps({
                "files": files_info[:limit],
                "total": len(files_info),
                "limit": limit,
                "filter": file_type
            }, indent=2)
        else:
            return _format_list_output(files_info, limit)

    except Exception as e:
        raise RuntimeError(f"Failed to list documents: {str(e)}")


def list_generated_documents_simple() -> str:
    """
    Lists all generated documents.
    """
    try:
        if not OUTPUT_DIR.exists():
            return "No documents generated yet."

        files = sorted(OUTPUT_DIR.glob("*"), key=lambda x: x.stat().st_mtime, reverse=True)
        if not files:
            return "No documents generated yet."

        result = "Generated Documents:\n"
        for i, file in enumerate(files[:20], 1):
            file_size = file.stat().st_size / 1024
            result += f"{i}. {file.name} ({file_size:.1f} KB)\n"
        return result
    except Exception as e:
        return f"✗ Error listing documents: {str(e)}"