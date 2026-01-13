from pathlib import Path

OUTPUT_DIR = Path("./generated_documents")

def list_generated_documents() -> str:
    """List all generated documents."""
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
