import os
import sys
from pathlib import Path
from typing import Any, Dict
from fastmcp import FastMCP

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "src"))

from src.generators.pdf_generator import generate_pdf
from src.generators.word_generator import generate_word
from src.generators.excel_generator import generate_excel
from src.generators.template_generator import TemplateDocumentGenerator
from src.tools.template_tools import TemplateTools, GenerateFromBase64TemplateReq


def get_output_dir() -> Path:
    out = os.getenv("MCP_DOC_OUTPUT_PATH", str(ROOT / "generated_documents"))
    p = Path(out).resolve()
    p.mkdir(parents=True, exist_ok=True)
    return p


mcp = FastMCP("mcp_doc_generator")


@mcp.tool
def generate_pdf(data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Generate a PDF using PDF generator.
    """
    return generate_pdf(data=data)


@mcp.tool
def generate_word(data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Generate a Word doc using Word generator.
    """
    return generate_word(data=data)


@mcp.tool
def generate_excel(data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Generate an Excel file using Excel generator.
    """
    import json

    def jsonify_if_needed(x: Any) -> Any:
        if isinstance(x, (dict, list)):
            return json.dumps(x)
        return x

    for k in ["headers", "data_rows"]:
        if k in data:
            data[k] = jsonify_if_needed(data[k])

    for k in [
        "styling_config",
        "sheet_config",
        "totals_config",
        "validation_config",
        "freeze_config",
        "filter_config",
        "conditional_config",
        "columns_config",
    ]:
        if k in data:
            data[k] = jsonify_if_needed(data[k])

    return generate_excel(**data)


@mcp.tool
def generate_from_template(data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Generate a document from a base64 template using TemplateTools.
    """
    tg = TemplateDocumentGenerator(output_dir=get_output_dir())
    tools = TemplateTools(template_generator=tg)

    req = GenerateFromBase64TemplateReq(**data)
    result = tools.generate_from_template(req)

    if hasattr(result, "model_dump"):
        return result.model_dump()
    if hasattr(result, "dict"):
        return result.dict()
    return result


def main():
    mcp.run()


if __name__ == "__main__":
    main()