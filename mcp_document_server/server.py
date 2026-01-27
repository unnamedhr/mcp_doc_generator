import sys
import json
from pathlib import Path
from typing import Any, Dict

from generators.pdf_generator import generate_pdf
from generators.excel_generator import generate_excel
from generators.word_generator import generate_word
from generators.template_generator import TemplateDocumentGenerator
from tools.template_tools import TemplateTools

OUTPUT_DIR = Path(__file__).resolve().parent / "generated_documents"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

tg = TemplateDocumentGenerator(output_dir=OUTPUT_DIR)
template_tools = TemplateTools(template_generator=tg)


def send(response: Dict[str, Any]):
    sys.stdout.write(json.dumps(response) + "\n")
    sys.stdout.flush()


def _tool_ok(_id, payload: Any):
    """
    MCP tool response: return structured JSON as text so the agent can parse it.
    """
    send({
        "jsonrpc": "2.0",
        "id": _id,
        "result": {
            "content": [
                {
                    "type": "text",
                    "text": json.dumps(payload)
                }
            ]
        }
    })


def _tool_error(_id, err: Exception):
    send({
        "jsonrpc": "2.0",
        "id": _id,
        "error": {
            "code": -32000,
            "message": str(err)
        }
    })


def handle_initialize(_id):
    send({
        "jsonrpc": "2.0",
        "id": _id,
        "result": {
            "protocolVersion": "2024-11-05",
            "capabilities": {
                "tools": {}
            }
        }
    })


def handle_tools_list(_id):
    send({
        "jsonrpc": "2.0",
        "id": _id,
        "result": {
            "tools": [
                {
                    "name": "pdf_generator",
                    "description": "Generate a PDF and return base64 content.",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "data": {"type": "object"},
                            "output_path": {"type": "string"}
                        },
                        "required": ["data"]
                    }
                },
                {
                    "name": "excel_generator",
                    "description": "Generate an Excel (.xlsx) and return base64 content.",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "title": {"type": "string"},
                            "sheet_name": {"type": "string"},
                            "headers": {"type": "string", "description": "JSON array string of headers"},
                            "data_rows": {"type": "string", "description": "JSON array string of row arrays"},
                            "styling_config": {"type": "string"},
                            "sheet_config": {"type": "string"},
                            "include_totals": {"type": "boolean"},
                            "totals_config": {"type": "string"},
                            "include_data_validation": {"type": "boolean"},
                            "validation_config": {"type": "string"},
                            "include_freeze_panes": {"type": "boolean"},
                            "freeze_config": {"type": "string"},
                            "include_autofilter": {"type": "boolean"},
                            "filter_config": {"type": "string"},
                            "include_conditional_formatting": {"type": "boolean"},
                            "conditional_config": {"type": "string"},
                            "columns_config": {"type": "string"},
                            "preset_theme": {"type": "string"}
                        },
                        "required": ["title", "sheet_name", "headers", "data_rows"]
                    }
                },
                {
                    "name": "word_generator",
                    "description": "Generate a Word (.docx) and return base64 content.",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "data": {"type": "object"},
                            "output_path": {"type": "string"}
                        },
                        "required": ["data"]
                    }
                },
                {
                    "name": "template_generator",
                    "description": "Generate DOCX/XLSX/PDF from a template using context data and return base64 content.",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "template_path": {"type": "string", "description": "Path to template on server filesystem"},
                            "context": {"type": "object"},
                            "output_filename": {"type": "string"},
                            "output_format": {"type": "string", "enum": ["docx", "xlsx", "pdf"]}
                        },
                        "required": ["template_path", "context"]
                    }
                },

                # Optional: keep your base64-template tool if your TemplateTools supports it
                {
                    "name": "generate_from_template",
                    "description": "Generate DOCX/XLSX/PDF from Base64 template + JSON data (returns base64 content).",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "base64_template": {"type": "string"},
                            "data": {"type": "object"},
                            "mapping": {"type": "array"},
                            "output_filename": {"type": "string"},
                            "output_format": {"type": "string", "enum": ["docx", "xlsx", "pdf"]},
                            "preview_suggestions": {"type": "boolean"},
                            "table_mappings": {"type": "array"},
                            "table_placeholder": {"type": "string"}
                        },
                        "required": ["base64_template", "data"]
                    }
                }
            ]
        }
    })


def handle_tool_call(_id, name: str, args: Dict[str, Any]):
    try:
        if name == "pdf_generator":
            result = generate_pdf(**args)

        elif name == "excel_generator":
            result = generate_excel(**args)

        elif name == "word_generator":
            result = generate_word(**args)

        elif name == "template_generator":
            # Uses filesystem template path and context
            template_path = Path(args["template_path"])
            context = args.get("context", {})
            output_filename = args.get("output_filename")
            output_format = args.get("output_format")

            result = tg.generate_base64(
                template_path=template_path,
                context=context,
                output_filename=output_filename,
                output_format=output_format,
            )

        elif name == "generate_from_template":
            result = template_tools.generate_from_template(**args)

        else:
            raise ValueError(f"Unknown tool: {name}")

        _tool_ok(_id, result)

    except Exception as e:
        _tool_error(_id, e)


def main():
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue

        message = json.loads(line)
        method = message.get("method")
        _id = message.get("id")

        if method == "initialize":
            handle_initialize(_id)

        elif method == "tools/list":
            handle_tools_list(_id)

        elif method == "tools/call":
            params = message.get("params", {})
            handle_tool_call(_id, params.get("name", ""), params.get("arguments", {}) or {})

        elif method == "shutdown":
            sys.exit(0)


if __name__ == "__main__":
    main()
