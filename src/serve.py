import sys
import json
from pathlib import Path
from generators.pdf_generator import generate_pdf_report
from generators.excel_generator import generate_excel
from generators.word_generator import generate_word_report
from generators.template_generator import TemplateDocumentGenerator
from tools.template_tools import TemplateTools

tg = TemplateDocumentGenerator(output_dir=Path("generated_documents"))
template_tools = TemplateTools(template_generator=tg)


def send(response):
    sys.stdout.write(json.dumps(response) + "\n")
    sys.stdout.flush()


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
                    "name": "generate_pdf",
                    "description": "Generate a PDF document.",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "template_path": {"type": "string"},
                            "output_path": {"type": "string"},
                            "data": {"type": "object"}
                        },
                        "required": ["template_path", "output_path", "data"]
                    }
                },
                {
                    "name": "generate_excel",
                    "description": "Generate an Excel document.",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "template_path": {"type": "string"},
                            "output_path": {"type": "string"},
                            "data": {"type": "object"}
                        },
                        "required": ["template_path", "output_path", "data"]
                    }
                },
                {
                    "name": "generate_word",
                    "description": "Generate a Word document.",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "template_path": {"type": "string"},
                            "output_path": {"type": "string"},
                            "data": {"type": "object"}
                        },
                        "required": ["template_path", "output_path", "data"]
                    }
                },
                {
                    "name": "generate_from_template",
                    "description": "Generate DOCX/XLSX/PDF from Base64 template + JSON data",
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


def handle_tool_call(_id, name, args):
    try:
        if name == "generate_pdf":
            result = generate_pdf_report(**args)
        elif name == "generate_excel":
            args["output_path"] = str(Path(args.get("output_path", "generated_documents/report.xlsx")).resolve())
            result = generate_excel(**args)
        elif name == "generate_word":
            result = generate_word_report(**args)
        elif name == "generate_from_template":
            result = template_tools.generate_from_template(**args)
        else:
            raise ValueError(f"Unknown tool: {name}")

        send({
            "jsonrpc": "2.0",
            "id": _id,
            "result": {
                "content": [
                    {
                        "type": "text",
                        "text": f"Document generated successfully: {result}"
                    }
                ]
            }
        })

    except Exception as e:
        send({
            "jsonrpc": "2.0",
            "id": _id,
            "error": {
                "code": -32000,
                "message": str(e)
            }
        })


def main():
    for line in sys.stdin:
        message = json.loads(line)
        method = message.get("method")
        _id = message.get("id")

        if method == "initialize":
            handle_initialize(_id)

        elif method == "tools/list":
            handle_tools_list(_id)

        elif method == "tools/call":
            params = message["params"]
            handle_tool_call(_id, params["name"], params.get("arguments", {}))

        elif method == "shutdown":
            sys.exit(0)


if __name__ == "__main__":
    main()
