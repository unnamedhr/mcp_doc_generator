import sys
import json
from pathlib import Path
from typing import Any
import os

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "src"))

from src.generators.excel_generator import generate_excel
from src.generators.pdf_generator import generate_pdf
from src.generators.word_generator import generate_word
from src.generators.template_generator import TemplateDocumentGenerator
from src.tools.template_tools import TemplateTools, GenerateFromBase64TemplateReq

OUTPUT_DIR = Path(os.getenv("MCP_DOC_OUTPUT_PATH", str(ROOT / "generated_documents")))
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

tg = TemplateDocumentGenerator(output_dir=OUTPUT_DIR)
template_tools = TemplateTools(template_generator=tg)


def _read_payload() -> dict:
    raw = sys.stdin.read()
    return json.loads(raw) if raw.strip() else {}


def _jsonify_if_needed(x: Any) -> Any:
    if isinstance(x, (dict, list)):
        return json.dumps(x)
    return x


def main():
    if len(sys.argv) < 2:
        sys.stderr.write("Missing tool name argument\n")
        sys.exit(2)

    tool = sys.argv[1]
    payload = _read_payload()
    data = payload.get("data", payload)

    try:
        # Excel Generator
        if tool == "excel_generator":
            if "headers" in data:
                data["headers"] = _jsonify_if_needed(data["headers"])
            if "data_rows" in data:
                data["data_rows"] = _jsonify_if_needed(data["data_rows"])
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
                    data[k] = _jsonify_if_needed(data[k])

            result = generate_excel(**data)

        # PDF Generator
        elif tool == "pdf_generator":
            result = generate_pdf(data=data)

        # Word Generator
        elif tool == "word_generator":
            result = generate_word(data=data)

        # Template Generator
        elif tool == "template_generator":
            req = GenerateFromBase64TemplateReq(**data)
            result = template_tools.generate_from_template(req)
            if hasattr(result, "model_dump"):
                result = result.model_dump()
            elif hasattr(result, "dict"):
                result = result.dict()

        else:
            raise ValueError(f"Unknown tool: {tool}")

        sys.stdout.write(json.dumps(result))
        sys.stdout.flush()

    except Exception as e:
        sys.stderr.write(str(e))
        sys.stderr.flush()
        sys.exit(1)


if __name__ == "__main__":
    main()
