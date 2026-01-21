from fastmcp import FastMCP
from pathlib import Path
import sys
import logging

# Loggers
logging.getLogger("mcp.server").setLevel(logging.ERROR)
logging.getLogger("mcp").setLevel(logging.ERROR)

# Initialize MCP server
mcp = FastMCP(name="Document Generator Server")

# Output directory
OUTPUT_DIR = Path("./generated_documents")
OUTPUT_DIR.mkdir(exist_ok=True)

# Import generator
from generators.pdf_generator import generate_pdf_report
from generators.excel_generator import generate_excel_report
from generators.word_generator import generate_word_report
from utils.document_utils import list_generated_documents

# Import template tools
from generators.template_generator import TemplateDocumentGenerator
from tools.template_tools import TemplateTools

# Instantiate
template_generator = TemplateDocumentGenerator(output_dir=OUTPUT_DIR)
template_tools = TemplateTools(template_generator=template_generator)

# Register tools
mcp.tool(generate_pdf_report)
mcp.tool(generate_excel_report)
mcp.tool(generate_word_report)
mcp.tool(list_generated_documents)
mcp.tool()(template_tools.generate_from_template)

if __name__ == "__main__":
    print("Document Generator MCP Server ready", file=sys.stderr)
    print(f"Output: {OUTPUT_DIR.absolute()}", file=sys.stderr)
    print("Template:", file=sys.stderr)
    mcp.run()