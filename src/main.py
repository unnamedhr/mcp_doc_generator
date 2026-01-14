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

# Template directory
TEMPLATES_DIR = Path("./templates")
TEMPLATES_DIR.mkdir(exist_ok=True)

# Import generator (separate files)
from generators.pdf_generator import generate_pdf_report
from generators.excel_generator import generate_excel_report
from generators.word_generator import generate_word_report
from utils.document_utils import list_generated_documents

# Import template tools
from components.templates.template_manager import TemplateManager
from generators.template_generator import TemplateDocumentGenerator
from tools.template_tools import TemplateTools
from utils.template_utils import cleanup_old_templates

# Instantiate template dependencies
template_manager = TemplateManager(templates_dir=TEMPLATES_DIR)
template_generator = TemplateDocumentGenerator(output_dir=OUTPUT_DIR)
template_tools = TemplateTools(
    template_manager=template_manager,
    template_generator=template_generator
)

# Template management
print("Running template cleanup...", file=sys.stderr)
cleanup_old_templates(template_manager, max_age_days=30)
print("Cleanup complete", file=sys.stderr)

# Register tools
mcp.tool(generate_pdf_report)
mcp.tool(generate_excel_report)
mcp.tool(generate_word_report)
mcp.tool(list_generated_documents)
mcp.tool()(template_tools.upload_template)
mcp.tool()(template_tools.list_templates)
mcp.tool()(template_tools.get_placeholders)
mcp.tool()(template_tools.generate_from_template)
mcp.tool()(template_tools.delete_template)

if __name__ == "__main__":
    print("Document Generator MCP Server ready", file=sys.stderr)
    print(f"Output: {OUTPUT_DIR.absolute()}", file=sys.stderr)
    print(f"Template system ({len(template_manager.list_templates())} templates)", file=sys.stderr)
    mcp.run()