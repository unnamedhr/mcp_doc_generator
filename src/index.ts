#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { spawnSync } from "node:child_process";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const projectRoot = resolve(__dirname, "..");
const bridgePath = resolve(projectRoot, "mcp_bridge.py");

function callPython(toolName: string, payload: unknown) {
  const python = process.env.PYTHON || "python";

  const res = spawnSync(python, [bridgePath, toolName], {
    input: JSON.stringify(payload ?? {}),
    encoding: "utf-8",
    cwd: projectRoot,
  });

  if (res.status !== 0) {
    const msg =
      [res.stderr?.trim(), res.stdout?.trim()].filter(Boolean).join("\n") ||
      "Document Generator tool failed";
    throw new Error(msg);
  }

  const out = (res.stdout || "").trim();
  if (!out) throw new Error("Empty response from Document Generator tool");

  try {
    return JSON.parse(out);
  } catch {
    throw new Error(`Python returned non-JSON output:\n${out}`);
  }
}

const server = new McpServer({
  name: "mcp_doc_generator",
  version: "0.1.0",
});

const ToolInputShape = { data: z.any() };

server.tool(
  "excel_generator",
  "Generate an Excel (.xlsx)",
  ToolInputShape,
  async ({ data }) => {
    const result = callPython("excel_generator", { data });
    return { content: [{ type: "text", text: JSON.stringify(result) }] };
  }
);

server.tool(
  "pdf_generator",
  "Generate a PDF",
  ToolInputShape,
  async ({ data }) => {
    const result = callPython("pdf_generator", { data });
    return { content: [{ type: "text", text: JSON.stringify(result) }] };
  }
);

server.tool(
  "word_generator",
  "Generate a Word (.docx)",
  ToolInputShape,
  async ({ data }) => {
    const result = callPython("word_generator", { data });
    return { content: [{ type: "text", text: JSON.stringify(result) }] };
  }
);

server.tool(
  "template_generator",
  "Generate a document from a template",
  ToolInputShape,
  async ({ data }) => {
    const result = callPython("template_generator", { data });
    return { content: [{ type: "text", text: JSON.stringify(result) }] };
  }
);

await server.connect(new StdioServerTransport());