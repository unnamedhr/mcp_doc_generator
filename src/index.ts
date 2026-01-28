#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { spawnSync } from "node:child_process";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// dist/index.js -> project root
const projectRoot = resolve(__dirname, "..");
const bridgePath = resolve(projectRoot, "mcp_bridge.py");

function callPython(toolName: string, payload: unknown) {
  const python = process.env.PYTHON || "python";

  const res = spawnSync(
    python,
    [bridgePath, toolName],
    { input: JSON.stringify(payload ?? {}), encoding: "utf-8" }
  );

  if (res.status !== 0) {
    const msg = [res.stderr?.trim(), res.stdout?.trim()]
      .filter(Boolean)
      .join("\n") || "Python tool failed";
    throw new Error(msg);
  }

  const out = (res.stdout || "").trim();
  if (!out) throw new Error("Empty response from Python tool");

  try {
    return JSON.parse(out);
  } catch (e) {
    throw new Error(`Python returned non-JSON output:\n${out}`);
  }
}

const server = new McpServer({
  name: "mcp_doc_generator",
  version: "0.1.0",
});

const ToolPayload = z.object({ data: z.any() });

function parseToolArgs(params: unknown) {
  const maybeObj = (params ?? {}) as any;
  const raw = maybeObj.arguments ?? maybeObj;
  return ToolPayload.parse(raw);
}

server.tool(
  "excel_generator",
  "Generate an Excel (.xlsx)",
  async (params: unknown) => {
    const { data } = parseToolArgs(params);
    const result = callPython("excel_generator", { data });
    return { content: [{ type: "text", text: JSON.stringify(result) }] };
  }
);

server.tool(
  "pdf_generator",
  "Generate a PDF.",
  async (params: unknown) => {
    const { data } = parseToolArgs(params);
    const result = callPython("pdf_generator", { data });
    return { content: [{ type: "text", text: JSON.stringify(result) }] };
  }
);

server.tool(
  "word_generator",
  "Generate a Word (.docx)",
  async (params: unknown) => {
    const { data } = parseToolArgs(params);
    const result = callPython("word_generator", { data });
    return { content: [{ type: "text", text: JSON.stringify(result) }] };
  }
);

server.tool(
  "template_generator",
  "Generate a document from a template.",
  async (params: unknown) => {
    const { data } = parseToolArgs(params);
    const result = callPython("template_generator", { data });
    return { content: [{ type: "text", text: JSON.stringify(result) }] };
  }
);

await server.connect(new StdioServerTransport());