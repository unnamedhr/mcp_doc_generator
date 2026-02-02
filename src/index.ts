#!/usr/bin/env node
import { spawn } from "node:child_process";

function run() {
  const uvx = process.env.MCP_UVX || "uvx";
  const DEFAULT_FROM =
    "git+https://github.com/unnamedhr/mcp_doc_generator.git@v0.1.0";

  const from = process.env.MCP_DOC_GENERATOR_FROM || DEFAULT_FROM;

  const args = ["--from", from, "mcp-doc-generator"];

  const child = spawn(uvx, args, {
    stdio: ["pipe", "pipe", "pipe"],
    env: process.env,
  });

  process.stdin.pipe(child.stdin);
  child.stdout.pipe(process.stdout);
  child.stderr.pipe(process.stderr);

  child.on("exit", (code) => process.exit(code ?? 1));
  child.on("error", (err) => {
    process.stderr.write(
      `Failed to start uvx.\nCommand: ${uvx} ${args.join(" ")}\nError: ${String(
        err
      )}\n\n` +
        `Make sure uv/uvx is installed and available on PATH.\n`
    );
    process.exit(1);
  });
}

run();