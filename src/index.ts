#!/usr/bin/env node
import { spawn } from "node:child_process";

function run() {
  const uvx = process.env.MCP_UVX || "uvx";
  const from = process.env.MCP_DOC_GENERATOR_FROM;

  if (!from) {
    process.stderr.write(
      "Missing MCP_DOC_GENERATOR_FROM.\n" +
        "Set it to your GitLab source, e.g.\n" +
        "  git+https://oauth2:${GITLAB_TOKEN}@gitlab.scheer-group.com/fatima.zivkovic/mcp_doc_generator.git@v0.1.3\n"
    );
    process.exit(1);
  }

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
      `Failed to start uvx.\nCommand: ${uvx} ${args.join(" ")}\nError: ${String(err)}\n`
    );
    process.exit(1);
  });
}

run();