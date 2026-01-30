#!/usr/bin/env node
import {spawn} from "node:child_process";

function run() {
    const cmd = process.env.MCP_UVX || "uvx";
    const args = ["mcp-doc-generator"];

    const child = spawn(cmd, args, {
        stdio: ["pipe", "pipe", "pipe"],
        env: process.env,
    });

    process.stdin.pipe(child.stdin);
    child.stdout.pipe(process.stdout);
    child.stderr.pipe(process.stderr);

    child.on("exit", (code) => process.exit(code ?? 1));
    child.on("error", (err) => {
        process.stderr.write(
            `Failed to start uvx wrapper.\n` +
            `Command: ${cmd} ${args.join(" ")}\n` +
            `Error: ${String(err)}\n\n` +
            `Make sure 'uv' is installed and 'uvx' is on PATH.\n`
        );
        process.exit(1);
    });
}

run();