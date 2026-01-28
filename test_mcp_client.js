import { spawn } from "node:child_process";
import assert from "node:assert";
import path from "node:path";
import fs from "node:fs";

const SERVER_PATH = path.resolve("dist/index.js");
const TEMPLATE_XLSX = path.resolve("src/test/test-template.xlsx");

function makeJsonLineReader(stream) {
  let buf = "";
  const queue = [];
  const waiters = [];

  stream.on("data", (chunk) => {
    buf += chunk.toString("utf8");
    while (true) {
      const idx = buf.indexOf("\n");
      if (idx < 0) break;
      const line = buf.slice(0, idx).trim();
      buf = buf.slice(idx + 1);
      if (!line) continue;

      let msg;
      try {
        msg = JSON.parse(line);
      } catch {
        continue;
      }

      if (waiters.length) waiters.shift()(msg);
      else queue.push(msg);
    }
  });

  return async function readMessage() {
    if (queue.length) return queue.shift();
    return await new Promise((resolve) => waiters.push(resolve));
  };
}

function startServer() {
  const proc = spawn("node", [SERVER_PATH], {
    stdio: ["pipe", "pipe", "pipe"],
  });
  const read = makeJsonLineReader(proc.stdout);

  proc.stderr.on("data", (d) => {
    process.stderr.write(d.toString());
  });

  return { proc, read };
}

function send(proc, msg) {
  proc.stdin.write(JSON.stringify(msg) + "\n");
}

async function rpc(read, proc, method, params) {
  const id = Math.floor(Math.random() * 1e9);

  send(proc, {
    jsonrpc: "2.0",
    id,
    method,
    params,
  });

  while (true) {
    const msg = await read();
    if (msg && msg.id === id) return msg;
  }
}

function decodeMagic(b64) {
  const buf = Buffer.from(b64, "base64");
  return buf.slice(0, 4).toString("utf8");
}

(async () => {
  const { proc, read } = startServer();

  // Tools/list
  const listRes = await rpc(read, proc, "tools/list", {});
  assert(listRes.result?.tools, "tools/list missing result.tools");

  const toolNames = listRes.result.tools.map((t) => t.name);
  const expectedTools = [
    "excel_generator",
    "pdf_generator",
    "word_generator",
    "template_generator",
  ];

  for (const t of expectedTools) {
    assert(toolNames.includes(t), `Missing tool: ${t}`);
  }
  console.log("✅ tools/list OK");

  // Excel generator
  const excelArgs = {
    data: {
      title: "MCP Excel Test",
      sheet_name: "Sheet1",
      headers: ["Name", "Score"],
      data_rows: [
        ["Ana", 10],
        ["Marko", 12],
      ],
      include_freeze_panes: true,
      include_autofilter: true,
    },
  };

  // PDF generator
  const pdfArgs = {
  data: {
    title: "MCP PDF Test",
    body: "This is a test PDF generated via MCP.",
    report_data: {
      sections: [
        {
          title: "Summary",
          text: "One small section to confirm PDF generation works.",
        },
        {
          title: "Table",
          table_data: [
            ["Item", "Qty"],
            ["Espresso", "2"],
            ["Croissant", "3"],
          ],
        },
      ],
    },
    styling_config: {
      document: {
        page_size: "letter",
        orientation: "portrait",
        margin: 0.75,
        color_profile: "RGB",
      },
      report_title: {
        font_size: 18,
        color: "#1F3A93",
        bold: true,
      },
      body: {
        font_size: 11,
        alignment: "justify",
        color: "#333333",
      },
      table: {
        header_bg: "#2C5282",
        header_text_color: "#FFFFFF",
        header_font_size: 11,
        font_size: 10,
        row_colors: ["#FFFFFF", "#F7FAFC"],
        grid_color: "#CCCCCC",
        cell_padding: 6,
      },
    },
  },
};


  // Word generator
  const wordArgs = {
  data: {
    title: "MCP Word Test",
    subtitle: "Generated through MCP",
    content: {
      sections: [
        { type: "heading1", text: "Intro" },
        { type: "paragraph", text: "This document confirms DOCX generation works." },
        { type: "heading2", text: "Data" },
        {
          type: "table",
          table_data: [
            ["City", "Country"],
            ["Zagreb", "Croatia"],
            ["Bangkok", "Thailand"],
          ],
        },

        { type: "paragraph", text: "End of report." },
      ],
    },
  },
};

  // Template generator
  const templateB64 = fs.readFileSync(TEMPLATE_XLSX).toString("base64");
  const templateArgs = {
    data: {
      base64_template: templateB64,
      data: {
        project_name: "MCP Rollout",
        project_manager: "Fatima",
        start_date: "2026-01-01",
        end_date: "2026-03-01",
        generated_date: "2026-01-28",
        project_status: "On Track",
        budget: "€12,500",
        total_team_size: 6,
        active_members: 5,
        on_leave_members: 1,
        completion_rate: 42,
        table_data: [
          { name: "Ana", role: "Engineer", department: "Platform", status: "Active" },
          { name: "Marko", role: "Designer", department: "UX", status: "Active" },
          { name: "Iva", role: "QA", department: "Testing", status: "On Leave" },
        ],
      },
      table_placeholder: "table_data",
    },
  };

  const cases = [
    ["excel_generator", excelArgs, "PK"],
    ["pdf_generator", pdfArgs, "%PDF"],
    ["word_generator", wordArgs, "PK"],
    ["template_generator", templateArgs, "PK"],
  ];

  for (const [tool, args, expectedMagic] of cases) {
    const res = await rpc(read, proc, "tools/call", {
      name: tool,
      arguments: args,
    });

    assert(res.result?.content?.length, `${tool}: missing result.content`);
    const item = res.result.content[0];
    assert(item.type === "text", `${tool}: expected text output`);
    assert(typeof item.text === "string", `${tool}: text not string`);

    let parsed;
    try {
      parsed = JSON.parse(item.text);
    } catch (e) {
      console.error(`\n--- ${tool} raw output (NOT JSON) ---\n${item.text}\n--- end ---\n`);
      throw e;
    }

    assert(parsed.type === "file_base64", `${tool}: expected parsed.type=file_base64`);
    assert(typeof parsed.base64 === "string" && parsed.base64.length > 100, `${tool}: base64 missing/too short`);
    assert(typeof parsed.filename === "string" && parsed.filename.length > 0, `${tool}: filename missing`);
    assert(typeof parsed.mime_type === "string" && parsed.mime_type.length > 0, `${tool}: mime_type missing`);

    const magic = decodeMagic(parsed.base64);
    assert(
      magic.startsWith(expectedMagic),
      `${tool}: magic bytes mismatch (got "${magic}", expected "${expectedMagic}")`
    );

    console.log(`✅ ${tool} OK -> ${parsed.filename}`);
  }

  proc.kill();
})().catch((err) => {
  console.error("❌ MCP smoke test failed:");
  console.error(err);
  process.exit(1);
});
