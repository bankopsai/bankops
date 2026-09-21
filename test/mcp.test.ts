import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { createMcpServer, budgetReport } from "../src/mcp/server.js";
import { annotateDeck } from "../src/annotate.js";

const deck = {
  title: "MCP test",
  style: "corporate",
  slides: [{ title: "Overview", body: { direction: "row", children: [
    { span: 6, header: "Highlights", content: { type: "text", runs: [{ text: "Revenue up 12%", bullet: true }] } },
    { span: 6, content: { type: "chart", chartType: "pie", items: [{ name: "A", value: 1 }, { name: "B", value: 2 }] } },
  ] } }],
};

async function connect(dir: string) {
  const server = createMcpServer({ basePath: dir });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await server.connect(serverTransport);
  const client = new Client({ name: "test", version: "0.0.0" });
  await client.connect(clientTransport);
  return { client, close: async () => { await client.close(); await server.close(); } };
}

const textOf = (r: any) => r.content.find((c: any) => c.type === "text").text as string;

test("lists the six tools and serves the guide", async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "bankops-mcp-"));
  const { client, close } = await connect(dir);
  try {
    const tools = (await client.listTools()).tools.map((t) => t.name).sort();
    assert.deepEqual(tools, ["bankops_annotate", "bankops_guide", "bankops_presets", "bankops_preview", "bankops_render", "bankops_schema", "bankops_validate"]);
    const guide = textOf(await client.callTool({ name: "bankops_guide", arguments: {} }));
    assert.ok(guide.startsWith("---\nname: bankops"));
    const fmt = textOf(await client.callTool({ name: "bankops_guide", arguments: { section: "format" } }));
    assert.ok(fmt.includes("statGrid"));
  } finally { await close(); }
});

test("validate returns fixes, annotate returns a budget report, render writes a file", async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "bankops-mcp-"));
  const { client, close } = await connect(dir);
  try {
    const bad = await client.callTool({ name: "bankops_validate", arguments: { deck: { slides: [{ body: { content: { type: "text", fontSize: 12 } } }] } } });
    assert.equal(bad.isError, true);
    assert.match(JSON.parse(textOf(bad)).errors[0].fix, /Multiply/);

    const ok = JSON.parse(textOf(await client.callTool({ name: "bankops_validate", arguments: { deck } })));
    assert.equal(ok.ok, true);
    assert.equal(ok.slides, 1);

    const report = textOf(await client.callTool({ name: "bankops_annotate", arguments: { deck } }));
    assert.match(report, /slides\[0\]\.title: 8\/\d+ chars/);
    assert.match(report, /content \(text\): 14\/\d+ chars/);

    const rendered = JSON.parse(textOf(await client.callTool({ name: "bankops_render", arguments: { deck, output_path: "out/test.pptx" } })));
    assert.equal(rendered.ok, true);
    assert.equal(rendered.slides, 1);
    assert.ok(fs.statSync(path.join(dir, "out", "test.pptx")).size > 10000);
  } finally { await close(); fs.rmSync(dir, { recursive: true, force: true }); }
});

test("budget report flags overflow", () => {
  const long = { slides: [{ title: "A very long title that should not fit on one slide title line at 28pt", body: { content: { type: "text", text: "x" } } }] };
  const { annotated } = annotateDeck(long as any);
  assert.match(budgetReport(annotated), /title: \d+\/\d+ chars  OVERFLOW/);
});
