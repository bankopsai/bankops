#!/usr/bin/env node
/**
 * bankops MCP server over stdio (local mode).
 *
 *   npx -y bankops mcp
 *   node --import tsx src/mcp/main.ts
 *
 * Only JSON-RPC may go to stdout; diagnostics go to stderr.
 */
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { createMcpServer } from "./server.js";

export async function main(): Promise<void> {
  const server = createMcpServer({ basePath: process.env.BANKOPS_DIR || process.cwd() });
  process.stderr.write(`[bankops] MCP server on stdio; files are written under ${process.env.BANKOPS_DIR || process.cwd()}\n`);
  await server.connect(new StdioServerTransport());
}

if (process.argv[1] && /main\.(ts|js)$/.test(process.argv[1])) {
  main().catch((e) => { process.stderr.write(`[bankops] ${e?.message || e}\n`); process.exit(1); });
}
