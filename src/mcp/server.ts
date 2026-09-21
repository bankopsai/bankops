/**
 * bankops MCP server: the tools an agent needs to go from an idea to a .pptx.
 *
 * Rules: every description leads with its decision criteria; results are compact;
 * errors are { path, message, fix } so the agent applies the fix instead of retrying.
 */

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { validateDeck, deckJsonSchema, DeckValidationError, CONTENT_TYPES, STYLE_PRESETS } from "../schema.js";
import { annotateDeck } from "../annotate.js";
import { renderDeck, type RenderOptions } from "../render.js";
import { renderSlidePng, findSoffice } from "../preview.js";
import SlideLayout from "../layout.js";
import SlideStyle from "../style.js";
import type { DeckJson } from "../types.js";

const require = createRequire(import.meta.url);
const VERSION: string = (() => { try { return require("../../package.json").version; } catch { return "0.0.0"; } })();

type Content = { type: "text"; text: string } | { type: "image"; data: string; mimeType: string };
type ToolResult = { content: Content[]; isError?: boolean };

const text = (s: string, isError = false): ToolResult => ({ content: [{ type: "text", text: s }], ...(isError ? { isError } : {}) });
const json = (body: unknown, isError = false): ToolResult => text(JSON.stringify(body), isError);

async function guard(fn: () => Promise<ToolResult> | ToolResult): Promise<ToolResult> {
  try {
    return await fn();
  } catch (e: any) {
    if (e instanceof DeckValidationError) return json({ ok: false, errors: e.errors, warnings: e.warnings }, true);
    return json({ ok: false, errors: [{ path: "deck", message: e?.message || String(e) }] }, true);
  }
}

const skillDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "..", "skill");
function readSkill(section: string): string {
  const file = section === "skill" ? path.join(skillDir, "SKILL.md") : path.join(skillDir, "reference", `${section}.md`);
  return fs.readFileSync(file, "utf8");
}

const deckArg = z.record(z.string(), z.unknown()).describe("The deck JSON: { style?, slides: [...] }");
const styleArg = z.union([z.string(), z.record(z.string(), z.unknown())]).optional().describe('Overrides deck.style: a preset name ("corporate" | "minimal" | "dark" | "warm") or a style object');
const themeArg = z.record(z.string(), z.unknown()).optional().describe("Brand tokens used as the base layer under deck.style");

/** Compact budget report: one line per text-bearing zone, flagging overflow. */
export function budgetReport(annotated: any): string {
  const lines: string[] = [];
  const slides = annotated.slides || [];
  const count = (c: any): number => {
    if (Array.isArray(c.runs)) return c.runs.reduce((n: number, r: any) => n + String(r.text || "").length, 0);
    return String(c.text || "").length;
  };
  slides.forEach((slide: any, i: number) => {
    const p = `slides[${i}]`;
    if (slide.title) lines.push(`${p}.title: ${slide.title.length}/${slide._titleMaxChars} chars${slide.title.length > slide._titleMaxChars ? "  OVERFLOW" : ""}`);
    const walk = (node: any, np: string) => {
      if (!node || typeof node !== "object") return;
      if (node.header && node._headerMaxChars != null) lines.push(`${np}.header: ${node.header.length}/${node._headerMaxChars} chars${node.header.length > node._headerMaxChars ? "  OVERFLOW" : ""}`);
      const c = node.content;
      if (c && c.type === "text" && c._maxChars != null) {
        const used = count(c);
        lines.push(`${np}.content (text): ${used}/${c._maxChars} chars, ${c._charsPerLine} per line, ${c._maxLines} lines${used > c._maxChars ? "  OVERFLOW" : ""}`);
      }
      if (c && c.type === "table" && c.data && c.data._maxRows != null) {
        const rows = Array.isArray(c.data.rows) ? c.data.rows.length : 0;
        lines.push(`${np}.content (table): ${rows}/${c.data._maxRows} rows, cell budgets ${JSON.stringify(c.data._cellMaxChars)} chars${rows > c.data._maxRows ? "  OVERFLOW" : ""}`);
      }
      if (Array.isArray(node.children)) node.children.forEach((ch: any, k: number) => walk(ch, `${np}.children[${k}]`));
    };
    if (slide.body) walk(slide.body, `${p}.body`);
  });
  return lines.length ? lines.join("\n") : "no text zones found";
}

export interface McpServerOptions {
  /** Directory that relative output paths and image paths resolve against (default cwd). */
  basePath?: string;
  /** Image resolver for AI/search images (hosted deployments). */
  resolveImage?: RenderOptions["resolveImage"];
  /** Called after a successful render in hosted mode to publish the deck; returns a URL. */
  share?: (deck: DeckJson, pptx: Buffer) => Promise<string>;
  /** Write the .pptx to disk (default true). Hosted deployments set false and rely on `share`. */
  writeFiles?: boolean;
  /** Allow `path` images from disk (default true). Hosted deployments set false. */
  allowLocalFiles?: boolean;
  /** Extra tools to register (hosted deployments add search/fetch). */
  extend?: (server: McpServer) => void;
}

export function createMcpServer(opts: McpServerOptions = {}): McpServer {
  const server = new McpServer({ name: "bankops", version: VERSION }, {
    instructions: "bankops renders deck JSON into investment-banking-grade PowerPoint files. Start with bankops_guide, write the deck JSON, bankops_validate it, fit copy to the budgets from bankops_annotate, then bankops_render and give the user the file or link.",
  });
  const basePath = opts.basePath || process.cwd();
  const writeFiles = opts.writeFiles !== false;
  const allowLocalFiles = opts.allowLocalFiles !== false;

  server.registerTool("bankops_guide", {
    title: "How to build a deck with bankops",
    description: "Read this first when the deliverable is a PowerPoint deck a human will open in PowerPoint (pitch, company profile, market overview, comps, credentials). Returns the bankops skill: the deck JSON format, workflow, IB conventions and layout recipes. Not for web pages, PDFs, spreadsheets or standalone chart images.",
    inputSchema: { section: z.enum(["skill", "format", "recipes", "errors"]).optional().describe("skill (default): workflow and rules; format: every field; recipes: IB layouts; errors: what each error means") },
  }, async ({ section }) => guard(() => text(readSkill(section || "skill"))));

  server.registerTool("bankops_schema", {
    title: "Deck JSON Schema",
    description: "The JSON Schema (draft 2020-12) of the deck format, for strict validation or tool definitions. Prefer bankops_guide with section=format for a readable reference.",
    inputSchema: {},
  }, async () => guard(() => json(deckJsonSchema())));

  server.registerTool("bankops_presets", {
    title: "Presets and content types",
    description: "Lists the style presets, layout presets and content types the engine supports.",
    inputSchema: {},
  }, async () => guard(() => json({
    stylePresets: STYLE_PRESETS,
    layoutPresets: Object.fromEntries(Object.entries(SlideLayout.presets).map(([k, v]) => [k, v.body])),
    contentTypes: CONTENT_TYPES,
    styleDefaults: SlideStyle.defaults,
  })));

  server.registerTool("bankops_validate", {
    title: "Validate a deck",
    description: "Check deck JSON before rendering. Returns { ok, errors, warnings }; each error has path, message and usually a fix. Apply every fix, then call bankops_annotate.",
    inputSchema: { deck: deckArg },
  }, async ({ deck }) => guard(() => {
    const r = validateDeck(deck);
    return json({ ok: r.ok, slides: r.ok ? (r.deck as DeckJson).slides.length : undefined, errors: r.errors, warnings: r.warnings }, !r.ok);
  }));

  server.registerTool("bankops_annotate", {
    title: "Text budgets per zone",
    description: "After validation and before rendering: returns how many characters fit in every title, header, text zone and table (and how many the deck currently uses), flagging OVERFLOW. Shorten copy, split slides or rebalance spans until nothing overflows. Set full=true to get the whole annotated deck JSON instead of the compact report.",
    inputSchema: { deck: deckArg, style: themeArg.describe("Optional theme tokens merged under deck.style when computing fonts"), full: z.boolean().optional() },
  }, async ({ deck, style, full }) => guard(() => {
    const v = validateDeck(deck);
    if (!v.ok) return json({ ok: false, errors: v.errors, warnings: v.warnings }, true);
    const { annotated } = annotateDeck(deck as unknown as DeckJson, style ? { ...(style as object), ...(typeof (deck as any).style === "object" ? (deck as any).style : {}) } : undefined);
    return full ? json(annotated) : text(budgetReport(annotated));
  }));

  server.registerTool("bankops_render", {
    title: "Render to PPTX",
    description: writeFiles
      ? "Render validated deck JSON to a .pptx file. Returns the file path, slide count and any warnings (missing images, skipped icons). Call bankops_validate and bankops_annotate first; a deck with validation errors is rejected with the same { path, message, fix } list."
      : "Render validated deck JSON and publish it. Returns a url the user opens to view and download the PowerPoint, plus slide count and warnings. Give the url to the user verbatim. Call bankops_validate and bankops_annotate first; a deck with validation errors is rejected with the same { path, message, fix } list.",
    inputSchema: {
      deck: deckArg,
      output_path: z.string().optional().describe("Where to write the .pptx (default: <title>.pptx in the working directory)"),
      style: styleArg,
      theme_tokens: themeArg,
    },
  }, async ({ deck, output_path, style, theme_tokens }) => guard(async () => {
    const r = await renderDeck(deck, { basePath, allowLocalFiles, style: style as any, themeTokens: theme_tokens as any, resolveImage: opts.resolveImage });
    const title = String((deck as any).title || "deck").replace(/[^a-z0-9_\-. ]/gi, "_").trim() || "deck";
    let out: string | null = null;
    if (writeFiles) {
      out = path.resolve(basePath, output_path || `${title}.pptx`);
      fs.mkdirSync(path.dirname(out), { recursive: true });
      fs.writeFileSync(out, r.buffer);
    }
    const url = opts.share ? await opts.share(deck as unknown as DeckJson, r.buffer) : null;
    const deliverable = url ? "give the user the url" : "hand the file to the user";
    return json({ ok: true, path: out, url, slides: r.slides, bytes: r.buffer.length, warnings: r.warnings, next: r.warnings.length ? `fix the warnings or accept them, then ${deliverable}` : deliverable });
  }));

  server.registerTool("bankops_preview", {
    title: "Preview one slide as an image",
    description: "Look at a rendered slide before handing over the deck: returns a PNG of one slide (1-based) so you can check overflow, empty zones and layout balance. Needs LibreOffice on the machine; the error says so when it is missing.",
    inputSchema: { deck: deckArg, slide: z.number().int().min(1).describe("1-based slide number"), style: styleArg, theme_tokens: themeArg, width: z.number().int().min(320).max(1920).optional() },
  }, async ({ deck, slide, style, theme_tokens, width }) => guard(async () => {
    const soffice = await findSoffice();
    if (!soffice) return json({ ok: false, errors: [{ path: "preview", message: "LibreOffice (soffice) is not installed", fix: "install LibreOffice, or skip previews and rely on bankops_annotate budgets" }] }, true);
    const { png, warnings } = await renderSlidePng(deck as unknown as DeckJson, slide - 1, { basePath, allowLocalFiles, style: style as any, themeTokens: theme_tokens as any, resolveImage: opts.resolveImage, width });
    return { content: [{ type: "image", data: png.toString("base64"), mimeType: "image/png" }, { type: "text", text: JSON.stringify({ ok: true, slide, warnings }) }] };
  }));

  opts.extend?.(server);
  return server;
}

/** The skill sections, for hosted search/fetch tools. */
export function skillSections(): { id: string; title: string; text: string }[] {
  return [
    { id: "skill", title: "bankops skill: workflow and rules", text: readSkill("skill") },
    { id: "format", title: "Deck format reference", text: readSkill("format") },
    { id: "recipes", title: "Investment-banking layout recipes", text: readSkill("recipes") },
    { id: "errors", title: "Validation errors and warnings", text: readSkill("errors") },
  ];
}
