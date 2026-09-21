#!/usr/bin/env node
/**
 * bankops CLI
 *   bankops render <deck.json> [out.pptx] [--style corporate] [--theme tokens.json] [--base-path dir] [--no-network]
 *   bankops validate <deck.json>
 *   bankops annotate <deck.json>            # prints the deck with text budgets (_maxChars ...)
 *   bankops schema                          # prints the JSON Schema
 *   bankops skill [--install <dir>]         # prints the agent skill, or copies skill/ into <dir>
 *   bankops presets                         # lists style presets and layout presets
 */

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { renderFile } from "../render.js";
import { validateDeck, deckJsonSchema } from "../schema.js";
import { annotateDeck } from "../annotate.js";
import SlideStyle from "../style.js";
import SlideLayout from "../layout.js";
import { DeckValidationError } from "../schema.js";

const HELP = `bankops — deck JSON → investment-banking-grade PPTX

Usage:
  bankops render <deck.json> [out.pptx] [options]   Render a deck to PPTX
  bankops validate <deck.json>                      Check a deck; prints errors with fixes
  bankops annotate <deck.json>                      Print the deck with text budgets (_maxChars, _maxLines...)
  bankops schema                                    Print the JSON Schema for deck files
  bankops presets                                   List style presets and layout presets
  bankops skill [--install <dir>]                   Print the agent skill, or install it into <dir>
  bankops serve [--port 5490] [--host 127.0.0.1]    Run the render service (HTTP)

Render options:
  --style <preset|json>   Override deck.style (corporate | minimal | dark | warm, or a JSON object)
  --theme <tokens.json>   Brand tokens used as the base layer under deck.style
  --base-path <dir>       Directory relative image paths resolve against (default: the deck's directory)
  --no-network            Do not fetch image URLs or icons
  --zone-borders          Draw debug borders around every zone

Free for noncommercial use (PolyForm Noncommercial 1.0.0). Commercial use needs a license: https://bankops.ai/license
`;

function readJson(p: string): unknown {
  let raw: string;
  try { raw = fs.readFileSync(p, "utf8"); } catch { fail(`cannot read ${p}`); }
  try { return JSON.parse(raw!); } catch (e: any) { fail(`${p} is not valid JSON: ${e?.message || e}`); }
}

function fail(msg: string): never {
  process.stderr.write(`error: ${msg}\n`);
  process.exit(1);
}

function parseArgs(argv: string[]): { positional: string[]; flags: Record<string, string | boolean> } {
  const positional: string[] = [];
  const flags: Record<string, string | boolean> = {};
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a.startsWith("--")) {
      const key = a.slice(2);
      const next = argv[i + 1];
      if (key.startsWith("no-")) { flags[key.slice(3)] = false; continue; }
      if (next !== undefined && !next.startsWith("--")) { flags[key] = next; i++; } else { flags[key] = true; }
    } else {
      positional.push(a);
    }
  }
  return { positional, flags };
}

function printIssues(label: string, issues: { path: string; message: string; fix?: string }[]): void {
  for (const i of issues) process.stdout.write(`${label} ${i.path}: ${i.message}${i.fix ? `\n    fix: ${i.fix}` : ""}\n`);
}

async function main(): Promise<void> {
  const [cmd, ...rest] = process.argv.slice(2);
  const { positional, flags } = parseArgs(rest);

  switch (cmd) {
    case "render": {
      const input = positional[0];
      if (!input) fail("render needs a deck JSON path");
      const opts: any = {};
      if (typeof flags.style === "string") {
        opts.style = flags.style.trim().startsWith("{") ? JSON.parse(flags.style) : flags.style;
      }
      if (typeof flags.theme === "string") opts.themeTokens = readJson(flags.theme);
      if (typeof flags["base-path"] === "string") opts.basePath = path.resolve(flags["base-path"]);
      if (flags.network === false) opts.network = false;
      if (flags["zone-borders"]) opts.zoneBorders = true;
      opts.onWarning = (m: string) => process.stderr.write(`warning: ${m}\n`);
      try {
        const r = await renderFile(input, positional[1], opts);
        process.stdout.write(`${r.path}: ${r.slides} slide${r.slides === 1 ? "" : "s"}${r.warnings.length ? `, ${r.warnings.length} warning${r.warnings.length === 1 ? "" : "s"}` : ""}\n`);
      } catch (e: any) {
        if (e instanceof DeckValidationError) {
          printIssues("error", e.errors);
          process.exit(1);
        }
        fail(e?.message || String(e));
      }
      return;
    }
    case "validate": {
      const input = positional[0];
      if (!input) fail("validate needs a deck JSON path");
      const r = validateDeck(readJson(input));
      printIssues("error", r.errors);
      printIssues("warning", r.warnings);
      if (r.ok) {
        const n = (r.deck as any).slides.length;
        process.stdout.write(`ok: ${n} slide${n === 1 ? "" : "s"}${r.warnings.length ? `, ${r.warnings.length} warning${r.warnings.length === 1 ? "" : "s"}` : ""}\n`);
      } else {
        process.exit(1);
      }
      return;
    }
    case "annotate": {
      const input = positional[0];
      if (!input) fail("annotate needs a deck JSON path");
      const deck = readJson(input) as any;
      const themeTokens = typeof flags.theme === "string" ? (readJson(flags.theme) as Record<string, unknown>) : undefined;
      const style = themeTokens ? { ...themeTokens, ...(typeof deck.style === "object" ? deck.style : {}) } : undefined;
      const { annotated } = annotateDeck(deck, style);
      process.stdout.write(JSON.stringify(annotated, null, 2) + "\n");
      return;
    }
    case "schema": {
      process.stdout.write(JSON.stringify(deckJsonSchema(), null, 2) + "\n");
      return;
    }
    case "presets": {
      process.stdout.write("Style presets:\n");
      for (const name of Object.keys(SlideStyle.presets)) process.stdout.write(`  ${name}\n`);
      process.stdout.write("Layout presets (SlideLayout.preset(name)):\n");
      for (const name of Object.keys(SlideLayout.presets)) process.stdout.write(`  ${name}\n`);
      return;
    }
    case "skill": {
      const here = path.dirname(fileURLToPath(import.meta.url));
      // dist/cli/bankops.js → ../../skill ; src/cli/bankops.ts → ../../skill
      const skillDir = path.resolve(here, "..", "..", "skill");
      if (!fs.existsSync(skillDir)) fail(`skill folder not found at ${skillDir}`);
      if (typeof flags.install === "string") {
        const dest = path.resolve(flags.install, "bankops");
        fs.mkdirSync(dest, { recursive: true });
        fs.cpSync(skillDir, dest, { recursive: true });
        process.stdout.write(`installed skill to ${dest}\n`);
      } else {
        process.stdout.write(fs.readFileSync(path.join(skillDir, "SKILL.md"), "utf8"));
      }
      return;
    }
    case "serve": {
      const mod = await import("../server/server.js");
      const port = typeof flags.port === "string" ? parseInt(flags.port, 10) : 5490;
      const host = typeof flags.host === "string" ? flags.host : "127.0.0.1";
      await mod.startServer({ port, host });
      return;
    }
    case undefined:
    case "help":
    case "--help":
    case "-h":
      process.stdout.write(HELP);
      return;
    default:
      fail(`unknown command "${cmd}"\n\n${HELP}`);
  }
}

main().catch((e) => fail(e?.message || String(e)));
