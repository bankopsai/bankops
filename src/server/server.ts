/**
 * bankops serve — a stateless HTTP render service. Same engine as the library,
 * so a hosted web app and a local agent produce identical files.
 *
 *   POST /render     { deck, options? }  → application/vnd.openxmlformats-officedocument.presentationml.presentation
 *   POST /validate   { deck }            → ValidationResult
 *   POST /annotate   { deck, style? }    → { annotated }
 *   POST /preview    { deck, slide, options? } → image/png of one slide (1-based; needs LibreOffice)
 *   GET  /schema                          → JSON Schema
 *   GET  /health                          → { ok, version }
 *
 * Only `options.style`, `themeTokens`, `network`, `zoneBorders` and `components` are
 * accepted over HTTP; file paths and reference PPTX are library-only.
 */

import http from "node:http";
import { renderDeck, type RenderOptions } from "../render.js";
import { validateDeck, deckJsonSchema, DeckValidationError } from "../schema.js";
import { annotateDeck } from "../annotate.js";
import { renderSlidePng } from "../preview.js";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const VERSION: string = (() => { try { return require("../../package.json").version; } catch { return "0.0.0"; } })();

export interface ServeOptions {
  port?: number;
  host?: string;
  /** Max request body in bytes (default 25 MB). */
  maxBodyBytes?: number;
  /** Supply an image resolver for the hosted deployment (AI images, logo search). */
  resolveImage?: RenderOptions["resolveImage"];
  /** Components available to every request by `$component` name. */
  components?: RenderOptions["components"];
  log?: (line: string) => void;
}

const HTTP_RENDER_OPTIONS = ["style", "themeTokens", "network", "zoneBorders", "components", "validate"] as const;

function json(res: http.ServerResponse, status: number, body: unknown): void {
  const text = JSON.stringify(body);
  res.writeHead(status, { "content-type": "application/json; charset=utf-8", "content-length": Buffer.byteLength(text) });
  res.end(text);
}

function readBody(req: http.IncomingMessage, max: number): Promise<string> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    let size = 0;
    req.on("data", (c: Buffer) => {
      size += c.length;
      if (size > max) { reject(new Error(`request body exceeds ${max} bytes`)); req.destroy(); return; }
      chunks.push(c);
    });
    req.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    req.on("error", reject);
  });
}

export function createHandler(opts: ServeOptions = {}) {
  const max = opts.maxBodyBytes || 25 * 1024 * 1024;
  const log = opts.log || (() => {});

  return async function handler(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    const url = new URL(req.url || "/", "http://localhost");
    const started = Date.now();
    try {
      if (req.method === "GET" && url.pathname === "/health") return json(res, 200, { ok: true, version: VERSION });
      if (req.method === "GET" && url.pathname === "/schema") return json(res, 200, deckJsonSchema());
      if (req.method !== "POST") return json(res, 405, { error: "method not allowed" });

      let body: any;
      try { body = JSON.parse(await readBody(req, max)); } catch (e: any) { return json(res, 400, { error: `invalid JSON body: ${e?.message || e}` }); }
      const deck = body?.deck ?? body;

      if (url.pathname === "/validate") return json(res, 200, validateDeck(deck));

      if (url.pathname === "/annotate") {
        const v = validateDeck(deck);
        if (!v.ok) return json(res, 422, { error: "invalid deck", errors: v.errors, warnings: v.warnings });
        return json(res, 200, annotateDeck(deck, body?.style));
      }

      if (url.pathname === "/render") {
        const options: RenderOptions = { network: true, resolveImage: opts.resolveImage, components: { ...(opts.components || {}) } };
        const reqOpts = body?.options || {};
        for (const k of HTTP_RENDER_OPTIONS) {
          if (reqOpts[k] !== undefined) {
            if (k === "components") Object.assign(options.components!, reqOpts.components);
            else (options as any)[k] = reqOpts[k];
          }
        }
        const r = await renderDeck(deck, options);
        res.writeHead(200, {
          "content-type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
          "content-length": r.buffer.length,
          "x-bankops-slides": String(r.slides),
          "x-bankops-warnings": String(r.warnings.length),
        });
        res.end(r.buffer);
        log(`render ${r.slides} slides ${r.buffer.length}B ${Date.now() - started}ms${r.warnings.length ? ` (${r.warnings.length} warnings)` : ""}`);
        return;
      }

      if (url.pathname === "/preview") {
        const slide = Number(body?.slide || 1);
        const options: RenderOptions = { network: true, resolveImage: opts.resolveImage, components: { ...(opts.components || {}) } };
        const reqOpts = body?.options || {};
        for (const k of HTTP_RENDER_OPTIONS) if (reqOpts[k] !== undefined && k !== "components") (options as any)[k] = reqOpts[k];
        const { png, warnings } = await renderSlidePng(deck, slide - 1, { ...options, width: Number(body?.width) || 960 });
        res.writeHead(200, { "content-type": "image/png", "content-length": png.length, "x-bankops-warnings": String(warnings.length) });
        res.end(png);
        log(`preview slide ${slide} ${png.length}B ${Date.now() - started}ms`);
        return;
      }

      return json(res, 404, { error: "not found", routes: ["POST /render", "POST /preview", "POST /validate", "POST /annotate", "GET /schema", "GET /health"] });
    } catch (e: any) {
      if (e instanceof DeckValidationError) return json(res, 422, { error: "invalid deck", errors: e.errors, warnings: e.warnings });
      log(`error ${url.pathname}: ${e?.message || e}`);
      return json(res, 500, { error: e?.message || String(e) });
    }
  };
}

export async function startServer(opts: ServeOptions = {}): Promise<http.Server> {
  const port = opts.port ?? 5490;
  const host = opts.host ?? "127.0.0.1";
  const log = opts.log || ((line: string) => process.stderr.write(line + "\n"));
  const server = http.createServer(createHandler({ ...opts, log }));
  await new Promise<void>((resolve) => server.listen(port, host, resolve));
  log(`bankops ${VERSION} render service listening on http://${host}:${port}`);
  return server;
}
