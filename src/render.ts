/**
 * renderPptx — deck JSON → PPTX bytes. The one function most callers need.
 * Node-only. Validates, resolves assets through pluggable resolvers, lays out
 * every slide with SlideLayout and writes OOXML through SlideBuilder.
 */

import fs from "node:fs";
import path from "node:path";
import SlideBuilder from "./slide-builder.js";
import SlideStyle, { type StyleOverrides } from "./style.js";
import { mergeStyleWithTheme } from "./resolve-style.js";
import { validateDeck, DeckValidationError, type DeckIssue } from "./schema.js";
import { resolveIcon, resolveImageBytes, type ImageResolver, type ResolverContext } from "./resolvers.js";
import type { DeckJson, GridNode, SlideSpec } from "./types.js";

export type ComponentFactory = GridNode | ((data: Record<string, unknown>) => GridNode);

export interface RenderOptions {
  /** Overrides deck.style (preset name or style overrides). */
  style?: string | StyleOverrides;
  /** Organization/brand tokens: base layer under deck.style. */
  themeTokens?: StyleOverrides;
  /** Directory that relative image paths resolve against. Default: process.cwd(). */
  basePath?: string;
  /** Allow fetching image URLs and Iconify icons. Default true. */
  network?: boolean;
  /** Allow `path` images read from disk under basePath. Default true; hosted services set false. */
  allowLocalFiles?: boolean;
  /** Per-request timeout for network fetches, ms. Default 15000. */
  timeoutMs?: number;
  /** Supplies bytes for images that need generation or search, and for logos/icons. */
  resolveImage?: ImageResolver;
  /** Named reusable nodes referenced by `$component`. */
  components?: Record<string, ComponentFactory>;
  /** Path to a reference .pptx whose theme/fonts/colors are analyzed and reused. */
  referencePptx?: string;
  /** Draw debug borders around every zone. */
  zoneBorders?: boolean;
  /** Validate before rendering (default true). Invalid decks throw DeckValidationError. */
  validate?: boolean;
  /** Called for each non-fatal problem (missing image, unknown key...). */
  onWarning?: (message: string) => void;
}

export interface RenderResult {
  buffer: Buffer;
  slides: number;
  warnings: string[];
  /** Lint warnings from validation (unknown keys, span sums). */
  lint: DeckIssue[];
}

const LAYOUT_PROPS = ["span", "header", "subheader", "subfooter", "background", "border", "margin", "padding", "id", "_task", "_sources"];

/** Walk a body tree: expand $component, resolve images/icons to buffers. */
async function resolveNode(node: any, ctx: ResolverContext, components: Record<string, ComponentFactory>, location: string): Promise<any> {
  if (!node || typeof node !== "object") return node;
  if (Array.isArray(node)) {
    const out: any[] = [];
    for (let i = 0; i < node.length; i++) out.push(await resolveNode(node[i], ctx, components, `${location}[${i}]`));
    return out;
  }

  let out: any = { ...node };

  if (out.$component) {
    const name = out.$component as string;
    const factory = components[name];
    if (!factory) {
      const known = Object.keys(components);
      throw new Error(`${location}: unknown $component "${name}"${known.length ? `; known components: ${known.join(", ")}` : "; pass options.components"}`);
    }
    const data: Record<string, unknown> = {};
    const layoutProps: Record<string, unknown> = {};
    for (const k of Object.keys(out)) {
      if (k === "$component") continue;
      if (LAYOUT_PROPS.includes(k)) layoutProps[k] = out[k]; else data[k] = out[k];
    }
    const expanded = typeof factory === "function" ? factory(data) : JSON.parse(JSON.stringify(factory));
    out = { ...expanded, ...layoutProps };
  }

  const c = out.content;
  if (c && typeof c === "object") {
    if (c.type === "icon") {
      const color = c.color || "#333333";
      const sizePct = c.size || 80;
      const buf = await resolveIcon(c.name, c.svg, color, ctx, `${location}.content`);
      if (buf) {
        const iconAnchor = c.align === "top" ? "t" : c.align === "bottom" ? "b" : "ctr";
        out.content = { type: "image", buffer: buf, width: sizePct + "%", height: sizePct + "%", objectFit: c.objectFit || "contain", anchor: iconAnchor };
      } else {
        out.content = undefined;
      }
    } else if (c.type === "image") {
      const buf = await resolveImageBytes(c, "image", ctx, `${location}.content`);
      const img: any = { ...c };
      delete img.path;
      if (buf) img.buffer = buf;
      out.content = img;
    } else if (c.type === "table" && c.data && Array.isArray(c.data.rows)) {
      const rows: any[] = [];
      for (let r = 0; r < c.data.rows.length; r++) {
        const row = c.data.rows[r];
        const resolvedRow: any[] = [];
        for (let ci = 0; ci < row.length; ci++) {
          const cell = row[ci];
          if (cell && typeof cell === "object" && cell.type === "image") {
            const buf = await resolveImageBytes(cell, "logo", ctx, `${location}.content.data.rows[${r}][${ci}]`);
            const img: any = { ...cell };
            delete img.path;
            if (buf) img.buffer = buf;
            resolvedRow.push(img);
          } else {
            resolvedRow.push(cell);
          }
        }
        rows.push(resolvedRow);
      }
      out.content = { ...c, data: { ...c.data, rows } };
    }
  }

  if (out.children) {
    const kids: any[] = [];
    for (let i = 0; i < out.children.length; i++) {
      if (!out.children[i]) continue;
      kids.push(await resolveNode(out.children[i], ctx, components, `${location}.children[${i}]`));
    }
    out.children = kids;
  }
  return out;
}

/** Render a deck and return the PPTX bytes plus warnings. */
export async function renderDeck(input: DeckJson | unknown, options: RenderOptions = {}): Promise<RenderResult> {
  const warnings: string[] = [];
  const warn = (m: string) => { warnings.push(m); options.onWarning?.(m); };

  let deck: DeckJson;
  let lint: DeckIssue[] = [];
  if (options.validate !== false) {
    const v = validateDeck(input);
    lint = v.warnings;
    for (const w of v.warnings) warn(`${w.path}: ${w.message}${w.fix ? ` (${w.fix})` : ""}`);
    if (!v.ok) throw new DeckValidationError(v.errors, v.warnings);
    deck = v.deck!;
  } else {
    deck = input as DeckJson;
  }

  const styleSource = options.style !== undefined ? options.style : deck.style;
  const style: SlideStyle = mergeStyleWithTheme(styleSource || {}, options.themeTokens);

  const builderOpts: any = { style, zoneBorders: !!options.zoneBorders };
  if (deck.slideSize && (deck.slideSize.width || deck.slideSize.height)) builderOpts.slideSize = deck.slideSize;

  let builder: SlideBuilder;
  if (options.referencePptx) {
    const refPath = path.resolve(options.basePath || process.cwd(), options.referencePptx);
    if (!fs.existsSync(refPath)) throw new Error(`referencePptx not found: ${refPath}`);
    builder = await SlideBuilder.fromReference(refPath, builderOpts);
  } else {
    builder = new SlideBuilder(builderOpts);
  }

  const ctx: ResolverContext = {
    basePath: options.basePath || process.cwd(),
    allowLocalFiles: options.allowLocalFiles !== false,
    network: options.network !== false,
    timeoutMs: options.timeoutMs || 15000,
    resolveImage: options.resolveImage,
    warn,
  };
  const components = options.components || {};

  const slides = deck.slides || [];
  for (let i = 0; i < slides.length; i++) {
    const slideDef = slides[i];
    const spec: SlideSpec = {};
    if (slideDef.title) spec.title = slideDef.title;
    if (slideDef.subtitle) spec.subtitle = slideDef.subtitle;
    if (slideDef.sectionLabel) spec.sectionLabel = slideDef.sectionLabel;
    if (slideDef.footer) spec.footer = slideDef.footer;
    if (slideDef.margin != null) spec.margin = slideDef.margin;
    if (slideDef.body) spec.body = await resolveNode(slideDef.body, ctx, components, `slides[${i}].body`);
    await builder.addGridSlide(spec);
  }

  const buffer: Buffer = await builder.toBuffer();
  return { buffer, slides: builder.pres.slides.length, warnings, lint };
}

/** Render a deck to PPTX bytes. */
export async function renderPptx(deck: DeckJson | unknown, options: RenderOptions = {}): Promise<Buffer> {
  return (await renderDeck(deck, options)).buffer;
}

/** Render a deck to a .pptx file. */
export async function renderPptxToFile(deck: DeckJson | unknown, outputPath: string, options: RenderOptions = {}): Promise<RenderResult & { path: string }> {
  const result = await renderDeck(deck, options);
  fs.mkdirSync(path.dirname(path.resolve(outputPath)), { recursive: true });
  fs.writeFileSync(outputPath, result.buffer);
  return { ...result, path: outputPath };
}

/** Render a deck JSON file to a .pptx file. Relative image paths resolve against the JSON's directory. */
export async function renderFile(jsonPath: string, outputPath?: string, options: RenderOptions = {}): Promise<RenderResult & { path: string }> {
  const raw = fs.readFileSync(jsonPath, "utf8");
  let deck: unknown;
  try { deck = JSON.parse(raw); } catch (e: any) { throw new Error(`${jsonPath} is not valid JSON: ${e?.message || e}`); }
  const out = outputPath || jsonPath.replace(/\.json$/i, "") + ".pptx";
  return renderPptxToFile(deck, out, { basePath: path.dirname(path.resolve(jsonPath)), ...options });
}
