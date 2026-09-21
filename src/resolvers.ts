/**
 * Asset resolvers. The engine never calls an AI or search API itself: images
 * that need generation or lookup are handed to the resolver you pass in.
 * Node-only (filesystem + fetch).
 */

import fs from "node:fs";
import path from "node:path";
import { isSvg, isWebP, svgToPng, toPng, hasSharp } from "./raster.js";

export interface ImageRequest {
  /** What kind of asset the deck asked for. */
  kind: "image" | "logo" | "icon";
  /** For images: how the deck says to obtain it. */
  source?: "ai" | "search" | "upload";
  prompt?: string;
  searchTerm?: string;
  url?: string;
  path?: string;
  /** For icons: Iconify id like "mdi:home" */
  name?: string;
  /** Where the request sits in the deck, for messages. */
  location: string;
}

/** Return PNG/JPEG bytes, or null to skip the asset. */
export type ImageResolver = (req: ImageRequest) => Promise<Buffer | null>;

export interface ResolverContext {
  basePath: string;
  /** Allow reading images from the local filesystem via `path` (false for hosted services). */
  allowLocalFiles: boolean;
  network: boolean;
  timeoutMs: number;
  resolveImage?: ImageResolver;
  warn: (message: string) => void;
}

/** Local file relative to basePath. */
export function loadLocalImage(p: string, ctx: ResolverContext): Buffer | null {
  if (!ctx.allowLocalFiles) { ctx.warn(`local image paths are disabled here; use a url instead of path "${p}"`); return null; }
  const resolved = path.resolve(ctx.basePath, p);
  if (!resolved.startsWith(path.resolve(ctx.basePath) + path.sep) && resolved !== path.resolve(ctx.basePath)) { ctx.warn(`image path escapes basePath: ${p}`); return null; }
  if (!fs.existsSync(resolved)) { ctx.warn(`image not found: ${resolved}`); return null; }
  return fs.readFileSync(resolved);
}

/** Fetch a URL; converts SVG and WebP to PNG when sharp is available. */
export async function fetchImage(url: string, ctx: ResolverContext): Promise<Buffer | null> {
  if (!ctx.network) { ctx.warn(`network disabled; skipped image URL ${url}`); return null; }
  try {
    const res = await fetch(url, { signal: AbortSignal.timeout(ctx.timeoutMs) });
    if (!res.ok) { ctx.warn(`failed to fetch image (${res.status}): ${url}`); return null; }
    let buf = Buffer.from(await res.arrayBuffer());
    const ct = res.headers.get("content-type") || "";
    return normalizeImage(buf, url, ct.includes("svg"), ctx);
  } catch (e: any) {
    ctx.warn(`failed to fetch image: ${url} (${e?.message || e})`);
    return null;
  }
}

/** PowerPoint accepts PNG/JPEG/GIF/BMP/TIFF but not SVG (reliably) or WebP. */
export async function normalizeImage(buf: Buffer, hint: string, svgHint: boolean, ctx: ResolverContext): Promise<Buffer | null> {
  if (svgHint || isSvg(buf, hint)) {
    if (!(await hasSharp())) { ctx.warn(`sharp is not installed; skipped SVG image ${hint} (npm install sharp)`); return null; }
    try { return await svgToPng(buf); } catch (e: any) { ctx.warn(`SVG to PNG failed for ${hint}: ${e?.message || e}`); return null; }
  }
  if (isWebP(buf)) {
    if (!(await hasSharp())) { ctx.warn(`sharp is not installed; skipped WebP image ${hint} (npm install sharp)`); return null; }
    try { return await toPng(buf); } catch (e: any) { ctx.warn(`WebP to PNG failed for ${hint}: ${e?.message || e}`); return null; }
  }
  return buf;
}

/** Iconify icon SVG → recolored PNG. */
export async function resolveIcon(name: string | undefined, svg: string | undefined, color: string, ctx: ResolverContext, location: string): Promise<Buffer | null> {
  let markup = svg;
  if (!markup && ctx.resolveImage) {
    const custom = await ctx.resolveImage({ kind: "icon", name, location });
    if (custom) return custom;
  }
  if (!markup && name) {
    if (!ctx.network) { ctx.warn(`${location}: network disabled; icon "${name}" needs inline svg`); return null; }
    const m = /^([a-z0-9-]+):([a-z0-9-]+)$/.exec(name);
    if (!m) { ctx.warn(`${location}: icon name "${name}" is not an Iconify id like "mdi:home"`); return null; }
    try {
      const res = await fetch(`https://api.iconify.design/${m[1]}/${m[2]}.svg`, { signal: AbortSignal.timeout(ctx.timeoutMs) });
      if (!res.ok) { ctx.warn(`${location}: icon "${name}" not found on Iconify (${res.status})`); return null; }
      markup = await res.text();
    } catch (e: any) {
      ctx.warn(`${location}: icon fetch failed for "${name}": ${e?.message || e}`);
      return null;
    }
  }
  if (!markup) return null;
  if (!(await hasSharp())) { ctx.warn(`${location}: sharp is not installed; skipped icon (npm install sharp)`); return null; }
  let out = markup.replace(/currentColor/g, color);
  out = out.replace(/\bwidth="[^"]*"/i, 'width="512"').replace(/\bheight="[^"]*"/i, 'height="512"');
  if (!/\bwidth="/i.test(out)) out = out.replace(/<svg\b/i, '<svg width="512" height="512"');
  try { return await svgToPng(out); } catch (e: any) { ctx.warn(`${location}: icon SVG to PNG failed: ${e?.message || e}`); return null; }
}

/** Resolve an image content spec (or table image cell) to bytes. */
export async function resolveImageBytes(spec: { source?: "ai" | "search" | "upload"; prompt?: string; searchTerm?: string; url?: string; path?: string; buffer?: unknown; $search?: string }, kind: "image" | "logo", ctx: ResolverContext, location: string): Promise<Buffer | null> {
  if (spec.buffer && Buffer.isBuffer(spec.buffer)) return spec.buffer;
  if (spec.path) return loadLocalImage(spec.path, ctx);
  if (spec.url) return fetchImage(spec.url, ctx);
  const searchTerm = spec.searchTerm || spec.$search;
  if (spec.prompt || searchTerm) {
    if (ctx.resolveImage) {
      const buf = await ctx.resolveImage({ kind, source: spec.source || (spec.prompt ? "ai" : "search"), prompt: spec.prompt, searchTerm, location });
      if (buf) return normalizeImage(buf, location, false, ctx);
      ctx.warn(`${location}: resolver returned nothing for ${spec.prompt ? `prompt "${spec.prompt}"` : `search "${searchTerm}"`}`);
      return null;
    }
    ctx.warn(`${location}: image needs ${spec.prompt ? "generation" : "search"} (${spec.prompt || searchTerm}); pass options.resolveImage or supply url/path`);
    return null;
  }
  return null;
}
