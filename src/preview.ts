/**
 * Slide previews: render one slide to PNG through LibreOffice (`soffice`), when installed.
 * Node-only, optional. Used by the MCP server and the HTTP service so agents can look
 * at what they built before handing it to a human.
 */

import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { renderDeck, type RenderOptions } from "./render.js";
import type { DeckJson } from "./types.js";

const execFileAsync = promisify(execFile);

let sofficeChecked: Promise<string | null> | null = null;

/** Path to the soffice binary, or null when LibreOffice is not installed. */
export function findSoffice(): Promise<string | null> {
  if (!sofficeChecked) {
    sofficeChecked = (async () => {
      const candidates = [process.env.BANKOPS_SOFFICE, "soffice", "libreoffice", "/usr/bin/soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"].filter(Boolean) as string[];
      for (const c of candidates) {
        try { await execFileAsync(c, ["--version"], { timeout: 20000 }); return c; } catch { /* try next */ }
      }
      return null;
    })();
  }
  return sofficeChecked;
}

// LibreOffice cannot run concurrently with one profile: serialize conversions.
let lock: Promise<void> = Promise.resolve();
function withLock<T>(fn: () => Promise<T>): Promise<T> {
  const prev = lock;
  let release!: () => void;
  lock = new Promise<void>((r) => { release = r; });
  return prev.then(fn).finally(() => release());
}

export interface PreviewOptions extends RenderOptions {
  /** Pixel width of the PNG (default 960; height follows the slide aspect ratio). */
  width?: number;
}

/** Convert PPTX bytes to a PNG of the first slide. */
export async function pptxToPng(pptx: Buffer, width = 960, height = 720): Promise<Buffer> {
  const soffice = await findSoffice();
  if (!soffice) throw new Error("LibreOffice (soffice) is not installed, so previews are unavailable. Fix: install LibreOffice, or open the .pptx in PowerPoint.");
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "bankops-preview-"));
  try {
    const src = path.join(dir, "slide.pptx");
    await fs.writeFile(src, pptx);
    const filter = `png:impress_png_Export:{"PixelWidth":{"type":"long","value":"${width}"},"PixelHeight":{"type":"long","value":"${height}"}}`;
    await withLock(() => execFileAsync(soffice, ["--headless", "--convert-to", filter, src, "--outdir", dir], { timeout: 60000 }));
    return await fs.readFile(path.join(dir, "slide.png"));
  } finally {
    await fs.rm(dir, { recursive: true, force: true }).catch(() => {});
  }
}

/** Render one slide (0-based index) of a deck to PNG. */
export async function renderSlidePng(deck: DeckJson, slideIndex: number, options: PreviewOptions = {}): Promise<{ png: Buffer; warnings: string[] }> {
  const slides = deck.slides || [];
  if (slideIndex < 0 || slideIndex >= slides.length) throw new Error(`slide ${slideIndex + 1} does not exist; the deck has ${slides.length} slide${slides.length === 1 ? "" : "s"}`);
  const single: DeckJson = { ...deck, slides: [slides[slideIndex]] };
  const { width = 960, ...renderOptions } = options;
  const r = await renderDeck(single, renderOptions);
  const w = deck.slideSize?.width || 10;
  const h = deck.slideSize?.height || 7.5;
  const png = await pptxToPng(r.buffer, width, Math.round((width * h) / w));
  return { png, warnings: r.warnings };
}
