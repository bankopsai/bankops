/**
 * Legacy content shapes still found in stored decks and older agent output.
 * Isomorphic; used by the PPTX builder, normalizeDeck and previews.
 */

import type { TextContent, TextRun } from "./types.js";

/** "single" | "double" | "1.5" | 1.2 → a numeric line-spacing multiplier, or undefined. */
export function lineSpacingNumber(v: unknown): number | undefined {
  if (typeof v === "number" && Number.isFinite(v)) return v;
  if (typeof v === "string") {
    const s = v.trim().toLowerCase();
    if (s === "single") return 1;
    if (s === "double") return 2;
    if (s === "1.5" || s === "one-and-a-half") return 1.5;
    const n = parseFloat(s);
    if (Number.isFinite(n)) return n;
  }
  return undefined;
}

/** True for the retired `bulletList` content type. */
export function isLegacyBulletList(c: any): boolean {
  return !!c && typeof c === "object" && c.type === "bulletList";
}

/**
 * Convert legacy `{ type: "bulletList", items }` (or a `text` with `items`/`bullets`)
 * to canonical `text` content with one bulleted run per item. Other fields carry over.
 */
export function legacyToText(c: any): TextContent {
  const items: unknown[] = Array.isArray(c.items) ? c.items : Array.isArray(c.bullets) ? c.bullets : [];
  const runs: TextRun[] = items.map((it: any) =>
    it && typeof it === "object" && "text" in it ? { bullet: true, ...it } : { text: String(it ?? ""), bullet: true },
  );
  const { items: _i, bullets: _b, type: _t, lineSpacing, ...rest } = c;
  const ls = lineSpacingNumber(lineSpacing);
  return { ...rest, type: "text", runs, ...(ls !== undefined ? { lineSpacing: ls } : {}) };
}
