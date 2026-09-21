/**
 * Fitting helpers shared by the PPTX builder and any preview: how a slide title should be
 * sized so it does not run under the body. Isomorphic.
 */

import { getFontMetrics } from "./text-metrics.js";

export interface TitleFit {
  /** Font size to use, points. */
  fontSizePt: number;
  /** Lines the title will occupy at that size (1 or 2). */
  lines: number;
  /** Characters that fit on one line at that size. */
  charsPerLine: number;
}

export interface FitTitleOptions {
  text: string;
  /** The style's title size, points. */
  fontSizePt: number;
  font?: string;
  /** Usable width for the title, inches. */
  widthInches: number;
  /** Smallest size to shrink to before wrapping (default 20pt). */
  minFontSizePt?: number;
  /** Text-box inset per side, inches (default 0.1). */
  padInches?: number;
}

function charsAt(sizePt: number, font: string | undefined, widthInches: number, pad: number): number {
  const m = getFontMetrics(font);
  // Titles carry more capitals and wide glyphs than body copy: 10% wider average.
  const charWidth = (sizePt * m.avgCharWidth * 1.1) / 72;
  return Math.max(1, Math.floor((widthInches - 2 * pad) / charWidth));
}

/**
 * Keep a title on one line by shrinking it in 2pt steps down to `minFontSizePt`;
 * past that, wrap to two lines at the minimum size.
 */
export function fitTitle(opts: FitTitleOptions): TitleFit {
  const pad = opts.padInches ?? 0.1;
  const min = opts.minFontSizePt ?? 20;
  const len = (opts.text || "").trim().length;
  let size = opts.fontSizePt;
  let cpl = charsAt(size, opts.font, opts.widthInches, pad);
  if (len <= cpl) return { fontSizePt: size, lines: 1, charsPerLine: cpl };
  while (size - 2 >= min) {
    size -= 2;
    cpl = charsAt(size, opts.font, opts.widthInches, pad);
    if (len <= cpl) return { fontSizePt: size, lines: 1, charsPerLine: cpl };
  }
  return { fontSizePt: size, lines: len <= cpl * 2 ? 2 : 2, charsPerLine: cpl };
}

/** True when every non-empty cell reads as a number: 1,234  (5.2)  12.4x  31%  $1.6tn  -8bps  n/a is not numeric. */
export function isNumericColumn(cells: unknown[]): boolean {
  const re = /^\(?[-+−]?[$€£¥]?\s?\d[\d,]*(\.\d+)?\s?(%|x|bps|pp|[kmbt]n?|mm)?\)?$/i;
  const filled = cells.filter((c) => typeof c === "string" && c.trim() !== "" && c.trim() !== "-" && c.trim() !== "—") as string[];
  if (filled.length === 0) return false;
  return filled.every((c) => re.test(c.trim()));
}
