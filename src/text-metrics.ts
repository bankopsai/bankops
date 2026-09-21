/**
 * Text metrics — approximate font metrics used to estimate how much text fits
 * in a zone or table cell. Isomorphic. Values are approximations for mixed-case
 * English; PowerPoint and browsers differ slightly.
 */

export interface FontMetrics { avgCharWidth: number; lineHeightFactor: number }

export const FONT_METRICS: Record<string, FontMetrics> = {
  "Calibri":         { avgCharWidth: 0.43,  lineHeightFactor: 1.2 },
  "Arial":           { avgCharWidth: 0.47,  lineHeightFactor: 1.2 },
  "Arial Bold":      { avgCharWidth: 0.49,  lineHeightFactor: 1.2 },
  "Helvetica Neue":  { avgCharWidth: 0.47,  lineHeightFactor: 1.2 },
  "Georgia":         { avgCharWidth: 0.455, lineHeightFactor: 1.2 },
  "Garamond":        { avgCharWidth: 0.42,  lineHeightFactor: 1.2 },
  "Calibri Light":   { avgCharWidth: 0.41,  lineHeightFactor: 1.2 },
  "Times New Roman": { avgCharWidth: 0.42,  lineHeightFactor: 1.2 },
};

export const DEFAULT_METRICS: FontMetrics = { avgCharWidth: 0.45, lineHeightFactor: 1.2 };

/** PPTX TextBody insets (document-model TextBody defaults), inches. */
export const TEXT_BODY_PAD_LR = 0.1;
export const TEXT_BODY_PAD_TB = 0.05;

/** Table cell padding, inches. */
export interface Padding { left?: number; right?: number; top?: number; bottom?: number }
export const DEFAULT_CELL_PADDING: Required<Padding> = { left: 0.1, right: 0.1, top: 0.05, bottom: 0.05 };

export const SLIDE_WIDTH = 10;
export const SLIDE_MARGIN = 0.5;
export const CONTENT_WIDTH = SLIDE_WIDTH - 2 * SLIDE_MARGIN;
export const GRID_COLUMNS = 12;

export function getFontMetrics(fontFamily?: string | null): FontMetrics {
  if (!fontFamily) return DEFAULT_METRICS;
  return FONT_METRICS[fontFamily] || DEFAULT_METRICS;
}

export function charWidthInches(fontFamily: string | undefined, fontSizePt: number): number {
  return (fontSizePt * getFontMetrics(fontFamily).avgCharWidth) / 72;
}

export function lineHeightInches(fontFamily: string | undefined, fontSizePt: number, lineSpacing?: number | null): number {
  const ls = lineSpacing != null ? lineSpacing : 1;
  return (fontSizePt * getFontMetrics(fontFamily).lineHeightFactor * ls) / 72;
}

/** Characters per line in a container, using cell padding by default. */
export function charsPerLine(containerWidthInches: number, fontFamily: string | undefined, fontSizePt: number, padding?: Padding): number {
  const pad = padding || DEFAULT_CELL_PADDING;
  const contentWidth = containerWidthInches - (pad.left || 0) - (pad.right || 0);
  if (contentWidth <= 0) return 0;
  return Math.floor(contentWidth / charWidthInches(fontFamily, fontSizePt));
}

/** Lines that fit in a container height, using cell padding by default. */
export function maxLines(containerHeightInches: number, fontFamily: string | undefined, fontSizePt: number, padding?: Padding, lineSpacing?: number | null): number {
  const pad = padding || DEFAULT_CELL_PADDING;
  const contentHeight = containerHeightInches - (pad.top || 0) - (pad.bottom || 0);
  if (contentHeight <= 0) return 0;
  return Math.floor(contentHeight / lineHeightInches(fontFamily, fontSizePt, lineSpacing));
}

/** Chars per line for a text body with standard TextBody insets (browser-preview convention). */
export function calcCharsPerLine(widthInches: number, fontSizePt: number, font?: string, extraMarginL = 0): number {
  const contentWidth = widthInches - 2 * TEXT_BODY_PAD_LR - extraMarginL;
  if (contentWidth <= 0) return 0;
  return Math.floor(contentWidth / charWidthInches(font, fontSizePt));
}

/** Max lines for a text body with standard TextBody insets. */
export function calcMaxLines(heightInches: number, fontSizePt: number, font?: string, lineSpacing = 1): number {
  const contentHeight = heightInches - 2 * TEXT_BODY_PAD_TB;
  if (contentHeight <= 0) return 0;
  return Math.floor(contentHeight / lineHeightInches(font, fontSizePt, lineSpacing));
}

export interface TextCapacity {
  charsPerLine: number;
  maxLines: number;
  maxChars: number;
  zoneWidthInches: number;
  zoneHeightInches: number;
}

/** Full capacity of a zone's content area (after header/subheader subtraction). */
export function computeTextCapacity(opts: { widthInches: number; heightInches: number; fontSizePt: number; font?: string; lineSpacing?: number; extraMarginL?: number }): TextCapacity {
  const cpl = calcCharsPerLine(opts.widthInches, opts.fontSizePt, opts.font, opts.extraMarginL);
  const ml = calcMaxLines(opts.heightInches, opts.fontSizePt, opts.font, opts.lineSpacing);
  return {
    charsPerLine: cpl,
    maxLines: ml,
    maxChars: cpl * ml,
    zoneWidthInches: Math.round(opts.widthInches * 100) / 100,
    zoneHeightInches: Math.round(opts.heightInches * 100) / 100,
  };
}

export interface TableCellCapacity { colWidthInches: number; charsPerLine: number; maxLines: number; maxChars: number }

/** Capacity of one table cell given its 12-column span. */
export function tableCellCapacity(opts: { colSpan?: number; rowHeight?: number; font?: string; fontSize?: number; lineSpacing?: number; tableWidthInches?: number } = {}): TableCellCapacity {
  const span = opts.colSpan || 1;
  const rowH = opts.rowHeight != null ? opts.rowHeight : 0.4;
  const font = opts.font || "Calibri";
  const fontSize = opts.fontSize != null ? opts.fontSize : 11;
  const ls = opts.lineSpacing != null ? opts.lineSpacing : 1;
  const tableW = opts.tableWidthInches != null ? opts.tableWidthInches : CONTENT_WIDTH;
  const colW = (span / GRID_COLUMNS) * tableW;
  const cpl = charsPerLine(colW, font, fontSize, DEFAULT_CELL_PADDING);
  const ml = maxLines(rowH, font, fontSize, DEFAULT_CELL_PADDING, ls);
  return { colWidthInches: Math.round(colW * 100) / 100, charsPerLine: cpl, maxLines: ml, maxChars: cpl * ml };
}

/** Markdown reference table of cell capacities, for agent instructions. */
export function generateReferenceTable(opts: { font?: string; fontSize?: number } = {}): string {
  const font = opts.font || "Calibri";
  const fontSize = opts.fontSize != null ? opts.fontSize : 11;
  const spans = [1, 2, 3, 4, 6, 9, 12];
  const rowHeights = [0.3, 0.4, 0.5, 0.65, 0.8];
  const pad = (val: unknown, width: number) => String(val).padStart(width, " ");
  const lines: string[] = [];
  lines.push(`Table cell capacity reference (${font} ${fontSize}pt, single-spaced):`, "");
  lines.push("Characters per line by column span:", "| Span | Width  | Chars/line |", "|------|--------|------------|");
  for (const s of spans) {
    const cap = tableCellCapacity({ colSpan: s, font, fontSize });
    lines.push(`| ${pad(s, 4)} | ${pad(cap.colWidthInches.toFixed(2) + '"', 6)} | ${pad(cap.charsPerLine, 10)} |`);
  }
  lines.push("", "Lines per cell by row height:", "| Row Height | Lines |", "|------------|-------|");
  for (const h of rowHeights) lines.push(`| ${pad(h.toFixed(2) + '"', 10)} | ${pad(maxLines(h, font, fontSize), 5)} |`);
  lines.push("", "Max characters per cell (span x row height):");
  lines.push("| Span \\ Height |" + rowHeights.map((h) => ` ${h.toFixed(1)}" |`).join(""));
  lines.push("|---------------|" + rowHeights.map(() => "------|").join(""));
  for (const s of spans) {
    lines.push(`| ${pad(s, 13)} |` + rowHeights.map((h) => ` ${pad(tableCellCapacity({ colSpan: s, rowHeight: h, font, fontSize }).maxChars, 4)} |`).join(""));
  }
  return lines.join("\n");
}

const TextMetrics = {
  FONT_METRICS, DEFAULT_METRICS, DEFAULT_CELL_PADDING,
  getFontMetrics, charWidthInches, lineHeightInches, charsPerLine, maxLines,
  calcCharsPerLine, calcMaxLines, computeTextCapacity, tableCellCapacity, generateReferenceTable,
};
export default TextMetrics;
