/**
 * Slide content annotation — injects text capacity metadata (_charsPerLine,
 * _maxLines, _maxChars, etc.) onto deck content so agents can size text to fit.
 *
 * The budgets are approximations derived from font metrics; they are what an agent
 * uses to size copy before rendering. Isomorphic.
 */

import { LINE_SPACING } from "./defaults.js";
import { getFontMetrics } from "./text-metrics.js";
import type { DeckJson } from "./types.js";

// ---- Padding constants (must match PPTX renderer) ----

// PPTX table cell padding (from document-model.js TableCell)
const CELL_PAD_LR = 0.1; // inches each side
const CELL_PAD_TB = 0.05; // inches each side

// PPTX TextBody inside margins (from document-model.js TextBody)
const TEXT_BODY_PAD_LR = 0.1;  // lIns/rIns = 91440 EMU = 0.1"
const TEXT_BODY_PAD_TB = 0.05; // tIns/bIns = 45720 EMU = 0.05"

// Bullet list paragraph margins (from slide-style.js default)
const BULLET_MAR_L = 0.375; // inches (marL on paragraph = 342360 EMU)

// Slide dimensions
const SLIDE_WIDTH = 10;
const SLIDE_MARGIN = 0.5;
const CONTENT_WIDTH = SLIDE_WIDTH - 2 * SLIDE_MARGIN; // 9"
const GRID_COLUMNS = 12;

// Title defaults (from slide-builder.js default style)
const TITLE_PAD_LR = 0.1;

// Zone header/subheader heights (from slide-layout.js)
const ZONE_HEADER_HEIGHT = 0.43; // headerBarHeight 0.35" + headerBarGap 0.08"
const ZONE_SUBHEADER_HEIGHT = 0.30; // subheaderHeight 0.22" + lineGap 0.02" + gap 0.06"

// ---- Resolved font defaults from deck style/theme ----

export interface StyleFontDefaults {
  bodyFont: string;
  bodyFontSizePt: number;
  bulletFont: string;
  bulletFontSizePt: number;
  titleFont: string;
  titleFontSizePt: number;
  titleHeightInches: number;
  sectionLabelFont: string;
  sectionLabelFontSizePt: number;
  footerFont: string;
  footerFontSizePt: number;
  headerFont: string;
  headerFontSizePt: number;
  subheaderFont: string;
  subheaderFontSizePt: number;
  tableFontSizePt: number;
  tableFont: string;
}

const HARDCODED_DEFAULTS: StyleFontDefaults = {
  bodyFont: "Calibri",
  bodyFontSizePt: 14,
  bulletFont: "Calibri",
  bulletFontSizePt: 13,
  titleFont: "Arial Bold",
  titleFontSizePt: 28,
  titleHeightInches: 0.62,
  sectionLabelFont: "Calibri",
  sectionLabelFontSizePt: 12,
  footerFont: "Calibri",
  footerFontSizePt: 8,
  headerFont: "Arial Bold",
  headerFontSizePt: 14,
  subheaderFont: "Calibri",
  subheaderFontSizePt: 14,
  tableFontSizePt: 11,
  tableFont: "Calibri",
};

/**
 * Resolve font defaults from the effective deck style (theme tokens merged
 * with deck.style overrides). If no style is provided, returns hardcoded
 * defaults matching slide-style.js DEFAULT_STYLE.
 */
export function resolveStyleFonts(deckStyle?: Record<string, unknown>): StyleFontDefaults {
  if (!deckStyle) return HARDCODED_DEFAULTS;

  const fonts = (deckStyle.fonts as Record<string, string> | undefined) || {};
  const body = (deckStyle.body as Record<string, unknown> | undefined) || {};
  const bullet = (deckStyle.bullet as Record<string, unknown> | undefined) || {};
  const title = (deckStyle.title as Record<string, unknown> | undefined) || {};
  const header = (deckStyle.header as Record<string, unknown> | undefined) || {};
  const subheader = (deckStyle.subheader as Record<string, unknown> | undefined) || {};
  const table = (deckStyle.table as Record<string, unknown> | undefined) || {};

  const bodyFont = (body.fontFamily as string) || fonts.body || HARDCODED_DEFAULTS.bodyFont;
  const bodyFontSizePt = (body.size as number) || HARDCODED_DEFAULTS.bodyFontSizePt;
  const titleFont = fonts.title || HARDCODED_DEFAULTS.titleFont;
  const titleFontSizePt = (title.size as number) || HARDCODED_DEFAULTS.titleFontSizePt;
  // SDK: titleH = titleSize * 914400/72 * 1.6 / 914400 = titleSize * 1.6 / 72
  const titleHeightInches = titleFontSizePt * 1.6 / 72;

  return {
    bodyFont,
    bodyFontSizePt,
    bulletFont: (bullet.fontFamily as string) || (bullet.font as string) || fonts.body || HARDCODED_DEFAULTS.bulletFont,
    bulletFontSizePt: (bullet.size as number) || HARDCODED_DEFAULTS.bulletFontSizePt,
    titleFont,
    titleFontSizePt,
    titleHeightInches,
    sectionLabelFont: fonts.label || HARDCODED_DEFAULTS.sectionLabelFont,
    sectionLabelFontSizePt: HARDCODED_DEFAULTS.sectionLabelFontSizePt,
    footerFont: fonts.body || HARDCODED_DEFAULTS.footerFont,
    footerFontSizePt: HARDCODED_DEFAULTS.footerFontSizePt,
    headerFont: (header.fontFamily as string) || fonts.title || HARDCODED_DEFAULTS.headerFont,
    headerFontSizePt: (header.size as number) || bodyFontSizePt,
    subheaderFont: (subheader.fontFamily as string) || fonts.body || HARDCODED_DEFAULTS.subheaderFont,
    subheaderFontSizePt: (subheader.size as number) || bodyFontSizePt,
    tableFontSizePt: (table.bodySize as number) || HARDCODED_DEFAULTS.tableFontSizePt,
    tableFont: (table.bodyFont as string) || fonts.body || HARDCODED_DEFAULTS.tableFont,
  };
}

// ---- Local capacity calculators (cell padding differs from TextBody padding) ----

/** Chars per line for a text body (uses TextBody margins). */
function calcTextCharsPerLine(containerWidthInches: number, fontSizePt: number, font?: string, extraMarginL = 0): number {
  const m = getFontMetrics(font);
  const contentWidth = containerWidthInches - 2 * TEXT_BODY_PAD_LR - extraMarginL;
  if (contentWidth <= 0) return containerWidthInches > 0 ? 1 : 0;
  const charWidth = fontSizePt * m.avgCharWidth / 72;
  return Math.max(1, Math.floor(contentWidth / charWidth));
}

/** Max lines for a text body (uses TextBody margins). */
function calcTextMaxLines(containerHeightInches: number, fontSizePt: number, font?: string, lineSpacing = 1): number {
  const m = getFontMetrics(font);
  const contentHeight = containerHeightInches - 2 * TEXT_BODY_PAD_TB;
  if (contentHeight <= 0) return containerHeightInches > 0 ? 1 : 0;
  const lineHeight = fontSizePt * m.lineHeightFactor * lineSpacing / 72;
  return Math.max(1, Math.round(contentHeight / lineHeight));
}

/** Chars per line for a table cell (uses cell padding). */
function calcCellCharsPerLine(containerWidthInches: number, fontSizePt: number, font?: string): number {
  const m = getFontMetrics(font);
  const contentWidth = containerWidthInches - 2 * CELL_PAD_LR;
  if (contentWidth <= 0) return containerWidthInches > 0 ? 1 : 0;
  const charWidth = fontSizePt * m.avgCharWidth / 72;
  return Math.max(1, Math.floor(contentWidth / charWidth));
}

/** Max lines for a table cell (uses cell padding). */
function calcCellMaxLines(containerHeightInches: number, fontSizePt: number, font?: string, lineSpacing = 1): number {
  const m = getFontMetrics(font);
  const contentHeight = containerHeightInches - 2 * CELL_PAD_TB;
  if (contentHeight <= 0) return containerHeightInches > 0 ? 1 : 0;
  const lineHeight = fontSizePt * m.lineHeightFactor * lineSpacing / 72;
  return Math.max(1, Math.round(contentHeight / lineHeight));
}

/** Max chars for a single-line text field (title, sectionLabel, footer, zone header). */
function calcSingleLineMaxChars(widthInches: number, fontSizePt: number, font?: string): number {
  const m = getFontMetrics(font);
  const usableWidth = widthInches - 2 * TEXT_BODY_PAD_LR;
  if (usableWidth <= 0) return 0;
  const charWidth = fontSizePt * m.avgCharWidth / 72;
  return Math.floor(usableWidth / charWidth);
}

// ---- Body height calculation ----

function calcAvailableBodyHeight(slide: any): number {
  let h = 7.5 - 1.0; // slide height minus top+bottom margins
  if (slide.sectionLabel) h -= 0.3;
  if (slide.title) h -= 0.65;
  if (slide.footer) h -= 0.4;
  return h;
}

// ---- Recursive annotation walker ----

/**
 * Walk the grid tree and inject capacity annotations on every node/content.
 * Does NOT detect issues — just writes _charsPerLine, _maxChars, etc.
 */
function annotateNode(node: any, parentWidth: number, parentHeight: number, fd: StyleFontDefaults): void {
  // Subtract margin and padding — they inset the zone's usable area on all sides
  const inset = (node.margin ?? 0) + (node.padding ?? 0);
  const nodeWidth = parentWidth - 2 * inset;
  const nodeHeight = parentHeight - 2 * inset;

  // Zone header — single line in headerBar
  if (node.header && typeof node.header === "string") {
    node._headerMaxChars = calcSingleLineMaxChars(nodeWidth, fd.headerFontSizePt, fd.headerFont);
  }

  // Zone subheader — single line
  if (node.subheader && typeof node.subheader === "string") {
    node._subheaderMaxChars = calcSingleLineMaxChars(nodeWidth, fd.subheaderFontSizePt, fd.subheaderFont);
  }

  // Content leaf — subtract header/subheader height from available content area
  if (node.content) {
    const c = node.content;
    const font = c.font || fd.bodyFont;
    let contentHeight = nodeHeight;
    if (node.header) contentHeight -= ZONE_HEADER_HEIGHT;
    if (node.subheader) contentHeight -= ZONE_SUBHEADER_HEIGHT;

    switch (c.type) {
      case "text": {
        const fontSizePt = c.fontSize ? c.fontSize / 100 : fd.bodyFontSizePt;
        const cpl = calcTextCharsPerLine(nodeWidth, fontSizePt, font);
        const ml = calcTextMaxLines(contentHeight, fontSizePt, font);
        c._charsPerLine = cpl;
        c._maxLines = ml;
        c._maxChars = cpl * ml;
        break;
      }
      case "callout": {
        const fontSizePt = c.fontSize ? c.fontSize / 100 : fd.bodyFontSizePt;
        const cpl = calcTextCharsPerLine(nodeWidth - 0.3, fontSizePt, font);
        const ml = calcTextMaxLines(contentHeight, fontSizePt, font) - (c.title ? 1 : 0);
        c._charsPerLine = cpl;
        c._maxLines = Math.max(1, ml);
        c._maxChars = cpl * Math.max(1, ml);
        break;
      }
      case "bulletList": {
        const blFont = c.font || fd.bulletFont;
        const fontSizePt = c.fontSize ? c.fontSize / 100 : fd.bulletFontSizePt;
        const lineSpacing: number = typeof c.lineSpacing === "number" ? c.lineSpacing : LINE_SPACING;
        const cpl = calcTextCharsPerLine(nodeWidth, fontSizePt, blFont, BULLET_MAR_L);
        const ml = calcTextMaxLines(contentHeight, fontSizePt, blFont, lineSpacing);
        const itemCount = Array.isArray(c.items) ? c.items.length : 1;
        c._charsPerLine = cpl;
        c._maxLines = ml;
        c._maxChars = cpl * ml;
        c._maxCharsPerItem = itemCount > 0 ? cpl * Math.floor(ml / itemCount) : cpl;
        break;
      }
      case "table": {
        if (c.data && Array.isArray(c.data.rows)) {
          const data = c.data;
          const tableFontSizePt = data.fontSize ? data.fontSize / 100 : fd.tableFontSizePt;
          const tblFont = fd.tableFont;
          const rowH = data.rowHeight ?? 0.4;
          const tblNumCols = data.headers?.length || (data.rows[0]?.length ?? 1);
          const colWidths = data.colWidths || Array.from({ length: tblNumCols }, () => GRID_COLUMNS / tblNumCols);
          const totalSpan = colWidths.reduce((a: number, b: number) => a + b, 0);

          const perColMaxChars: number[] = colWidths.map((span: number) => {
            const w = (span / totalSpan) * nodeWidth;
            const cpl = calcCellCharsPerLine(w, tableFontSizePt, tblFont);
            const ml = calcCellMaxLines(rowH, tableFontSizePt, tblFont);
            return cpl * ml;
          });

          // Header row may have different height/font size
          const headerFontSizePt = data.headerFontSize ? data.headerFontSize / 100 : tableFontSizePt;
          const headerH = data.headerHeight ?? 0.35;
          const perHeaderMaxChars: number[] = colWidths.map((span: number) => {
            const w = (span / totalSpan) * nodeWidth;
            const cpl = calcCellCharsPerLine(w, headerFontSizePt, tblFont);
            const ml = calcCellMaxLines(headerH, headerFontSizePt, tblFont);
            return cpl * ml;
          });

          data._headersMaxChars = perHeaderMaxChars;
          data._cellMaxChars = perColMaxChars;

          // _maxRows: how many data rows fit in the available vertical space
          const hasHeaders = data.headers && data.headers.length > 0;
          const usableHeight = contentHeight - (hasHeaders ? headerH : 0);
          data._maxRows = Math.max(0, Math.floor(usableHeight / rowH));
        }
        break;
      }
    }
  }

  // Recurse into children
  if (Array.isArray(node.children)) {
    const children = node.children.filter(Boolean);
    const direction = node.direction || "row";
    const gap = node.gap ?? 0.15;
    const totalGap = gap * (children.length - 1);

    let childrenHeight = nodeHeight;
    if (node.header) childrenHeight -= ZONE_HEADER_HEIGHT;
    if (node.subheader) childrenHeight -= ZONE_SUBHEADER_HEIGHT;

    let totalSpan = 0;
    for (const child of children) {
      totalSpan += child.span || (GRID_COLUMNS / children.length);
    }

    for (const child of children) {
      const childSpan = child.span || (GRID_COLUMNS / children.length);
      let childWidth: number;
      let childHeight: number;

      if (direction === "row") {
        childWidth = (childSpan / totalSpan) * (nodeWidth - totalGap);
        childHeight = childrenHeight;
      } else {
        childWidth = nodeWidth;
        childHeight = (childSpan / totalSpan) * (childrenHeight - totalGap);
      }

      annotateNode(child, childWidth, childHeight, fd);
    }
  }
}

// ---- Top-level entry point ----

/**
 * Deep-clone the deck content and inject capacity annotations on every
 * text/table/bulletList zone. Returns the annotated clone (original is untouched).
 */
export function annotateDeck(content: DeckJson | Record<string, unknown>, deckStyle?: Record<string, unknown>): { annotated: any } {
  const annotated = JSON.parse(JSON.stringify(content));

  const slides = annotated.slides;
  if (!Array.isArray(slides)) {
    return { annotated };
  }

  // Resolve effective style from the deck (may include theme tokens merged in by caller)
  const effectiveStyle = deckStyle || (content.style as Record<string, unknown> | undefined);
  const fd = resolveStyleFonts(effectiveStyle);

  for (let i = 0; i < slides.length; i++) {
    const slide = slides[i];

    // Inject slide-level annotations
    if (slide.title) {
      const m = getFontMetrics(fd.titleFont);
      const availableWidth = CONTENT_WIDTH - 2 * TITLE_PAD_LR;
      // Titles have more uppercase/wide characters — use 10% wider avg char width
      const charWidth = fd.titleFontSizePt * m.avgCharWidth * 1.1 / 72;
      const cpl = Math.floor(availableWidth / charWidth);
      const lineHeight = fd.titleFontSizePt * m.lineHeightFactor / 72;
      const ml = Math.floor(fd.titleHeightInches / lineHeight);
      slide._titleMaxChars = cpl * ml;
      slide._titleCharsPerLine = cpl;
      slide._titleMaxLines = ml;
    }
    if (slide.sectionLabel) {
      slide._sectionLabelMaxChars = calcSingleLineMaxChars(CONTENT_WIDTH, fd.sectionLabelFontSizePt, fd.sectionLabelFont);
    }
    if (slide.footer) {
      slide._footerMaxChars = calcSingleLineMaxChars(CONTENT_WIDTH, fd.footerFontSizePt, fd.footerFont);
    }

    const availableH = calcAvailableBodyHeight(slide);

    // Annotate type: "table" slide data (top-level table slides)
    if (slide.type === "table" && slide.data && Array.isArray(slide.data.rows)) {
      const data = slide.data;
      const fontSizePt = data.fontSize ? data.fontSize / 100 : fd.tableFontSizePt;
      const rowH = data.rowHeight ?? 0.4;
      const headerH = data.headerHeight ?? 0.35;
      const slideNumCols = data.headers?.length || (data.rows[0]?.length ?? 1);
      const colWidths = data.colWidths || Array.from({ length: slideNumCols }, () => GRID_COLUMNS / slideNumCols);
      const totalSpan = colWidths.reduce((a: number, b: number) => a + b, 0);

      data._cellMaxChars = colWidths.map((span: number) => {
        const w = (span / totalSpan) * CONTENT_WIDTH;
        return calcCellCharsPerLine(w, fontSizePt, fd.tableFont) * calcCellMaxLines(rowH, fontSizePt, fd.tableFont);
      });

      const hdrFontSizePt = data.headerFontSize ? data.headerFontSize / 100 : fontSizePt;
      data._headersMaxChars = colWidths.map((span: number) => {
        const w = (span / totalSpan) * CONTENT_WIDTH;
        return calcCellCharsPerLine(w, hdrFontSizePt, fd.tableFont) * calcCellMaxLines(headerH, hdrFontSizePt, fd.tableFont);
      });

      const hasHeaders = data.headers && data.headers.length > 0;
      const usableHeight = availableH - (hasHeaders ? headerH : 0);
      data._maxRows = Math.max(0, Math.floor(usableHeight / rowH));
    }

    // Annotate body tree
    if (slide.body) {
      annotateNode(slide.body, CONTENT_WIDTH, availableH, fd);
    }
  }

  return { annotated };
}
