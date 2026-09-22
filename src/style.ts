/**
 * SlideStyle — CSS-like stylesheet for decks. Users write human-friendly units
 * (pt, inches, hex); toInternal() converts to the EMU / hundredths-pt tokens the
 * renderer and browser preview consume. Isomorphic.
 */

import U from "./units.js";
import type { StyleTokens } from "./types.js";

/** Any subset of DEFAULT_STYLE (two-level deep partial). */
export type StyleOverrides = { [section: string]: unknown };

// ---- Deep merge utility (two-level) ----
function deepMerge(base: any, overrides: any): any {
  if (!overrides) return base;
  const result: any = {};
  for (const k of Object.keys(base)) {
    if (overrides[k] !== undefined && typeof base[k] === "object" && base[k] !== null && !Array.isArray(base[k])) {
      result[k] = {};
      for (const sk of Object.keys(base[k])) {
        result[k][sk] = overrides[k] && overrides[k][sk] !== undefined ? overrides[k][sk] : base[k][sk];
      }
      if (overrides[k]) {
        for (const ek of Object.keys(overrides[k])) {
          if (result[k][ek] === undefined) result[k][ek] = overrides[k][ek];
        }
      }
    } else {
      result[k] = overrides[k] !== undefined ? overrides[k] : base[k];
    }
  }
  for (const ok of Object.keys(overrides)) {
    if (result[ok] === undefined) result[ok] = overrides[ok];
  }
  return result;
}

// ---- Default style (human-friendly units): the "corporate" IB look ----
export const DEFAULT_STYLE = {
  slide: {
    background: '#FFFFFF',
    margin: 0.5                   // inches
  },

  colors: {
    primary: '#333333',           // main text color
    secondary: '#666666',         // muted text
    accent: '#4472C4',            // highlights, section labels
    darkFill: '#164556',          // header bar backgrounds
    border: '#BFBFBF',           // card/table borders
    headerBarText: '#FFFFFF',     // text on dark header bars
    cardFill: '#FFFFFF'           // card backgrounds
  },

  chart: {
    colors: null,                 // null = use default chart palette
    fontFamily: null,             // null = inherit from fonts.body
    fontSize: null,               // null = default 11pt
    barBorderColor: null,         // null = no bar borders
    barBorderWidth: null,         // null = no bar borders
    bar: null,                    // null = defaults; { showXAxis, showYAxis, showSplitLine, gridLeft, gridRight, gridTop, gridBottom }
    pie: null,                    // null = defaults; { borderColor, borderWidth }
    combo: null                   // null = defaults; { showPrimaryYAxis, showSecondaryYAxis, showSplitLine, showBarValues, showLineValues }
  },

  fonts: {
    title: 'Arial Bold',         // slide titles, header bars
    body: 'Calibri',             // body text, bullets, tables
    label: 'Calibri',            // section labels (eyebrow)
    dividerNumber: 'Calibri Light' // giant section numbers
  },

  title: {
    size: 28,                     // pt
    bold: true,
    color: '#405363',             // dark slate
    y: 0.9                        // inches from top
  },

  sectionLabel: {
    size: 12,                     // pt
    bold: true,
    italic: false,
    color: null,                  // null = inherit from colors.accent
    y: 0.43                       // inches from top
  },

  body: {
    size: 14,                     // pt
    color: null,                  // null = inherit from colors.primary
    align: 'l',                   // 'l' | 'ctr' | 'r'
    fontFamily: null,             // null = inherit from fonts.body
    lineSpacing: null,            // null = use Defaults.LINE_SPACING
    bold: null,                   // null = false
    italic: null                  // null = false
  },

  footnote: {
    size: 8,                      // pt
    color: null,                  // null = inherit from colors.secondary
    y: null                       // null = auto (0 internally)
  },

  bullet: {
    char: '\u2022',
    indent: 0.375,                // inches
    marginLeft: 0.375,            // inches
    font: null,                   // null = inherit from fonts.body
    bulletFont: null,             // null = same as text font
    bulletColor: null,            // null = same as text color
    size: 13,                     // pt (default 13pt for bullet lists)
    color: null,                  // null = inherit from colors.primary
    fontFamily: null,             // null = inherit from fonts.body
    lineSpacing: null             // null = use Defaults.LINE_SPACING
  },

  divider: {
    numberSize: 115,              // pt (giant section number)
    numberColor: null,            // null = inherit from colors.primary
    numberX: 0.5,                 // inches
    numberY: 2.4,                 // inches
    titleSize: 24,                // pt
    titleColor: null,             // null = inherit from colors.accent
    titleY: null                  // null = auto (0 internally)
  },

  header: {
    size: null,                   // null = inherit from body.size
    color: null,                  // null = inherit from colors.headerBarText
    background: null,             // null = inherit from colors.darkFill
    align: 'ctr',
    borderRadius: null,           // null = PPTX roundRect default (16.667%), 0 = sharp corners
    fontFamily: null,             // null = inherit from fonts.title
    bold: null,                   // null = derived from font string
    italic: null                  // null = false
  },

  subheader: {
    size: null,                   // null = inherit from body.size
    bold: true,
    color: null,                  // null = inherit from colors.primary
    align: 'ctr',
    lineColor: null,              // null = inherit from colors.primary
    lineWidth: 0.75,              // pt
    fontFamily: null,             // null = inherit from fonts.body
    italic: null                  // null = false
  },

  subfooter: {
    size: null,                   // null = inherit from footnote.size * 0.9
    italic: true,
    color: null,                  // null = inherit from colors.secondary
    fontFamily: null              // null = inherit from fonts.body
  },

  table: {
    headerFill: null,             // null = inherit from colors.darkFill (matches zone header bars)
    headerRuleColor: null,        // rule under the header row; null = inherit from colors.accent
    headerRuleWidth: 1.25,        // pt; 0 = no rule
    borderMode: 'horizontal',     // 'horizontal' (rules between rows only) | 'grid' (every cell) | 'none'
    headerTextColor: '#FFFFFF',
    headerSize: 12,               // pt
    headerBold: true,
    headerFont: null,             // null = inherit from fonts.body
    verticalHeaderFill: null,     // null = inherit from headerFill (same as horizontal by default)
    verticalHeaderTextColor: null, // null = inherit from headerTextColor
    verticalHeaderSize: null,     // null = inherit from headerSize
    verticalHeaderBold: null,     // null = inherit from headerBold
    verticalHeaderFont: null,     // null = inherit from headerFont
    bodySize: 11,                 // pt
    bodyTextColor: null,          // null = inherit from colors.primary
    bodyFont: null,               // null = inherit from fonts.body
    alternateRows: true,           // even/odd row backgrounds differ
    evenRowFill: '#F4F6F9',       // subtle banding
    oddRowFill: '#FFFFFF',
    borderColor: '#D9DEE5',       // hairline between rows
    borderWidth: 0.5,             // pt
    summaryRowFill: '#E9EEF6',      // light accent tint (median, mean, total rows)
    summaryRowTextColor: null,      // null = inherit from colors.primary
    summaryRowSize: null,           // null = inherit from bodySize
    summaryRowBold: true,
    summaryRowFont: null,           // null = inherit from bodyFont
    summaryRowBorderTopColor: null,    // null = inherit from colors.accent
    summaryRowBorderTopWidth: 1,       // pt
    summaryRowBorderBottomColor: null, // null = inherit from colors.accent
    summaryRowBorderBottomWidth: 1,
    summaryRowBorderLeftColor: null,
    summaryRowBorderLeftWidth: 0,      // no vertical rules by default
    summaryRowBorderRightColor: null,
    summaryRowBorderRightWidth: 0
  },

  statGrid: {
    valueFontFamily: null,        // null = inherit from fonts.title
    valueSize: 16,                // pt
    valueColor: null,             // null = inherit from colors.accent
    valueBold: true,
    valueItalic: false,
    labelFontFamily: null,        // null = inherit from fonts.body
    labelSize: null,              // null = 85% of body.size
    labelColor: null,             // null = inherit from colors.primary
    labelBold: false,
    labelItalic: false,
    sublabelFontFamily: null,     // null = inherit from fonts.body
    sublabelSize: null,           // null = inherit from footnote.size
    sublabelColor: null,          // null = inherit from colors.secondary
    sublabelBold: false,
    sublabelItalic: false
  },

  cardGrid: {
    titleFontFamily: null,        // null = inherit from fonts.body
    titleSize: null,              // null = inherit from body.size
    titleColor: null,             // null = inherit from colors.accent
    titleBold: true,
    titleItalic: false,
    lineFontFamily: null,         // null = inherit from fonts.body
    lineSize: null,               // null = 85% of body.size
    lineColor: null,              // null = inherit from colors.primary
    lineBold: false,
    lineItalic: false,
    lineBullet: false,            // show bullet char on lines
    lineBulletIndent: 0.05        // inches
  },

  spacing: {
    columnGap: 0.25,              // inches
    bulletSpacing: 0.125,         // inches
    cardGap: 0.08,                // inches
    cardBorderWidth: 0.75,        // pt
    cardPaddingLeft: 0.075,       // inches
    cardPaddingTop: 0.05,         // inches
    subheaderLineWidth: 0.75,     // pt
    headerBarHeight: 0.35,        // inches
    contentGap: 0.1,              // inches
    zoneGap: 0.15                 // inches
  },

  titleSlide: {
    titleSize: 36,                // pt
    subtitleSize: 18              // pt
  }
};


export type StyleSheet = typeof DEFAULT_STYLE;

export class SlideStyle {
  _style: any;

  constructor(userStyle?: StyleOverrides | null) {
    this._style = deepMerge(DEFAULT_STYLE, userStyle || {});
  }

  /** The merged stylesheet in human units. */
  get sheet(): StyleSheet { return this._style as StyleSheet; }

  /** Convert to internal format (EMU, hundredths-pt). */
  toInternal(): StyleTokens {
    const s = this._style;
    const primary = s.colors.primary;
    const secondary = s.colors.secondary;
    const accent = s.colors.accent;
  
    return {
      sectionLabel: {
        font: s.fonts.label,
        fontSize: U.pointsToFontSize(s.sectionLabel.size),
        bold: s.sectionLabel.bold,
        italic: s.sectionLabel.italic || false,
        color: s.sectionLabel.color || accent,
        x: Math.round(U.inchesToEmu(s.slide.margin)),
        y: Math.round(U.inchesToEmu(s.sectionLabel.y))
      },
      mainTitle: {
        font: s.fonts.title,
        fontSize: U.pointsToFontSize(s.title.size),
        bold: s.title.bold,
        color: s.title.color || primary,
        x: Math.round(U.inchesToEmu(s.slide.margin)),
        y: Math.round(U.inchesToEmu(s.title.y))
      },
      bodyText: {
        font: s.body.fontFamily || s.fonts.body,
        fontSize: U.pointsToFontSize(s.body.size),
        color: s.body.color || primary,
        align: s.body.align,
        lineSpacing: s.body.lineSpacing,
        bold: s.body.bold || false,
        italic: s.body.italic || false
      },
      bullet: {
        char: s.bullet.char,
        indent: -Math.round(U.inchesToEmu(s.bullet.indent)),
        marginLeft: Math.round(U.inchesToEmu(s.bullet.marginLeft)),
        bulletFont: s.bullet.bulletFont,
        bulletColor: s.bullet.bulletColor,
        font: s.bullet.fontFamily || s.bullet.font,
        fontSize: U.pointsToFontSize(s.bullet.size || s.body.size),
        color: s.bullet.color || primary,
        lineSpacing: s.bullet.lineSpacing
      },
      footnote: {
        font: s.fonts.body,
        fontSize: U.pointsToFontSize(s.footnote.size),
        color: s.footnote.color || secondary,
        y: s.footnote.y != null ? Math.round(U.inchesToEmu(s.footnote.y)) : 0
      },
      divider: {
        numberFont: s.fonts.dividerNumber,
        numberFontSize: U.pointsToFontSize(s.divider.numberSize),
        numberColor: s.divider.numberColor || primary,
        numberX: Math.round(U.inchesToEmu(s.divider.numberX)),
        numberY: Math.round(U.inchesToEmu(s.divider.numberY)),
        titleFont: s.fonts.body,
        titleFontSize: U.pointsToFontSize(s.divider.titleSize),
        titleColor: s.divider.titleColor || accent,
        titleY: s.divider.titleY != null ? Math.round(U.inchesToEmu(s.divider.titleY)) : 0
      },
      palette: {
        primaryText: primary,
        secondaryText: secondary,
        accent: accent,
        darkFill: s.colors.darkFill,
        border: s.colors.border
      },
      tableHeader: {
        fillColor: s.table.headerFill || s.colors.darkFill,
        textColor: s.table.headerTextColor,
        font: s.table.headerFont || null,
        fontSize: U.pointsToFontSize(s.table.headerSize),
        bold: s.table.headerBold,
        borderBottom: { color: s.table.headerRuleColor || accent, width: Math.round(U.pointsToEmu(s.table.headerRuleWidth != null ? s.table.headerRuleWidth : 0)) }
      },
      tableVerticalHeader: {
        fillColor: s.table.verticalHeaderFill || s.table.headerFill || s.colors.darkFill,
        textColor: s.table.verticalHeaderTextColor || s.table.headerTextColor,
        font: s.table.verticalHeaderFont || s.table.headerFont || null,
        fontSize: U.pointsToFontSize(s.table.verticalHeaderSize || s.table.headerSize),
        bold: s.table.verticalHeaderBold != null ? s.table.verticalHeaderBold : s.table.headerBold
      },
      tableBody: {
        alternateRows: s.table.alternateRows !== false,
        fillEven: s.table.evenRowFill,
        fillOdd: s.table.oddRowFill,
        textColor: s.table.bodyTextColor || primary,
        font: s.table.bodyFont || null,
        fontSize: U.pointsToFontSize(s.table.bodySize)
      },
      tableSummaryRow: {
        fillColor: s.table.summaryRowFill || s.table.headerFill || s.colors.darkFill,
        textColor: s.table.summaryRowTextColor || (s.table.summaryRowFill ? primary : s.table.headerTextColor),
        font: s.table.summaryRowFont || s.table.bodyFont || null,
        fontSize: U.pointsToFontSize(s.table.summaryRowSize || s.table.bodySize),
        bold: s.table.summaryRowBold != null ? s.table.summaryRowBold : true,
        borderTop: { color: s.table.summaryRowBorderTopColor || accent, width: U.pointsToEmu(s.table.summaryRowBorderTopWidth != null ? s.table.summaryRowBorderTopWidth : s.table.borderWidth) },
        borderBottom: { color: s.table.summaryRowBorderBottomColor || accent, width: U.pointsToEmu(s.table.summaryRowBorderBottomWidth != null ? s.table.summaryRowBorderBottomWidth : s.table.borderWidth) },
        borderLeft: { color: s.table.summaryRowBorderLeftColor || s.table.borderColor || s.colors.border, width: U.pointsToEmu(s.table.summaryRowBorderLeftWidth != null ? s.table.summaryRowBorderLeftWidth : s.table.borderWidth) },
        borderRight: { color: s.table.summaryRowBorderRightColor || s.table.borderColor || s.colors.border, width: U.pointsToEmu(s.table.summaryRowBorderRightWidth != null ? s.table.summaryRowBorderRightWidth : s.table.borderWidth) }
      },
      tableBorder: {
        color: s.table.borderColor || s.colors.border,
        width: U.pointsToEmu(s.table.borderWidth),
        mode: s.table.borderMode === 'grid' || s.table.borderMode === 'none' ? s.table.borderMode : 'horizontal'
      },
      // New categories (previously hardcoded)
      slide: {
        background: s.slide.background
      },
      headerBar: {
        textColor: s.colors.headerBarText
      },
      header: {
        fontSize: s.header.size != null ? U.pointsToFontSize(s.header.size) : null,
        color: s.header.color,
        background: s.header.background,
        align: s.header.align,
        borderRadius: s.header.borderRadius,
        fontFamily: s.header.fontFamily,
        bold: s.header.bold,
        italic: s.header.italic
      },
      subheader: {
        fontSize: s.subheader.size != null ? U.pointsToFontSize(s.subheader.size) : null,
        bold: s.subheader.bold,
        color: s.subheader.color,
        align: s.subheader.align,
        lineColor: s.subheader.lineColor,
        lineWidth: Math.round(U.pointsToEmu(s.subheader.lineWidth)),
        fontFamily: s.subheader.fontFamily,
        italic: s.subheader.italic
      },
      subfooter: {
        fontSize: s.subfooter.size != null ? U.pointsToFontSize(s.subfooter.size) : null,
        italic: s.subfooter.italic,
        color: s.subfooter.color,
        fontFamily: s.subfooter.fontFamily
      },
      card: {
        fill: s.colors.cardFill,
        borderWidth: Math.round(U.pointsToEmu(s.spacing.cardBorderWidth)),
        paddingLeft: Math.round(U.inchesToEmu(s.spacing.cardPaddingLeft)),
        paddingTop: Math.round(U.inchesToEmu(s.spacing.cardPaddingTop))
      },
      spacing: {
        columnGap: Math.round(U.inchesToEmu(s.spacing.columnGap)),
        bulletSpacing: Math.round(U.inchesToEmu(s.spacing.bulletSpacing)),
        cardGap: Math.round(U.inchesToEmu(s.spacing.cardGap)),
        headerBarHeight: Math.round(U.inchesToEmu(s.spacing.headerBarHeight)),
        contentGap: Math.round(U.inchesToEmu(s.spacing.contentGap)),
        zoneGap: Math.round(U.inchesToEmu(s.spacing.zoneGap)),
        subheaderLineWidth: Math.round(U.pointsToEmu(s.spacing.subheaderLineWidth))
      },
      titleSlide: {
        titleSize: U.pointsToFontSize(s.titleSlide.titleSize),
        subtitleSize: U.pointsToFontSize(s.titleSlide.subtitleSize)
      },
      statGrid: {
        valueFont: s.statGrid.valueFontFamily || s.fonts.title,
        valueSize: U.pointsToFontSize(s.statGrid.valueSize),
        valueColor: s.statGrid.valueColor || accent,
        valueBold: s.statGrid.valueBold,
        valueItalic: s.statGrid.valueItalic,
        labelFont: s.statGrid.labelFontFamily || s.fonts.body,
        labelSize: s.statGrid.labelSize != null ? U.pointsToFontSize(s.statGrid.labelSize) : null,
        labelColor: s.statGrid.labelColor || primary,
        labelBold: s.statGrid.labelBold,
        labelItalic: s.statGrid.labelItalic,
        sublabelFont: s.statGrid.sublabelFontFamily || s.fonts.body,
        sublabelSize: s.statGrid.sublabelSize != null ? U.pointsToFontSize(s.statGrid.sublabelSize) : null,
        sublabelColor: s.statGrid.sublabelColor || secondary,
        sublabelBold: s.statGrid.sublabelBold,
        sublabelItalic: s.statGrid.sublabelItalic
      },
      cardGrid: {
        titleFont: s.cardGrid.titleFontFamily || s.fonts.body,
        titleSize: s.cardGrid.titleSize != null ? U.pointsToFontSize(s.cardGrid.titleSize) : null,
        titleColor: s.cardGrid.titleColor || accent,
        titleBold: s.cardGrid.titleBold,
        titleItalic: s.cardGrid.titleItalic,
        lineFont: s.cardGrid.lineFontFamily || s.fonts.body,
        lineSize: s.cardGrid.lineSize != null ? U.pointsToFontSize(s.cardGrid.lineSize) : null,
        lineColor: s.cardGrid.lineColor || primary,
        lineBold: s.cardGrid.lineBold,
        lineItalic: s.cardGrid.lineItalic,
        lineBullet: s.cardGrid.lineBullet || false,
        lineBulletIndent: Math.round(U.inchesToEmu(s.cardGrid.lineBulletIndent || 0.05))
      },
      chart: {
        colors: s.chart.colors || null,
        fontFamily: s.chart.fontFamily || null,
        fontSize: s.chart.fontSize || null,
        barBorderColor: s.chart.barBorderColor || null,
        barBorderWidth: s.chart.barBorderWidth != null ? s.chart.barBorderWidth : null,
        bar: s.chart.bar || null,
        pie: s.chart.pie || null,
        combo: s.chart.combo || null
      }
    };
  }

  /** New SlideStyle with overrides merged on top. */
  merge(overrides: StyleOverrides): SlideStyle {
    return new SlideStyle(deepMerge(this._style, overrides));
  }

  // ---- Built-in presets ----
  static presets: Record<string, StyleOverrides> = {
    corporate: {},  // empty = all defaults (the IB-safe look)
  
    minimal: {
      colors: {
        primary: '#2c3e50', secondary: '#7f8c8d', accent: '#2980b9',
        darkFill: '#2c3e50', border: '#ecf0f1'
      },
      fonts: { title: 'Helvetica Neue', body: 'Helvetica Neue', label: 'Helvetica Neue', dividerNumber: 'Helvetica Neue' },
      spacing: { zoneGap: 0.2, cardGap: 0.1 }
    },
  
    dark: {
      slide: { background: '#1a1a2e' },
      title: { color: '#EAEAEA' },   // DEFAULT_STYLE.title.color is explicit slate, not cascaded: unreadable on the dark background
      colors: {
        primary: '#EAEAEA', secondary: '#A0A0A0', accent: '#E94560',
        darkFill: '#16213e', border: '#444444',
        headerBarText: '#FFFFFF', cardFill: '#16213e'
      },
      fonts: { title: 'Arial Bold', body: 'Arial', label: 'Arial', dividerNumber: 'Arial' },
      table: { evenRowFill: '#16213e', oddRowFill: '#1a1a2e', headerTextColor: '#FFFFFF', borderColor: '#33405c', summaryRowFill: '#243357', summaryRowTextColor: '#EAEAEA' }
    },
  
    warm: {
      colors: {
        primary: '#4a3728', secondary: '#8b7355', accent: '#D57F5B',
        darkFill: '#4a3728', border: '#d4c5a9'
      },
      fonts: { title: 'Georgia', body: 'Garamond', label: 'Georgia', dividerNumber: 'Georgia' }
    }
  };

  /** Style from a preset name. Throws on unknown names. */
  static from(name: string): SlideStyle {
    const preset = SlideStyle.presets[name];
    if (!preset) throw new Error(`Unknown style preset: ${name}. Known presets: ${Object.keys(SlideStyle.presets).join(", ")}`);
    return new SlideStyle(preset);
  }

  static defaults = DEFAULT_STYLE;
}

export default SlideStyle;
