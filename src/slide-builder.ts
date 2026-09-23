import DM from "./pptx/document-model.js";
import U from "./units.js";
import PptxParser from "./pptx/pptx-parser.js";
import PptxWriter from "./pptx/pptx-writer.js";
import LayoutAnalyzer from "./pptx/layout-analyzer.js";
import LC from "./pptx/layout-catalog.js";
import SlideLayout from "./layout.js";
import Defaults from "./defaults.js";
import { fitTitle } from "./fit.js";
import SlideStyle from "./style.js";
import { convertToNative } from "./charts/chart-to-native.js";
import ChartGenerator from "./charts/chart-generator.js";
import { generateChartParts } from "./charts/pptxchart-generator.js";
import type { Bounds, SlideSpec, GridNode, StyleTokens } from "./types.js";

// ---- Statistical helpers ----
function _mostCommon(arr: any[]): any {
  const counts = {};
  for (let i = 0; i < arr.length; i++) {
    if (arr[i] != null && arr[i] !== '') {
      counts[arr[i]] = (counts[arr[i]] || 0) + 1;
    }
  }
  let best: any = null, bestCount = 0;
  for (const k in counts) {
    if (counts[k] > bestCount) { best = k; bestCount = counts[k]; }
  }
  return best;
}

function _medianNum(arr: number[]): number {
  const nums = arr.filter(function(n) { return n > 0; }).sort(function(a, b) { return a - b; });
  if (nums.length === 0) return 0;
  return nums[Math.floor(nums.length / 2)];
}

// Responsive column count based on container aspect ratio
function _responsiveCols(itemCount: number, cx: number, cy: number): number {
  if (itemCount <= 1) return 1;
  const aspect = cy > 0 ? cx / cy : 999;
  if (itemCount === 2) return aspect > 1 ? 2 : 1;
  if (itemCount <= 4) {
    if (aspect > 3) return itemCount;
    if (aspect > 1) return 2;
    return 1;
  }
  if (aspect > 3) return itemCount;
  if (aspect > 1.5) return 3;
  if (aspect > 0.8) return 2;
  return 1;
}

// ---- Read image dimensions from buffer header (PNG, JPEG, WebP) ----
function readImageDimensions(buffer: any): { width: number; height: number } | null {
  if (!buffer || buffer.length < 24) return null;

  // PNG: signature 89 50 4E 47, IHDR at offset 16/20
  if (buffer[0] === 0x89 && buffer[1] === 0x50 && buffer[2] === 0x4E && buffer[3] === 0x47) {
    return { width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20) };
  }

  // JPEG: starts with FF D8, scan for SOF markers (C0, C1, C2, C3)
  if (buffer[0] === 0xFF && buffer[1] === 0xD8) {
    let i = 2;
    while (i + 1 < buffer.length) {
      if (buffer[i] !== 0xFF) break;
      const marker = buffer[i + 1];
      // Skip padding FF bytes
      if (marker === 0xFF) { i++; continue; }
      // SOF0-SOF3: contains dimensions
      if (marker >= 0xC0 && marker <= 0xC3 && i + 9 < buffer.length) {
        const height = buffer.readUInt16BE(i + 5);
        const width = buffer.readUInt16BE(i + 7);
        if (width > 0 && height > 0) return { width: width, height: height };
      }
      // Skip to next marker (2 byte marker + 2 byte length + payload)
      if (i + 3 < buffer.length) {
        const segLen = buffer.readUInt16BE(i + 2);
        i += 2 + segLen;
      } else {
        break;
      }
    }
    return null;
  }

  // WebP: RIFF header, 'WEBP' at offset 8
  if (buffer[0] === 0x52 && buffer[1] === 0x49 && buffer[2] === 0x46 && buffer[3] === 0x46 &&
      buffer[8] === 0x57 && buffer[9] === 0x45 && buffer[10] === 0x42 && buffer[11] === 0x50) {
    // VP8 lossy: 'VP8 ' at offset 12
    if (buffer[12] === 0x56 && buffer[13] === 0x50 && buffer[14] === 0x38 && buffer[15] === 0x20 && buffer.length >= 30) {
      return { width: buffer.readUInt16LE(26) & 0x3FFF, height: buffer.readUInt16LE(28) & 0x3FFF };
    }
    // VP8L lossless: 'VP8L' at offset 12
    if (buffer[12] === 0x56 && buffer[13] === 0x50 && buffer[14] === 0x38 && buffer[15] === 0x4C && buffer.length >= 25) {
      const bits = buffer.readUInt32LE(21);
      return { width: (bits & 0x3FFF) + 1, height: ((bits >> 14) & 0x3FFF) + 1 };
    }
    // VP8X extended: 'VP8X' at offset 12, canvas size at offset 24 (3+3 bytes, little-endian)
    if (buffer[12] === 0x56 && buffer[13] === 0x50 && buffer[14] === 0x38 && buffer[15] === 0x58 && buffer.length >= 30) {
      const xw = (buffer[24] | (buffer[25] << 8) | (buffer[26] << 16)) + 1;
      const xh = (buffer[27] | (buffer[28] << 8) | (buffer[29] << 16)) + 1;
      return { width: xw, height: xh };
    }
  }

  return null;
}

function _lightenHex(hex: string, amount: number): string {
  hex = hex.replace('#', '');
  let r = parseInt(hex.substring(0, 2), 16);
  let g = parseInt(hex.substring(2, 4), 16);
  let b = parseInt(hex.substring(4, 6), 16);
  r = Math.round(r + (255 - r) * amount);
  g = Math.round(g + (255 - g) * amount);
  b = Math.round(b + (255 - b) * amount);
  return '#' + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
}

const DEBUG_COLORS = ['#E53935', '#1E88E5', '#43A047', '#FB8C00', '#8E24AA', '#00ACC1'];

// ===== SlideBuilder =====
// High-level API for AI agents to programmatically build presentations.
// When loaded from a reference .pptx, deeply analyzes visual patterns
// (fonts, colors, sizes, positioning) and replicates them faithfully.

class SlideBuilder {
  pres: any;
  _shapeIdCounter: number;
  _refMedia: any;
  _catalog: any;
  _zoneBorders: boolean;
  _style: any;

  constructor(options?: any) {
    this.pres = new DM.Presentation();
    this.pres.theme = new DM.Theme();
    this._shapeIdCounter = 1;
    this._refMedia = null;
    this._catalog = null;
    this._zoneBorders = !!(options && options.zoneBorders);

    // Default style tokens (overridden by reference analysis or custom SlideStyle)
    this._style = {
      sectionLabel: {
        font: 'Calibri', fontSize: 1200, bold: true, color: '#4472C4',
        x: 457200, y: 393192
      },
      mainTitle: {
        font: 'Arial Bold', fontSize: 2800, bold: true, color: '#405363',
        x: 457200, y: 822960
      },
      bodyText: {
        font: 'Calibri', fontSize: 1400, color: '#333333', align: 'l'
      },
      bullet: {
        char: '\u2022', indent: -342900, marginLeft: 342900,
        bulletFont: null, bulletColor: null,
        font: null, fontSize: 1400, color: '#333333'
      },
      footnote: {
        font: 'Calibri', fontSize: 800, color: '#666666', y: 0
      },
      divider: {
        numberFont: 'Calibri Light', numberFontSize: 11500, numberColor: '#333333',
        numberX: 457200, numberY: 2194560,
        titleFont: 'Calibri', titleFontSize: 2400, titleColor: '#4472C4',
        titleY: 0
      },
      palette: {
        primaryText: '#333333',
        secondaryText: '#666666',
        accent: '#4472C4',
        darkFill: '#164556',
        border: '#BFBFBF'
      },
      tableHeader: {
        fillColor: '#4472C4', textColor: '#FFFFFF', font: null, fontSize: 1200, bold: true
      },
      tableBody: {
        alternateRows: true, fillEven: '#F2F2F2', fillOdd: '#FFFFFF', textColor: '#333333', font: null, fontSize: 1100
      },
      tableBorder: { color: '#BFBFBF', width: 9525 },
      // New categories (previously hardcoded values)
      slide: { background: '#FFFFFF' },
      headerBar: { textColor: '#FFFFFF' },
      header: { fontSize: null, color: null, background: null, align: 'ctr' },
      subheader: { fontSize: null, bold: true, color: null, align: 'ctr', lineColor: null, lineWidth: 9525 },
      subfooter: { fontSize: null, italic: true, color: null },
      card: { fill: '#FFFFFF', borderWidth: 9525, paddingLeft: 68580, paddingTop: 45720 },
      spacing: {
        columnGap: 228600, bulletSpacing: 114300, cardGap: 73152,
        headerBarHeight: 320040, contentGap: 91440, zoneGap: 137160,
        subheaderLineWidth: 9525
      },
      titleSlide: { titleSize: 3600, subtitleSize: 1800 }
    };

    if (options) {
      if (options.slideSize) {
        if (options.slideSize.width) this.pres.slideSize.cx = U.inchesToEmu(options.slideSize.width);
        if (options.slideSize.height) this.pres.slideSize.cy = U.inchesToEmu(options.slideSize.height);
      }
      if (options.style) {
        const styleObj = options.style instanceof SlideStyle ? options.style : new SlideStyle(options.style);
        this._style = styleObj.toInternal();
      }
    }
  }

  // ---- Static factory: load from reference .pptx ----
  static async fromReference(pptxPath: string, options?: any): Promise<SlideBuilder> {
    const parser = new PptxParser();
    const refPres = await parser.parseFile(pptxPath);
    const builder = new SlideBuilder();
    builder._applyReferenceAnalysis(refPres);
    // Apply custom style overrides on top of reference-extracted styles
    if (options && options.style) {
      const styleObj = options.style instanceof SlideStyle ? options.style : new SlideStyle(options.style);
      const overrides = styleObj.toInternal();
      for (const key in overrides) {
        if (typeof overrides[key] === 'object' && overrides[key] !== null) {
          if (!builder._style[key]) builder._style[key] = {};
          for (const sk in overrides[key]) {
            if (overrides[key][sk] !== undefined) builder._style[key][sk] = overrides[key][sk];
          }
        } else if (overrides[key] !== undefined) {
          builder._style[key] = overrides[key];
        }
      }
    }
    return builder;
  }

  // ---- Static factory: load from reference .pptx with layout analysis ----
  static async fromReferenceWithLayouts(pptxPath: string, options?: any): Promise<SlideBuilder> {
    const parser = new PptxParser();
    const refPres = await parser.parseFile(pptxPath);
    const builder = new SlideBuilder();
    builder._applyReferenceAnalysis(refPres);
    if (options && options.style) {
      const styleObj = options.style instanceof SlideStyle ? options.style : new SlideStyle(options.style);
      const overrides = styleObj.toInternal();
      for (const key in overrides) {
        if (typeof overrides[key] === 'object' && overrides[key] !== null) {
          if (!builder._style[key]) builder._style[key] = {};
          for (const sk in overrides[key]) {
            if (overrides[key][sk] !== undefined) builder._style[key][sk] = overrides[key][sk];
          }
        } else if (overrides[key] !== undefined) {
          builder._style[key] = overrides[key];
        }
      }
    }
    const analyzer = new LayoutAnalyzer();
    builder._catalog = analyzer.analyze(refPres);
    builder._catalog.metadata.sourceFile = pptxPath;
    return builder;
  }

  // ---- Reference analysis: deep pattern extraction ----
  _applyReferenceAnalysis(refPres: any) {
    // Copy theme
    if (refPres.theme) {
      this.pres.theme = refPres.theme;
    }

    // Copy slide dimensions
    this.pres.slideSize = { cx: refPres.slideSize.cx, cy: refPres.slideSize.cy };

    const slideH = refPres.slideSize.cy;
    const slideHIn = slideH / 914400;

    // Theme font resolution
    const majorLatin = this.pres.theme.fontScheme.majorFont.latin || 'Calibri Light';
    const minorLatin = this.pres.theme.fontScheme.minorFont.latin || 'Calibri';

    function resolveFont(name) {
      if (!name) return null;
      if (name === '+mj-lt' || name === '+mj-ea' || name === '+mj-cs') return majorLatin;
      if (name === '+mn-lt' || name === '+mn-ea' || name === '+mn-cs') return minorLatin;
      return name;
    }

    // ---- Pass 1: Collect all text runs with context ----
    const allRuns: any[] = [];
    const dividerSlideIndices = {};

    for (let si = 0; si < refPres.slides.length; si++) {
      const slide = refPres.slides[si];
      for (let shi = 0; shi < slide.shapes.length; shi++) {
        const shape = slide.shapes[shi];
        if (!shape.textBody) continue;

        const topIn = shape.xfrm.off.y / 914400;
        const bottomIn = (shape.xfrm.off.y + shape.xfrm.ext.cy) / 914400;

        for (let pi = 0; pi < shape.textBody.paragraphs.length; pi++) {
          const para = shape.textBody.paragraphs[pi];

          for (let ri = 0; ri < para.runs.length; ri++) {
            const run = para.runs[ri];
            if (!run.text || !run.text.trim()) continue;

            const info = {
              font: resolveFont(run.rPr.fontFamily),
              bold: run.rPr.b,
              color: run.rPr.color,
              size: run.rPr.sz,
              x: shape.xfrm.off.x,
              y: shape.xfrm.off.y,
              topIn: topIn,
              bottomIn: bottomIn,
              slideIdx: si,
              para: para
            };

            allRuns.push(info);

            // Mark slides that contain very large text (section divider numbers)
            if (info.size >= 8000) {
              dividerSlideIndices[si] = true;
            }
          }
        }
      }
    }

    // ---- Pass 2: Categorize runs ----
    const labelRuns: any[] = [];      // section labels (small text at top of slide)
    const headerRuns: any[] = [];     // main titles (large text near top)
    const bodyRuns: any[] = [];       // body content text
    const footnoteRuns: any[] = [];   // small text at bottom
    const dividerNumbers: any[] = []; // very large section numbers
    const dividerTitles: any[] = [];  // section titles on divider slides
    const bulletParas: any[] = [];    // paragraphs with bullet markers (deduplicated)
    const seenParas = new Set ? new Set() : {};

    for (let i = 0; i < allRuns.length; i++) {
      const r = allRuns[i];

      if (r.size >= 8000) {
        dividerNumbers.push(r);
        continue;
      }

      // Runs on divider slides: medium text is the section title
      if (dividerSlideIndices[r.slideIdx] && r.size >= 2000 && r.size <= 4000) {
        dividerTitles.push(r);
        continue;
      }

      // Section label (eyebrow): small text at very top of slide
      if (r.topIn < 0.8 && r.size > 0 && r.size <= 1400) {
        labelRuns.push(r);
        continue;
      }

      // Main title: large text near top
      if (r.topIn < 1.5 && r.size >= 2400) {
        headerRuns.push(r);
        continue;
      }

      // Footnote: small text at bottom
      if (r.bottomIn > slideHIn - 1.5 && r.size > 0 && r.size <= 1000) {
        footnoteRuns.push(r);
        continue;
      }

      // Everything else is body text (in the content zone, reasonable size)
      if (r.size >= 1050 && r.topIn >= 1.0 && r.bottomIn < slideHIn - 1.0) {
        bodyRuns.push(r);
      }

      // Track bullet paragraphs (avoid counting same para multiple times)
      const p = r.para;
      if (p && (p.pPr.buChar || p.pPr.buAutoNum) && !p.pPr.buNone) {
        const paraKey = r.slideIdx + ':' + r.y + ':' + (p.runs[0] ? p.runs[0].text : '');
        if (seenParas instanceof Set) {
          if (!seenParas.has(paraKey)) {
            seenParas.add(paraKey);
            bulletParas.push(p);
          }
        } else {
          if (!seenParas[paraKey]) {
            seenParas[paraKey] = true;
            bulletParas.push(p);
          }
        }
      }
    }

    // ---- Pass 3: Extract style tokens ----
    const style = this._style;

    // Section label (eyebrow)
    if (labelRuns.length >= 3) {
      style.sectionLabel = {
        font: _mostCommon(labelRuns.map(function(r) { return r.font; })) || style.sectionLabel.font,
        fontSize: parseInt(_mostCommon(labelRuns.map(function(r) { return r.size; }))) || style.sectionLabel.fontSize,
        bold: labelRuns.filter(function(r) { return r.bold; }).length > labelRuns.length / 2,
        color: _mostCommon(labelRuns.map(function(r) { return r.color; })) || style.sectionLabel.color,
        x: _medianNum(labelRuns.map(function(r) { return r.x; })) || style.sectionLabel.x,
        y: _medianNum(labelRuns.map(function(r) { return r.y; })) || style.sectionLabel.y
      };
    }

    // Main title
    if (headerRuns.length >= 3) {
      style.mainTitle = {
        font: _mostCommon(headerRuns.map(function(r) { return r.font; })) || style.mainTitle.font,
        fontSize: parseInt(_mostCommon(headerRuns.map(function(r) { return r.size; }))) || style.mainTitle.fontSize,
        bold: headerRuns.filter(function(r) { return r.bold; }).length > headerRuns.length / 2,
        color: _mostCommon(headerRuns.map(function(r) { return r.color; })) || style.mainTitle.color,
        x: _medianNum(headerRuns.map(function(r) { return r.x; })) || style.mainTitle.x,
        y: _medianNum(headerRuns.map(function(r) { return r.y; })) || style.mainTitle.y
      };
    }

    // Body text
    if (bodyRuns.length >= 5) {
      style.bodyText = {
        font: _mostCommon(bodyRuns.map(function(r) { return r.font; })) || style.bodyText.font,
        fontSize: parseInt(_mostCommon(bodyRuns.map(function(r) { return r.size; }))) || style.bodyText.fontSize,
        color: _mostCommon(bodyRuns.map(function(r) { return r.color; })) || style.bodyText.color,
        align: 'l'
      };
    }

    // Footnotes
    if (footnoteRuns.length >= 2) {
      style.footnote = {
        font: _mostCommon(footnoteRuns.map(function(r) { return r.font; })) || style.footnote.font,
        fontSize: parseInt(_mostCommon(footnoteRuns.map(function(r) { return r.size; }))) || style.footnote.fontSize,
        color: _mostCommon(footnoteRuns.map(function(r) { return r.color; })) || style.footnote.color,
        y: _medianNum(footnoteRuns.map(function(r) { return r.y; }))
      };
    }

    // Section divider numbers
    if (dividerNumbers.length >= 2) {
      style.divider.numberFont = _mostCommon(dividerNumbers.map(function(r) { return r.font; })) || style.divider.numberFont;
      style.divider.numberFontSize = parseInt(_mostCommon(dividerNumbers.map(function(r) { return r.size; }))) || style.divider.numberFontSize;
      style.divider.numberColor = _mostCommon(dividerNumbers.map(function(r) { return r.color; })) || style.divider.numberColor;
      style.divider.numberX = _medianNum(dividerNumbers.map(function(r) { return r.x; })) || style.divider.numberX;
      style.divider.numberY = _medianNum(dividerNumbers.map(function(r) { return r.y; })) || style.divider.numberY;
    }

    // Section divider titles
    if (dividerTitles.length >= 2) {
      style.divider.titleFont = _mostCommon(dividerTitles.map(function(r) { return r.font; })) || style.divider.titleFont;
      style.divider.titleFontSize = parseInt(_mostCommon(dividerTitles.map(function(r) { return r.size; }))) || style.divider.titleFontSize;
      style.divider.titleColor = _mostCommon(dividerTitles.map(function(r) { return r.color; })) || style.divider.titleColor;
      style.divider.titleY = _medianNum(dividerTitles.map(function(r) { return r.y; })) || style.divider.titleY;
    }

    // Bullet styling
    if (bulletParas.length >= 3) {
      const buChars = bulletParas.map(function(p) { return p.pPr.buChar; }).filter(Boolean);
      const buIndents = bulletParas.map(function(p) { return Math.abs(p.pPr.indent || 0); }).filter(function(n) { return n > 0; });
      const buMarLs = bulletParas.map(function(p) { return p.pPr.marL || 0; }).filter(function(n) { return n > 0; });
      const buFonts = bulletParas.map(function(p) { return p.pPr.buFont; }).filter(Boolean);
      const buColors = bulletParas.map(function(p) { return p.pPr.buClr; }).filter(Boolean);

      // First run of each bullet paragraph for text styling
      const bulletFirstRuns = bulletParas.map(function(p) {
        return p.runs.length > 0 ? p.runs[0] : null;
      }).filter(Boolean);

      style.bullet = {
        char: _mostCommon(buChars) || style.bullet.char,
        indent: -(_medianNum(buIndents) || Math.abs(style.bullet.indent)),
        marginLeft: _medianNum(buMarLs) || style.bullet.marginLeft,
        bulletFont: _mostCommon(buFonts) || null,
        bulletColor: _mostCommon(buColors) || null,
        font: resolveFont(_mostCommon(bulletFirstRuns.map(function(r) { return r.rPr.fontFamily; }))) || style.bodyText.font,
        fontSize: parseInt(_mostCommon(bulletFirstRuns.map(function(r) { return r.rPr.sz; }))) || style.bodyText.fontSize,
        color: _mostCommon(bulletFirstRuns.map(function(r) { return r.rPr.color; })) || style.bodyText.color
      };
    }

    // ---- Color palette from actual usage ----
    const colorCounts = {};
    allRuns.forEach(function(r) {
      if (r.color) colorCounts[r.color] = (colorCounts[r.color] || 0) + 1;
    });
    const sortedColors = Object.keys(colorCounts).sort(function(a, b) { return colorCounts[b] - colorCounts[a]; });

    style.palette.primaryText = sortedColors[0] || style.palette.primaryText;
    if (sortedColors.length > 1) style.palette.secondaryText = sortedColors[1];
    style.palette.accent = style.sectionLabel.color || (sortedColors.length > 2 ? sortedColors[2] : style.palette.accent);

    // Derive table styling from main style
    style.tableHeader.font = style.mainTitle.font;
    style.tableHeader.fillColor = style.palette.accent;
    style.tableBody.font = style.bodyText.font;
    style.tableBody.textColor = style.palette.primaryText;

    this._refMedia = refPres.media;
  }

  // ---- Computed layout helpers ----
  _slideW(): number { return this.pres.slideSize.cx; }

  _slideH(): number { return this.pres.slideSize.cy; }

  _marginLeft(): number { return this._style.sectionLabel.x || 457200; }

  _marginRight(): number { return this._marginLeft(); }

  _contentTop(): number {
    const titleY = this._style.mainTitle.y;
    const titleSizePt = this._style.mainTitle.fontSize / 100;
    const titleHeightEmu = Math.round(titleSizePt * 914400 / 72 * 1.4);
    return titleY + titleHeightEmu + 182880; // title bottom + 0.2" gap
  }

  _bodyWidth(): number {
    return this._slideW() - this._marginLeft() - this._marginRight();
  }

  _bodyHeight(): number {
    const fnY = this._style.footnote.y;
    const bottomLimit = fnY > 0 ? fnY - 91440 : this._slideH() - 457200;
    return bottomLimit - this._contentTop();
  }

  // ---- Shape ID ----
  _nextId(): number {
    return this._shapeIdCounter++;
  }

  // ---- Run factory ----
  _makeRun(text: string, opts: any) {
    const run = new DM.Run(text);
    run.rPr = {
      sz: opts.fontSize != null ? opts.fontSize : 1800,
      b: opts.bold || false,
      i: opts.italic || false,
      u: opts.underline || 'none',
      strike: 'noStrike',
      cap: opts.cap || null,
      color: opts.color || this._style.palette.primaryText,
      highlight: null,
      fontFamily: opts.fontFamily || this._style.bodyText.font,
      baseline: 0
    };
    return run;
  }

  // ---- Core shape builder ----
  _createTextShape(name: string, text: string | string[], rect: Bounds, opts: any) {
    const shape = new DM.Shape();
    shape.id = this._nextId();
    shape.name = name;
    shape.xfrm = {
      off: { x: rect.x, y: rect.y },
      ext: { cx: rect.cx, cy: rect.cy },
      rot: 0, flipH: false, flipV: false
    };
    shape.geometry = { type: 'rect' };
    shape.fill = opts.fill || new DM.Fill('none');
    if (opts.line) shape.line = opts.line;
    shape.textBody = new DM.TextBody();
    shape.textBody.bodyPr.anchor = opts.anchor || 't';

    const lines = Array.isArray(text) ? text : [text];
    for (let i = 0; i < lines.length; i++) {
      const para = new DM.Paragraph();
      para.pPr.algn = opts.align || 'l';
      if (i > 0 && opts.spacing) para.pPr.spcBef = opts.spacing;
      const run = this._makeRun(lines[i], opts);
      para.runs.push(run);
      shape.textBody.paragraphs.push(para);
    }

    return shape;
  }

  // ---- Two-part header (section label + main title) ----
  _addSlideHeader(slide: any, sectionLabel: string | null, title: string) {
    const s = this._style;
    const bodyW = this._bodyWidth();

    // Section label (eyebrow text at top)
    if (sectionLabel) {
      const labelH = Math.round(s.sectionLabel.fontSize / 100 * 914400 / 72 * 1.6);
      const labelShape = this._createTextShape(
        'Section Label', sectionLabel,
        { x: s.sectionLabel.x, y: s.sectionLabel.y, cx: bodyW, cy: labelH },
        {
          fontSize: s.sectionLabel.fontSize,
          bold: s.sectionLabel.bold,
          italic: s.sectionLabel.italic || false,
          color: s.sectionLabel.color,
          fontFamily: s.sectionLabel.font,
          align: 'l', anchor: 't'
        }
      );
      slide.shapes.push(labelShape);
    }

    // Main title
    const titleH = Math.round(s.mainTitle.fontSize / 100 * 914400 / 72 * 1.6);
    const titleShape = this._createTextShape(
      'Title', title,
      { x: s.mainTitle.x, y: s.mainTitle.y, cx: bodyW, cy: titleH },
      {
        fontSize: s.mainTitle.fontSize,
        bold: s.mainTitle.bold,
        color: s.mainTitle.color,
        fontFamily: s.mainTitle.font,
        align: 'l', anchor: 't'
      }
    );
    slide.shapes.push(titleShape);
  }

  // ---- Footnote shape ----
  _addFootnote(slide: any, text: string) {
    const s = this._style;
    const fnY = s.footnote.y > 0 ? s.footnote.y : (this._slideH() - 365760);
    const fnH = Math.round(s.footnote.fontSize / 100 * 914400 / 72 * 2.5);

    const shape = this._createTextShape(
      'Footnote', text,
      { x: this._marginLeft(), y: fnY, cx: this._bodyWidth(), cy: fnH },
      {
        fontSize: s.footnote.fontSize,
        color: s.footnote.color,
        fontFamily: s.footnote.font,
        align: 'l', anchor: 'b'
      }
    );
    slide.shapes.push(shape);
  }

  _addSlideNumber(slide: any) {
    const s = this._style;
    const numW = 457200;  // 0.5 inches
    const numH = 228600;  // 0.25 inches
    const shape = this._createTextShape(
      'Slide Number', String(this.pres.slides.length + 1),
      {
        x: this._slideW() - this._marginRight() - numW,
        y: this._slideH() - numH - 91440,   // 0.1" from bottom
        cx: numW,
        cy: numH
      },
      {
        fontSize: 800,
        color: s.footnote.color,
        fontFamily: s.footnote.font,
        align: 'r', anchor: 'b'
      }
    );
    slide.shapes.push(shape);
  }

  // ===== PUBLIC SLIDE METHODS (all return `this` for chaining) =====

  // Content slide with two-part header + bullet list
  addContentSlide(title: string, bullets: any[], options?: any): this {
    const opts = options || {};
    const slide = new DM.Slide(this.pres.slides.length + 1);
    const s = this._style;

    slide.background = new DM.Fill('solid', { color: this._style.slide.background });
    this._addSlideHeader(slide, opts.sectionLabel || null, title);

    const contentTop = this._contentTop();
    const bodyShape = new DM.Shape();
    bodyShape.id = this._nextId();
    bodyShape.name = 'Content';
    bodyShape.xfrm = {
      off: { x: this._marginLeft(), y: contentTop },
      ext: { cx: this._bodyWidth(), cy: this._bodyHeight() },
      rot: 0, flipH: false, flipV: false
    };
    bodyShape.geometry = { type: 'rect' };
    bodyShape.fill = new DM.Fill('none');
    bodyShape.textBody = new DM.TextBody();
    bodyShape.textBody.bodyPr.anchor = 't';

    for (let i = 0; i < bullets.length; i++) {
      const para = new DM.Paragraph();
      para.pPr.algn = s.bodyText.align || 'l';
      para.pPr.buNone = false;
      para.pPr.buChar = s.bullet.char;
      if (s.bullet.bulletFont) para.pPr.buFont = s.bullet.bulletFont;
      if (s.bullet.bulletColor) para.pPr.buClr = s.bullet.bulletColor;
      para.pPr.marL = Math.abs(s.bullet.marginLeft);
      para.pPr.indent = s.bullet.indent;
      if (i > 0) para.pPr.spcBef = s.spacing.bulletSpacing || 114300;

      const run = this._makeRun(
        typeof bullets[i] === 'string' ? bullets[i] : String(bullets[i]),
        {
          fontSize: s.bullet.fontSize,
          color: s.bullet.color,
          fontFamily: s.bullet.font || s.bodyText.font
        }
      );
      para.runs.push(run);
      bodyShape.textBody.paragraphs.push(para);
    }

    slide.shapes.push(bodyShape);
    if (opts.footnote) this._addFootnote(slide, opts.footnote);

    this.pres.slides.push(slide);
    return this;
  }

  // Two-column layout with header
  addTwoColumnSlide(title: string, leftContent: any, rightContent: any, options?: any): this {
    const opts = options || {};
    const slide = new DM.Slide(this.pres.slides.length + 1);
    const s = this._style;

    slide.background = new DM.Fill('solid', { color: this._style.slide.background });
    this._addSlideHeader(slide, opts.sectionLabel || null, title);

    const contentTop = this._contentTop();
    const colGap = s.spacing.columnGap || 228600;
    const colWidth = Math.round((this._bodyWidth() - colGap) / 2);
    const colHeight = this._bodyHeight();
    const leftX = this._marginLeft();
    const rightX = leftX + colWidth + colGap;

    const leftLines = Array.isArray(leftContent) ? leftContent : [leftContent];
    const leftShape = this._createTextShape(
      'Left Column', leftLines,
      { x: leftX, y: contentTop, cx: colWidth, cy: colHeight },
      {
        fontSize: s.bodyText.fontSize, color: s.bodyText.color,
        fontFamily: s.bodyText.font, align: 'l', anchor: 't', spacing: s.spacing.bulletSpacing || 114300
      }
    );
    slide.shapes.push(leftShape);

    const rightLines = Array.isArray(rightContent) ? rightContent : [rightContent];
    const rightShape = this._createTextShape(
      'Right Column', rightLines,
      { x: rightX, y: contentTop, cx: colWidth, cy: colHeight },
      {
        fontSize: s.bodyText.fontSize, color: s.bodyText.color,
        fontFamily: s.bodyText.font, align: 'l', anchor: 't', spacing: s.spacing.bulletSpacing || 114300
      }
    );
    slide.shapes.push(rightShape);

    if (opts.footnote) this._addFootnote(slide, opts.footnote);

    this.pres.slides.push(slide);
    return this;
  }

  // Company profile slide with key-value pairs
  addCompanyProfileSlide(name: string, details: Record<string, any>, options?: any): this {
    const opts = options || {};
    const slide = new DM.Slide(this.pres.slides.length + 1);
    const s = this._style;

    slide.background = new DM.Fill('solid', { color: this._style.slide.background });
    this._addSlideHeader(slide, opts.sectionLabel || null, name);

    const contentTop = this._contentTop();
    const keys = Object.keys(details);

    const bodyShape = new DM.Shape();
    bodyShape.id = this._nextId();
    bodyShape.name = 'Company Details';
    bodyShape.xfrm = {
      off: { x: this._marginLeft(), y: contentTop },
      ext: { cx: this._bodyWidth(), cy: this._bodyHeight() },
      rot: 0, flipH: false, flipV: false
    };
    bodyShape.geometry = { type: 'rect' };
    bodyShape.fill = new DM.Fill('none');
    bodyShape.textBody = new DM.TextBody();
    bodyShape.textBody.bodyPr.anchor = 't';

    for (let i = 0; i < keys.length; i++) {
      const para = new DM.Paragraph();
      para.pPr.algn = 'l';
      if (i > 0) para.pPr.spcBef = s.spacing.bulletSpacing || 114300;
      const keyRun = this._makeRun(keys[i] + ': ', {
        fontSize: s.bodyText.fontSize,
        bold: true,
        color: s.palette.accent,
        fontFamily: s.bodyText.font
      });
      para.runs.push(keyRun);
      const valRun = this._makeRun(String(details[keys[i]]), {
        fontSize: s.bodyText.fontSize,
        color: s.palette.primaryText,
        fontFamily: s.bodyText.font
      });
      para.runs.push(valRun);
      bodyShape.textBody.paragraphs.push(para);
    }

    slide.shapes.push(bodyShape);
    if (opts.footnote) this._addFootnote(slide, opts.footnote);

    this.pres.slides.push(slide);
    return this;
  }

  // Section divider slide with large number + title
  addSectionDividerSlide(sectionNumber: number | string, sectionTitle: string): this {
    const slide = new DM.Slide(this.pres.slides.length + 1);
    const s = this._style;
    const d = s.divider;

    slide.background = new DM.Fill('solid', { color: this._style.slide.background });

    // Format section number (e.g., 1 -> "01.", "01" -> "01.")
    let numStr: string;
    if (typeof sectionNumber === 'number') {
      numStr = (sectionNumber < 10 ? '0' + sectionNumber : String(sectionNumber)) + '.';
    } else {
      numStr = String(sectionNumber);
      if (numStr.indexOf('.') === -1) numStr += '.';
    }

    // Large section number
    const numH = Math.round(d.numberFontSize / 100 * 914400 / 72 * 1.3);
    const numShape = this._createTextShape(
      'Section Number', numStr,
      { x: d.numberX, y: d.numberY, cx: Math.round(this._slideW() * 0.6), cy: numH },
      {
        fontSize: d.numberFontSize,
        bold: true,
        color: d.numberColor,
        fontFamily: d.numberFont,
        align: 'l', anchor: 't'
      }
    );
    slide.shapes.push(numShape);

    // Section title below the number
    const titleY = d.titleY > 0 ? d.titleY : (d.numberY + numH + 91440);
    const titleH = Math.round(d.titleFontSize / 100 * 914400 / 72 * 1.6);
    const titleShape = this._createTextShape(
      'Section Title', sectionTitle,
      { x: d.numberX, y: titleY, cx: Math.round(this._slideW() * 0.6), cy: titleH },
      {
        fontSize: d.titleFontSize,
        bold: true,
        color: d.titleColor,
        fontFamily: d.titleFont,
        align: 'l', anchor: 't'
      }
    );
    slide.shapes.push(titleShape);

    this.pres.slides.push(slide);
    return this;
  }

  // Blank slide
  addBlankSlide(): this {
    const slide = new DM.Slide(this.pres.slides.length + 1);
    slide.background = new DM.Fill('solid', { color: this._style.slide.background });
    this.pres.slides.push(slide);
    return this;
  }

  // Add footnote to the most recently added slide
  addFootnote(text: string): this {
    const slides = this.pres.slides;
    if (slides.length === 0) return this;
    this._addFootnote(slides[slides.length - 1], text);
    return this;
  }

  // Advanced: add raw shape to the last slide
  addShapeToLastSlide(shape: any): this {
    const slides = this.pres.slides;
    if (slides.length === 0) this.addBlankSlide();
    slides[slides.length - 1].shapes.push(shape);
    return this;
  }

  // ===== LAYOUT METHODS =====
  // Complex multi-zone slide layouts with header bars and polymorphic content.

  // ---- Internal: Header Bar ----
  // Creates a roundRect shape with dark fill and white text, matching reference deck style.
  _createHeaderBar(text: string, bounds: Bounds) {
    const s = this._style;
    const shape = new DM.Shape();
    shape.id = this._nextId();
    shape.name = 'Header Bar';
    shape.xfrm = {
      off: { x: bounds.x, y: bounds.y },
      ext: { cx: bounds.cx, cy: bounds.cy },
      rot: 0, flipH: false, flipV: false
    };
    const hdrStyle = s.header || {};
    shape.geometry = { type: hdrStyle.borderRadius === 0 ? 'rect' : 'roundRect' };
    shape.fill = new DM.Fill('solid', { color: hdrStyle.background || s.palette.darkFill || '#164556' });
    shape.textBody = new DM.TextBody();
    shape.textBody.bodyPr.anchor = 'ctr';

    const para = new DM.Paragraph();
    para.pPr.algn = hdrStyle.align || 'ctr';
    const run = this._makeRun(text, {
      fontSize: hdrStyle.fontSize || s.bodyText.fontSize || 1100,
      bold: true,
      color: hdrStyle.color || s.headerBar.textColor || '#FFFFFF',
      fontFamily: s.mainTitle.font || s.bodyText.font
    });
    para.runs.push(run);
    shape.textBody.paragraphs.push(para);
    return shape;
  }

  // ---- Internal: Subheader (bold text + separate line shape below) ----
  // Returns { textShape, lineShape } — caller pushes both to slide
  _createSubheader(text: string, textBounds: Bounds, lineBounds: Bounds) {
    const s = this._style;
    const shStyle = s.subheader || {};
    const color = shStyle.color || s.palette.primaryText || '#333333';

    // Text shape
    const textShape = new DM.Shape();
    textShape.id = this._nextId();
    textShape.name = 'Subheader';
    textShape.xfrm = {
      off: { x: textBounds.x, y: textBounds.y },
      ext: { cx: textBounds.cx, cy: textBounds.cy },
      rot: 0, flipH: false, flipV: false
    };
    textShape.geometry = { type: 'rect' };
    textShape.fill = new DM.Fill('none');
    textShape.textBody = new DM.TextBody();
    textShape.textBody.bodyPr.anchor = 'b';
    textShape.textBody.bodyPr.bIns = 0;

    const para = new DM.Paragraph();
    para.pPr.algn = shStyle.align || 'ctr';
    para.runs.push(this._makeRun(text, {
      fontSize: shStyle.fontSize || s.bodyText.fontSize || 1050,
      bold: shStyle.bold !== false,
      color: color,
      fontFamily: s.bodyText.font || 'Calibri'
    }));
    textShape.textBody.paragraphs.push(para);

    // Line shape (horizontal connector at bottom of subheader)
    const lineShape = new DM.Shape();
    lineShape.type = 'cxnSp';
    lineShape.id = this._nextId();
    lineShape.name = 'Subheader Line';
    lineShape.xfrm = {
      off: { x: lineBounds.x, y: lineBounds.y },
      ext: { cx: lineBounds.cx, cy: 0 },
      rot: 0, flipH: false, flipV: false
    };
    lineShape.line = new DM.LineProps({ width: shStyle.lineWidth || s.spacing.subheaderLineWidth || 9525, color: shStyle.lineColor || color });

    return { textShape: textShape, lineShape: lineShape };
  }

  // ---- Internal: Subfooter (small text at bottom of container) ----
  _createSubfooter(text: string, bounds: Bounds) {
    const s = this._style;
    const sfStyle = s.subfooter || {};
    const shape = new DM.Shape();
    shape.id = this._nextId();
    shape.name = 'Subfooter';
    shape.xfrm = {
      off: { x: bounds.x, y: bounds.y },
      ext: { cx: bounds.cx, cy: bounds.cy },
      rot: 0, flipH: false, flipV: false
    };
    shape.geometry = { type: 'rect' };
    shape.fill = new DM.Fill('none');
    shape.textBody = new DM.TextBody();
    shape.textBody.bodyPr.anchor = 't';
    shape.textBody.bodyPr.lIns = 45720;

    const para = new DM.Paragraph();
    para.pPr.algn = 'l';
    const run = this._makeRun(text, {
      fontSize: sfStyle.fontSize || Math.round((s.footnote.fontSize || 800) * 0.9),
      bold: false,
      italic: sfStyle.italic !== false,
      color: sfStyle.color || s.palette.secondaryText || '#666666',
      fontFamily: s.footnote.font || s.bodyText.font || 'Calibri'
    });
    para.runs.push(run);
    shape.textBody.paragraphs.push(para);
    return shape;
  }

  // ---- Internal: Get zone bounds from catalog or defaults ----
  _getZoneBounds(layoutType: string, zoneRole: string) {
    if (this._catalog) {
      const template = this._catalog.getTemplate(layoutType);
      if (template) {
        for (let i = 0; i < template.zones.length; i++) {
          if (template.zones[i].role === zoneRole) {
            return {
              zone: template.zones[i].bounds,
              headerBar: template.zones[i].headerBar ? template.zones[i].headerBar.bounds : null
            };
          }
        }
      }
    }
    return null;
  }

  // ---- Internal: Default quadrant positions ----
  _defaultQuadrantBounds() {
    const sw = this._slideW();
    const sh = this._slideH();
    const margin = this._marginLeft();
    const gap = this._style.spacing.zoneGap || Math.round(0.15 * 914400);
    const headerTop = this._contentTop();
    const colW = Math.round((sw - 2 * margin - gap) / 2);
    const barH = this._style.spacing.headerBarHeight || Math.round(0.35 * 914400);
    const midY = Math.round((headerTop + sh - margin) / 2);
    const contentGap = this._style.spacing.contentGap || Math.round(0.1 * 914400);

    return {
      topLeft: {
        bar: { x: margin, y: headerTop, cx: colW, cy: barH },
        content: { x: margin, y: headerTop + barH + contentGap, cx: colW, cy: midY - headerTop - barH - contentGap - gap }
      },
      topRight: {
        bar: { x: margin + colW + gap, y: headerTop, cx: colW, cy: barH },
        content: { x: margin + colW + gap, y: headerTop + barH + contentGap, cx: colW, cy: midY - headerTop - barH - contentGap - gap }
      },
      bottomLeft: {
        bar: { x: margin, y: midY, cx: colW, cy: barH },
        content: { x: margin, y: midY + barH + contentGap, cx: colW, cy: sh - margin - midY - barH - contentGap }
      },
      bottomRight: {
        bar: { x: margin + colW + gap, y: midY, cx: colW, cy: barH },
        content: { x: margin + colW + gap, y: midY + barH + contentGap, cx: colW, cy: sh - margin - midY - barH - contentGap }
      }
    };
  }

  // ---- Internal: Default three-column positions ----
  _defaultThreeColBounds() {
    const sw = this._slideW();
    const sh = this._slideH();
    const margin = this._marginLeft();
    const gap = this._style.spacing.zoneGap || Math.round(0.15 * 914400);
    const headerTop = this._contentTop();
    const colW = Math.round((sw - 2 * margin - 2 * gap) / 3);
    const barH = this._style.spacing.headerBarHeight || Math.round(0.35 * 914400);
    const contentGap = this._style.spacing.contentGap || Math.round(0.1 * 914400);
    const contentH = sh - margin - headerTop - barH - contentGap;

    return [
      {
        bar: { x: margin, y: headerTop, cx: colW, cy: barH },
        content: { x: margin, y: headerTop + barH + contentGap, cx: colW, cy: contentH }
      },
      {
        bar: { x: margin + colW + gap, y: headerTop, cx: colW, cy: barH },
        content: { x: margin + colW + gap, y: headerTop + barH + contentGap, cx: colW, cy: contentH }
      },
      {
        bar: { x: margin + 2 * (colW + gap), y: headerTop, cx: colW, cy: barH },
        content: { x: margin + 2 * (colW + gap), y: headerTop + barH + contentGap, cx: colW, cy: contentH }
      }
    ];
  }

  // ---- Internal: Polymorphic zone content renderer ----
  async _renderZoneContent(slide: any, bounds: Bounds, content: any): Promise<void> {
    if (!content || !content.type) return;

    switch (content.type) {
      case 'statGrid':
        this._renderStatGrid(slide, bounds, content.items || []);
        break;
      case 'cardGrid':
        this._renderCardGrid(slide, bounds, content.items || []);
        break;
      case 'table':
        this._renderTableInZone(slide, bounds, content.data || {});
        break;
      case 'text':
        if (content.runs && content.runs.length > 0) {
          this._renderRichTextBlock(slide, bounds, content);
        } else {
          this._renderTextBlock(slide, bounds, content.text || '', content);
        }
        break;
      case 'image':
        this._renderImage(slide, bounds, content);
        break;
      case 'chart':
        await this._renderChart(slide, bounds, content);
        break;
      case 'pptxChart':
        await this._renderPptxChart(slide, bounds, content);
        break;
      case 'timeline':
        this._renderTimeline(slide, bounds, content);
        break;
      case 'callout':
        this._renderCallout(slide, bounds, content);
        break;
      case 'profile':
        this._renderProfile(slide, bounds, content);
        break;
      case 'icon':
        // Icon content is resolved to image buffer by render.ts before reaching here
        if (content.buffer) {
          this._renderImage(slide, bounds, { type: 'image', buffer: content.buffer, width: content.width || '80%', height: content.height || '80%', objectFit: 'contain' });
        }
        break;
      case 'line': {
        const lineS = new DM.Shape();
        lineS.type = 'cxnSp';
        lineS.id = this._nextId();
        lineS.name = 'Divider Line';
        const lineY = bounds.y + Math.round(bounds.cy / 2);
        lineS.xfrm = { off: { x: bounds.x, y: lineY }, ext: { cx: bounds.cx, cy: 0 }, rot: 0, flipH: false, flipV: false };
        lineS.line = new DM.LineProps({ width: content.width || 9525, color: content.color || this._style.palette.primaryText || '#333333' });
        slide.shapes.push(lineS);
        break;
      }
      default:
        if (content.text) this._renderTextBlock(slide, bounds, content.text);
        break;
    }
  }

  // ---- Content Renderers ----
  _renderStatGrid(slide: any, bounds: Bounds, items: any[]) {
    if (items.length === 0) return;
    const s = this._style;
    const sg = s.statGrid || {};
    const cols = _responsiveCols(items.length, bounds.cx, bounds.cy);
    const rows = Math.ceil(items.length / cols);
    const cellW = Math.round(bounds.cx / cols);
    const cellH = Math.round(bounds.cy / rows);

    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      const col = i % cols;
      const row = Math.floor(i / cols);
      const cellX = bounds.x + col * cellW;
      const cellY = bounds.y + row * cellH;

      // Value (large text)
      const valueH = Math.round(cellH * 0.55);
      const valueShape = this._createTextShape(
        'Stat Value', String(item.value || ''),
        { x: cellX, y: cellY, cx: cellW, cy: valueH },
        {
          fontSize: sg.valueSize || 1600,
          bold: sg.valueBold !== false,
          italic: sg.valueItalic || false,
          color: sg.valueColor || s.palette.accent,
          fontFamily: sg.valueFont || s.mainTitle.font,
          align: 'ctr', anchor: 'b'
        }
      );
      slide.shapes.push(valueShape);

      // Label (small text below)
      const labelShape = this._createTextShape(
        'Stat Label', String(item.label || ''),
        { x: cellX, y: cellY + valueH, cx: cellW, cy: Math.round(cellH * 0.3) },
        {
          fontSize: sg.labelSize || Math.round(s.bodyText.fontSize * 0.85),
          bold: sg.labelBold || false,
          italic: sg.labelItalic || false,
          color: sg.labelColor || s.palette.primaryText,
          fontFamily: sg.labelFont || s.bodyText.font,
          align: 'ctr', anchor: 't'
        }
      );
      slide.shapes.push(labelShape);

      // Sublabel if present
      if (item.sublabel) {
        const subShape = this._createTextShape(
          'Stat Sublabel', String(item.sublabel),
          { x: cellX, y: cellY + valueH + Math.round(cellH * 0.3), cx: cellW, cy: Math.round(cellH * 0.15) },
          {
            fontSize: sg.sublabelSize || Math.round(s.footnote.fontSize),
            bold: sg.sublabelBold || false,
            italic: sg.sublabelItalic || false,
            color: sg.sublabelColor || s.palette.secondaryText,
            fontFamily: sg.sublabelFont || s.bodyText.font,
            align: 'ctr', anchor: 't'
          }
        );
        slide.shapes.push(subShape);
      }
    }
  }

  _renderCardGrid(slide: any, bounds: Bounds, items: any[]) {
    if (items.length === 0) return;
    const s = this._style;
    const cg = s.cardGrid || {};
    const cols = _responsiveCols(items.length, bounds.cx, bounds.cy);
    const rows = Math.ceil(items.length / cols);
    const cardStyle = s.card || {};
    const gap = s.spacing.cardGap || Math.round(0.08 * 914400);
    const cellW = Math.round((bounds.cx - (cols - 1) * gap) / cols);
    const cellH = Math.round((bounds.cy - (rows - 1) * gap) / rows);
    const bulletChar = '\u2022';
    const bulletIndent = cg.lineBulletIndent || Math.round(0.05 * 914400);

    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      const col = i % cols;
      const row = Math.floor(i / cols);
      const cellX = bounds.x + col * (cellW + gap);
      const cellY = bounds.y + row * (cellH + gap);

      // Card border
      const card = new DM.Shape();
      card.id = this._nextId();
      card.name = 'Card';
      card.xfrm = {
        off: { x: cellX, y: cellY },
        ext: { cx: cellW, cy: cellH },
        rot: 0, flipH: false, flipV: false
      };
      card.geometry = { type: 'roundRect' };
      card.fill = new DM.Fill('solid', { color: cardStyle.fill || '#FFFFFF' });
      card.line = new DM.LineProps({ width: cardStyle.borderWidth || 9525, color: s.palette.border || '#BFBFBF' });

      // Card content
      card.textBody = new DM.TextBody();
      card.textBody.bodyPr.anchor = 't';
      card.textBody.bodyPr.lIns = cardStyle.paddingLeft || 68580;
      card.textBody.bodyPr.tIns = cardStyle.paddingTop || 45720;

      // Title paragraph
      const titlePara = new DM.Paragraph();
      titlePara.pPr.algn = 'l';
      const titleRun = this._makeRun(String(item.title || ''), {
        fontSize: cg.titleSize || s.bodyText.fontSize,
        bold: cg.titleBold !== false,
        italic: cg.titleItalic || false,
        color: cg.titleColor || s.palette.accent,
        fontFamily: cg.titleFont || s.bodyText.font
      });
      titlePara.runs.push(titleRun);
      card.textBody.paragraphs.push(titlePara);

      // Detail lines
      const lines = item.lines || [];
      for (let li = 0; li < lines.length; li++) {
        const linePara = new DM.Paragraph();
        linePara.pPr.algn = 'l';
        linePara.pPr.spcBef = 45720;
        if (cg.lineBullet) {
          linePara.pPr.indent = -bulletIndent;
          linePara.pPr.marL = bulletIndent;
          linePara.pPr.buChar = bulletChar;
        }
        const lineRun = this._makeRun(String(lines[li]), {
          fontSize: cg.lineSize || Math.round(s.bodyText.fontSize * 0.85),
          bold: cg.lineBold || false,
          italic: cg.lineItalic || false,
          color: cg.lineColor || s.palette.primaryText,
          fontFamily: cg.lineFont || s.bodyText.font
        });
        linePara.runs.push(lineRun);
        card.textBody.paragraphs.push(linePara);
      }

      slide.shapes.push(card);
    }
  }

  _renderTableInZone(slide: any, bounds: Bounds, rawData: any) {
    if (!rawData.rows) return;
    // Normalize: colAlign array → align record, headers:[] → no headers
    const data = Object.assign({}, rawData);
    if (data.headers && data.headers.length === 0) data.headers = undefined;
    if (data.colAlign && !data.align) {
      const alignFromArr = {};
      for (let ai = 0; ai < data.colAlign.length; ai++) alignFromArr[ai] = data.colAlign[ai];
      data.align = alignFromArr;
    }
    const hasHeaders = data.headers && data.headers.length > 0;
    const hasVerticalHeaders = data.verticalHeaders === true;
    const s = this._style;

    const shape = new DM.Shape();
    shape.id = this._nextId();
    shape.name = 'Zone Table';
    shape.type = 'graphicFrame';
    shape.geometry = null;

    const td = new DM.TableData();
    const numCols = hasHeaders ? data.headers.length : (data.rows[0] ? data.rows[0].length : 1);

    // Column widths: use 12-column grid spans if provided, else equal split
    if (data.colWidths && data.colWidths.length === numCols) {
      let totalSpan = 0;
      for (let si = 0; si < data.colWidths.length; si++) totalSpan += data.colWidths[si];
      for (let ci = 0; ci < numCols; ci++) {
        td.cols.push({ w: Math.floor(data.colWidths[ci] / totalSpan * bounds.cx) });
      }
    } else {
      const colW = Math.floor(bounds.cx / numCols);
      for (let ci2 = 0; ci2 < numCols; ci2++) {
        td.cols.push({ w: colW });
      }
    }
    td.tblPr = { bandRow: true, firstRow: hasHeaders };

    // Row heights: fixed defaults (all body rows same height, header has its own)
    // Default: 0.4" body rows, 0.35" header row
    const DEFAULT_BODY_ROW_H = 365760;   // 0.4 inches in EMU
    const DEFAULT_HEADER_ROW_H = 320040; // 0.35 inches in EMU

    const bodyRowH = data.rowHeight != null
      ? Math.round(data.rowHeight * 914400)
      : DEFAULT_BODY_ROW_H;
    const headerH = data.headerHeight != null
      ? Math.round(data.headerHeight * 914400)
      : DEFAULT_HEADER_ROW_H;

    // Font size overrides
    const hdrFontSize = data.headerFontSize || s.tableHeader.fontSize;
    const bodyFontSize = data.fontSize || s.tableBody.fontSize;

    // Per-column alignment map
    const alignMap = data.align || {};

    // Cell border style. "horizontal" (default) draws rules between rows only; "grid" every edge; "none" nothing.
    const borderMode = (s.tableBorder && s.tableBorder.mode) || 'horizontal';
    const hairline = s.tableBorder && borderMode !== 'none' && s.tableBorder.width > 0 ? new DM.LineProps({ width: s.tableBorder.width, color: s.tableBorder.color }) : null;
    const cellBorder = hairline;
    const vertical = borderMode === 'grid' ? hairline : null;
    const headerRule = s.tableHeader.borderBottom && s.tableHeader.borderBottom.width > 0
      ? new DM.LineProps({ width: s.tableHeader.borderBottom.width, color: s.tableHeader.borderBottom.color })
      : hairline;
    const srBorders = s.tableSummaryRow;
    const lineOrNull = (b: any) => (b && b.width > 0 ? new DM.LineProps({ width: b.width, color: b.color }) : (borderMode === 'grid' ? hairline : null));

    // Track image cells for overlay rendering
    const imageCells: any[] = [];

    // Header row (only if headers are provided)
    if (hasHeaders) {
      const headerRow = new DM.TableRow(headerH);
      for (let hi = 0; hi < data.headers.length; hi++) {
        const hCell = new DM.TableCell();
        hCell.txBody = new DM.TextBody();
        const hPara = new DM.Paragraph();
        hPara.pPr.algn = alignMap[hi] || 'l';
        hPara.runs.push(this._makeRun(String(data.headers[hi]), {
          fontSize: hdrFontSize, bold: true,
          color: s.tableHeader.textColor, fontFamily: s.tableHeader.font || s.bodyText.font
        }));
        hCell.txBody.paragraphs.push(hPara);
        hCell.tcPr.fill = new DM.Fill('solid', { color: s.tableHeader.fillColor });
        hCell.tcPr.borders = { l: vertical, r: vertical, t: borderMode === 'grid' ? hairline : null, b: headerRule };
        headerRow.cells.push(hCell);
      }
      td.rows.push(headerRow);
    }

    // Summary row detection
    const summaryRowCount = data.summaryRows || 0;
    const summaryStartIdx = data.rows.length - summaryRowCount;

    // Data rows
    for (let ri = 0; ri < data.rows.length; ri++) {
      const row = new DM.TableRow(bodyRowH);
      const isSummaryRow = summaryRowCount > 0 && ri >= summaryStartIdx;
      const isFirstSummaryRow = isSummaryRow && ri === summaryStartIdx;
      const isLastSummaryRow = isSummaryRow && ri === data.rows.length - 1;
      for (let rci = 0; rci < data.rows[ri].length; rci++) {
        const cellData = data.rows[ri][rci];
        const cell = new DM.TableCell();
        cell.txBody = new DM.TextBody();
        const cPara = new DM.Paragraph();
        cPara.pPr.algn = alignMap[rci] || 'l';

        // Determine if this cell is a vertical header (first column when verticalHeaders is on)
        const isVerticalHeaderCell = hasVerticalHeaders && rci === 0;
        const cellFontSize = isVerticalHeaderCell ? (s.tableVerticalHeader.fontSize || hdrFontSize) : isSummaryRow ? s.tableSummaryRow.fontSize : bodyFontSize;
        const cellTextColor = isVerticalHeaderCell ? s.tableVerticalHeader.textColor : isSummaryRow ? s.tableSummaryRow.textColor : s.tableBody.textColor;
        const cellFontFamily = isVerticalHeaderCell ? (s.tableVerticalHeader.font || s.bodyText.font) : isSummaryRow ? (s.tableSummaryRow.font || s.bodyText.font) : (s.tableBody.font || s.bodyText.font);
        const cellBold = isVerticalHeaderCell ? s.tableVerticalHeader.bold : isSummaryRow ? s.tableSummaryRow.bold : false;

        // Check if cell is an image object
        const isImageCell = cellData && typeof cellData === 'object' && cellData.type === 'image' && cellData.buffer;
        if (isImageCell) {
          // Show the search term / prompt as text in the cell; image goes off-slide
          let cellLabel = cellData.searchTerm || cellData.prompt || cellData.name || '';
          // Strip trailing " logo" suffix for cleaner display
          if (cellLabel.toLowerCase().endsWith(' logo')) cellLabel = cellLabel.slice(0, -5);
          cPara.runs.push(this._makeRun(cellLabel, {
            fontSize: cellFontSize, bold: cellBold, color: cellTextColor,
            fontFamily: cellFontFamily
          }));
          imageCells.push({ ri: ri, ci: rci, content: cellData });
        } else {
          cPara.runs.push(this._makeRun(String(cellData || ''), {
            fontSize: cellFontSize, bold: cellBold, color: cellTextColor,
            fontFamily: cellFontFamily
          }));
        }

        cell.txBody.paragraphs.push(cPara);
        const cellFillColor = isVerticalHeaderCell ? s.tableVerticalHeader.fillColor : isSummaryRow ? s.tableSummaryRow.fillColor : ((s.tableBody.alternateRows !== false && ri % 2 === 0) ? s.tableBody.fillEven : s.tableBody.fillOdd);
        cell.tcPr.fill = new DM.Fill('solid', { color: cellFillColor });
        const topRule = ri === 0 && hasHeaders ? headerRule : cellBorder;
        if (isSummaryRow) {
          // Summary row group: accent rule on top of the first and under the last; verticals only in grid mode
          const srBorderT = isFirstSummaryRow ? lineOrNull(srBorders.borderTop) : cellBorder;
          const srBorderB = isLastSummaryRow ? lineOrNull(srBorders.borderBottom) : cellBorder;
          cell.tcPr.borders = { l: lineOrNull(srBorders.borderLeft), r: lineOrNull(srBorders.borderRight), t: srBorderT, b: srBorderB };
        } else if (summaryRowCount > 0 && ri === summaryStartIdx - 1) {
          // Row just above summary: bottom border = summary top rule so it takes precedence
          cell.tcPr.borders = { l: vertical, r: vertical, t: topRule, b: lineOrNull(srBorders.borderTop) };
        } else {
          cell.tcPr.borders = { l: vertical, r: vertical, t: topRule, b: cellBorder };
        }
        row.cells.push(cell);
      }
      td.rows.push(row);
    }

    const totalH = headerH + bodyRowH * data.rows.length;
    shape.xfrm = {
      off: { x: bounds.x, y: bounds.y },
      ext: { cx: bounds.cx, cy: Math.min(totalH, bounds.cy) },
      rot: 0, flipH: false, flipV: false
    };
    shape.tableData = td;
    slide.shapes.push(shape);

    // Place image shapes off-slide to the left (pasteboard area).
    // Users can drag them into position in PowerPoint.
    const IMG_SIZE = 457200; // 0.5 inches — compact logo size
    const IMG_GAP = 73152;   // ~0.08 inches gap between stacked logos
    const IMG_X = -IMG_SIZE - 91440; // 0.5" wide + 0.1" margin from slide edge
    for (let ii = 0; ii < imageCells.length; ii++) {
      const ic = imageCells[ii];
      const offBounds = {
        x: IMG_X,
        y: bounds.y + (ii * (IMG_SIZE + IMG_GAP)),
        cx: IMG_SIZE,
        cy: IMG_SIZE
      };
      // Use contain mode so logo fits entirely within the square
      const cellContent = Object.assign({}, ic.content, { objectFit: 'contain' });
      this._renderImage(slide, offBounds, cellContent);
    }
  }

  _renderTextBlock(slide: any, bounds: Bounds, text: any, opts?: any) {
    if (!text) return;
    const s = this._style;
    const o = opts || {};
    const lines = Array.isArray(text) ? text : [text];

    const shape = this._createTextShape(
      'Text Block', lines,
      { x: bounds.x, y: bounds.y, cx: bounds.cx, cy: bounds.cy },
      {
        fontSize: o.fontSize || s.bodyText.fontSize,
        bold: o.bold || false,
        color: o.color || s.palette.primaryText,
        fontFamily: o.font || s.bodyText.font,
        align: o.align || 'l',
        anchor: o.anchor || 't',
        spacing: 0
      }
    );
    // Apply line spacing if specified
    if (o.lineSpacing != null) {
      const lnSpc = Math.round(o.lineSpacing * 100);
      const paras = shape.textBody!.paragraphs;
      for (let li = 0; li < paras.length; li++) {
        paras[li].pPr.lnSpc = lnSpc;
      }
    }
    slide.shapes.push(shape);
  }

  // ---- Rich text block (multiple runs per paragraph) ----
  _renderRichTextBlock(slide: any, bounds: Bounds, content: any) {
    const s = this._style;
    const runs = content.runs;

    const blockFontSize = content.fontSize || s.bodyText.fontSize;
    const blockBold = content.bold || false;
    const blockItalic = content.italic || false;
    const blockUnderline = content.underline || false;
    const blockCap = content.textTransform === 'uppercase' ? 'all' : null;
    const blockColor = content.color || s.palette.primaryText;
    const blockFont = content.font || s.bodyText.font;

    // Split runs into paragraphs on \n, preserving bullet flag and level
    const paragraphs: any[][] = [[]];
    const paraBullets = [false]; // track bullet state per paragraph
    const paraBulletLevels = [0]; // track bullet nesting level per paragraph
    for (let ri = 0; ri < runs.length; ri++) {
      const run = runs[ri];
      // A bulleted run starts its own paragraph when the current one already has text,
      // so agents can write one run per bullet without inserting "\n".
      if (run.bullet && paragraphs[paragraphs.length - 1].length > 0) {
        paragraphs.push([]);
        paraBullets.push(false);
        paraBulletLevels.push(0);
      }
      const parts = run.text.split('\n');
      for (let pi = 0; pi < parts.length; pi++) {
        if (pi > 0) {
          paragraphs.push([]);
          paraBullets.push(false);
          paraBulletLevels.push(0);
        }
        if (parts[pi].length > 0 || parts.length === 1) {
          const runObj = {
            text: parts[pi],
            bold: run.bold, italic: run.italic, underline: run.underline,
            color: run.color, fontSize: run.fontSize, font: run.font
          };
          paragraphs[paragraphs.length - 1].push(runObj);
          // First run in a paragraph sets bullet state
          if (run.bullet && paragraphs[paragraphs.length - 1].length === 1) {
            paraBullets[paraBullets.length - 1] = true;
            paraBulletLevels[paraBulletLevels.length - 1] = run.bulletLevel || 0;
          }
        }
      }
    }

    const shape = new DM.Shape();
    shape.id = this._nextId();
    shape.name = 'Text Block';
    shape.xfrm = {
      off: { x: bounds.x, y: bounds.y },
      ext: { cx: bounds.cx, cy: bounds.cy },
      rot: 0, flipH: false, flipV: false
    };
    shape.geometry = { type: 'rect' };
    shape.fill = new DM.Fill('none');
    shape.textBody = new DM.TextBody();
    shape.textBody.bodyPr.anchor = content.anchor || 't';
    if (content.paddingTop != null) shape.textBody.bodyPr.tIns = Math.round(content.paddingTop * 914400);
    if (content.paddingRight != null) shape.textBody.bodyPr.rIns = Math.round(content.paddingRight * 914400);
    if (content.paddingBottom != null) shape.textBody.bodyPr.bIns = Math.round(content.paddingBottom * 914400);
    if (content.paddingLeft != null) shape.textBody.bodyPr.lIns = Math.round(content.paddingLeft * 914400);

    for (let i = 0; i < paragraphs.length; i++) {
      const para = new DM.Paragraph();
      para.pPr.algn = content.align || 'l';
      if (content.lineSpacing != null) para.pPr.lnSpc = Math.round(content.lineSpacing * 100);

      // Apply bullet formatting for bulleted paragraphs
      if (paraBullets[i]) {
        const bLevel = paraBulletLevels[i] || 0;
        const bs = content.bulletStyle || {};
        para.pPr.buNone = false;

        // Bullet character: override → style default (en-dash for sub-bullets)
        if (bLevel >= 1) {
          para.pPr.buChar = bs.char || '\u2013';
          para.pPr.lvl = bLevel;
        } else {
          para.pPr.buChar = bs.char || s.bullet.char;
        }

        // Margin left: override (inches→EMU) or style default (already EMU)
        if (bs.marginLeft != null) {
          para.pPr.marL = Math.round(bs.marginLeft * 914400) * (bLevel >= 1 ? bLevel + 1 : 1);
        } else {
          para.pPr.marL = Math.abs(s.bullet.marginLeft) * (bLevel >= 1 ? bLevel + 1 : 1);
        }

        // Indent: override (inches→negative EMU) or style default (already negative EMU)
        if (bs.indent != null) {
          para.pPr.indent = -Math.round(bs.indent * 914400);
        } else {
          para.pPr.indent = s.bullet.indent;
        }

        // Bullet font & color: override → style default
        para.pPr.buFont = s.bullet.bulletFont || null;
        para.pPr.buClr = bs.color || s.bullet.bulletColor || null;

        // Space before bullets (override in points → EMU)
        if (bs.spaceBefore != null && i > 0) {
          para.pPr.spcBef = Math.round(bs.spaceBefore * 12700);
        }
      }

      const paraRuns = paragraphs[i];
      if (paraRuns.length === 0) {
        // Empty paragraph (from \n\n). Use blockFontSize (zone-level font)
        // to match the web render where empty lines inherit the container font.
        para.runs.push(this._makeRun('', {
          fontSize: blockFontSize, bold: blockBold, italic: blockItalic,
          underline: blockUnderline ? 'sng' : 'none',
          cap: blockCap,
          color: blockColor, fontFamily: blockFont
        }));
      } else {
        for (let j = 0; j < paraRuns.length; j++) {
          const r = paraRuns[j];
          para.runs.push(this._makeRun(r.text, {
            fontSize: r.fontSize != null ? r.fontSize : blockFontSize,
            bold: r.bold != null ? r.bold : blockBold,
            italic: r.italic != null ? r.italic : blockItalic,
            underline: (r.underline != null ? r.underline : blockUnderline) ? 'sng' : 'none',
            cap: blockCap,
            color: r.color || blockColor,
            fontFamily: r.font || blockFont
          }));
        }
      }

      shape.textBody.paragraphs.push(para);
    }

    slide.shapes.push(shape);
  }

  // ---- Image content renderer ----
  // content = { type: 'image', buffer: Buffer, width?: px, height?: px, ext?: 'png' }
  _renderImage(slide: any, bounds: Bounds, content: any) {
    if (!content || !content.buffer) return;

    const ext = content.ext || 'png';
    const mediaPath = 'ppt/media/img_' + this._nextId() + '.' + ext;

    // Store the image buffer in presentation media
    this.pres.media[mediaPath] = content.buffer;

    // Apply imagePadding (inches → EMU) to inset bounds before fit calculation
    if (content.imagePadding) {
      const pad = Math.round(content.imagePadding * 914400);
      bounds = { x: bounds.x + pad, y: bounds.y + pad, cx: Math.max(0, bounds.cx - 2 * pad), cy: Math.max(0, bounds.cy - 2 * pad) };
    }

    // Resolve percentage dimensions against container bounds
    let wRaw = content.width;
    let hRaw = content.height;
    let wIsPercent = typeof wRaw === 'string' && wRaw.charAt(wRaw.length - 1) === '%';
    let hIsPercent = typeof hRaw === 'string' && hRaw.charAt(hRaw.length - 1) === '%';

    // Compute display bounds
    let dispX = bounds.x;
    let dispY = bounds.y;
    let dispCx = bounds.cx;
    let dispCy = bounds.cy;
    let cropRect: any = null; // { l, t, r, b } in 1/1000th of a percent (OOXML srcRect)

    // Explicit crop from user — convert percentages (0-100) to OOXML (0-100000)
    if (content.crop && (content.crop.l > 0 || content.crop.t > 0 || content.crop.r > 0 || content.crop.b > 0)) {
      cropRect = {
        l: Math.round((content.crop.l || 0) * 1000),
        t: Math.round((content.crop.t || 0) * 1000),
        r: Math.round((content.crop.r || 0) * 1000),
        b: Math.round((content.crop.b || 0) * 1000),
      };

      // Read natural dimensions to preserve aspect ratio of the cropped region
      const cropDims = readImageDimensions(content.buffer);
      if (cropDims && cropDims.width > 0 && cropDims.height > 0) {
        const cropVisW = (100 - (content.crop.l || 0) - (content.crop.r || 0)) / 100;
        const cropVisH = (100 - (content.crop.t || 0) - (content.crop.b || 0)) / 100;
        const croppedW = cropDims.width * cropVisW;
        const croppedH = cropDims.height * cropVisH;
        const croppedAR = croppedW / croppedH;
        let cellAR = bounds.cx / bounds.cy;
        const fit = content.objectFit || 'cover';
        let anchor = content.anchor || 'ctr';

        if (fit === 'contain') {
          if (croppedAR > cellAR) {
            // Cropped image is wider — fit to width, letterbox vertically
            dispCx = bounds.cx;
            dispCy = Math.round(bounds.cx / croppedAR);
            if (anchor === 't') dispY = bounds.y;
            else if (anchor === 'b') dispY = bounds.y + (bounds.cy - dispCy);
            else dispY = bounds.y + Math.round((bounds.cy - dispCy) / 2);
          } else if (croppedAR < cellAR) {
            // Cropped image is taller — fit to height, pillarbox horizontally
            dispCy = bounds.cy;
            dispCx = Math.round(bounds.cy * croppedAR);
            dispX = bounds.x + Math.round((bounds.cx - dispCx) / 2);
          }
        }
        // 'cover' and 'fill': use full container bounds (default behavior)
      }

      let shape = new DM.Shape();
      shape.type = 'pic';
      shape.id = this._nextId();
      shape.name = content.name || 'Image';
      shape.xfrm = {
        off: { x: dispX, y: dispY },
        ext: { cx: dispCx, cy: dispCy },
        rot: 0, flipH: false, flipV: false
      };
      shape._imagePath = mediaPath;
      shape._cropRect = cropRect;
      shape.lockAspectRatio = true;
      slide.shapes.push(shape);
      return;
    }

    // Check if both dimensions are 100% — treat as cover mode (same as no dimensions)
    const isFillBoth = wIsPercent && hIsPercent && parseFloat(wRaw) >= 100 && parseFloat(hRaw) >= 100;

    if (isFillBoth) {
      // 100%/100% = cover mode: fall through to the no-dimensions branch below
      wRaw = undefined;
      hRaw = undefined;
      wIsPercent = false;
      hIsPercent = false;
    }

    if (wIsPercent || hIsPercent) {
      // Percentage mode: compute sub-box within container
      if (wIsPercent) dispCx = Math.round(bounds.cx * parseFloat(wRaw) / 100);
      if (hIsPercent) dispCy = Math.round(bounds.cy * parseFloat(hRaw) / 100);
      // Center the sub-box within the container
      if (dispCx < bounds.cx) dispX = bounds.x + Math.round((bounds.cx - dispCx) / 2);
      if (dispCy < bounds.cy) dispY = bounds.y + Math.round((bounds.cy - dispCy) / 2);

      // Apply objectFit within the sub-box to preserve aspect ratio
      const pctFit = content.objectFit || 'fill';
      if (pctFit === 'contain' || pctFit === 'cover') {
        const pctDims = readImageDimensions(content.buffer);
        if (pctDims && pctDims.width > 0 && pctDims.height > 0) {
          const pctImgAR = pctDims.width / pctDims.height;
          const pctBoxAR = dispCx / dispCy;
          const pctAnchor = content.anchor || 'ctr';
          let subX = dispX, subY = dispY, subCx = dispCx, subCy = dispCy;

          if (pctFit === 'contain') {
            if (pctImgAR > pctBoxAR) {
              subCy = Math.round(dispCx / pctImgAR);
              if (pctAnchor === 't') subY = dispY;
              else if (pctAnchor === 'b') subY = dispY + (dispCy - subCy);
              else subY = dispY + Math.round((dispCy - subCy) / 2);
            } else if (pctImgAR < pctBoxAR) {
              subCx = Math.round(dispCy * pctImgAR);
              subX = dispX + Math.round((dispCx - subCx) / 2);
            }
          }
          // cover: fill sub-box, let PPTX crop overflow (default shape behavior)

          dispX = subX; dispY = subY; dispCx = subCx; dispCy = subCy;
        }
      }
    } else if (wRaw && hRaw) {
      // Pixel mode: fit within cell, preserve aspect ratio
      const imgAspect = wRaw / hRaw;
      const cellAspect = bounds.cx / bounds.cy;

      const vAlign = content.anchor || 'ctr';
      if (imgAspect > cellAspect) {
        // Image is wider than cell — fit to width
        dispCx = bounds.cx;
        dispCy = Math.round(bounds.cx / imgAspect);
        if (vAlign === 't') {
          dispY = bounds.y;
        } else if (vAlign === 'b') {
          dispY = bounds.y + (bounds.cy - dispCy);
        } else {
          dispY = bounds.y + Math.round((bounds.cy - dispCy) / 2);
        }
      } else {
        // Image is taller than cell — fit to height, center horizontally
        dispCy = bounds.cy;
        dispCx = Math.round(bounds.cy * imgAspect);
        dispX = bounds.x + Math.round((bounds.cx - dispCx) / 2);
      }
    } else if (content.objectFit === 'contain') {
      // Contain mode: fit entire image within container, no crop
      let dims = readImageDimensions(content.buffer);
      if (dims && dims.width > 0 && dims.height > 0) {
        let imgAR = dims.width / dims.height;
        let cellAR = bounds.cx / bounds.cy;
        let anchor = content.anchor || 'ctr';

        if (imgAR > cellAR) {
          // Image is wider than container — fit to width, letterbox top/bottom
          dispCx = bounds.cx;
          dispCy = Math.round(bounds.cx / imgAR);
          if (anchor === 't') {
            dispY = bounds.y;
          } else if (anchor === 'b') {
            dispY = bounds.y + (bounds.cy - dispCy);
          } else {
            dispY = bounds.y + Math.round((bounds.cy - dispCy) / 2);
          }
        } else if (imgAR < cellAR) {
          // Image is taller — fit to height, pillarbox left/right
          dispCy = bounds.cy;
          dispCx = Math.round(bounds.cy * imgAR);
          dispX = bounds.x + Math.round((bounds.cx - dispCx) / 2);
        }
        // If aspect ratios match, dispX/Y/Cx/Cy stay at full container bounds
      }
    } else {
      // No explicit dimensions: "cover" mode (default) — fill container, crop overflow
      let dims = readImageDimensions(content.buffer);
      if (dims && dims.width > 0 && dims.height > 0) {
        let imgAR = dims.width / dims.height;
        let cellAR = bounds.cx / bounds.cy;
        let anchor = content.anchor || 'ctr';

        if (imgAR > cellAR) {
          // Image is wider — crop left/right
          const visibleFraction = cellAR / imgAR;
          const cropTotal = Math.round((1 - visibleFraction) * 100000);
          if (anchor === 'ctr') {
            cropRect = { l: Math.round(cropTotal / 2), r: Math.round(cropTotal / 2), t: 0, b: 0 };
          } else {
            cropRect = { l: 0, r: cropTotal, t: 0, b: 0 };
          }
        } else if (imgAR < cellAR) {
          // Image is taller — crop top/bottom
          const visibleFractionV = imgAR / cellAR;
          const cropTotalV = Math.round((1 - visibleFractionV) * 100000);
          if (anchor === 't') {
            cropRect = { l: 0, r: 0, t: 0, b: cropTotalV };
          } else if (anchor === 'b') {
            cropRect = { l: 0, r: 0, t: cropTotalV, b: 0 };
          } else {
            cropRect = { l: 0, r: 0, t: Math.round(cropTotalV / 2), b: Math.round(cropTotalV / 2) };
          }
        }
        // dispX/Y/Cx/Cy stay at full container bounds — cropRect handles the clipping
      }
    }

    // Create picture shape
    let shape = new DM.Shape();
    shape.type = 'pic';
    shape.id = this._nextId();
    shape.name = content.name || 'Image';
    shape.xfrm = {
      off: { x: dispX, y: dispY },
      ext: { cx: dispCx, cy: dispCy },
      rot: 0, flipH: false, flipV: false
    };
    shape._imagePath = mediaPath;
    shape._cropRect = cropRect;
    shape.lockAspectRatio = true;

    slide.shapes.push(shape);
  }

  // ---- Chart content renderer ----
  // content = { type: 'chart', chartType: 'bar'|'line'|'pie'|'combo'|..., ...spec }
  // Generates a chart image at the exact container dimensions, then renders it.
  async _renderChart(slide: any, bounds: Bounds, content: any): Promise<void> {
    // Native Excel export mode — convert chart to native PPTX chart
    if (content.pptxExportMode === 'nativeExcel') {
      const nativeSpec = convertToNative(content, this._style.chart || {});
      if (nativeSpec) {
        await this._renderPptxChart(slide, bounds, nativeSpec);
        return;
      }
    }

    const chartStyle = this._style.chart || {};

    // Convert container EMU to pixels (96 DPI logical size)
    const widthPx = Math.round(bounds.cx / 914400 * 96);
    const heightPx = Math.round(bounds.cy / 914400 * 96);

    const gen = new ChartGenerator({
      width: widthPx, height: heightPx, dpiScale: 2,
      colors: chartStyle.colors || undefined,
      fontFamily: chartStyle.fontFamily || this._style.bodyText.font || undefined,
      fontSize: chartStyle.fontSize || undefined,
      barBorderColor: chartStyle.barBorderColor || undefined,
      barBorderWidth: chartStyle.barBorderWidth != null ? chartStyle.barBorderWidth : undefined,
      bar: chartStyle.bar || undefined,
      pie: chartStyle.pie || undefined,
      combo: chartStyle.combo || undefined,
    });
    const chartType = content.chartType || 'bar';
    if (typeof gen[chartType] !== 'function') {
      console.warn('[SlideBuilder] Unknown chart type: ' + chartType);
      return;
    }

    const buffer = await gen[chartType](content);
    this._renderImage(slide, bounds, {
      type: 'image',
      buffer: buffer,
      width: '100%',
      height: '100%'
    });
  }

  // ---- Native PPTX Chart renderer ----
  // Renders a PptxChartContent as a native editable chart (with embedded Excel workbook).
  async _renderPptxChart(slide: any, bounds: Bounds, content: any): Promise<void> {
    // Assign a unique chart index for this presentation
    if (!this.pres._nextChartIdx) this.pres._nextChartIdx = 1;
    const chartIdx = this.pres._nextChartIdx++;

    // Generate chart parts (chart XML, embedded Excel, rels)
    const result = await generateChartParts(content, bounds, chartIdx);

    // Store chart parts and content types on the presentation for the writer
    if (!this.pres._chartParts) this.pres._chartParts = {};
    for (const partPath in result.parts) {
      this.pres._chartParts[partPath] = result.parts[partPath];
    }
    if (!this.pres._chartContentTypes) this.pres._chartContentTypes = [];
    for (let cti = 0; cti < result.contentTypeOverrides.length; cti++) {
      this.pres._chartContentTypes.push(result.contentTypeOverrides[cti]);
    }

    // Create a nativeChart shape
    const shape = new DM.Shape();
    shape.type = 'nativeChart';
    shape.id = this._nextId();
    shape.name = 'Chart ' + chartIdx;
    shape.xfrm = { off: { x: bounds.x, y: bounds.y }, ext: { cx: bounds.cx, cy: bounds.cy }, rot: 0, flipH: false, flipV: false };
    shape._chartIdx = chartIdx;
    shape._graphicFrameXml = result.graphicFrameXml ?? undefined;
    slide.shapes.push(shape);
  }

  // ---- Callout renderer: tinted panel + accent bar + optional title + text ----
  _renderCallout(slide: any, bounds: Bounds, content: any) {
    const s = this._style;
    const co = s.callout || { fill: '#EEF3FA', textColor: s.palette.primaryText, accentColor: s.palette.accent, titleColor: s.palette.accent, font: s.bodyText.font, fontSize: s.bodyText.fontSize, titleSize: s.bodyText.fontSize, barWidth: 54864, padding: 109728 };

    const panel = new DM.Shape();
    panel.id = this._nextId();
    panel.name = 'Callout';
    panel.xfrm = { off: { x: bounds.x, y: bounds.y }, ext: { cx: bounds.cx, cy: bounds.cy }, rot: 0, flipH: false, flipV: false };
    panel.geometry = { type: 'rect' };
    panel.fill = new DM.Fill('solid', { color: content.background || co.fill });
    slide.shapes.push(panel);

    const bar = new DM.Shape();
    bar.id = this._nextId();
    bar.name = 'Callout Bar';
    bar.xfrm = { off: { x: bounds.x, y: bounds.y }, ext: { cx: co.barWidth, cy: bounds.cy }, rot: 0, flipH: false, flipV: false };
    bar.geometry = { type: 'rect' };
    bar.fill = new DM.Fill('solid', { color: content.accentColor || co.accentColor });
    slide.shapes.push(bar);

    const runs: any[] = [];
    if (content.title) runs.push({ text: content.title + '\n', bold: true, color: content.accentColor || co.titleColor, fontSize: co.titleSize });
    if (Array.isArray(content.runs) && content.runs.length) runs.push(...content.runs);
    else runs.push({ text: String(content.text || '') });
    const inset = { x: bounds.x + co.barWidth + co.padding, y: bounds.y, cx: Math.max(0, bounds.cx - co.barWidth - 2 * co.padding), cy: bounds.cy };
    this._renderRichTextBlock(slide, inset, {
      type: 'text', runs,
      font: content.font || co.font,
      fontSize: content.fontSize || co.fontSize,
      color: content.color || co.textColor,
      align: content.align || 'l',
      anchor: content.anchor || 'ctr',
    });
  }

  // ---- Timeline content renderer ----
  // content = { type: 'timeline', events: [{date, name, detail}], ... }
  // Renders a horizontal timeline with text boxes connected by elbow connectors.
  _renderTimeline(slide: any, bounds: Bounds, content: any) {
    const events = content.events || [];
    if (events.length === 0) return;

    const s = this._style;
    const bx = bounds.x, by = bounds.y, bw = bounds.cx, bh = bounds.cy;
    const margin = Math.round(bw * 0.04);

    // Config with defaults
    const lineColor = content.lineColor || s.palette.accent || '#4472C4';
    const lineWidth = content.lineWidth || 25400;
    const nodeColor = content.nodeColor || '#FFFFFF';
    const nodeBorderColor = content.nodeBorderColor || lineColor;
    const nodeBorderWidth = content.nodeBorderWidth || 19050;
    const connectorColor = content.connectorColor || _lightenHex(lineColor, 0.3);
    const connectorWidth = content.connectorWidth || 9525;
    const dateFormat = content.dateFormat || 'month-year';
    const textMode = content.textMode || 'plain';
    const placement = content.placement || 'split';
    const count = events.length;
    const topCount = placement === 'bottom' ? 0 : placement === 'top' ? count : (content.topCount != null ? content.topCount : Math.ceil(count / 2));
    const botCount = count - topCount;

    // Timeline line — matches web: h * 0.48
    const timelineY = by + Math.round(bh * 0.48);
    const tlLeft = bx + margin;
    const tlRight = bx + bw - margin;
    const tlWidth = tlRight - tlLeft;

    // Date → X mapping
    const timestamps = events.map(function(ev) { return new Date(ev.date).getTime(); });
    const validTs = timestamps.filter(function(t) { return !isNaN(t); });
    const minTs = validTs.length ? Math.min.apply(null, validTs) : 0;
    const maxTs = validTs.length ? Math.max.apply(null, validTs) : 1;
    const dateRange = maxTs - minTs || 1;
    const marginTs = dateRange * 0.05;
    const rangeStart = minTs - marginTs;
    const totalRange = dateRange + 2 * marginTs;
    function dateToX(ts) { return tlLeft + Math.round(((ts - rangeStart) / totalRange) * tlWidth); }

    // Box spacing solver — matches web exactly
    const BOX_H = content.boxHeight || Math.round(bh * ((content.boxHeightPct || 28) / 100));
    // boxGap in inches → EMU
    const fixedGapEmu = content.boxGap != null ? Math.round(content.boxGap * 914400) : undefined;
    // When boxGap is set, compute boxW from the most constrained row so both rows match
    let fixedBoxW: any = undefined;
    if (fixedGapEmu != null) {
      const maxN = Math.max(topCount, botCount, 1);
      fixedBoxW = Math.round((tlWidth - (maxN + 1) * fixedGapEmu) / maxN);
      if (fixedBoxW < 91440) fixedBoxW = 91440;
    }
    function solve(n) {
      if (n === 0) return { boxW: 0, gap: 0, positions: [] };
      let boxW = fixedBoxW || content.boxWidth || Math.min(Math.round(tlWidth * 0.18), Math.round(tlWidth / n * 0.78));
      let gap = fixedGapEmu != null ? fixedGapEmu : Math.round((tlWidth - n * boxW) / (n + 1));
      if (gap < 19050) { gap = 19050; boxW = Math.round((tlWidth - (n + 1) * gap) / n); }
      // Center the row: compute total used width then offset from left
      const usedWidth = n * boxW + (n - 1) * gap;
      const startX = tlLeft + Math.round((tlWidth - usedWidth) / 2);
      const positions: number[] = [];
      for (let i = 0; i < n; i++) positions.push(startX + i * (boxW + gap));
      return { boxW: boxW, gap: gap, positions: positions };
    }

    const topSol = solve(topCount);
    const botSol = solve(botCount);

    // Vertical positions
    const boxOffsetEmu = content.boxOffset != null ? Math.round(content.boxOffset * 914400) : Math.round(bh * 0.09);
    const topBoxY = timelineY - boxOffsetEmu - BOX_H;
    const botBoxY = timelineY + boxOffsetEmu;
    const topRouteBase = topBoxY + BOX_H + Math.round(bh * 0.01);
    const topRouteEnd = timelineY - Math.round(bh * 0.02);
    const botRouteBase = timelineY + Math.round(bh * 0.02);
    const botRouteEnd = botBoxY - Math.round(bh * 0.01);
    const topLaneH = topCount > 1 ? Math.round((topRouteEnd - topRouteBase) / topCount) : 0;
    const botLaneH = botCount > 1 ? Math.round((botRouteEnd - botRouteBase) / botCount) : 0;
    // Scale node size proportionally — matches web: Math.min(w * 0.012, h * 0.02)
    const nodeSize = Math.round(Math.min(bw * 0.012, bh * 0.02));

    // Assign placements
    const placements: any[] = [];
    let tIdx = 0, bIdx = 0;
    for (let i = 0; i < count; i++) {
      const ts = timestamps[i];
      const dateX = isNaN(ts) ? tlLeft + Math.round(i / Math.max(count - 1, 1) * tlWidth) : dateToX(ts);
      const isTop = i < topCount;
      if (isTop) {
        placements.push({ isTop: true, boxX: topSol.positions[tIdx], boxY: topBoxY, boxW: topSol.boxW, boxH: BOX_H, boxCenterX: topSol.positions[tIdx] + Math.round(topSol.boxW / 2), dateX: dateX, routeY: topRouteBase + tIdx * topLaneH });
        tIdx++;
      } else {
        placements.push({ isTop: false, boxX: botSol.positions[bIdx], boxY: botBoxY, boxW: botSol.boxW, boxH: BOX_H, boxCenterX: botSol.positions[bIdx] + Math.round(botSol.boxW / 2), dateX: dateX, routeY: botRouteBase + bIdx * botLaneH });
        bIdx++;
      }
    }

    const self = this;

    // Helper: create a simple line shape (not cxnSp — avoids PowerPoint connector snapping)
    function addLine(x1, y1, x2, y2, width, color) {
      const lx = Math.min(x1, x2), ly = Math.min(y1, y2);
      const lcx = Math.abs(x2 - x1) || 1; // minimum 1 EMU to avoid zero-extent
      const lcy = Math.abs(y2 - y1) || 1;
      const sh = new DM.Shape();
      sh.id = self._nextId(); sh.name = 'Line'; sh.type = 'cxnSp';
      sh.xfrm = { off: { x: lx, y: ly }, ext: { cx: lcx, cy: lcy }, rot: 0, flipH: x2 < x1, flipV: y2 < y1 };
      sh.geometry = { type: 'straightConnector1' };
      sh.line = new DM.LineProps({ width: width, color: color });
      slide.shapes.push(sh);
    }

    // Draw timeline line
    addLine(tlLeft, timelineY, tlRight, timelineY, lineWidth, lineColor);

    // End caps — scale with line width, matches web: r = lineW + 1
    const capSize = Math.round(lineWidth * 2.2);
    [tlLeft, tlRight].forEach(function(cx) {
      const cap = new DM.Shape();
      cap.id = self._nextId(); cap.name = 'Cap';
      cap.xfrm = { off: { x: cx - capSize / 2, y: timelineY - capSize / 2 }, ext: { cx: capSize, cy: capSize }, rot: 0, flipH: false, flipV: false };
      cap.geometry = { type: 'ellipse' };
      cap.fill = new DM.Fill('solid', { color: lineColor });
      slide.shapes.push(cap);
    });

    // Date formatter
    const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    function fmtDate(iso) {
      const d = new Date(iso);
      if (isNaN(d.getTime())) return iso;
      const m = MONTHS[d.getMonth()], y = d.getFullYear(), q = Math.floor(d.getMonth() / 3) + 1;
      switch (dateFormat) {
        case 'month': return m;
        case 'year': return String(y);
        case 'quarter': return 'Q' + q;
        case 'quarter-year': return 'Q' + q + ' ' + y;
        default: return m + ' ' + y;
      }
    }

    // Responsive font sizes — match web: Math.max(6, Math.min(9, boxW/14)) for name
    // Web px → PPTX hundredths-of-point: multiply by ~10 (approximate px-to-pt at 96dpi ≈ 0.75, ×100)
    function nameFontSize(boxW) {
      let pxApprox = boxW / 914400 * 96; // EMU to px
      let webPx = Math.max(6, Math.min(9, pxApprox / 14));
      return Math.round(webPx * 100); // px ≈ pt for these small sizes, ×100 for hundredths
    }
    function detailFontSize(boxW) {
      let pxApprox = boxW / 914400 * 96;
      let webPx = Math.max(5, Math.min(7, pxApprox / 18));
      return Math.round(webPx * 100);
    }

    // Date label size — proportional to node, matches web: width=60px, fontSize=5px
    const dateLabelW = Math.round(nodeSize * 8);
    const dateLabelH = Math.round(nodeSize * 2.5);

    // Render each event
    for (let i = 0; i < count; i++) {
      const ev = events[i];
      const p = placements[i];

      // Elbow connector: 3 segments (draw first, behind node and text box)
      // Seg 1: vertical from text box to routeY
      const v1Start = p.isTop ? p.boxY + p.boxH : p.routeY;
      const v1End = p.isTop ? p.routeY : p.boxY;
      if (Math.abs(v1End - v1Start) > 1000) {
        addLine(p.boxCenterX, v1Start, p.boxCenterX, v1End, connectorWidth, connectorColor);
      }

      // Seg 2: horizontal from boxCenter to dateX at routeY
      if (Math.abs(p.boxCenterX - p.dateX) > 5000) {
        addLine(p.boxCenterX, p.routeY, p.dateX, p.routeY, connectorWidth, connectorColor);
      }

      // Seg 3: vertical from routeY to timeline node
      const v3Start = p.isTop ? p.routeY : timelineY + nodeSize / 2;
      const v3End = p.isTop ? timelineY - nodeSize / 2 : p.routeY;
      if (Math.abs(v3End - v3Start) > 1000) {
        addLine(p.dateX, v3Start, p.dateX, v3End, connectorWidth, connectorColor);
      }

      // Node circle (on top of connectors)
      const nd = new DM.Shape();
      nd.id = self._nextId(); nd.name = 'Node';
      nd.xfrm = { off: { x: p.dateX - nodeSize / 2, y: timelineY - nodeSize / 2 }, ext: { cx: nodeSize, cy: nodeSize }, rot: 0, flipH: false, flipV: false };
      nd.geometry = { type: 'ellipse' };
      nd.fill = new DM.Fill('solid', { color: nodeColor });
      nd.line = new DM.LineProps({ width: nodeBorderWidth, color: nodeBorderColor });
      slide.shapes.push(nd);

      // Text box (on top of everything)
      const nameFs = nameFontSize(p.boxW);
      const detailFs = detailFontSize(p.boxW);
      const tb = new DM.Shape();
      tb.id = self._nextId(); tb.name = 'EventText';
      tb.xfrm = { off: { x: p.boxX, y: p.boxY }, ext: { cx: p.boxW, cy: p.boxH }, rot: 0, flipH: false, flipV: false };
      tb.geometry = { type: 'roundRect' };
      tb.fill = new DM.Fill('solid', { color: content.boxBackground || '#F0F4F8' });
      tb.line = new DM.LineProps({ width: content.boxBorderWidth != null ? Math.round(content.boxBorderWidth * 12700) : 6350, color: content.boxBorderColor || '#CBD5E1' });
      tb.textBody = new DM.TextBody();
      tb.textBody.bodyPr.anchor = 'ctr';
      tb.textBody.bodyPr.lIns = 45720; tb.textBody.bodyPr.rIns = 45720;
      tb.textBody.bodyPr.tIns = 27432; tb.textBody.bodyPr.bIns = 27432;

      // Name paragraph
      const np = new DM.Paragraph();
      np.pPr.algn = content.nameAlign || 'ctr';
      np.pPr.spcAft = 36000;
      np.runs.push(self._makeRun(String(ev.name || ''), {
        fontSize: content.nameFontSize != null ? content.nameFontSize * 100 : nameFs,
        bold: content.nameBold != null ? content.nameBold : true,
        italic: content.nameItalic || false,
        underline: content.nameUnderline || false,
        color: content.nameColor || s.palette.primaryText || '#1F2937',
        fontFamily: content.nameFont || s.bodyText.font || 'Calibri'
      }));
      tb.textBody.paragraphs.push(np);

      // Detail paragraph(s)
      const detailRunOpts = {
        fontSize: content.detailFontSize != null ? content.detailFontSize * 100 : detailFs,
        bold: content.detailBold || false,
        italic: content.detailItalic || false,
        underline: content.detailUnderline || false,
        color: content.detailColor || s.palette.secondaryText || '#6B7280',
        fontFamily: content.detailFont || s.bodyText.font || 'Calibri'
      };
      if (textMode === 'bullet') {
        const lines = String(ev.detail || '').split('\n').filter(function(l) { return l.trim(); });
        for (let li = 0; li < lines.length; li++) {
          const bp = new DM.Paragraph();
          bp.pPr.algn = 'l';
          bp.pPr.lnSpc = 115;
          bp.pPr.lnSpcType = 'pct';
          bp.pPr.buNone = false;
          bp.pPr.buChar = '\u2022';
          bp.pPr.indent = -171450;
          bp.pPr.marL = 171450;
          bp.runs.push(self._makeRun(lines[li], detailRunOpts));
          tb.textBody.paragraphs.push(bp);
        }
      } else {
        const dp2 = new DM.Paragraph();
        dp2.pPr.algn = 'l';
        dp2.pPr.lnSpc = 115;
        dp2.pPr.lnSpcType = 'pct';
        dp2.runs.push(self._makeRun(String(ev.detail || ''), detailRunOpts));
        tb.textBody.paragraphs.push(dp2);
      }

      slide.shapes.push(tb);

      // Date label near the node
      const dp = content.datePlacement || 'auto';
      const dateBelow = dp === 'bottom' ? true : dp === 'top' ? false : p.isTop;
      const dateLabelY = dateBelow ? timelineY + nodeSize / 2 + Math.round(nodeSize * 0.15) : timelineY - nodeSize / 2 - dateLabelH;
      slide.shapes.push(self._createTextShape('Date', fmtDate(ev.date), {
        x: p.dateX - dateLabelW / 2, y: dateLabelY, cx: dateLabelW, cy: dateLabelH
      }, {
        fontSize: content.dateFontSize != null ? content.dateFontSize * 100 : 500,
        bold: content.dateBold != null ? content.dateBold : true,
        italic: content.dateItalic || false,
        color: content.dateColor || lineColor,
        fontFamily: content.dateFont || s.bodyText.font || 'Calibri', align: 'ctr', anchor: dateBelow ? 't' : 'b'
      }));
    }
  }

  // ---- Profile content renderer ----
  // content = { type: 'profile', name: 'Entity Name', items: ['Point 1', 'Point 2', ...] }
  // Renders bold name at top, followed by small bullet points.
  _renderProfile(slide: any, bounds: Bounds, content: any) {
    if (!content) return;
    const s = this._style;

    const shape = new DM.Shape();
    shape.id = this._nextId();
    shape.name = 'Profile';
    shape.xfrm = {
      off: { x: bounds.x, y: bounds.y },
      ext: { cx: bounds.cx, cy: bounds.cy },
      rot: 0, flipH: false, flipV: false
    };
    shape.geometry = { type: 'rect' };
    shape.fill = new DM.Fill('none');
    shape.textBody = new DM.TextBody();
    shape.textBody.bodyPr.anchor = 't';
    shape.textBody.bodyPr.lIns = 45720;
    shape.textBody.bodyPr.tIns = 22860;

    // Bold name
    if (content.name) {
      const namePara = new DM.Paragraph();
      namePara.pPr.algn = 'l';
      namePara.runs.push(this._makeRun(String(content.name), {
        fontSize: Math.round((s.bodyText.fontSize || 1050) * 0.9),
        bold: true,
        color: s.palette.primaryText || '#333333',
        fontFamily: s.bodyText.font || 'Calibri'
      }));
      shape.textBody.paragraphs.push(namePara);
    }

    // Bullet items
    const items = content.items || [];
    const bulletSize = Math.round((s.bodyText.fontSize || 1050) * 0.75);
    for (let i = 0; i < items.length; i++) {
      const para = new DM.Paragraph();
      para.pPr.algn = 'l';
      para.pPr.buNone = false;
      para.pPr.buChar = '\u2022';
      para.pPr.indent = -171450;
      para.pPr.marL = 171450;
      if (i === 0) para.pPr.spcBef = 45720;
      para.runs.push(this._makeRun(String(items[i]), {
        fontSize: bulletSize,
        color: s.palette.secondaryText || '#666666',
        fontFamily: s.bodyText.font || 'Calibri'
      }));
      shape.textBody.paragraphs.push(para);
    }

    slide.shapes.push(shape);
  }

  // ===== PUBLIC LAYOUT METHODS (all return `this` for chaining) =====

  // Quadrant slide: 4 zones in 2x2 grid with header bars
  // zones = { topLeft: {header, content}, topRight: {header, content}, bottomLeft: {header, content}, bottomRight: {header, content} }
  // Each content = { type: 'statGrid'|'cardGrid'|'text'|'table', items|text|data: ... }
  addQuadrantSlide(title: string, zones: any, options?: any): this {
    const opts = options || {};
    const slide = new DM.Slide(this.pres.slides.length + 1);
    slide.background = new DM.Fill('solid', { color: this._style.slide.background });
    this._addSlideHeader(slide, opts.sectionLabel || null, title);

    const defaults = this._defaultQuadrantBounds();
    const roles = ['topLeft', 'topRight', 'bottomLeft', 'bottomRight'];

    // Use catalog bounds if available
    for (let ri = 0; ri < roles.length; ri++) {
      const role = roles[ri];
      const zoneData = zones[role];
      if (!zoneData) continue;

      const b = defaults[role];

      // Try catalog bounds
      const catBounds = this._getZoneBounds(LC.LAYOUT_TYPES.QUADRANT, role);
      if (catBounds) {
        if (catBounds.headerBar) b.bar = catBounds.headerBar;
        if (catBounds.zone) b.content = catBounds.zone;
      }

      // Header bar
      if (zoneData.header) {
        slide.shapes.push(this._createHeaderBar(zoneData.header, b.bar));
      }

      // Zone content
      if (zoneData.content) {
        this._renderZoneContent(slide, b.content, zoneData.content);
      }
    }

    if (opts.footnote) this._addFootnote(slide, opts.footnote);
    this.pres.slides.push(slide);
    return this;
  }

  // Three-column slide: 3 columns with header bars
  // columns = [ {header, content}, {header, content}, {header, content} ]
  addThreeColumnSlide(title: string, columns: any[], options?: any): this {
    const opts = options || {};
    const slide = new DM.Slide(this.pres.slides.length + 1);
    slide.background = new DM.Fill('solid', { color: this._style.slide.background });
    this._addSlideHeader(slide, opts.sectionLabel || null, title);

    const defaults = this._defaultThreeColBounds();
    const catRoles = [LC.ZONE_ROLES.LEFT, LC.ZONE_ROLES.CENTER, LC.ZONE_ROLES.RIGHT];

    for (let i = 0; i < Math.min(columns.length, 3); i++) {
      const colData = columns[i];
      const b = defaults[i];

      const catBounds = this._getZoneBounds(LC.LAYOUT_TYPES.THREE_COLUMN, catRoles[i]);
      if (catBounds) {
        if (catBounds.headerBar) b.bar = catBounds.headerBar;
        if (catBounds.zone) b.content = catBounds.zone;
      }

      if (colData.header) {
        slide.shapes.push(this._createHeaderBar(colData.header, b.bar));
      }
      if (colData.content) {
        this._renderZoneContent(slide, b.content, colData.content);
      }
    }

    if (opts.footnote) this._addFootnote(slide, opts.footnote);
    this.pres.slides.push(slide);
    return this;
  }

  // Two-panel slide: 2 zones (horizontal or vertical)
  // panels = { orientation: 'horizontal'|'vertical', first: {header, content}, second: {header, content} }
  addTwoPanelSlide(title: string, panels: any, options?: any): this {
    const opts = options || {};
    const slide = new DM.Slide(this.pres.slides.length + 1);
    slide.background = new DM.Fill('solid', { color: this._style.slide.background });
    this._addSlideHeader(slide, opts.sectionLabel || null, title);

    const sw = this._slideW();
    const sh = this._slideH();
    const margin = this._marginLeft();
    const gap = this._style.spacing.zoneGap || Math.round(0.15 * 914400);
    const headerTop = this._contentTop();
    const barH = this._style.spacing.headerBarHeight || Math.round(0.35 * 914400);
    const contentGap = this._style.spacing.contentGap || Math.round(0.1 * 914400);
    const isVert = (panels.orientation || 'horizontal') === 'vertical';

    let bounds: any;
    if (isVert) {
      const topH = Math.round((sh - margin - headerTop - barH * 2 - contentGap * 2 - gap) / 2);
      const midY = headerTop + barH + contentGap + topH + gap;
      bounds = [
        {
          bar: { x: margin, y: headerTop, cx: sw - 2 * margin, cy: barH },
          content: { x: margin, y: headerTop + barH + contentGap, cx: sw - 2 * margin, cy: topH }
        },
        {
          bar: { x: margin, y: midY, cx: sw - 2 * margin, cy: barH },
          content: { x: margin, y: midY + barH + contentGap, cx: sw - 2 * margin, cy: topH }
        }
      ];
    } else {
      const colW = Math.round((sw - 2 * margin - gap) / 2);
      const contentH = sh - margin - headerTop - barH - contentGap;
      bounds = [
        {
          bar: { x: margin, y: headerTop, cx: colW, cy: barH },
          content: { x: margin, y: headerTop + barH + contentGap, cx: colW, cy: contentH }
        },
        {
          bar: { x: margin + colW + gap, y: headerTop, cx: colW, cy: barH },
          content: { x: margin + colW + gap, y: headerTop + barH + contentGap, cx: colW, cy: contentH }
        }
      ];
    }

    // Try catalog bounds
    const catRoles = isVert ? [LC.ZONE_ROLES.TOP, LC.ZONE_ROLES.BOTTOM] : [LC.ZONE_ROLES.LEFT, LC.ZONE_ROLES.RIGHT];
    for (let pi = 0; pi < 2; pi++) {
      const catBounds = this._getZoneBounds(LC.LAYOUT_TYPES.TWO_PANEL, catRoles[pi]);
      if (catBounds) {
        if (catBounds.headerBar) bounds[pi].bar = catBounds.headerBar;
        if (catBounds.zone) bounds[pi].content = catBounds.zone;
      }
    }

    const panelData = [panels.first || {}, panels.second || {}];
    for (let i = 0; i < 2; i++) {
      if (panelData[i].header) {
        slide.shapes.push(this._createHeaderBar(panelData[i].header, bounds[i].bar));
      }
      if (panelData[i].content) {
        this._renderZoneContent(slide, bounds[i].content, panelData[i].content);
      }
    }

    if (opts.footnote) this._addFootnote(slide, opts.footnote);
    this.pres.slides.push(slide);
    return this;
  }

  // Timeline slide: horizontal timeline with events
  // events = [ {date, label, detail?} ]
  addTimelineSlide(title: string, events: any[], options?: any): this {
    const opts = options || {};
    const slide = new DM.Slide(this.pres.slides.length + 1);
    const s = this._style;
    slide.background = new DM.Fill('solid', { color: this._style.slide.background });
    this._addSlideHeader(slide, opts.sectionLabel || null, title);

    const sw = this._slideW();
    const sh = this._slideH();
    const margin = this._marginLeft();
    const contentTop = this._contentTop();
    const timelineY = Math.round((contentTop + sh - margin) / 2); // center vertically
    const lineWidth = sw - 2 * margin;

    // Horizontal timeline line
    const lineShape = new DM.Shape();
    lineShape.id = this._nextId();
    lineShape.name = 'Timeline Line';
    lineShape.type = 'cxnSp';
    lineShape.xfrm = {
      off: { x: margin, y: timelineY },
      ext: { cx: lineWidth, cy: 0 },
      rot: 0, flipH: false, flipV: false
    };
    lineShape.geometry = { type: 'straightConnector1' };
    lineShape.line = new DM.LineProps({ width: 19050, color: s.palette.accent });
    slide.shapes.push(lineShape);

    // Events along the timeline
    const count = events.length;
    if (count === 0) {
      this.pres.slides.push(slide);
      return this;
    }

    const spacing = Math.round(lineWidth / (count + 1));

    for (let i = 0; i < count; i++) {
      const ev = events[i];
      const evX = margin + spacing * (i + 1);
      const nodeSize = Math.round(0.15 * 914400);

      // Vertical connector from timeline to node
      const vLine = new DM.Shape();
      vLine.id = this._nextId();
      vLine.name = 'Event Connector';
      vLine.type = 'cxnSp';
      vLine.xfrm = {
        off: { x: evX, y: timelineY - Math.round(0.3 * 914400) },
        ext: { cx: 0, cy: Math.round(0.6 * 914400) },
        rot: 0, flipH: false, flipV: false
      };
      vLine.geometry = { type: 'straightConnector1' };
      vLine.line = new DM.LineProps({ width: 12700, color: s.palette.accent });
      slide.shapes.push(vLine);

      // Node circle
      const node = new DM.Shape();
      node.id = this._nextId();
      node.name = 'Event Node';
      node.xfrm = {
        off: { x: evX - Math.round(nodeSize / 2), y: timelineY - Math.round(nodeSize / 2) },
        ext: { cx: nodeSize, cy: nodeSize },
        rot: 0, flipH: false, flipV: false
      };
      node.geometry = { type: 'ellipse' };
      node.fill = new DM.Fill('solid', { color: s.palette.accent });
      slide.shapes.push(node);

      // Date label above timeline
      const dateW = Math.round(spacing * 0.9);
      const dateShape = this._createTextShape(
        'Event Date', String(ev.date || ''),
        { x: evX - Math.round(dateW / 2), y: timelineY - Math.round(1.2 * 914400), cx: dateW, cy: Math.round(0.6 * 914400) },
        {
          fontSize: Math.round(s.bodyText.fontSize * 0.9),
          bold: true,
          color: s.palette.accent,
          fontFamily: s.bodyText.font,
          align: 'ctr', anchor: 'b'
        }
      );
      slide.shapes.push(dateShape);

      // Label below timeline
      const labelShape = this._createTextShape(
        'Event Label', String(ev.label || ''),
        { x: evX - Math.round(dateW / 2), y: timelineY + Math.round(0.5 * 914400), cx: dateW, cy: Math.round(0.4 * 914400) },
        {
          fontSize: s.bodyText.fontSize,
          bold: true,
          color: s.palette.primaryText,
          fontFamily: s.bodyText.font,
          align: 'ctr', anchor: 't'
        }
      );
      slide.shapes.push(labelShape);

      // Detail text
      if (ev.detail) {
        const detailShape = this._createTextShape(
          'Event Detail', String(ev.detail),
          { x: evX - Math.round(dateW / 2), y: timelineY + Math.round(0.9 * 914400), cx: dateW, cy: Math.round(0.6 * 914400) },
          {
            fontSize: Math.round(s.bodyText.fontSize * 0.85),
            color: s.palette.secondaryText,
            fontFamily: s.bodyText.font,
            align: 'ctr', anchor: 't'
          }
        );
        slide.shapes.push(detailShape);
      }
    }

    if (opts.footnote) this._addFootnote(slide, opts.footnote);
    this.pres.slides.push(slide);
    return this;
  }

  // Card grid slide: adaptive grid of bordered cards
  // cards = [ {title, lines: []} ]
  addCardGridSlide(title: string, cards: any[], options?: any): this {
    const opts = options || {};
    const slide = new DM.Slide(this.pres.slides.length + 1);
    slide.background = new DM.Fill('solid', { color: this._style.slide.background });
    this._addSlideHeader(slide, opts.sectionLabel || null, title);

    const contentBounds = {
      x: this._marginLeft(),
      y: this._contentTop(),
      cx: this._bodyWidth(),
      cy: this._bodyHeight()
    };

    this._renderCardGrid(slide, contentBounds, cards);

    if (opts.footnote) this._addFootnote(slide, opts.footnote);
    this.pres.slides.push(slide);
    return this;
  }

  // Stat grid slide: grid of stat boxes (value + label)
  // stats = [ {value, label, sublabel?} ]
  addStatGridSlide(title: string, stats: any[], options?: any): this {
    const opts = options || {};
    const slide = new DM.Slide(this.pres.slides.length + 1);
    slide.background = new DM.Fill('solid', { color: this._style.slide.background });
    this._addSlideHeader(slide, opts.sectionLabel || null, title);

    const contentBounds = {
      x: this._marginLeft(),
      y: this._contentTop(),
      cx: this._bodyWidth(),
      cy: this._bodyHeight()
    };

    this._renderStatGrid(slide, contentBounds, stats);

    if (opts.footnote) this._addFootnote(slide, opts.footnote);
    this.pres.slides.push(slide);
    return this;
  }

  // Dashboard slide: 4 panels in 2x2, like quadrant but semantically for data/charts
  addDashboardSlide(title: string, panels: any, options?: any): this {
    return this.addQuadrantSlide(title, panels, options);
  }

  // Get layout catalog (for inspection/debugging)
  getCatalog(): any {
    return this._catalog;
  }

  // ===== GRID LAYOUT SYSTEM =====
  // Describe slides as nested rows/columns (like MUI Grid), compute positions automatically.

  // Create a SlideLayout configured with this builder's dimensions and style
  _createLayout(slideMargin?: number, titleFontSizePt?: number, titleLines?: number) {
    const s = this._style;
    const labelH = s.sectionLabel.fontSize / 100 * 914400 / 72 * 1.6;
    const titlePt = titleFontSizePt || s.mainTitle.fontSize / 100;
    const titleH = titlePt * 914400 / 72 * 1.6 * (titleLines || 1);

    // Subheader height = font size in inches + small padding
    const subFontPt = (s.bodyText.fontSize || 1050) / 100; // hundredths of pt → pt
    const subH = subFontPt / 72 + 0.06; // font height in inches + padding

    return new SlideLayout({
      slideWidth: this._slideW(),
      slideHeight: this._slideH(),
      margin: slideMargin != null ? slideMargin : this._marginLeft() / 914400,
      sectionLabelHeight: labelH / 914400,
      titleHeight: titleH / 914400,
      footerHeight: 0.35,
      defaultGap: 0.12,
      headerBarHeight: 0.35,
      subheaderHeight: subH
    });
  }

  // Add a slide from a grid layout spec
  // spec = { title?, sectionLabel?, footer?, body: { direction, children: [{span, header?, content?, children?}] } }
  // options = { debug: bool }
  async addGridSlide(spec: SlideSpec, options?: any): Promise<this> {
    const opts = options || {};
    const slide = new DM.Slide(this.pres.slides.length + 1);
    slide.background = new DM.Fill('solid', { color: this._style.slide.background });

    // Fit the title: shrink before wrapping, and reserve a second line when it still wraps.
    const marginIn = spec && spec.margin != null ? spec.margin : this._marginLeft() / 914400;
    const fit = spec.title
      ? fitTitle({ text: spec.title, fontSizePt: this._style.mainTitle.fontSize / 100, font: this._style.mainTitle.font, widthInches: this._slideW() / 914400 - 2 * marginIn })
      : null;
    const layout = this._createLayout(spec && spec.margin, fit?.fontSizePt, fit?.lines);
    const tree = layout.compute(spec);

    // Render section label
    if (tree.sectionLabelBounds && spec.sectionLabel) {
      let s = this._style;
      const lb = tree.sectionLabelBounds;
      const labelShape = this._createTextShape(
        'Section Label', spec.sectionLabel,
        { x: lb.x, y: lb.y, cx: lb.cx, cy: lb.cy },
        {
          fontSize: s.sectionLabel.fontSize,
          bold: s.sectionLabel.bold,
          color: s.sectionLabel.color,
          fontFamily: s.sectionLabel.font,
          align: 'l', anchor: 't'
        }
      );
      slide.shapes.push(labelShape);
    }

    // Render title
    if (tree.titleBounds && spec.title) {
      let s = this._style;
      const tb = tree.titleBounds;
      const titleShape = this._createTextShape(
        'Title', spec.title,
        { x: tb.x, y: tb.y, cx: tb.cx, cy: tb.cy },
        {
          fontSize: fit ? Math.round(fit.fontSizePt * 100) : s.mainTitle.fontSize,
          bold: s.mainTitle.bold,
          color: s.mainTitle.color,
          fontFamily: s.mainTitle.font,
          align: 'l', anchor: 't'
        }
      );
      slide.shapes.push(titleShape);
    }

    // Render footer
    if (tree.footerBounds && spec.footer) {
      let s = this._style;
      const fb = tree.footerBounds;
      const footerShape = this._createTextShape(
        'Footer', spec.footer,
        { x: fb.x, y: fb.y, cx: fb.cx, cy: fb.cy },
        {
          fontSize: s.footnote.fontSize,
          color: s.footnote.color,
          fontFamily: s.footnote.font,
          align: 'l', anchor: 'b'
        }
      );
      slide.shapes.push(footerShape);
    }

    // Walk the layout tree and render shapes
    await this._renderLayoutNode(slide, tree.body, opts.debug ? 0 : -1);

    this._addSlideNumber(slide);
    this.pres.slides.push(slide);
    return this;
  }

  // Recursively render a layout node
  async _renderLayoutNode(slide: any, node: any, debugDepth: number, parentChildShape?: string | null, childIndex?: number, parentChildShapeColor?: string | null, siblingCount?: number, parentBounds?: Bounds | null, pyramidSpan?: { top: number; height: number } | null): Promise<void> {
    if (!node) return;

    // Build shadow descriptor from node if present
    const nodeShadow = node.shadow || (node._node && node._node.shadow) || null;
    let shadowDesc: any = null;
    if (nodeShadow) {
      shadowDesc = {
        blur: nodeShadow.blur != null ? nodeShadow.blur : 4,
        dist: nodeShadow.dist != null ? nodeShadow.dist : 3,
        dir: nodeShadow.dir != null ? nodeShadow.dir : 45,
        color: nodeShadow.color || '#000000',
        alpha: nodeShadow.alpha != null ? nodeShadow.alpha : 0.5,
        size: nodeShadow.size != null ? nodeShadow.size : 100
      };
    }

    // Render background fill (colored rectangle behind all content)
    const chevronFill = parentChildShape === 'chevron' && !node.background ? (parentChildShapeColor || '#E5E7EB') : null;
    const pyramidFill = parentChildShape === 'pyramid' && !node.background ? (parentChildShapeColor || '#E5E7EB') : null;
    // Compute PPTX chevron adj: 8% of width, converted to short-side units (adj is relative to min(w,h))
    let chevronAdj: any = null;
    if (parentChildShape === 'chevron' && node.bounds) {
      const ss = Math.min(node.bounds.cx, node.bounds.cy);
      chevronAdj = ss > 0 ? Math.round(0.08 * node.bounds.cx / ss * 100000) : 8000;
    }
    // Compute PPTX pyramid: custom polygon geometry per child
    let pyramidXfrm: any = null;
    let pyramidGeom: any = null;
    if (parentChildShape === 'pyramid' && node.bounds && parentBounds && pyramidSpan && pyramidSpan.height > 0) {
      const maxInset = 50; // max inset % per side (50 = pointy top)
      const topFrac = (node.bounds.y - pyramidSpan.top) / pyramidSpan.height;
      const bottomFrac = (node.bounds.y + node.bounds.cy - pyramidSpan.top) / pyramidSpan.height;
      const topInsetPct = maxInset * (1 - topFrac);
      const bottomInsetPct = maxInset * (1 - bottomFrac);
      // Shape bounds span the full parent width so polygon coordinates map cleanly
      pyramidXfrm = {
        off: { x: parentBounds.x, y: node.bounds.y },
        ext: { cx: parentBounds.cx, cy: node.bounds.cy },
        rot: 0, flipH: false, flipV: false
      };
      // Custom polygon in 0-100000 coordinate space
      const tL = Math.round(topInsetPct * 1000);      // top-left x
      const tR = Math.round((100 - topInsetPct) * 1000); // top-right x
      const bL = Math.round(bottomInsetPct * 1000);    // bottom-left x
      const bR = Math.round((100 - bottomInsetPct) * 1000); // bottom-right x
      pyramidGeom = { type: 'custom', path: [
        { x: tL, y: 0 }, { x: tR, y: 0 },
        { x: bR, y: 100000 }, { x: bL, y: 100000 }
      ]};
    }
    if ((node.background || shadowDesc || chevronFill || pyramidFill) && node.bounds) {
      const bgShape = new DM.Shape();
      bgShape.id = this._nextId();
      bgShape.name = 'Background';
      bgShape.xfrm = pyramidXfrm || {
        off: { x: node.bounds.x, y: node.bounds.y },
        ext: { cx: node.bounds.cx, cy: node.bounds.cy },
        rot: 0, flipH: false, flipV: false
      };
      bgShape.geometry = parentChildShape === 'chevron' ? { type: childIndex === 0 ? 'homePlate' : 'chevron', avLst: [{ name: 'adj', val: chevronAdj }] } : pyramidGeom ? pyramidGeom : { type: 'rect' };
      // Shadow needs a visible fill to render properly in PowerPoint — use white if no explicit background
      bgShape.fill = node.background ? new DM.Fill('solid', { color: node.background, alpha: node.backgroundOpacity != null ? node.backgroundOpacity : 100 }) : chevronFill ? new DM.Fill('solid', { color: chevronFill }) : pyramidFill ? new DM.Fill('solid', { color: pyramidFill }) : (shadowDesc ? new DM.Fill('solid', { color: '#FFFFFF' }) : new DM.Fill('none'));
      if (shadowDesc) bgShape.shadow = shadowDesc;
      slide.shapes.push(bgShape);
    }

    // Render border (solid outline, no fill)
    if (node.bounds && (node.border || node.borderTop || node.borderRight || node.borderBottom || node.borderLeft)) {
      const hasSideBorders = node.borderTop || node.borderRight || node.borderBottom || node.borderLeft;

      if (!hasSideBorders && node.border) {
        // All-sides border as a rectangle outline
        const borderColor = typeof node.border === 'string' ? node.border : (node.border.color || '#999999');
        const borderWidth = (node.border && node.border.width) || 9525;
        const borderShape = new DM.Shape();
        borderShape.id = this._nextId();
        borderShape.name = 'Border';
        borderShape.xfrm = pyramidXfrm || {
          off: { x: node.bounds.x, y: node.bounds.y },
          ext: { cx: node.bounds.cx, cy: node.bounds.cy },
          rot: 0, flipH: false, flipV: false
        };
        borderShape.geometry = parentChildShape === 'chevron' ? { type: childIndex === 0 ? 'homePlate' : 'chevron', avLst: [{ name: 'adj', val: chevronAdj }] } : pyramidGeom ? pyramidGeom : { type: 'rect' };
        borderShape.fill = new DM.Fill('none');
        const borderDash = (typeof node.border === 'object' && node.border.style) || 'solid';
        borderShape.line = new DM.LineProps({ width: borderWidth, color: borderColor, dash: borderDash });
        slide.shapes.push(borderShape);
      } else {
        // Per-side borders as individual line connectors
        const bx = node.bounds.x, by = node.bounds.y, bcx = node.bounds.cx, bcy = node.bounds.cy;
        const sides = [
          { def: node.borderTop || (node.border && !hasSideBorders ? node.border : null), name: 'Border Top',
            x: bx, y: by, cx: bcx, cy: 0 },
          { def: node.borderRight || (node.border && !hasSideBorders ? node.border : null), name: 'Border Right',
            x: bx + bcx, y: by, cx: 0, cy: bcy },
          { def: node.borderBottom || (node.border && !hasSideBorders ? node.border : null), name: 'Border Bottom',
            x: bx, y: by + bcy, cx: bcx, cy: 0 },
          { def: node.borderLeft || (node.border && !hasSideBorders ? node.border : null), name: 'Border Left',
            x: bx, y: by, cx: 0, cy: bcy },
        ];

        for (let si = 0; si < sides.length; si++) {
          const sd = sides[si];
          if (!sd.def) continue;
          const sideObj = typeof sd.def === 'object' ? sd.def : { color: sd.def };
          const sColor = sideObj.color || '#999999';
          const sWidth = sideObj.width || 9525;
          const sDash = sideObj.style || 'solid';
          const lineShape = new DM.Shape();
          lineShape.id = this._nextId();
          lineShape.name = sd.name;
          lineShape.xfrm = {
            off: { x: sd.x, y: sd.y },
            ext: { cx: sd.cx, cy: sd.cy },
            rot: 0, flipH: false, flipV: false
          };
          lineShape.geometry = { type: 'line' };
          lineShape.fill = new DM.Fill('none');
          lineShape.line = new DM.LineProps({ width: sWidth, color: sColor, dash: sDash });
          slide.shapes.push(lineShape);
        }
      }
    }

    // Debug: render cell border
    if (debugDepth >= 0) {
      const debugLabel = node.header || node.subheader || '';
      this._renderDebugCell(slide, node.bounds, debugLabel, debugDepth);
    }

    // Render header bar (dark background, white text)
    if (node.headerBounds && node.header) {
      slide.shapes.push(this._createHeaderBar(node.header, node.headerBounds));
    }

    // Render subheader (bold text + line separator)
    if (node.subheaderBounds && node.subheader) {
      const sub = this._createSubheader(
        node.subheader,
        node.subheaderBounds,
        node.subheaderLineBounds || { x: node.subheaderBounds.x, y: node.subheaderBounds.y + node.subheaderBounds.cy, cx: node.subheaderBounds.cx, cy: 0 }
      );
      slide.shapes.push(sub.textShape);
      slide.shapes.push(sub.lineShape);
    }

    // Render subfooter (small text at bottom)
    if (node.subfooterBounds && node.subfooter) {
      slide.shapes.push(this._createSubfooter(node.subfooter, node.subfooterBounds));
    }

    // Leaf node: render content
    if (node.content && node.contentBounds) {
      await this._renderZoneContent(slide, node.contentBounds, node.content);
    }

    // Container: recurse into children
    if (node.children) {
      const nodeChildShape = node.childShape || (node._node && node._node.childShape) || null;
      const nodeChildShapeColor = node.childShapeColor || (node._node && node._node.childShapeColor) || null;
      // Pre-compute pyramid span (top to bottom of all children including gaps)
      let childPyramidSpan: any = null;
      if (nodeChildShape === 'pyramid' && node.children.length > 1) {
        const first = node.children[0];
        const last = node.children[node.children.length - 1];
        if (first && first.bounds && last && last.bounds) {
          childPyramidSpan = { top: first.bounds.y, height: (last.bounds.y + last.bounds.cy) - first.bounds.y };
        }
      }
      for (let i = 0; i < node.children.length; i++) {
        await this._renderLayoutNode(slide, node.children[i], debugDepth >= 0 ? debugDepth + 1 : -1, nodeChildShape, i, nodeChildShapeColor, node.children.length, node.bounds, childPyramidSpan);
      }
    }

    // Zone borders: render on top of all content so borders are visible over tables/images
    if (this._zoneBorders && node.bounds) {
      this._renderZoneBorder(slide, node.bounds);
    }
  }

  // Debug: render a colored bordered rectangle for a grid cell
  _renderDebugCell(slide: any, bounds: Bounds, label: string, depth: number) {
    const color = DEBUG_COLORS[depth % DEBUG_COLORS.length];
    const shape = new DM.Shape();
    shape.id = this._nextId();
    shape.name = 'Debug Cell';
    shape.xfrm = {
      off: { x: bounds.x, y: bounds.y },
      ext: { cx: bounds.cx, cy: bounds.cy },
      rot: 0, flipH: false, flipV: false
    };
    shape.geometry = { type: 'rect' };
    shape.fill = new DM.Fill('none');
    shape.line = new DM.LineProps({ width: 12700, color: color, dash: 'dash' });

    // Small label in top-left corner
    if (label) {
      shape.textBody = new DM.TextBody();
      shape.textBody.bodyPr.anchor = 't';
      shape.textBody.bodyPr.lIns = 45720;
      shape.textBody.bodyPr.tIns = 22860;
      const para = new DM.Paragraph();
      para.pPr.algn = 'l';
      const run = new DM.Run(String(label));
      run.rPr = {
        sz: 700, b: false, i: true, u: 'none', strike: 'noStrike',
        color: color, highlight: null, fontFamily: 'Arial', baseline: 0
      };
      para.runs.push(run);
      shape.textBody.paragraphs.push(para);
    }

    slide.shapes.push(shape);
  }

  // Zone border: thin dashed border for AI visual inspection
  _renderZoneBorder(slide: any, bounds: Bounds) {
    const shape = new DM.Shape();
    shape.id = this._nextId();
    shape.name = 'Zone Border';
    shape.xfrm = {
      off: { x: bounds.x, y: bounds.y },
      ext: { cx: bounds.cx, cy: bounds.cy },
      rot: 0, flipH: false, flipV: false
    };
    shape.geometry = { type: 'rect' };
    shape.fill = new DM.Fill('none');
    shape.line = new DM.LineProps({ width: 6350, color: '#5b9aed', dash: 'dash' });
    slide.shapes.push(shape);
  }

  // Theme override
  setTheme(themeConfig: any): this {
    if (themeConfig.colors) {
      const cs = this.pres.theme.clrScheme;
      const keys = Object.keys(themeConfig.colors);
      for (let i = 0; i < keys.length; i++) {
        if (cs.hasOwnProperty(keys[i])) cs[keys[i]] = themeConfig.colors[keys[i]];
      }
    }
    if (themeConfig.fonts) {
      if (themeConfig.fonts.title) this.pres.theme.fontScheme.majorFont.latin = themeConfig.fonts.title;
      if (themeConfig.fonts.body) this.pres.theme.fontScheme.minorFont.latin = themeConfig.fonts.body;
    }
    return this;
  }

  // Output
  async save(outputPath: string): Promise<this> {
    const writer = new PptxWriter();
    await writer.writeToFile(this.pres, outputPath);
    return this;
  }

  async toBuffer(): Promise<Buffer> {
    const writer = new PptxWriter();
    return writer.writeToBuffer(this.pres);
  }

  static DM = DM;
  static Units = U;
}

export default SlideBuilder;
export { SlideBuilder };
