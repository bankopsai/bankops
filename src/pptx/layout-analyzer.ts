/** Discovers reusable layout patterns (header bars, grids, zones) from a parsed presentation. Node-only. */

import LC, { LayoutCatalog } from "./layout-catalog.js";
import PptxParser from "./pptx-parser.js";

const EMU_PER_INCH = 914400;

interface SlideAnalysis {
  slideIndex: number;
  classification: string;
  headerBars: any[];
  connectors: any[];
  tables: any[];
  images: any[];
  textShapes: any[];
  zones: any[];
  shapeCount: number;
}

// ===== LayoutAnalyzer =====
// Multi-pass analysis engine that discovers layout patterns from parsed .pptx files.
// Identifies header bars, classifies slides by layout type, discovers grid structures,
// and builds a LayoutCatalog of reusable layout templates.
export class LayoutAnalyzer {
  // ---- Main entry point ----
  analyze(parsedPres: any): LayoutCatalog {
    const catalog = new LC.LayoutCatalog();
    catalog.metadata.slideCount = parsedPres.slides.length;
    catalog.metadata.analyzedAt = new Date().toISOString();

    const slideSize = parsedPres.slideSize;
    const analyses: SlideAnalysis[] = [];

    // Pass 1: Analyze each slide independently
    for (let si = 0; si < parsedPres.slides.length; si++) {
      analyses.push(this._analyzeSlide(parsedPres.slides[si], si, slideSize));
    }

    // Pass 2: Build layout templates from classified slides
    this._buildLayoutTemplates(analyses, catalog, slideSize);

    return catalog;
  }

  // Convenience: analyze from file path
  async analyzeFile(pptxPath: string): Promise<LayoutCatalog> {
    const parser = new PptxParser();
    const pres = await parser.parseFile(pptxPath);
    const catalog = this.analyze(pres);
    catalog.metadata.sourceFile = pptxPath;
    return catalog;
  }

  // ---- Slide Analysis ----

  _analyzeSlide(slide: any, slideIndex: number, slideSize: any): SlideAnalysis {
    const shapes = slide.shapes;
    const headerBars: any[] = [];
    const connectors: any[] = [];
    const tables: any[] = [];
    const images: any[] = [];
    const textShapes: any[] = [];
    let hasBigText = false;
    let maxTextSize = 0;

    for (let i = 0; i < shapes.length; i++) {
      const shape = shapes[i];

      if (shape.type === 'cxnSp') {
        connectors.push(shape);
        continue;
      }
      if (shape.type === 'pic') {
        images.push(shape);
        continue;
      }
      if (shape.tableData) {
        tables.push(shape);
        continue;
      }

      // Check for header bar pattern
      if (this._isHeaderBar(shape)) {
        headerBars.push(shape);
        continue;
      }

      // Check for big text (section divider signal)
      if (shape.textBody) {
        const maxSz = this._maxFontSize(shape);
        if (maxSz > maxTextSize) maxTextSize = maxSz;
        if (maxSz >= 8000) hasBigText = true;
        textShapes.push(shape);
      } else {
        textShapes.push(shape);
      }
    }

    const classification = this._classifySlide(
      headerBars, connectors, tables, images, textShapes,
      hasBigText, maxTextSize, slideIndex, slideSize, shapes.length
    );

    const zones = this._discoverZones(headerBars, slideSize, classification, shapes);

    return {
      slideIndex: slideIndex,
      classification: classification,
      headerBars: headerBars,
      connectors: connectors,
      tables: tables,
      images: images,
      textShapes: textShapes,
      zones: zones,
      shapeCount: shapes.length
    };
  }

  // ---- Header Bar Detection ----
  // A header bar is a wide, short roundRect with solid fill and text content.
  // In the reference deck these use dark teal (#164556) or accent scheme colors.

  _isHeaderBar(shape: any): boolean {
    if (!shape.geometry || shape.geometry.type !== 'roundRect') return false;
    if (!shape.fill || shape.fill.type !== 'solid') return false;

    // Must be wide (>2") and short (<1")
    const cx = shape.xfrm.ext.cx;
    const cy = shape.xfrm.ext.cy;
    if (cx < 1828800 || cy > 914400) return false;

    // Fill must not be white/very light
    const fillColor = _colorVal(shape.fill.color);
    if (_isLightColor(fillColor)) return false;

    // Must contain text
    if (!shape.textBody || shape.textBody.paragraphs.length === 0) return false;
    let hasText = false;
    for (let pi = 0; pi < shape.textBody.paragraphs.length; pi++) {
      for (let ri = 0; ri < shape.textBody.paragraphs[pi].runs.length; ri++) {
        if (shape.textBody.paragraphs[pi].runs[ri].text &&
            shape.textBody.paragraphs[pi].runs[ri].text.trim()) {
          hasText = true;
          break;
        }
      }
      if (hasText) break;
    }

    return hasText;
  }

  // ---- Slide Classification ----

  _classifySlide(
    headerBars: any[], connectors: any[], tables: any[], images: any[], textShapes: any[],
    hasBigText: boolean, maxTextSize: number, slideIndex: number, slideSize: any, totalShapes: number
  ): string {
    const hbCount = headerBars.length;

    // Section divider: very large text, few shapes
    if (hasBigText && totalShapes <= 10) {
      return LC.LAYOUT_TYPES.SECTION_DIVIDER;
    }

    // Title slide: first slide or few shapes with centered medium-large text
    if (slideIndex === 0 && totalShapes <= 10 && !hasBigText && hbCount === 0) {
      return LC.LAYOUT_TYPES.TITLE;
    }

    // Table slide: has table shapes, no header bars or just 1
    if (tables.length > 0 && hbCount <= 1) {
      return LC.LAYOUT_TYPES.TABLE;
    }

    // Quadrant: 4 header bars in 2x2 grid
    if (hbCount === 4 && this._isGridArrangement(headerBars, 2, 2, slideSize)) {
      return LC.LAYOUT_TYPES.QUADRANT;
    }

    // Dashboard: 4 header bars (same as quadrant but may have different content)
    if (hbCount === 4) {
      return LC.LAYOUT_TYPES.QUADRANT;
    }

    // Three column: 3 header bars in a row (same Y)
    if (hbCount === 3 && this._isRowArrangement(headerBars, slideSize)) {
      return LC.LAYOUT_TYPES.THREE_COLUMN;
    }

    // Partial quadrant: 3 header bars in 2x2 grid (one corner empty)
    if (hbCount === 3 && this._isPartialGrid(headerBars, 2, 2, slideSize)) {
      return LC.LAYOUT_TYPES.QUADRANT;
    }

    // Two panel: 2 header bars
    if (hbCount === 2) {
      return LC.LAYOUT_TYPES.TWO_PANEL;
    }

    // Timeline: many connectors (>=5), shapes spread horizontally
    if (connectors.length >= 5) {
      return LC.LAYOUT_TYPES.TIMELINE;
    }

    // Single header bar with table → table
    if (hbCount === 1 && tables.length > 0) {
      return LC.LAYOUT_TYPES.TABLE;
    }

    // Image heavy: many images relative to text
    if (images.length >= 5 && images.length > textShapes.length) {
      return LC.LAYOUT_TYPES.IMAGE_HEAVY;
    }

    // Default: content slide
    return LC.LAYOUT_TYPES.CONTENT;
  }

  // ---- Grid & Arrangement Detection ----

  // Check if header bars form an N x M grid
  _isGridArrangement(headerBars: any[], expectedCols: number, expectedRows: number, slideSize: any): boolean {
    const tolerance = 0.5 * EMU_PER_INCH; // 0.5" tolerance

    const xPositions = headerBars.map(function(s) { return s.xfrm.off.x; });
    const yPositions = headerBars.map(function(s) { return s.xfrm.off.y; });

    const xClusters = _clusterValues(xPositions, tolerance);
    const yClusters = _clusterValues(yPositions, tolerance);

    return xClusters.length === expectedCols && yClusters.length === expectedRows;
  }

  // Check if header bars form a partial N x M grid (one or more corners empty)
  _isPartialGrid(headerBars: any[], expectedCols: number, expectedRows: number, slideSize: any): boolean {
    const tolerance = 0.5 * EMU_PER_INCH;
    const xPositions = headerBars.map(function(s) { return s.xfrm.off.x; });
    const yPositions = headerBars.map(function(s) { return s.xfrm.off.y; });
    const xClusters = _clusterValues(xPositions, tolerance);
    const yClusters = _clusterValues(yPositions, tolerance);
    // Must have the right number of columns and rows
    return xClusters.length <= expectedCols && xClusters.length >= 2 &&
           yClusters.length <= expectedRows && yClusters.length >= 2;
  }

  // Check if header bars are in a single row (similar Y, different X)
  _isRowArrangement(headerBars: any[], slideSize: any): boolean {
    const tolerance = 0.5 * EMU_PER_INCH;
    const yPositions = headerBars.map(function(s) { return s.xfrm.off.y; });
    const yClusters = _clusterValues(yPositions, tolerance);
    return yClusters.length === 1;
  }

  // For 2 header bars, determine if they're side-by-side or stacked
  _twoBarOrientation(headerBars: any[], slideSize: any): string {
    const tolerance = 0.5 * EMU_PER_INCH;
    const y0 = headerBars[0].xfrm.off.y;
    const y1 = headerBars[1].xfrm.off.y;

    if (Math.abs(y0 - y1) < tolerance) {
      return 'horizontal'; // side by side
    }
    return 'vertical'; // stacked
  }

  // ---- Zone Discovery ----
  // Given header bars and slide size, compute rectangular zones

  _discoverZones(headerBars: any[], slideSize: any, classification: string, allShapes: any[]): any[] {
    let zones: any[] = [];
    const sw = slideSize.cx;
    const sh = slideSize.cy;
    const margin = Math.round(0.5 * EMU_PER_INCH);
    const gap = Math.round(0.1 * EMU_PER_INCH);

    if (classification === LC.LAYOUT_TYPES.QUADRANT && headerBars.length >= 3) {
      zones = this._quadrantZones(headerBars, sw, sh);
    } else if (classification === LC.LAYOUT_TYPES.THREE_COLUMN && headerBars.length === 3) {
      zones = this._threeColumnZones(headerBars, sw, sh);
    } else if (classification === LC.LAYOUT_TYPES.TWO_PANEL && headerBars.length === 2) {
      zones = this._twoPanelZones(headerBars, sw, sh);
    } else if (classification === LC.LAYOUT_TYPES.TIMELINE) {
      // Single full zone for timeline
      const tz = new LC.ZoneTemplate(LC.ZONE_ROLES.FULL);
      tz.bounds = { x: margin, y: Math.round(2 * EMU_PER_INCH), cx: sw - 2 * margin, cy: sh - Math.round(3 * EMU_PER_INCH) };
      tz.contentType = LC.CONTENT_TYPES.TIMELINE;
      zones.push(tz);
    } else if (classification === LC.LAYOUT_TYPES.TABLE) {
      const fz = new LC.ZoneTemplate(LC.ZONE_ROLES.FULL);
      fz.bounds = { x: margin, y: Math.round(2 * EMU_PER_INCH), cx: sw - 2 * margin, cy: sh - Math.round(3 * EMU_PER_INCH) };
      fz.contentType = LC.CONTENT_TYPES.TABLE;
      zones.push(fz);
    } else {
      // Default: single full zone
      const dz = new LC.ZoneTemplate(LC.ZONE_ROLES.FULL);
      dz.bounds = { x: margin, y: Math.round(2 * EMU_PER_INCH), cx: sw - 2 * margin, cy: sh - Math.round(3 * EMU_PER_INCH) };
      dz.contentType = LC.CONTENT_TYPES.TEXT;
      zones.push(dz);
    }

    // Try to determine content types within zones from shape analysis
    this._classifyZoneContent(zones, allShapes);

    return zones;
  }

  // Quadrant zones: 4 zones in 2x2 grid based on header bar positions
  _quadrantZones(headerBars: any[], sw: number, sh: number): any[] {
    const tolerance = 0.5 * EMU_PER_INCH;
    const xPositions = headerBars.map(function(s) { return s.xfrm.off.x; });
    const yPositions = headerBars.map(function(s) { return s.xfrm.off.y; });

    const xClusters = _clusterValues(xPositions, tolerance);
    const yClusters = _clusterValues(yPositions, tolerance);

    xClusters.sort(function(a, b) { return a - b; });
    yClusters.sort(function(a, b) { return a - b; });

    const margin = Math.round(0.5 * EMU_PER_INCH);
    const gap = Math.round(0.15 * EMU_PER_INCH);
    const barHeight = headerBars[0].xfrm.ext.cy;

    // Sort bars into grid positions
    const grid: Record<string, any> = {}; // 'row,col' → headerBar
    for (let i = 0; i < headerBars.length; i++) {
      const col = _nearestCluster(headerBars[i].xfrm.off.x, xClusters, tolerance);
      const row = _nearestCluster(headerBars[i].xfrm.off.y, yClusters, tolerance);
      grid[row + ',' + col] = headerBars[i];
    }

    const roles: Array<[string, number, number]> = [
      [LC.ZONE_ROLES.TOP_LEFT, 0, 0],
      [LC.ZONE_ROLES.TOP_RIGHT, 0, 1],
      [LC.ZONE_ROLES.BOTTOM_LEFT, 1, 0],
      [LC.ZONE_ROLES.BOTTOM_RIGHT, 1, 1]
    ];

    const zones: any[] = [];
    for (let ri = 0; ri < roles.length; ri++) {
      const role = roles[ri][0];
      const row = roles[ri][1];
      const col = roles[ri][2];

      const bar = grid[row + ',' + col];
      const zone = new LC.ZoneTemplate(role);

      if (bar) {
        const barX = bar.xfrm.off.x;
        const barY = bar.xfrm.off.y;
        const barW = bar.xfrm.ext.cx;

        // Zone extends from header bar X to header bar X + width
        // and from below header bar down to the next row's header bar (or slide bottom)
        const zoneTop = barY + barHeight + gap;
        let zoneBottom: number;
        if (row === 0 && yClusters.length > 1) {
          zoneBottom = yClusters[1] - gap;
        } else {
          zoneBottom = sh - margin;
        }

        zone.bounds = {
          x: barX,
          y: zoneTop,
          cx: barW,
          cy: Math.max(0, zoneBottom - zoneTop)
        };

        // Store header bar spec
        const hbSpec = new LC.ShapeSpec();
        hbSpec.bounds = { x: barX, y: barY, cx: barW, cy: barHeight };
        hbSpec.geometry = 'roundRect';
        hbSpec.fillColor = _colorVal(bar.fill ? bar.fill.color : null);
        zone.headerBar = hbSpec;
      }

      zones.push(zone);
    }

    return zones;
  }

  // Three-column zones
  _threeColumnZones(headerBars: any[], sw: number, sh: number): any[] {
    // Sort bars by X position
    const sorted = headerBars.slice().sort(function(a, b) {
      return a.xfrm.off.x - b.xfrm.off.x;
    });

    const margin = Math.round(0.5 * EMU_PER_INCH);
    const gap = Math.round(0.15 * EMU_PER_INCH);
    const barHeight = sorted[0].xfrm.ext.cy;
    const roleNames = [LC.ZONE_ROLES.LEFT, LC.ZONE_ROLES.CENTER, LC.ZONE_ROLES.RIGHT];

    const zones: any[] = [];
    for (let i = 0; i < 3; i++) {
      const bar = sorted[i];
      const zone = new LC.ZoneTemplate(roleNames[i]);

      const zoneTop = bar.xfrm.off.y + barHeight + gap;
      zone.bounds = {
        x: bar.xfrm.off.x,
        y: zoneTop,
        cx: bar.xfrm.ext.cx,
        cy: sh - margin - zoneTop
      };

      const hbSpec = new LC.ShapeSpec();
      hbSpec.bounds = { x: bar.xfrm.off.x, y: bar.xfrm.off.y, cx: bar.xfrm.ext.cx, cy: barHeight };
      hbSpec.geometry = 'roundRect';
      hbSpec.fillColor = _colorVal(bar.fill ? bar.fill.color : null);
      zone.headerBar = hbSpec;

      zones.push(zone);
    }

    return zones;
  }

  // Two-panel zones
  _twoPanelZones(headerBars: any[], sw: number, sh: number): any[] {
    const tolerance = 0.5 * EMU_PER_INCH;
    const margin = Math.round(0.5 * EMU_PER_INCH);
    const gap = Math.round(0.15 * EMU_PER_INCH);
    const barHeight = headerBars[0].xfrm.ext.cy;

    const y0 = headerBars[0].xfrm.off.y;
    const y1 = headerBars[1].xfrm.off.y;
    const isHorizontal = Math.abs(y0 - y1) < tolerance;

    let sorted: any[];
    let roleNames: string[];

    if (isHorizontal) {
      // Side by side — sort by X
      sorted = headerBars.slice().sort(function(a, b) { return a.xfrm.off.x - b.xfrm.off.x; });
      roleNames = [LC.ZONE_ROLES.LEFT, LC.ZONE_ROLES.RIGHT];
    } else {
      // Stacked — sort by Y
      sorted = headerBars.slice().sort(function(a, b) { return a.xfrm.off.y - b.xfrm.off.y; });
      roleNames = [LC.ZONE_ROLES.TOP, LC.ZONE_ROLES.BOTTOM];
    }

    const zones: any[] = [];
    for (let i = 0; i < 2; i++) {
      const bar = sorted[i];
      const zone = new LC.ZoneTemplate(roleNames[i]);

      const zoneTop = bar.xfrm.off.y + barHeight + gap;
      let zoneBottom: number;

      if (!isHorizontal && i === 0 && sorted.length > 1) {
        zoneBottom = sorted[1].xfrm.off.y - gap;
      } else {
        zoneBottom = sh - margin;
      }

      zone.bounds = {
        x: bar.xfrm.off.x,
        y: zoneTop,
        cx: bar.xfrm.ext.cx,
        cy: Math.max(0, zoneBottom - zoneTop)
      };

      const hbSpec = new LC.ShapeSpec();
      hbSpec.bounds = { x: bar.xfrm.off.x, y: bar.xfrm.off.y, cx: bar.xfrm.ext.cx, cy: barHeight };
      hbSpec.geometry = 'roundRect';
      hbSpec.fillColor = _colorVal(bar.fill ? bar.fill.color : null);
      zone.headerBar = hbSpec;

      zones.push(zone);
    }

    return zones;
  }

  // ---- Zone Content Classification ----

  _classifyZoneContent(zones: any[], allShapes: any[]): void {
    for (let zi = 0; zi < zones.length; zi++) {
      const zone = zones[zi];
      const b = zone.bounds;
      if (!b || b.cx === 0) continue;

      // Find shapes within this zone's bounds
      const contained: any[] = [];
      for (let si = 0; si < allShapes.length; si++) {
        const s = allShapes[si];
        const sx = s.xfrm.off.x;
        const sy = s.xfrm.off.y;
        // Shape center must be within zone bounds
        const cx = sx + s.xfrm.ext.cx / 2;
        const cy = sy + s.xfrm.ext.cy / 2;
        if (cx >= b.x && cx <= b.x + b.cx && cy >= b.y && cy <= b.y + b.cy) {
          contained.push(s);
        }
      }

      // Classify based on contained shapes
      let hasTable = false;
      let hasBullets = false;
      let imageCount = 0;
      let smallTextCount = 0;
      let borderedRectCount = 0;
      let connectorCount = 0;

      for (let ci = 0; ci < contained.length; ci++) {
        const cs = contained[ci];
        if (cs.tableData) hasTable = true;
        if (cs.type === 'pic') imageCount++;
        if (cs.type === 'cxnSp') connectorCount++;
        if (cs.line && cs.geometry && cs.geometry.type === 'roundRect' && cs.textBody) {
          borderedRectCount++;
        }
        if (cs.textBody) {
          for (let pi = 0; pi < cs.textBody.paragraphs.length; pi++) {
            const p = cs.textBody.paragraphs[pi];
            if (p.pPr.buChar || p.pPr.buAutoNum) hasBullets = true;
          }
          // Small shapes with number-like text → stat
          if (cs.xfrm.ext.cx < 3 * EMU_PER_INCH && cs.xfrm.ext.cy < 1.5 * EMU_PER_INCH) {
            smallTextCount++;
          }
        }
      }

      if (hasTable) {
        zone.contentType = LC.CONTENT_TYPES.TABLE;
      } else if (connectorCount >= 3) {
        zone.contentType = LC.CONTENT_TYPES.TIMELINE;
      } else if (borderedRectCount >= 2) {
        zone.contentType = LC.CONTENT_TYPES.CARD_GRID;
      } else if (hasBullets) {
        // NOTE: BULLET_LIST is not defined in CONTENT_TYPES (legacy behavior: resolves to undefined)
        zone.contentType = LC.CONTENT_TYPES.BULLET_LIST;
      } else if (smallTextCount >= 3) {
        zone.contentType = LC.CONTENT_TYPES.STAT_GRID;
      } else if (imageCount >= 3) {
        zone.contentType = LC.CONTENT_TYPES.IMAGE;
      } else if (!zone.contentType) {
        zone.contentType = LC.CONTENT_TYPES.TEXT;
      }
    }
  }

  // ---- Build Layout Templates ----

  _buildLayoutTemplates(analyses: SlideAnalysis[], catalog: any, slideSize: any): void {
    // Group analyses by classification
    const groups: Record<string, SlideAnalysis[]> = {};
    for (let i = 0; i < analyses.length; i++) {
      const a = analyses[i];
      const type = a.classification;
      if (!groups[type]) groups[type] = [];
      groups[type].push(a);
    }

    // Build a template for each layout type
    for (const type in groups) {
      const group = groups[type];
      const template = new LC.LayoutTemplate(type);
      template.slideSize = { cx: slideSize.cx, cy: slideSize.cy };
      template.frequency = group.length;
      template.slideIndices = group.map(function(a) { return a.slideIndex; });

      // Average zone positions across all slides of this type
      if (group[0].zones.length > 0) {
        // Use first slide's zones as base, then average
        const baseZones = group[0].zones;
        for (let zi = 0; zi < baseZones.length; zi++) {
          const avgZone = new LC.ZoneTemplate(baseZones[zi].role);
          avgZone.contentType = baseZones[zi].contentType;

          let sumX = 0, sumY = 0, sumCX = 0, sumCY = 0;
          let count = 0;
          let hbSumX = 0, hbSumY = 0, hbSumCX = 0, hbSumCY = 0;
          let hbCount = 0;
          let hbColor: any = null;

          for (let gi = 0; gi < group.length; gi++) {
            if (gi < group.length && group[gi].zones.length > zi) {
              const gz = group[gi].zones[zi];
              sumX += gz.bounds.x;
              sumY += gz.bounds.y;
              sumCX += gz.bounds.cx;
              sumCY += gz.bounds.cy;
              count++;

              if (gz.headerBar) {
                hbSumX += gz.headerBar.bounds.x;
                hbSumY += gz.headerBar.bounds.y;
                hbSumCX += gz.headerBar.bounds.cx;
                hbSumCY += gz.headerBar.bounds.cy;
                hbCount++;
                if (!hbColor) hbColor = gz.headerBar.fillColor;
              }
            }
          }

          if (count > 0) {
            avgZone.bounds = {
              x: Math.round(sumX / count),
              y: Math.round(sumY / count),
              cx: Math.round(sumCX / count),
              cy: Math.round(sumCY / count)
            };
          }

          if (hbCount > 0) {
            const hbSpec = new LC.ShapeSpec();
            hbSpec.bounds = {
              x: Math.round(hbSumX / hbCount),
              y: Math.round(hbSumY / hbCount),
              cx: Math.round(hbSumCX / hbCount),
              cy: Math.round(hbSumCY / hbCount)
            };
            hbSpec.geometry = 'roundRect';
            hbSpec.fillColor = hbColor;
            avgZone.headerBar = hbSpec;
          }

          template.zones.push(avgZone);
        }
      }

      catalog.addTemplate(template);
    }
  }

  // Get the maximum font size in a shape's text body
  _maxFontSize(shape: any): number {
    let max = 0;
    if (!shape.textBody) return 0;
    for (let pi = 0; pi < shape.textBody.paragraphs.length; pi++) {
      for (let ri = 0; ri < shape.textBody.paragraphs[pi].runs.length; ri++) {
        const sz = shape.textBody.paragraphs[pi].runs[ri].rPr.sz;
        if (sz > max) max = sz;
      }
    }
    return max;
  }
}

// ===== Helper Functions =====

// Extract the raw color value from a color that might be an object
function _colorVal(c: any): any {
  if (!c) return null;
  if (typeof c === 'object' && c.val !== undefined) return c.val;
  return c;
}

// Check if a color is "light" (white, near-white, or known light scheme colors)
function _isLightColor(color: any): boolean {
  if (!color) return true;
  const c = String(color);

  // Known light scheme colors
  if (c === 'lt1' || c === 'lt2' || c === 'bg1' || c === 'bg2') return true;

  // Hex color check
  if (c.charAt(0) === '#') {
    const hex = c.substring(1);
    if (hex.length === 6) {
      const r = parseInt(hex.substr(0, 2), 16);
      const g = parseInt(hex.substr(2, 2), 16);
      const b = parseInt(hex.substr(4, 2), 16);
      const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
      return luminance > 0.85;
    }
  }

  // Unknown scheme colors (accent1, dk1, etc.) — assume not light
  return false;
}

// Cluster numeric values with given tolerance
// Returns array of cluster centers (sorted)
function _clusterValues(values: number[], tolerance: number): number[] {
  if (values.length === 0) return [];

  const sorted = values.slice().sort(function(a, b) { return a - b; });
  const clusters = [sorted[0]];

  for (let i = 1; i < sorted.length; i++) {
    let found = false;
    for (let ci = 0; ci < clusters.length; ci++) {
      if (Math.abs(sorted[i] - clusters[ci]) < tolerance) {
        // Update cluster center (running average)
        clusters[ci] = Math.round((clusters[ci] + sorted[i]) / 2);
        found = true;
        break;
      }
    }
    if (!found) {
      clusters.push(sorted[i]);
    }
  }

  return clusters.sort(function(a, b) { return a - b; });
}

// Find which cluster center a value is nearest to
function _nearestCluster(value: number, clusters: number[], tolerance: number): number {
  let best = 0;
  let bestDist = Math.abs(value - clusters[0]);
  for (let i = 1; i < clusters.length; i++) {
    const dist = Math.abs(value - clusters[i]);
    if (dist < bestDist) {
      best = i;
      bestDist = dist;
    }
  }
  return best;
}

export default LayoutAnalyzer;
