/** Parses a .pptx (OOXML zip) into the document model. Node-only (jszip + @xmldom/xmldom). */

import fs from "node:fs";
import JSZip from "jszip";
import { DOMParser } from "@xmldom/xmldom";
import U from "../units.js";
import DM from "./document-model.js";

// ---- XML Namespace URIs ----
const NS = {
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types'
};

// ---- XML Query Helpers ----
// Note: Using manual loop instead of Array.from() for @xmldom/xmldom compatibility
function qn(parent: any, nsUri: string, localName: string): any[] {
  if (!parent) return [];
  const nl = parent.getElementsByTagNameNS(nsUri, localName);
  const arr: any[] = [];
  for (let i = 0; i < nl.length; i++) arr.push(nl[i]);
  return arr;
}

function q1(parent: any, nsUri: string, localName: string): any {
  if (!parent) return null;
  const nl = parent.getElementsByTagNameNS(nsUri, localName);
  return (nl.length > 0) ? nl[0] : null;
}

// Direct child only (not deep descendants)
function q1d(parent: any, nsUri: string, localName: string): any {
  if (!parent) return null;
  for (let i = 0; i < parent.childNodes.length; i++) {
    const c = parent.childNodes[i];
    if (c.nodeType === 1 && c.namespaceURI === nsUri && c.localName === localName) return c;
  }
  return null;
}

function attr(el: any, name: string): string { return el ? (el.getAttribute(name) || '') : ''; }
function intAttr(el: any, name: string, def?: number): number { const v = el ? el.getAttribute(name) : null; return v != null ? parseInt(v, 10) : (def || 0); }

// ---- Preset Colors ----
const COLOR_TRANSFORM_NAMES = ['lumMod','lumOff','tint','shade','satMod','satOff','alpha'];
const PRESET_COLORS: Record<string, string> = {
  white:'#FFFFFF',black:'#000000',red:'#FF0000',green:'#008000',blue:'#0000FF',yellow:'#FFFF00',
  cyan:'#00FFFF',magenta:'#FF00FF',gray:'#808080',grey:'#808080',silver:'#C0C0C0',maroon:'#800000',
  olive:'#808000',navy:'#000080',purple:'#800080',teal:'#008080',aqua:'#00FFFF',fuchsia:'#FF00FF',
  lime:'#00FF00',orange:'#FFA500',darkBlue:'#00008B',darkCyan:'#008B8B',darkGoldenrod:'#B8860B',
  darkGray:'#A9A9A9',darkGreen:'#006400',darkKhaki:'#BDB76B',darkMagenta:'#8B008B',
  darkOliveGreen:'#556B2F',darkOrange:'#FF8C00',darkOrchid:'#9932CC',darkRed:'#8B0000',
  darkSalmon:'#E9967A',darkSeaGreen:'#8FBC8F',darkSlateBlue:'#483D8B',darkSlateGray:'#2F4F4F',
  darkTurquoise:'#00CED1',darkViolet:'#9400D3',lightBlue:'#ADD8E6',lightCoral:'#F08080',
  lightCyan:'#E0FFFF',lightGray:'#D3D3D3',lightGreen:'#90EE90',lightPink:'#FFB6C1',
  lightSalmon:'#FFA07A',lightSeaGreen:'#20B2AA',lightSkyBlue:'#87CEFA',lightSlateGray:'#778899',
  lightSteelBlue:'#B0C4DE',lightYellow:'#FFFFE0',medBlue:'#0000CD',mediumAquamarine:'#66CDAA',
  mediumBlue:'#0000CD',mediumOrchid:'#BA55D3',mediumPurple:'#9370DB',mediumSeaGreen:'#3CB371',
  mediumSlateBlue:'#7B68EE',mediumSpringGreen:'#00FA9A',mediumTurquoise:'#48D1CC',
  mediumVioletRed:'#C71585',cornflowerBlue:'#6495ED',coral:'#FF7F50',crimson:'#DC143C',
  dkBlue:'#00008B',dkCyan:'#008B8B',dkGoldenrod:'#B8860B',dkGray:'#A9A9A9',dkGreen:'#006400',
  dkMagenta:'#8B008B',dkOliveGreen:'#556B2F',dkOrange:'#FF8C00',dkOrchid:'#9932CC',dkRed:'#8B0000',
  dkSalmon:'#E9967A',dkSeaGreen:'#8FBC8F',dkSlateBlue:'#483D8B',dkSlateGray:'#2F4F4F',
  dkTurquoise:'#00CED1',dkViolet:'#9400D3',ltBlue:'#ADD8E6',ltCoral:'#F08080',ltCyan:'#E0FFFF',
  ltGray:'#D3D3D3',ltGreen:'#90EE90',ltPink:'#FFB6C1',ltSalmon:'#FFA07A',ltSeaGreen:'#20B2AA',
  ltSkyBlue:'#87CEFA',ltSlateGray:'#778899',ltSteelBlue:'#B0C4DE',ltYellow:'#FFFFE0'
};

export interface Relationship { id: string; type: string; target: string; }

// ===== PptxParser =====
export class PptxParser {
  zip: JSZip | null;
  rels: Record<string, any>;
  media: Record<string, Buffer>;

  constructor() {
    this.zip = null;
    this.rels = {};
    this.media = {};
  }

  // Parse from file path (Node.js convenience)
  async parseFile(filePath: string): Promise<any> {
    const data = await fs.promises.readFile(filePath);
    return this.parse(data);
  }

  /** Parse from a Buffer (alias of parse()). */
  async parseBuffer(buf: Buffer | ArrayBuffer | Uint8Array): Promise<any> {
    return this.parse(buf);
  }

  // Parse from Buffer or ArrayBuffer
  async parse(arrayBuffer: Buffer | ArrayBuffer | Uint8Array): Promise<any> {
    this.zip = await JSZip.loadAsync(arrayBuffer);
    this.media = {};

    const pres = new DM.Presentation();

    // Parse root rels
    const rootRels = await this._parseRels('_rels/.rels');

    // Find presentation path
    const presRel = rootRels.find(function(r) {
      return r.type.endsWith('/officeDocument');
    });
    let presPath = presRel ? presRel.target : 'ppt/presentation.xml';
    if (presPath.charAt(0) === '/') presPath = presPath.substring(1);

    // Parse presentation
    const presXml = await this._readXml(presPath);
    if (!presXml) throw new Error('Cannot read presentation.xml');

    // Slide size
    const sldSz = q1(presXml.documentElement, NS.p, 'sldSz');
    if (sldSz) {
      pres.slideSize.cx = intAttr(sldSz, 'cx', DM.DEFAULT_SLIDE_WIDTH);
      pres.slideSize.cy = intAttr(sldSz, 'cy', DM.DEFAULT_SLIDE_HEIGHT);
    }

    // Presentation rels
    const presDir = presPath.substring(0, presPath.lastIndexOf('/') + 1);
    const presRelsPath = presDir + '_rels/' + presPath.substring(presPath.lastIndexOf('/') + 1) + '.rels';
    const presRels = await this._parseRels(presRelsPath);

    // Parse theme
    const themeRel = presRels.find(function(r) { return r.type.indexOf('/theme') !== -1; });
    if (themeRel) {
      const themePath = this._resolvePath(presDir, themeRel.target);
      pres.theme = await this._parseTheme(themePath);
    }
    if (!pres.theme) pres.theme = new DM.Theme();

    // Parse slide masters
    const masterRels = presRels.filter(function(r) { return r.type.indexOf('/slideMaster') !== -1; });
    for (let mi = 0; mi < masterRels.length; mi++) {
      const masterPath = this._resolvePath(presDir, masterRels[mi].target);
      const master = await this._parseSlideMaster(masterPath, pres.theme);
      if (master) {
        pres.slideMasters.push(master);
        const masterDir = masterPath.substring(0, masterPath.lastIndexOf('/') + 1);
        const masterRelsPath = masterDir + '_rels/' + masterPath.substring(masterPath.lastIndexOf('/') + 1) + '.rels';
        const mRels = await this._parseRels(masterRelsPath);
        const layoutRels = mRels.filter(function(r) { return r.type.indexOf('/slideLayout') !== -1; });
        for (let li = 0; li < layoutRels.length; li++) {
          const layoutPath = this._resolvePath(masterDir, layoutRels[li].target);
          const layout = await this._parseSlideLayout(layoutPath, master);
          if (layout) {
            master.slideLayouts.push(layout);
            pres.slideLayouts.push(layout);
          }
        }
      }
    }

    // Parse slides in order
    const sldIdLst = q1(presXml.documentElement, NS.p, 'sldIdLst');
    const sldIds = sldIdLst ? qn(sldIdLst, NS.p, 'sldId') : [];
    for (let si = 0; si < sldIds.length; si++) {
      const rId = attr(sldIds[si], 'r:id');
      const slideRel = presRels.find(function(r) { return r.id === rId; });
      if (!slideRel) continue;
      const slidePath = this._resolvePath(presDir, slideRel.target);
      const slide = await this._parseSlide(slidePath, si + 1, pres);
      if (slide) {
        slide.rId = rId;
        pres.slides.push(slide);
      }
    }

    pres.media = this.media;
    return pres;
  }

  // ---- XML Utilities ----
  async _readXml(path: string): Promise<any> {
    const file = this.zip!.file(path);
    if (!file) return null;
    const text = await file.async('string');
    return new DOMParser().parseFromString(text, 'application/xml');
  }

  async _parseRels(path: string): Promise<Relationship[]> {
    const doc = await this._readXml(path);
    if (!doc) return [];
    const rels: Relationship[] = [];
    const els = doc.getElementsByTagName('Relationship');
    for (let i = 0; i < els.length; i++) {
      rels.push({
        id: attr(els[i], 'Id'),
        type: attr(els[i], 'Type'),
        target: attr(els[i], 'Target')
      });
    }
    return rels;
  }

  _resolvePath(baseDir: string, target: string): string {
    if (target.charAt(0) === '/') return target.substring(1);
    const parts = (baseDir + target).split('/');
    const resolved: string[] = [];
    for (let i = 0; i < parts.length; i++) {
      if (parts[i] === '..') resolved.pop();
      else if (parts[i] !== '.' && parts[i] !== '') resolved.push(parts[i]);
    }
    return resolved.join('/');
  }

  // ---- Theme ----
  async _parseTheme(path: string): Promise<any> {
    const doc = await this._readXml(path);
    if (!doc) return null;
    const theme = new DM.Theme();
    const themeEl = doc.documentElement;
    theme.name = attr(themeEl, 'name') || 'Theme';

    const clrScheme = q1(themeEl, NS.a, 'clrScheme');
    if (clrScheme) {
      theme.clrScheme.name = attr(clrScheme, 'name');
      const colorNames = ['dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink'];
      for (let ci = 0; ci < colorNames.length; ci++) {
        const cEl = q1(clrScheme, NS.a, colorNames[ci]);
        if (cEl) {
          const srgb = q1(cEl, NS.a, 'srgbClr');
          const sys = q1(cEl, NS.a, 'sysClr');
          if (srgb) theme.clrScheme[colorNames[ci]] = '#' + attr(srgb, 'val');
          else if (sys) theme.clrScheme[colorNames[ci]] = '#' + (attr(sys, 'lastClr') || attr(sys, 'val') || '000000');
        }
      }
    }

    const fontScheme = q1(themeEl, NS.a, 'fontScheme');
    if (fontScheme) {
      theme.fontScheme.name = attr(fontScheme, 'name');
      const majorFont = q1(fontScheme, NS.a, 'majorFont');
      const minorFont = q1(fontScheme, NS.a, 'minorFont');
      if (majorFont) {
        const mjLatin = q1(majorFont, NS.a, 'latin');
        if (mjLatin) theme.fontScheme.majorFont.latin = attr(mjLatin, 'typeface');
      }
      if (minorFont) {
        const mnLatin = q1(minorFont, NS.a, 'latin');
        if (mnLatin) theme.fontScheme.minorFont.latin = attr(mnLatin, 'typeface');
      }
    }

    return theme;
  }

  // ---- Slide Master ----
  async _parseSlideMaster(path: string, theme: any): Promise<any> {
    const doc = await this._readXml(path);
    if (!doc) return null;
    const master = new DM.SlideMaster(path);
    master.theme = theme;
    const root = doc.documentElement;

    const clrMap = q1(root, NS.p, 'clrMap');
    if (clrMap) {
      const mapAttrs = ['dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink'];
      for (let i = 0; i < mapAttrs.length; i++) {
        const v = attr(clrMap, mapAttrs[i]);
        if (v) master.clrMap[mapAttrs[i]] = v;
      }
    }

    const cSld = q1(root, NS.p, 'cSld');
    if (cSld) {
      master.background = this._parseBackground(cSld);
      const spTree = q1(cSld, NS.p, 'spTree');
      if (spTree) master.shapes = this._parseShapes(spTree, null);
    }

    return master;
  }

  // ---- Slide Layout ----
  async _parseSlideLayout(path: string, master: any): Promise<any> {
    const doc = await this._readXml(path);
    if (!doc) return null;
    const layout = new DM.SlideLayout(path);
    const root = doc.documentElement;
    layout.name = attr(root, 'name') || '';
    layout.type = attr(root, 'type') || '';
    layout.slideMaster = master;

    const cSld = q1(root, NS.p, 'cSld');
    if (cSld) {
      layout.background = this._parseBackground(cSld);
      layout.name = layout.name || attr(cSld, 'name') || '';
      const spTree = q1(cSld, NS.p, 'spTree');
      if (spTree) layout.shapes = this._parseShapes(spTree, null);
    }

    return layout;
  }

  // ---- Slide ----
  async _parseSlide(path: string, slideNum: number, pres: any): Promise<any> {
    const doc = await this._readXml(path);
    if (!doc) return null;
    const slide = new DM.Slide(slideNum);
    const root = doc.documentElement;

    const slideDir = path.substring(0, path.lastIndexOf('/') + 1);
    const slideRelsPath = slideDir + '_rels/' + path.substring(path.lastIndexOf('/') + 1) + '.rels';
    const slideRels = await this._parseRels(slideRelsPath);
    slide.relationships = slideRels;

    // Find layout
    const layoutRel = slideRels.find(function(r) { return r.type.indexOf('/slideLayout') !== -1; });
    if (layoutRel) {
      const layoutPath = this._resolvePath(slideDir, layoutRel.target);
      slide.slideLayout = pres.slideLayouts.find(function(l: any) { return l.id === layoutPath; }) || null;
    }

    // Load media as Buffers (Node.js — no blob URLs)
    const imageRels = slideRels.filter(function(r) {
      return r.type.indexOf('/image') !== -1;
    });
    for (let ii = 0; ii < imageRels.length; ii++) {
      const imgPath = this._resolvePath(slideDir, imageRels[ii].target);
      if (!this.media[imgPath]) {
        const imgFile = this.zip!.file(imgPath);
        if (imgFile) {
          const buffer = await imgFile.async('nodebuffer');
          this.media[imgPath] = buffer;
        }
      }
    }

    const cSld = q1(root, NS.p, 'cSld');
    if (cSld) {
      slide.background = this._parseBackground(cSld);
      const spTree = q1(cSld, NS.p, 'spTree');
      if (spTree) slide.shapes = this._parseShapes(spTree, slideRels);
    }

    return slide;
  }

  // ---- Background ----
  _parseBackground(cSld: any): any {
    const bg = q1(cSld, NS.p, 'bg');
    if (!bg) return null;
    const bgPr = q1(bg, NS.p, 'bgPr');
    if (bgPr) {
      const solidFill = q1(bgPr, NS.a, 'solidFill');
      if (solidFill) return new DM.Fill('solid', { color: this._parseColor(solidFill) });
    }
    const bgRef = q1(bg, NS.p, 'bgRef');
    if (bgRef) {
      const solidFill2 = q1(bgRef, NS.a, 'solidFill');
      if (solidFill2) return new DM.Fill('solid', { color: this._parseColor(solidFill2) });
    }
    return null;
  }

  // ---- Shape Parsing ----
  _unwrapMC(node: any): any[] {
    if (node.localName !== 'AlternateContent') return [node];
    let choice: any = null, fallback: any = null;
    for (let c = 0; c < node.childNodes.length; c++) {
      const ch = node.childNodes[c];
      if (ch.nodeType !== 1) continue;
      if (ch.localName === 'Choice' && !choice) choice = ch;
      else if (ch.localName === 'Fallback' && !fallback) fallback = ch;
    }
    const container = choice || fallback;
    if (!container) return [];
    const result: any[] = [];
    for (let j = 0; j < container.childNodes.length; j++) {
      if (container.childNodes[j].nodeType === 1) result.push(container.childNodes[j]);
    }
    return result;
  }

  _parseShapes(spTree: any, slideRels: Relationship[] | null): any[] {
    const shapes: any[] = [];
    let shapeId = 1;
    const self = this;
    const children = spTree.childNodes;
    for (let i = 0; i < children.length; i++) {
      const node = children[i];
      if (node.nodeType !== 1) continue;
      const nodes = self._unwrapMC(node);
      for (let ni = 0; ni < nodes.length; ni++) {
        const n = nodes[ni];
        const localName = n.localName;
        let shape: any = null;
        if (localName === 'sp') shape = self._parseSp(n);
        else if (localName === 'pic') shape = self._parsePic(n, slideRels);
        else if (localName === 'cxnSp') shape = self._parseCxnSp(n);
        else if (localName === 'graphicFrame') shape = self._parseGraphicFrame(n, slideRels);
        else if (localName === 'grpSp') {
          const grpShapes = self._parseGroupShape(n, slideRels);
          for (let gi = 0; gi < grpShapes.length; gi++) shapes.push(grpShapes[gi]);
          continue;
        }
        if (shape) {
          if (!shape.id) shape.id = shapeId++;
          shapes.push(shape);
        }
      }
    }
    return shapes;
  }

  _parseGroupShape(grpSp: any, slideRels: Relationship[] | null): any[] {
    const grpSpPr = q1(grpSp, NS.p, 'grpSpPr') || q1(grpSp, NS.a, 'grpSpPr');
    const grpOff = { x: 0, y: 0 };
    const chOff = { x: 0, y: 0 };
    let scaleX = 1, scaleY = 1;
    if (grpSpPr) {
      const xfrmEl = q1(grpSpPr, NS.a, 'xfrm');
      if (xfrmEl) {
        const offEl = q1(xfrmEl, NS.a, 'off');
        const extEl = q1(xfrmEl, NS.a, 'ext');
        const chOffEl = q1(xfrmEl, NS.a, 'chOff');
        const chExtEl = q1(xfrmEl, NS.a, 'chExt');
        if (offEl) { grpOff.x = intAttr(offEl, 'x'); grpOff.y = intAttr(offEl, 'y'); }
        if (chOffEl) { chOff.x = intAttr(chOffEl, 'x'); chOff.y = intAttr(chOffEl, 'y'); }
        if (extEl && chExtEl) {
          const extCx = intAttr(extEl, 'cx') || 1;
          const extCy = intAttr(extEl, 'cy') || 1;
          const chExtCx = intAttr(chExtEl, 'cx') || 1;
          const chExtCy = intAttr(chExtEl, 'cy') || 1;
          scaleX = extCx / chExtCx;
          scaleY = extCy / chExtCy;
        }
      }
    }
    const shapes: any[] = [];
    const self = this;
    const children = grpSp.childNodes;
    for (let i = 0; i < children.length; i++) {
      const node = children[i];
      if (node.nodeType !== 1) continue;
      const nodes = self._unwrapMC(node);
      for (let ni = 0; ni < nodes.length; ni++) {
        const n = nodes[ni];
        let parsed: any = null;
        if (n.localName === 'sp') parsed = self._parseSp(n);
        else if (n.localName === 'pic') parsed = self._parsePic(n, slideRels);
        else if (n.localName === 'cxnSp') parsed = self._parseCxnSp(n);
        else if (n.localName === 'graphicFrame') parsed = self._parseGraphicFrame(n, slideRels);
        else if (n.localName === 'grpSp') {
          const nested = self._parseGroupShape(n, slideRels);
          for (let gi = 0; gi < nested.length; gi++) {
            const ns = nested[gi];
            ns.xfrm.off.x = grpOff.x + (ns.xfrm.off.x - chOff.x) * scaleX;
            ns.xfrm.off.y = grpOff.y + (ns.xfrm.off.y - chOff.y) * scaleY;
            ns.xfrm.ext.cx = Math.round(ns.xfrm.ext.cx * scaleX);
            ns.xfrm.ext.cy = Math.round(ns.xfrm.ext.cy * scaleY);
            shapes.push(ns);
          }
          continue;
        }
        if (parsed) {
          parsed.xfrm.off.x = grpOff.x + (parsed.xfrm.off.x - chOff.x) * scaleX;
          parsed.xfrm.off.y = grpOff.y + (parsed.xfrm.off.y - chOff.y) * scaleY;
          parsed.xfrm.ext.cx = Math.round(parsed.xfrm.ext.cx * scaleX);
          parsed.xfrm.ext.cy = Math.round(parsed.xfrm.ext.cy * scaleY);
          shapes.push(parsed);
        }
      }
    }
    return shapes;
  }

  _parseSp(spEl: any): any {
    const shape = new DM.Shape();
    shape.type = 'sp';

    const nvSpPr = q1(spEl, NS.p, 'nvSpPr');
    if (nvSpPr) {
      const cNvPr = q1(nvSpPr, NS.p, 'cNvPr');
      if (cNvPr) {
        shape.id = intAttr(cNvPr, 'id');
        shape.name = attr(cNvPr, 'name');
        shape.hidden = attr(cNvPr, 'hidden') === '1';
      }
      const nvPr = q1(nvSpPr, NS.p, 'nvPr');
      if (nvPr) {
        const ph = q1(nvPr, NS.p, 'ph');
        if (ph) {
          shape.placeholder = { type: attr(ph, 'type') || null, idx: attr(ph, 'idx') || null };
        }
      }
    }

    const spPr = q1(spEl, NS.p, 'spPr');
    if (spPr) this._parseShapeProps(spPr, shape);

    const txBody = q1(spEl, NS.p, 'txBody');
    if (txBody) shape.textBody = this._parseTextBody(txBody);

    const pStyle = q1(spEl, NS.p, 'style');
    if (pStyle) {
      const fontRef = q1(pStyle, NS.a, 'fontRef');
      if (fontRef) {
        const schemeClr = q1(fontRef, NS.a, 'schemeClr');
        if (schemeClr) {
          shape.defaultTextColor = attr(schemeClr, 'val');
        }
      }
    }

    return shape;
  }

  _parsePic(picEl: any, slideRels: Relationship[] | null): any {
    const shape = new DM.Shape();
    shape.type = 'pic';

    const nvPicPr = q1(picEl, NS.p, 'nvPicPr');
    if (nvPicPr) {
      const cNvPr = q1(nvPicPr, NS.p, 'cNvPr');
      if (cNvPr) {
        shape.id = intAttr(cNvPr, 'id');
        shape.name = attr(cNvPr, 'name');
      }
    }

    const blipFill = q1(picEl, NS.p, 'blipFill');
    if (blipFill) {
      const blip = q1(blipFill, NS.a, 'blip');
      if (blip) {
        const embed = blip.getAttribute('r:embed');
        if (embed) shape.imageRId = embed;
      }
      // Parse srcRect crop values
      const srcRect = q1(blipFill, NS.a, 'srcRect');
      if (srcRect) {
        const cl = parseInt(srcRect.getAttribute('l') || '0', 10);
        const ct = parseInt(srcRect.getAttribute('t') || '0', 10);
        const cr = parseInt(srcRect.getAttribute('r') || '0', 10);
        const cb = parseInt(srcRect.getAttribute('b') || '0', 10);
        if (cl > 0 || ct > 0 || cr > 0 || cb > 0) {
          shape._cropRect = { l: cl, t: ct, r: cr, b: cb };
        }
      }
    }

    const spPr = q1(picEl, NS.p, 'spPr');
    if (spPr) this._parseShapeProps(spPr, shape);

    // Resolve image path via rel (store resolved path for writer to use)
    if (shape.imageRId && slideRels) {
      const imgRel = slideRels.find(function(r) { return r.id === shape.imageRId; });
      if (imgRel) {
        const resolvedPath = this._resolvePath('ppt/slides/', imgRel.target);
        shape.imageUrl = resolvedPath; // In Node: this is the ZIP path, not a blob URL
        shape._imagePath = resolvedPath; // Internal: for writer to find Buffer in media map
      }
    }

    if (shape.xfrm && shape.xfrm.ext) {
      shape.imageOrigSize = { cx: shape.xfrm.ext.cx, cy: shape.xfrm.ext.cy };
      shape.lockAspectRatio = true;
    }

    return shape;
  }

  _parseCxnSp(cxnEl: any): any {
    const shape = new DM.Shape();
    shape.type = 'cxnSp';

    const nvCxnSpPr = q1(cxnEl, NS.p, 'nvCxnSpPr');
    if (nvCxnSpPr) {
      const cNvPr = q1(nvCxnSpPr, NS.p, 'cNvPr');
      if (cNvPr) {
        shape.id = intAttr(cNvPr, 'id');
        shape.name = attr(cNvPr, 'name');
      }
    }

    const spPr = q1(cxnEl, NS.p, 'spPr');
    if (spPr) this._parseShapeProps(spPr, shape);

    if (!shape.geometry) shape.geometry = { type: 'line' };
    return shape;
  }

  // ---- GraphicFrame (Tables) ----
  _parseGraphicFrame(gfEl: any, slideRels: Relationship[] | null): any {
    const graphic = q1(gfEl, NS.a, 'graphic');
    if (!graphic) return null;
    const graphicData = q1(graphic, NS.a, 'graphicData');
    if (!graphicData) return null;
    const uri = attr(graphicData, 'uri');
    if (uri !== 'http://schemas.openxmlformats.org/drawingml/2006/table') return null;

    const tbl = q1(graphicData, NS.a, 'tbl');
    if (!tbl) return null;

    const shape = new DM.Shape();
    shape.type = 'graphicFrame';
    shape.geometry = null;

    const nvGfPr = q1(gfEl, NS.p, 'nvGraphicFramePr');
    if (nvGfPr) {
      const cNvPr = q1(nvGfPr, NS.p, 'cNvPr');
      if (cNvPr) {
        shape.id = intAttr(cNvPr, 'id');
        shape.name = attr(cNvPr, 'name');
      }
    }

    const xfrm = q1(gfEl, NS.p, 'xfrm');
    if (xfrm) {
      const off = q1(xfrm, NS.a, 'off');
      const ext = q1(xfrm, NS.a, 'ext');
      if (off) { shape.xfrm.off.x = intAttr(off, 'x'); shape.xfrm.off.y = intAttr(off, 'y'); }
      if (ext) { shape.xfrm.ext.cx = intAttr(ext, 'cx'); shape.xfrm.ext.cy = intAttr(ext, 'cy'); }
    }

    const tableData = this._parseTable(tbl);
    shape.tableData = tableData;

    let totalW = 0, totalH = 0;
    for (let c = 0; c < tableData.cols.length; c++) totalW += tableData.cols[c].w;
    for (let r = 0; r < tableData.rows.length; r++) totalH += tableData.rows[r].h;
    if (totalW > shape.xfrm.ext.cx) shape.xfrm.ext.cx = totalW;
    if (totalH > shape.xfrm.ext.cy) shape.xfrm.ext.cy = totalH;

    return shape;
  }

  _parseTable(tblEl: any): any {
    const td = new DM.TableData();

    const tblPr = q1(tblEl, NS.a, 'tblPr');
    if (tblPr) {
      td.tblPr.bandRow = attr(tblPr, 'bandRow') === '1';
      td.tblPr.bandCol = attr(tblPr, 'bandCol') === '1';
      td.tblPr.firstRow = attr(tblPr, 'firstRow') === '1';
      td.tblPr.firstCol = attr(tblPr, 'firstCol') === '1';
      td.tblPr.lastRow = attr(tblPr, 'lastRow') === '1';
      td.tblPr.lastCol = attr(tblPr, 'lastCol') === '1';
    }

    const tblGrid = q1(tblEl, NS.a, 'tblGrid');
    if (tblGrid) {
      const gridCols = qn(tblGrid, NS.a, 'gridCol');
      for (let ci = 0; ci < gridCols.length; ci++) {
        td.cols.push({ w: intAttr(gridCols[ci], 'w') });
      }
    }

    const rows = qn(tblEl, NS.a, 'tr');
    for (let ri = 0; ri < rows.length; ri++) {
      const row = new DM.TableRow(intAttr(rows[ri], 'h'));
      const cells = qn(rows[ri], NS.a, 'tc');
      for (let ci2 = 0; ci2 < cells.length; ci2++) {
        row.cells.push(this._parseTableCell(cells[ci2]));
      }
      td.rows.push(row);
    }

    return td;
  }

  _parseTableCell(tcEl: any): any {
    const cell = new DM.TableCell();
    cell.gridSpan = intAttr(tcEl, 'gridSpan', 1);
    cell.rowSpan = intAttr(tcEl, 'rowSpan', 1);
    cell.hMerge = attr(tcEl, 'hMerge') === '1';
    cell.vMerge = attr(tcEl, 'vMerge') === '1';

    const txBody = q1(tcEl, NS.a, 'txBody');
    if (txBody) cell.txBody = this._parseTextBody(txBody);

    const tcPr = q1(tcEl, NS.a, 'tcPr');
    if (tcPr) {
      if (tcPr.hasAttribute('marL')) cell.tcPr.marL = intAttr(tcPr, 'marL');
      if (tcPr.hasAttribute('marR')) cell.tcPr.marR = intAttr(tcPr, 'marR');
      if (tcPr.hasAttribute('marT')) cell.tcPr.marT = intAttr(tcPr, 'marT');
      if (tcPr.hasAttribute('marB')) cell.tcPr.marB = intAttr(tcPr, 'marB');
      cell.tcPr.anchor = attr(tcPr, 'anchor') || 't';

      let solidFill: any = null, noFill: any = null;
      for (let ci4 = 0; ci4 < tcPr.childNodes.length; ci4++) {
        const cn = tcPr.childNodes[ci4];
        if (cn.nodeType !== 1) continue;
        if (cn.localName === 'solidFill' && cn.namespaceURI === NS.a) solidFill = cn;
        if (cn.localName === 'noFill' && cn.namespaceURI === NS.a) noFill = cn;
      }
      if (solidFill) cell.tcPr.fill = new DM.Fill('solid', { color: this._parseColor(solidFill) });
      else if (noFill) cell.tcPr.fill = new DM.Fill('none');

      cell.tcPr.borders = {
        l: this._parseCellBorder(q1(tcPr, NS.a, 'lnL')),
        r: this._parseCellBorder(q1(tcPr, NS.a, 'lnR')),
        t: this._parseCellBorder(q1(tcPr, NS.a, 'lnT')),
        b: this._parseCellBorder(q1(tcPr, NS.a, 'lnB'))
      };
    }

    return cell;
  }

  _parseCellBorder(lnEl: any): any {
    if (!lnEl) return null;
    const w = intAttr(lnEl, 'w', 12700);
    const solidFill = q1(lnEl, NS.a, 'solidFill');
    const noFill = q1(lnEl, NS.a, 'noFill');
    if (noFill) return null;
    return new DM.LineProps({
      width: w,
      color: solidFill ? this._parseColor(solidFill) : '#000000',
      dash: attr(q1(lnEl, NS.a, 'prstDash'), 'val') || 'solid'
    });
  }

  // ---- Shape Properties ----
  _parseShapeProps(spPr: any, shape: any): void {
    const xfrm = q1(spPr, NS.a, 'xfrm');
    if (xfrm) {
      const off = q1(xfrm, NS.a, 'off');
      const ext = q1(xfrm, NS.a, 'ext');
      if (off) { shape.xfrm.off.x = intAttr(off, 'x'); shape.xfrm.off.y = intAttr(off, 'y'); }
      if (ext) { shape.xfrm.ext.cx = intAttr(ext, 'cx'); shape.xfrm.ext.cy = intAttr(ext, 'cy'); }
      shape.xfrm.rot = intAttr(xfrm, 'rot');
      shape.xfrm.flipH = attr(xfrm, 'flipH') === '1';
      shape.xfrm.flipV = attr(xfrm, 'flipV') === '1';
    }

    const prstGeom = q1(spPr, NS.a, 'prstGeom');
    const custGeom = q1(spPr, NS.a, 'custGeom');
    if (prstGeom) {
      const prst = attr(prstGeom, 'prst');
      shape.geometry = { type: this._mapGeometry(prst), avLst: [] };
    } else if (custGeom) {
      shape.geometry = { type: 'custom', pathData: this._parseCustGeom(custGeom) };
    }

    // Fill — use q1d (direct children only) to avoid picking up fills from nested <a:ln>
    const solidFill = q1d(spPr, NS.a, 'solidFill');
    const noFill = q1d(spPr, NS.a, 'noFill');
    const gradFill = q1d(spPr, NS.a, 'gradFill');
    if (solidFill) {
      shape.fill = new DM.Fill('solid', { color: this._parseColor(solidFill) });
    } else if (noFill) {
      shape.fill = new DM.Fill('none');
    } else if (gradFill) {
      shape.fill = this._parseGradientFill(gradFill);
    }

    const ln = q1(spPr, NS.a, 'ln');
    if (ln) {
      const lnWidth = intAttr(ln, 'w', 12700);
      const lnFill = q1(ln, NS.a, 'solidFill');
      const lnNoFill = q1(ln, NS.a, 'noFill');
      if (lnNoFill) {
        shape.line = null;
      } else {
        shape.line = new DM.LineProps({
          width: lnWidth,
          color: lnFill ? this._parseColor(lnFill) : '#000000',
          dash: attr(q1(ln, NS.a, 'prstDash'), 'val') || 'solid'
        });
      }
    }

    const effectLst = q1(spPr, NS.a, 'effectLst');
    if (effectLst) {
      const outerShdw = q1(effectLst, NS.a, 'outerShdw');
      if (outerShdw) {
        const blurRad = intAttr(outerShdw, 'blurRad', 0);
        const dist = intAttr(outerShdw, 'dist', 0);
        const dir = intAttr(outerShdw, 'dir', 0);
        const shdwColor = this._parseColor(outerShdw);
        let shdwAlpha = 1;
        const shdwColorEl = q1(outerShdw, NS.a, 'srgbClr') || q1(outerShdw, NS.a, 'prstClr') || q1(outerShdw, NS.a, 'schemeClr');
        if (shdwColorEl) {
          const alphaEl = q1(shdwColorEl, NS.a, 'alpha');
          if (alphaEl) shdwAlpha = intAttr(alphaEl, 'val', 100000) / 100000;
        }
        const sx = intAttr(outerShdw, 'sx', 0);
        shape.shadow = {
          blur: U.emuToPoints(blurRad),
          dist: U.emuToPoints(dist),
          dir: dir / 60000,
          color: shdwColor,
          alpha: shdwAlpha
        };
        if (sx > 0 && sx !== 100000) shape.shadow.size = Math.round(sx / 1000);
      }
    }
  }

  _mapGeometry(prst: string): string {
    const map: Record<string, string> = {
      'rect': 'rect', 'roundRect': 'roundRect', 'ellipse': 'ellipse',
      'triangle': 'triangle', 'rtTriangle': 'rtTriangle', 'isoscelesTriangle': 'triangle',
      'diamond': 'diamond', 'pentagon': 'pentagon', 'hexagon': 'hexagon',
      'line': 'line', 'straightConnector1': 'line', 'bentConnector3': 'line',
      'rightArrow': 'rightArrow', 'leftArrow': 'leftArrow',
      'upArrow': 'upArrow', 'downArrow': 'downArrow',
      'leftRightArrow': 'leftRightArrow',
      'flowChartProcess': 'rect', 'flowChartDecision': 'diamond',
      'flowChartTerminator': 'roundRect',
      'flowChartConnector': 'ellipse',
      'round2SameRect': 'round2SameRect',
      'leftBrace': 'leftBrace', 'rightBrace': 'rightBrace',
      'star5': 'star5', 'star4': 'star4',
      'cloud': 'cloud', 'heart': 'heart',
      'plus': 'plus', 'cross': 'plus',
      'homePlate': 'homePlate', 'chevron': 'chevron',
      'octagon': 'octagon', 'trapezoid': 'trapezoid',
      'parallelogram': 'parallelogram',
      'snip1Rect': 'rect', 'snip2SameRect': 'rect',
      'round1Rect': 'roundRect', 'round2DiagRect': 'roundRect'
    };
    return map[prst] || 'rect';
  }

  _parseCustGeom(custGeom: any): string {
    const pathLst = q1(custGeom, NS.a, 'pathLst');
    if (!pathLst) return '';
    const paths = qn(pathLst, NS.a, 'path');
    const allD: string[] = [];
    for (let pi = 0; pi < paths.length; pi++) {
      const pathEl = paths[pi];
      const pw = intAttr(pathEl, 'w', 1) || 1;
      const ph = intAttr(pathEl, 'h', 1) || 1;
      let d = '';
      for (let ci6 = 0; ci6 < pathEl.childNodes.length; ci6++) {
        const cmd = pathEl.childNodes[ci6];
        if (cmd.nodeType !== 1) continue;
        const tag = cmd.localName;
        if (tag === 'moveTo') {
          const pt = q1(cmd, NS.a, 'pt');
          if (pt) d += 'M' + (intAttr(pt, 'x', 0) / pw) + ' ' + (intAttr(pt, 'y', 0) / ph) + ' ';
        } else if (tag === 'lnTo') {
          const pt2 = q1(cmd, NS.a, 'pt');
          if (pt2) d += 'L' + (intAttr(pt2, 'x', 0) / pw) + ' ' + (intAttr(pt2, 'y', 0) / ph) + ' ';
        } else if (tag === 'cubicBezTo') {
          const pts = qn(cmd, NS.a, 'pt');
          if (pts.length >= 3) {
            d += 'C' + (intAttr(pts[0], 'x', 0) / pw) + ' ' + (intAttr(pts[0], 'y', 0) / ph) + ' ' +
              (intAttr(pts[1], 'x', 0) / pw) + ' ' + (intAttr(pts[1], 'y', 0) / ph) + ' ' +
              (intAttr(pts[2], 'x', 0) / pw) + ' ' + (intAttr(pts[2], 'y', 0) / ph) + ' ';
          }
        } else if (tag === 'quadBezTo') {
          const pts2 = qn(cmd, NS.a, 'pt');
          if (pts2.length >= 2) {
            d += 'Q' + (intAttr(pts2[0], 'x', 0) / pw) + ' ' + (intAttr(pts2[0], 'y', 0) / ph) + ' ' +
              (intAttr(pts2[1], 'x', 0) / pw) + ' ' + (intAttr(pts2[1], 'y', 0) / ph) + ' ';
          }
        } else if (tag === 'arcTo') {
          d += 'L';
        } else if (tag === 'close') {
          d += 'Z ';
        }
      }
      allD.push(d.trim());
    }
    return allD.join(' ');
  }

  // ---- Color ----
  _parseColor(fillEl: any): any {
    let colorEl: any = null, val = '#000000';
    const srgb = q1(fillEl, NS.a, 'srgbClr');
    if (srgb) { colorEl = srgb; val = '#' + attr(srgb, 'val'); }
    else {
      const schemeClr = q1(fillEl, NS.a, 'schemeClr');
      if (schemeClr) { colorEl = schemeClr; val = attr(schemeClr, 'val'); }
      else {
        const sysClr = q1(fillEl, NS.a, 'sysClr');
        if (sysClr) { colorEl = sysClr; val = '#' + (attr(sysClr, 'lastClr') || '000000'); }
        else {
          const prstClr = q1(fillEl, NS.a, 'prstClr');
          if (prstClr) { colorEl = prstClr; val = PRESET_COLORS[attr(prstClr, 'val')] || '#000000'; }
        }
      }
    }
    const transforms: Array<{ type: string; val: number }> = [];
    if (colorEl) {
      for (let ci5 = 0; ci5 < colorEl.childNodes.length; ci5++) {
        const cn = colorEl.childNodes[ci5];
        if (cn.nodeType !== 1) continue;
        const tName = cn.localName;
        if (COLOR_TRANSFORM_NAMES.indexOf(tName) !== -1) {
          transforms.push({ type: tName, val: intAttr(cn, 'val', 100000) });
        }
      }
    }
    if (transforms.length > 0) return { val: val, transforms: transforms };
    return val;
  }

  _parseGradientFill(gradFill: any): any {
    const gsLst = q1(gradFill, NS.a, 'gsLst');
    const stops: Array<{ pos: number; color: any }> = [];
    if (gsLst) {
      const gs = qn(gsLst, NS.a, 'gs');
      for (let i = 0; i < gs.length; i++) {
        stops.push({
          pos: intAttr(gs[i], 'pos', 0) / 1000,
          color: this._parseColor(gs[i])
        });
      }
    }
    let angle = 0;
    const lin = q1(gradFill, NS.a, 'lin');
    if (lin) angle = intAttr(lin, 'ang', 0) / 60000;

    return new DM.Fill('gradient', { stops: stops, angle: angle });
  }

  // ---- Text Body ----
  _parseTextBody(txBodyEl: any): any {
    const tb = new DM.TextBody();

    const bodyPr = q1(txBodyEl, NS.a, 'bodyPr');
    if (bodyPr) {
      tb.bodyPr.wrap = attr(bodyPr, 'wrap') || 'square';
      tb.bodyPr.anchor = attr(bodyPr, 'anchor') || 't';
      tb.bodyPr.anchorCtr = attr(bodyPr, 'anchorCtr') === '1';
      if (bodyPr.hasAttribute('lIns')) tb.bodyPr.lIns = intAttr(bodyPr, 'lIns');
      if (bodyPr.hasAttribute('tIns')) tb.bodyPr.tIns = intAttr(bodyPr, 'tIns');
      if (bodyPr.hasAttribute('rIns')) tb.bodyPr.rIns = intAttr(bodyPr, 'rIns');
      if (bodyPr.hasAttribute('bIns')) tb.bodyPr.bIns = intAttr(bodyPr, 'bIns');
      const spAutoFit = q1(bodyPr, NS.a, 'spAutoFit');
      const normAutofit = q1(bodyPr, NS.a, 'normAutofit');
      if (spAutoFit) tb.bodyPr.autoFit = 'spAutoFit';
      else if (normAutofit) tb.bodyPr.autoFit = 'shrink';
      const vert = attr(bodyPr, 'vert');
      if (vert) tb.bodyPr.vert = vert;
      const rot2 = bodyPr.getAttribute('rot');
      if (rot2) tb.bodyPr.rot = parseInt(rot2, 10);
    }

    const paragraphs = qn(txBodyEl, NS.a, 'p');
    for (let pi = 0; pi < paragraphs.length; pi++) {
      tb.paragraphs.push(this._parseParagraph(paragraphs[pi]));
    }

    return tb;
  }

  _parseParagraph(pEl: any): any {
    const para = new DM.Paragraph();

    const pPr = q1(pEl, NS.a, 'pPr');
    if (pPr) {
      para.pPr.algn = attr(pPr, 'algn') || 'l';
      para.pPr.lvl = intAttr(pPr, 'lvl');
      para.pPr.indent = intAttr(pPr, 'indent');
      para.pPr.marL = intAttr(pPr, 'marL');

      const spcBef = q1(pPr, NS.a, 'spcBef');
      if (spcBef) {
        const spcPts = q1(spcBef, NS.a, 'spcPts');
        const spcPctB = q1(spcBef, NS.a, 'spcPct');
        if (spcPts) para.pPr.spcBef = intAttr(spcPts, 'val') * 12700 / 100;
        else if (spcPctB) para.pPr.spcBefPct = intAttr(spcPctB, 'val');
      }
      const spcAft = q1(pPr, NS.a, 'spcAft');
      if (spcAft) {
        const spcPts2 = q1(spcAft, NS.a, 'spcPts');
        const spcPctA = q1(spcAft, NS.a, 'spcPct');
        if (spcPts2) para.pPr.spcAft = intAttr(spcPts2, 'val') * 12700 / 100;
        else if (spcPctA) para.pPr.spcAftPct = intAttr(spcPctA, 'val');
      }

      const lnSpcEl = q1(pPr, NS.a, 'lnSpc');
      if (lnSpcEl) {
        const lnPct = q1(lnSpcEl, NS.a, 'spcPct');
        const lnPts = q1(lnSpcEl, NS.a, 'spcPts');
        if (lnPct) {
          para.pPr.lnSpc = intAttr(lnPct, 'val') / 1000;
          para.pPr.lnSpcType = 'pct';
        } else if (lnPts) {
          para.pPr.lnSpc = intAttr(lnPts, 'val');
          para.pPr.lnSpcType = 'pts';
        }
      }

      const buNone = q1(pPr, NS.a, 'buNone');
      const buChar = q1(pPr, NS.a, 'buChar');
      const buAutoNum = q1(pPr, NS.a, 'buAutoNum');
      const buClr = q1(pPr, NS.a, 'buClr');
      const buFont = q1(pPr, NS.a, 'buFont');
      const buSzPct = q1(pPr, NS.a, 'buSzPct');
      if (buNone) para.pPr.buNone = true;
      if (buChar) { para.pPr.buNone = false; para.pPr.buChar = attr(buChar, 'char'); }
      if (buAutoNum) { para.pPr.buNone = false; para.pPr.buAutoNum = attr(buAutoNum, 'type') || 'arabicPeriod'; }
      if (buClr) para.pPr.buClr = this._parseColor(buClr);
      if (buFont) para.pPr.buFont = attr(buFont, 'typeface');
      if (buSzPct) para.pPr.buSzPct = intAttr(buSzPct, 'val');
    }

    const children = pEl.childNodes;
    for (let i = 0; i < children.length; i++) {
      const node = children[i];
      if (node.nodeType !== 1) continue;
      if (node.localName === 'r') {
        para.runs.push(this._parseRun(node));
      } else if (node.localName === 'br') {
        const br = new DM.Run('\n');
        para.runs.push(br);
      }
    }

    return para;
  }

  _parseRun(rEl: any): any {
    const tEl = q1(rEl, NS.a, 't');
    const text = tEl ? (tEl.textContent || '') : '';
    const run = new DM.Run(text);

    const rPr = q1(rEl, NS.a, 'rPr');
    if (rPr) {
      run.rPr.sz = rPr.hasAttribute('sz') ? intAttr(rPr, 'sz') : 1800;
      run.rPr.b = attr(rPr, 'b') === '1';
      run.rPr.i = attr(rPr, 'i') === '1';
      run.rPr.u = attr(rPr, 'u') || 'none';
      run.rPr.strike = attr(rPr, 'strike') || 'noStrike';
      run.rPr.baseline = intAttr(rPr, 'baseline');

      const solidFill = q1(rPr, NS.a, 'solidFill');
      if (solidFill) run.rPr.color = this._parseColor(solidFill);

      const highlight = q1(rPr, NS.a, 'highlight');
      if (highlight) run.rPr.highlight = this._parseColor(highlight);

      const latin = q1(rPr, NS.a, 'latin');
      if (latin) run.rPr.fontFamily = attr(latin, 'typeface');
      const ea = q1(rPr, NS.a, 'ea');
      if (ea && !run.rPr.fontFamily) run.rPr.fontFamily = attr(ea, 'typeface');
      const cs = q1(rPr, NS.a, 'cs');
      if (cs && !run.rPr.fontFamily) run.rPr.fontFamily = attr(cs, 'typeface');
    }

    return run;
  }
}

export default PptxParser;
