/** Serializes a document-model Presentation into a .pptx (OOXML zip). Node-only. */

import fs from "node:fs";
import JSZip from "jszip";
import DM from "./document-model.js";
import U from "../units.js";

// XML Namespace URIs
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';

const RT_SLIDE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
const RT_SLIDE_LAYOUT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout';
const RT_SLIDE_MASTER = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster';
const RT_THEME = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme';
const RT_PRES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const RT_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

// Geometry reverse map (internal name → OOXML prst)
const GEOM_MAP: Record<string, string> = {
  'rect': 'rect', 'roundRect': 'roundRect', 'ellipse': 'ellipse',
  'triangle': 'triangle', 'line': 'line', 'rightArrow': 'rightArrow',
  'leftArrow': 'leftArrow', 'upArrow': 'upArrow', 'downArrow': 'downArrow',
  'diamond': 'diamond', 'pentagon': 'pentagon', 'hexagon': 'hexagon',
  'leftRightArrow': 'leftRightArrow', 'rtTriangle': 'rtTriangle',
  'round2SameRect': 'round2SameRect', 'leftBrace': 'leftBrace', 'rightBrace': 'rightBrace',
  'star5': 'star5', 'star4': 'star4', 'cloud': 'cloud', 'heart': 'heart',
  'plus': 'plus', 'homePlate': 'homePlate', 'chevron': 'chevron',
  'octagon': 'octagon', 'trapezoid': 'trapezoid', 'parallelogram': 'parallelogram'
};

// Image extension → content type mapping
const IMAGE_CONTENT_TYPES: Record<string, string> = {
  'png': 'image/png',
  'jpg': 'image/jpeg',
  'jpeg': 'image/jpeg',
  'gif': 'image/gif',
  'bmp': 'image/bmp',
  'tiff': 'image/tiff',
  'tif': 'image/tiff',
  'svg': 'image/svg+xml',
  'emf': 'image/x-emf',
  'wmf': 'image/x-wmf'
};

function esc(str: any): string {
  // Strip characters invalid in XML 1.0 (allowed: #x9, #xA, #xD, #x20+)
  return (str || '').replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function colorVal(c: any): any {
  return (c && typeof c === 'object' && c.val !== undefined) ? c.val : c;
}
function colorTransforms(c: any): any[] {
  return (c && typeof c === 'object' && c.transforms) ? c.transforms : [];
}
function colorToHex(c: any): string {
  const v = colorVal(c);
  if (!v) return '000000';
  if (v.charAt(0) === '#') return v.substring(1).toUpperCase();
  const m = v.match(/rgba?\(\s*(\d+),\s*(\d+),\s*(\d+)/);
  if (m) {
    return (('0' + parseInt(m[1]).toString(16)).slice(-2)
          + ('0' + parseInt(m[2]).toString(16)).slice(-2)
          + ('0' + parseInt(m[3]).toString(16)).slice(-2)).toUpperCase();
  }
  return v;
}

function isSchemeColor(c: any): boolean {
  const v = colorVal(c);
  if (!v || v.charAt(0) === '#') return false;
  const schemes = ['dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink',
    'bg1','bg2','tx1','tx2'];
  return schemes.indexOf(v) !== -1;
}

function colorTransformsXml(transforms: any[]): string {
  if (!transforms || transforms.length === 0) return '';
  let xml = '';
  for (let i = 0; i < transforms.length; i++) {
    xml += '<a:' + transforms[i].type + ' val="' + transforms[i].val + '"/>';
  }
  return xml;
}

// Resolve a relative path from one ZIP directory to a target
function resolvePath(baseDir: string, target: string): string {
  if (target.charAt(0) === '/') return target.substring(1);
  const parts = (baseDir + target).split('/');
  const resolved: string[] = [];
  for (let i = 0; i < parts.length; i++) {
    if (parts[i] === '..') resolved.pop();
    else if (parts[i] !== '.' && parts[i] !== '') resolved.push(parts[i]);
  }
  return resolved.join('/');
}

interface ImageTracker {
  allImages: Record<string, boolean>;
  slideImages: Array<Array<{ rId: string; imgPath: string; relTarget: string }>>;
  slideCharts: Array<Array<{ rId: string; chartIdx: number; relTarget: string }>>;
  imageExtensions: Record<string, boolean>;
}

// ===== PptxWriter =====
export class PptxWriter {
  // Build the ZIP object (shared between writeToBuffer and writeToFile)
  _buildZip(pres: any): JSZip {
    const zip = new JSZip();

    // Collect all image paths referenced by slides
    const imageTracker = this._collectImages(pres);

    // [Content_Types].xml
    zip.file('[Content_Types].xml', this._contentTypes(pres, imageTracker));

    // _rels/.rels
    zip.file('_rels/.rels', this._rootRels());

    // ppt/presentation.xml
    zip.file('ppt/presentation.xml', this._presentationXml(pres));

    // ppt/_rels/presentation.xml.rels
    zip.file('ppt/_rels/presentation.xml.rels', this._presentationRels(pres));

    // ppt/theme/theme1.xml
    zip.file('ppt/theme/theme1.xml', this._themeXml(pres.theme));

    // ppt/slideMasters/slideMaster1.xml
    zip.file('ppt/slideMasters/slideMaster1.xml', this._slideMasterXml(pres));
    zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', this._slideMasterRels(pres));

    // ppt/slideLayouts/slideLayout1.xml
    zip.file('ppt/slideLayouts/slideLayout1.xml', this._slideLayoutXml());
    zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', this._slideLayoutRels());

    // Slides
    for (let i = 0; i < pres.slides.length; i++) {
      const slideNum = i + 1;
      const slideImages = imageTracker.slideImages[i] || [];
      const slideCharts = imageTracker.slideCharts[i] || [];
      zip.file('ppt/slides/slide' + slideNum + '.xml', this._slideXml(pres.slides[i], pres, slideImages));
      zip.file('ppt/slides/_rels/slide' + slideNum + '.xml.rels', this._slideRels(pres.slides[i], slideNum, slideImages, slideCharts));
    }

    // Write image files into ZIP
    const writtenMedia: Record<string, boolean> = {};
    for (const imgPath in imageTracker.allImages) {
      if (!writtenMedia[imgPath] && pres.media && pres.media[imgPath]) {
        zip.file(imgPath, pres.media[imgPath]);
        writtenMedia[imgPath] = true;
      }
    }

    // Write native chart parts into ZIP (chart XML, embedded Excel, rels)
    if (pres._chartParts) {
      for (const chartPartPath in pres._chartParts) {
        zip.file(chartPartPath, pres._chartParts[chartPartPath]);
      }
    }

    // docProps/app.xml and core.xml
    zip.file('docProps/app.xml', this._appXml(pres));
    zip.file('docProps/core.xml', this._coreXml());

    return zip;
  }

  // Collect all images and native charts referenced across slides, assign rIds
  _collectImages(pres: any): ImageTracker {
    const tracker: ImageTracker = {
      allImages: {},       // imgPath → true
      slideImages: [],     // per-slide: [{rId, imgPath, relTarget}]
      slideCharts: [],     // per-slide: [{rId, chartIdx, relTarget}]
      imageExtensions: {}  // extension → true (for content types)
    };

    for (let si = 0; si < pres.slides.length; si++) {
      const slide = pres.slides[si];
      const slideImgs: Array<{ rId: string; imgPath: string; relTarget: string }> = [];
      const slideCharts: Array<{ rId: string; chartIdx: number; relTarget: string }> = [];
      let rIdCounter = 2; // rId1 = slideLayout

      for (let shi = 0; shi < slide.shapes.length; shi++) {
        const shape = slide.shapes[shi];
        if (shape.type === 'pic' && shape._imagePath && pres.media && pres.media[shape._imagePath]) {
          const imgPath = shape._imagePath;
          const ext = imgPath.split('.').pop().toLowerCase();
          // Compute relative target from ppt/slides/ to ppt/media/
          const relTarget = '../media/' + imgPath.split('/').pop();
          const newRId = 'rId' + rIdCounter++;

          slideImgs.push({ rId: newRId, imgPath: imgPath, relTarget: relTarget });
          // Update shape's imageRId to the new rId for this slide
          shape._exportRId = newRId;

          tracker.allImages[imgPath] = true;
          if (ext) tracker.imageExtensions[ext] = true;
        }
        // Native chart shapes
        if (shape.type === 'nativeChart' && shape._chartIdx) {
          const chartRId = 'rId' + rIdCounter++;
          slideCharts.push({ rId: chartRId, chartIdx: shape._chartIdx, relTarget: '../charts/chart' + shape._chartIdx + '.xml' });
          shape._exportRId = chartRId;
        }
      }

      tracker.slideImages.push(slideImgs);
      tracker.slideCharts.push(slideCharts);
    }

    return tracker;
  }

  // Write to Node.js Buffer
  async writeToBuffer(pres: any): Promise<Buffer> {
    const zip = this._buildZip(pres);
    return zip.generateAsync({ type: 'nodebuffer' });
  }

  // Write to file
  async writeToFile(pres: any, filePath: string): Promise<void> {
    const buffer = await this.writeToBuffer(pres);
    await fs.promises.writeFile(filePath, buffer);
  }

  // Legacy write() method — returns Buffer in Node.js
  async write(pres: any): Promise<Buffer> {
    return this.writeToBuffer(pres);
  }

  /** Alias of writeToBuffer(pres). */
  async toBuffer(pres: any): Promise<Buffer> {
    return this.writeToBuffer(pres);
  }

  /** Alias of writeToFile(pres, filePath). */
  async save(pres: any, filePath: string): Promise<void> {
    return this.writeToFile(pres, filePath);
  }

  // ---- [Content_Types].xml ----
  _contentTypes(pres: any, imageTracker: ImageTracker): string {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += '<Types xmlns="' + NS_CT + '">';
    xml += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
    xml += '<Default Extension="xml" ContentType="application/xml"/>';

    // Image content types
    if (imageTracker) {
      for (const ext in imageTracker.imageExtensions) {
        const ct = IMAGE_CONTENT_TYPES[ext];
        if (ct) xml += '<Default Extension="' + ext + '" ContentType="' + ct + '"/>';
      }
    }

    // Native chart default extensions
    if (pres._chartParts) {
      let hasXlsx = false;
      for (const cpPath in pres._chartParts) {
        if (cpPath.match(/\.xlsx$/)) hasXlsx = true;
      }
      if (hasXlsx) xml += '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>';
    }

    xml += '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>';
    xml += '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>';
    xml += '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>';
    xml += '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>';
    for (let i = 0; i < pres.slides.length; i++) {
      xml += '<Override PartName="/ppt/slides/slide' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>';
    }

    // Native chart content type overrides
    if (pres._chartContentTypes) {
      for (let cti = 0; cti < pres._chartContentTypes.length; cti++) {
        const ct = pres._chartContentTypes[cti];
        xml += '<Override PartName="' + ct.partName + '" ContentType="' + ct.contentType + '"/>';
      }
    }

    xml += '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
    xml += '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
    xml += '</Types>';
    return xml;
  }

  // ---- _rels/.rels ----
  _rootRels(): string {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += '<Relationships xmlns="' + NS_REL + '">';
    xml += '<Relationship Id="rId1" Type="' + RT_PRES + '" Target="ppt/presentation.xml"/>';
    xml += '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>';
    xml += '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>';
    xml += '</Relationships>';
    return xml;
  }

  // ---- ppt/presentation.xml ----
  _presentationXml(pres: any): string {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += '<p:presentation xmlns:a="' + NS_A + '" xmlns:r="' + NS_R + '" xmlns:p="' + NS_P + '"';
    xml += ' saveSubsetFonts="1">';

    xml += '<p:sldMasterIdLst>';
    xml += '<p:sldMasterId id="2147483648" r:id="rId1"/>';
    xml += '</p:sldMasterIdLst>';

    xml += '<p:sldIdLst>';
    for (let i = 0; i < pres.slides.length; i++) {
      xml += '<p:sldId id="' + (256 + i) + '" r:id="rId' + (i + 3) + '"/>';
    }
    xml += '</p:sldIdLst>';

    xml += '<p:sldSz cx="' + pres.slideSize.cx + '" cy="' + pres.slideSize.cy + '" type="screen4x3"/>';
    xml += '<p:notesSz cx="' + pres.slideSize.cy + '" cy="' + pres.slideSize.cx + '"/>';

    xml += '</p:presentation>';
    return xml;
  }

  // ---- ppt/_rels/presentation.xml.rels ----
  _presentationRels(pres: any): string {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += '<Relationships xmlns="' + NS_REL + '">';
    xml += '<Relationship Id="rId1" Type="' + RT_SLIDE_MASTER + '" Target="slideMasters/slideMaster1.xml"/>';
    xml += '<Relationship Id="rId2" Type="' + RT_THEME + '" Target="theme/theme1.xml"/>';
    for (let i = 0; i < pres.slides.length; i++) {
      xml += '<Relationship Id="rId' + (i + 3) + '" Type="' + RT_SLIDE + '" Target="slides/slide' + (i + 1) + '.xml"/>';
    }
    xml += '</Relationships>';
    return xml;
  }

  // ---- Theme ----
  _themeXml(theme: any): string {
    if (!theme) theme = new DM.Theme();
    const cs = theme.clrScheme;
    const fsch = theme.fontScheme;

    function clrEl(name: string, hex: any): string {
      return '<a:' + name + '><a:srgbClr val="' + colorToHex(hex) + '"/></a:' + name + '>';
    }
    function sysClr(name: string, val: string, lastClr: any): string {
      return '<a:' + name + '><a:sysClr val="' + val + '" lastClr="' + colorToHex(lastClr) + '"/></a:' + name + '>';
    }

    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += '<a:theme xmlns:a="' + NS_A + '" name="' + esc(theme.name) + '">';
    xml += '<a:themeElements>';

    xml += '<a:clrScheme name="' + esc(cs.name || 'Office') + '">';
    xml += sysClr('dk1', 'windowText', cs.dk1);
    xml += sysClr('lt1', 'window', cs.lt1);
    xml += clrEl('dk2', cs.dk2);
    xml += clrEl('lt2', cs.lt2);
    xml += clrEl('accent1', cs.accent1);
    xml += clrEl('accent2', cs.accent2);
    xml += clrEl('accent3', cs.accent3);
    xml += clrEl('accent4', cs.accent4);
    xml += clrEl('accent5', cs.accent5);
    xml += clrEl('accent6', cs.accent6);
    xml += clrEl('hlink', cs.hlink);
    xml += clrEl('folHlink', cs.folHlink);
    xml += '</a:clrScheme>';

    xml += '<a:fontScheme name="' + esc(fsch.name || 'Office') + '">';
    xml += '<a:majorFont><a:latin typeface="' + esc(fsch.majorFont.latin) + '"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>';
    xml += '<a:minorFont><a:latin typeface="' + esc(fsch.minorFont.latin) + '"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>';
    xml += '</a:fontScheme>';

    xml += '<a:fmtScheme name="Office">';
    xml += '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>';
    xml += '<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>';
    xml += '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>';
    xml += '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>';
    xml += '</a:fmtScheme>';

    xml += '</a:themeElements>';
    xml += '<a:objectDefaults/><a:extraClrSchemeLst/>';
    xml += '</a:theme>';
    return xml;
  }

  // ---- Slide Master ----
  _slideMasterXml(pres: any): string {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += '<p:sldMaster xmlns:a="' + NS_A + '" xmlns:r="' + NS_R + '" xmlns:p="' + NS_P + '">';
    xml += '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>';
    xml += '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
    xml += '</p:spTree></p:cSld>';
    xml += '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';
    xml += '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>';
    xml += '</p:sldMaster>';
    return xml;
  }

  _slideMasterRels(pres: any): string {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += '<Relationships xmlns="' + NS_REL + '">';
    xml += '<Relationship Id="rId1" Type="' + RT_SLIDE_LAYOUT + '" Target="../slideLayouts/slideLayout1.xml"/>';
    xml += '<Relationship Id="rId2" Type="' + RT_THEME + '" Target="../theme/theme1.xml"/>';
    xml += '</Relationships>';
    return xml;
  }

  // ---- Slide Layout ----
  _slideLayoutXml(): string {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += '<p:sldLayout xmlns:a="' + NS_A + '" xmlns:r="' + NS_R + '" xmlns:p="' + NS_P + '" type="blank" preserve="1">';
    xml += '<p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>';
    xml += '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
    xml += '</p:spTree></p:cSld>';
    xml += '</p:sldLayout>';
    return xml;
  }

  _slideLayoutRels(): string {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += '<Relationships xmlns="' + NS_REL + '">';
    xml += '<Relationship Id="rId1" Type="' + RT_SLIDE_MASTER + '" Target="../slideMasters/slideMaster1.xml"/>';
    xml += '</Relationships>';
    return xml;
  }

  // ---- Slide ----
  _slideXml(slide: any, pres: any, slideImages: any[]): string {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += '<p:sld xmlns:a="' + NS_A + '" xmlns:r="' + NS_R + '" xmlns:p="' + NS_P + '">';
    xml += '<p:cSld>';

    if (slide.background) {
      xml += '<p:bg><p:bgPr>';
      xml += this._fillXml(slide.background);
      xml += '<a:effectLst/>';
      xml += '</p:bgPr></p:bg>';
    }

    xml += '<p:spTree>';
    xml += '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>';
    xml += '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

    for (let i = 0; i < slide.shapes.length; i++) {
      xml += this._shapeXml(slide.shapes[i], i + 2);
    }

    xml += '</p:spTree>';
    xml += '</p:cSld>';
    xml += '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>';
    xml += '</p:sld>';
    return xml;
  }

  _slideRels(slide: any, slideNum: number, slideImages: any[], slideCharts: any[]): string {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += '<Relationships xmlns="' + NS_REL + '">';
    xml += '<Relationship Id="rId1" Type="' + RT_SLIDE_LAYOUT + '" Target="../slideLayouts/slideLayout1.xml"/>';

    // Image relationships
    if (slideImages) {
      for (let i = 0; i < slideImages.length; i++) {
        const img = slideImages[i];
        xml += '<Relationship Id="' + img.rId + '" Type="' + RT_IMAGE + '" Target="' + img.relTarget + '"/>';
      }
    }

    // Native chart relationships
    if (slideCharts) {
      for (let ci = 0; ci < slideCharts.length; ci++) {
        const chart = slideCharts[ci];
        xml += '<Relationship Id="' + chart.rId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="' + chart.relTarget + '"/>';
      }
    }

    xml += '</Relationships>';
    return xml;
  }

  // ---- Shape XML ----
  _shapeXml(shape: any, globalId: number): string {
    if (shape.type === 'pic') return this._picXml(shape, globalId);
    if (shape.type === 'cxnSp') return this._cxnSpXml(shape, globalId);
    if (shape.type === 'graphicFrame' && shape.tableData) return this._graphicFrameXml(shape, globalId);
    if (shape.type === 'nativeChart') return this._nativeChartXml(shape, globalId);
    let xml = '<p:sp>';

    xml += '<p:nvSpPr>';
    xml += '<p:cNvPr id="' + (globalId || shape.id || 1) + '" name="' + esc(shape.name) + '"';
    if (shape.hidden) xml += ' hidden="1"';
    xml += '/>';
    xml += '<p:cNvSpPr';
    if (shape.textBody) xml += ' txBox="1"';
    xml += '/>';
    xml += '<p:nvPr>';
    if (shape.placeholder) {
      xml += '<p:ph';
      if (shape.placeholder.type) xml += ' type="' + shape.placeholder.type + '"';
      if (shape.placeholder.idx) xml += ' idx="' + shape.placeholder.idx + '"';
      xml += '/>';
    }
    xml += '</p:nvPr>';
    xml += '</p:nvSpPr>';

    xml += '<p:spPr>';
    xml += this._xfrmXml(shape.xfrm);
    xml += this._geometryXml(shape.geometry);
    if (shape.fill) xml += this._fillXml(shape.fill);
    if (shape.line) xml += this._lineXml(shape.line);
    if (shape.shadow) {
      const sh = shape.shadow;
      let shdwAttrs = 'blurRad="' + U.pointsToEmu(sh.blur) + '" dist="' + U.pointsToEmu(sh.dist) + '" dir="' + Math.round(sh.dir * 60000) + '" rotWithShape="0"';
      if (sh.size != null && sh.size !== 100) {
        const sizeVal = Math.round(sh.size * 1000);
        shdwAttrs += ' sx="' + sizeVal + '" sy="' + sizeVal + '"';
      }
      xml += '<a:effectLst><a:outerShdw ' + shdwAttrs + '>';
      let shColorStr = sh.color || '#000000';
      if (typeof shColorStr !== 'string') shColorStr = '#000000';
      const shHex = shColorStr.replace('#', '');
      xml += '<a:srgbClr val="' + shHex + '">';
      if (sh.alpha < 1) xml += '<a:alpha val="' + Math.round(sh.alpha * 100000) + '"/>';
      xml += '</a:srgbClr></a:outerShdw></a:effectLst>';
    }
    xml += '</p:spPr>';

    if (shape.textBody) {
      xml += this._textBodyXml(shape.textBody);
    }

    xml += '</p:sp>';
    return xml;
  }

  // ---- Picture XML ----
  _picXml(shape: any, globalId: number): string {
    const rId = shape._exportRId || shape.imageRId || 'rId1';
    let xml = '<p:pic>';
    xml += '<p:nvPicPr>';
    xml += '<p:cNvPr id="' + (globalId || shape.id || 1) + '" name="' + esc(shape.name) + '"/>';
    xml += '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>';
    xml += '<p:nvPr/>';
    xml += '</p:nvPicPr>';
    xml += '<p:blipFill>';
    xml += '<a:blip r:embed="' + rId + '"/>';
    if (shape._cropRect) {
      xml += '<a:srcRect l="' + shape._cropRect.l + '" t="' + shape._cropRect.t + '" r="' + shape._cropRect.r + '" b="' + shape._cropRect.b + '"/>';
    }
    xml += '<a:stretch><a:fillRect/></a:stretch>';
    xml += '</p:blipFill>';
    xml += '<p:spPr>';
    xml += this._xfrmXml(shape.xfrm);
    xml += '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
    if (shape.line) xml += this._lineXml(shape.line);
    xml += '</p:spPr>';
    xml += '</p:pic>';
    return xml;
  }

  // ---- Native Chart XML ----
  _nativeChartXml(shape: any, globalId: number): string {
    const rId = shape._exportRId || 'rId1';
    const chartIdx = shape._chartIdx;

    // If we have a graphicFrame XML snippet from pptxgenjs, use it with rId and id rewritten
    if (shape._graphicFrameXml) {
      let gfXml = shape._graphicFrameXml;
      // Replace the placeholder with actual rId
      gfXml = gfXml.replace('__CHART_RID_' + chartIdx + '__', rId);
      // Also replace any generic placeholder patterns
      gfXml = gfXml.replace(/__CHART_RID_\d+__/g, rId);
      // Rewrite the cNvPr id to avoid collisions with other shapes on the slide
      const safeId = globalId || shape.id || 1;
      gfXml = gfXml.replace(/<p:cNvPr\s+id="\d+"/, '<p:cNvPr id="' + safeId + '"');
      return gfXml;
    }

    // Fallback: build a graphicFrame from scratch
    let xml = '<p:graphicFrame>';
    xml += '<p:nvGraphicFramePr>';
    xml += '<p:cNvPr id="' + (globalId || shape.id || 1) + '" name="' + esc(shape.name) + '"/>';
    xml += '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>';
    xml += '<p:nvPr/>';
    xml += '</p:nvGraphicFramePr>';
    xml += '<p:xfrm>';
    xml += '<a:off x="' + Math.round(shape.xfrm.off.x) + '" y="' + Math.round(shape.xfrm.off.y) + '"/>';
    xml += '<a:ext cx="' + Math.round(shape.xfrm.ext.cx) + '" cy="' + Math.round(shape.xfrm.ext.cy) + '"/>';
    xml += '</p:xfrm>';
    xml += '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">';
    xml += '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="' + NS_R + '" r:id="' + rId + '"/>';
    xml += '</a:graphicData></a:graphic>';
    xml += '</p:graphicFrame>';
    return xml;
  }

  _cxnSpXml(shape: any, globalId: number): string {
    let xml = '<p:cxnSp>';
    xml += '<p:nvCxnSpPr>';
    xml += '<p:cNvPr id="' + (globalId || shape.id || 1) + '" name="' + esc(shape.name) + '"/>';
    xml += '<p:cNvCxnSpPr/>';
    xml += '<p:nvPr/>';
    xml += '</p:nvCxnSpPr>';
    xml += '<p:spPr>';
    xml += this._xfrmXml(shape.xfrm);
    xml += '<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>';
    if (shape.fill) xml += this._fillXml(shape.fill);
    if (shape.line) xml += this._lineXml(shape.line);
    xml += '</p:spPr>';
    xml += '</p:cxnSp>';
    return xml;
  }

  // ---- GraphicFrame / Table ----
  _graphicFrameXml(shape: any, globalId: number): string {
    let xml = '<p:graphicFrame>';
    xml += '<p:nvGraphicFramePr>';
    xml += '<p:cNvPr id="' + (globalId || shape.id || 1) + '" name="' + esc(shape.name) + '"/>';
    xml += '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>';
    xml += '<p:nvPr/>';
    xml += '</p:nvGraphicFramePr>';
    xml += '<p:xfrm>';
    xml += '<a:off x="' + Math.round(shape.xfrm.off.x) + '" y="' + Math.round(shape.xfrm.off.y) + '"/>';
    xml += '<a:ext cx="' + Math.round(shape.xfrm.ext.cx) + '" cy="' + Math.round(shape.xfrm.ext.cy) + '"/>';
    xml += '</p:xfrm>';
    xml += '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">';
    xml += this._tableXml(shape.tableData);
    xml += '</a:graphicData></a:graphic>';
    xml += '</p:graphicFrame>';
    return xml;
  }

  _tableXml(td: any): string {
    let xml = '<a:tbl>';
    xml += '<a:tblPr';
    if (td.tblPr.bandRow) xml += ' bandRow="1"';
    if (td.tblPr.bandCol) xml += ' bandCol="1"';
    if (td.tblPr.firstRow) xml += ' firstRow="1"';
    if (td.tblPr.firstCol) xml += ' firstCol="1"';
    if (td.tblPr.lastRow) xml += ' lastRow="1"';
    if (td.tblPr.lastCol) xml += ' lastCol="1"';
    xml += '/>';

    xml += '<a:tblGrid>';
    for (let ci = 0; ci < td.cols.length; ci++) {
      xml += '<a:gridCol w="' + Math.round(td.cols[ci].w) + '"/>';
    }
    xml += '</a:tblGrid>';

    for (let ri = 0; ri < td.rows.length; ri++) {
      const row = td.rows[ri];
      xml += '<a:tr h="' + Math.round(row.h) + '">';
      for (let ci2 = 0; ci2 < row.cells.length; ci2++) {
        xml += this._tableCellXml(row.cells[ci2]);
      }
      xml += '</a:tr>';
    }

    xml += '</a:tbl>';
    return xml;
  }

  _tableCellXml(cell: any): string {
    let xml = '<a:tc';
    if (cell.gridSpan > 1) xml += ' gridSpan="' + cell.gridSpan + '"';
    if (cell.rowSpan > 1) xml += ' rowSpan="' + cell.rowSpan + '"';
    if (cell.hMerge) xml += ' hMerge="1"';
    if (cell.vMerge) xml += ' vMerge="1"';
    xml += '>';

    if (cell.txBody) {
      xml += this._cellTextBodyXml(cell.txBody);
    } else {
      xml += '<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></a:txBody>';
    }

    const tcPr = cell.tcPr;
    xml += '<a:tcPr';
    xml += ' marL="' + Math.round(tcPr.marL) + '"';
    xml += ' marR="' + Math.round(tcPr.marR) + '"';
    xml += ' marT="' + Math.round(tcPr.marT) + '"';
    xml += ' marB="' + Math.round(tcPr.marB) + '"';
    if (tcPr.anchor && tcPr.anchor !== 't') xml += ' anchor="' + tcPr.anchor + '"';
    xml += '>';

    // Absent borders are written as explicit noFill so PowerPoint and LibreOffice do not
    // fall back to the table style's gridlines.
    const borders = tcPr.borders;
    const none = (name: string) => '<a:' + name + ' w="0"><a:noFill/></a:' + name + '>';
    xml += borders.l ? this._cellBorderXml('lnL', borders.l) : none('lnL');
    xml += borders.r ? this._cellBorderXml('lnR', borders.r) : none('lnR');
    xml += borders.t ? this._cellBorderXml('lnT', borders.t) : none('lnT');
    xml += borders.b ? this._cellBorderXml('lnB', borders.b) : none('lnB');

    if (tcPr.fill) xml += this._fillXml(tcPr.fill);

    xml += '</a:tcPr>';
    xml += '</a:tc>';
    return xml;
  }

  _cellTextBodyXml(tb: any): string {
    let xml = '<a:txBody>';
    xml += '<a:bodyPr/>';
    xml += '<a:lstStyle/>';
    for (let pi = 0; pi < tb.paragraphs.length; pi++) {
      xml += this._paragraphXml(tb.paragraphs[pi]);
    }
    if (tb.paragraphs.length === 0) {
      xml += '<a:p><a:endParaRPr lang="en-US"/></a:p>';
    }
    xml += '</a:txBody>';
    return xml;
  }

  _cellBorderXml(tagName: string, lineProps: any): string {
    let xml = '<a:' + tagName + ' w="' + Math.round(lineProps.width) + '">';
    const clr = lineProps.color || '#000000';
    xml += '<a:solidFill>';
    if (isSchemeColor(clr)) {
      xml += '<a:schemeClr val="' + clr + '"/>';
    } else {
      xml += '<a:srgbClr val="' + colorToHex(clr) + '"/>';
    }
    xml += '</a:solidFill>';
    if (lineProps.dash && lineProps.dash !== 'solid') {
      xml += '<a:prstDash val="' + lineProps.dash + '"/>';
    }
    xml += '</a:' + tagName + '>';
    return xml;
  }

  // ---- Transform ----
  _xfrmXml(xfrm: any): string {
    if (!xfrm) return '';
    let xml = '<a:xfrm';
    if (xfrm.rot) xml += ' rot="' + xfrm.rot + '"';
    if (xfrm.flipH) xml += ' flipH="1"';
    if (xfrm.flipV) xml += ' flipV="1"';
    xml += '>';
    xml += '<a:off x="' + Math.round(xfrm.off.x) + '" y="' + Math.round(xfrm.off.y) + '"/>';
    xml += '<a:ext cx="' + Math.round(xfrm.ext.cx) + '" cy="' + Math.round(xfrm.ext.cy) + '"/>';
    xml += '</a:xfrm>';
    return xml;
  }

  // ---- Geometry ----
  _geometryXml(geom: any): string {
    if (!geom) return '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
    // Custom polygon geometry (array of {x,y} points in 0-100000 coordinate space)
    if (geom.type === 'custom' && geom.path && geom.path.length) {
      const pts = geom.path;
      let pathXml = '<a:moveTo><a:pt x="' + pts[0].x + '" y="' + pts[0].y + '"/></a:moveTo>';
      for (let pi = 1; pi < pts.length; pi++) {
        pathXml += '<a:lnTo><a:pt x="' + pts[pi].x + '" y="' + pts[pi].y + '"/></a:lnTo>';
      }
      pathXml += '<a:close/>';
      return '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="100000" h="100000">' + pathXml + '</a:path></a:pathLst></a:custGeom>';
    }
    const prst = GEOM_MAP[geom.type] || 'rect';
    let avXml = '';
    if (geom.avLst && geom.avLst.length) {
      for (let i = 0; i < geom.avLst.length; i++) {
        avXml += '<a:gd name="' + geom.avLst[i].name + '" fmla="val ' + geom.avLst[i].val + '"/>';
      }
    }
    return '<a:prstGeom prst="' + prst + '"><a:avLst>' + avXml + '</a:avLst></a:prstGeom>';
  }

  // ---- Fill ----
  _fillXml(fill: any): string {
    if (!fill || fill.type === 'none') return '<a:noFill/>';
    if (fill.type === 'solid') {
      let xml = '<a:solidFill>';
      const cv = colorVal(fill.color);
      const ct = colorTransforms(fill.color);
      if (isSchemeColor(fill.color)) {
        if (ct.length > 0) {
          xml += '<a:schemeClr val="' + cv + '">' + colorTransformsXml(ct) + '</a:schemeClr>';
        } else {
          xml += '<a:schemeClr val="' + cv + '"/>';
        }
      } else {
        xml += '<a:srgbClr val="' + colorToHex(fill.color) + '"';
        let childXml = colorTransformsXml(ct);
        if (fill.alpha != null && fill.alpha < 100) {
          childXml += '<a:alpha val="' + (fill.alpha * 1000) + '"/>';
        }
        if (childXml) {
          xml += '>' + childXml + '</a:srgbClr>';
        } else {
          xml += '/>';
        }
      }
      xml += '</a:solidFill>';
      return xml;
    }
    if (fill.type === 'gradient' && fill.stops) {
      let xml = '<a:gradFill><a:gsLst>';
      for (let i = 0; i < fill.stops.length; i++) {
        const sc = fill.stops[i].color;
        const scv = colorVal(sc), sct = colorTransforms(sc);
        xml += '<a:gs pos="' + Math.round(fill.stops[i].pos * 1000) + '">';
        if (isSchemeColor(sc)) {
          if (sct.length > 0) {
            xml += '<a:schemeClr val="' + scv + '">' + colorTransformsXml(sct) + '</a:schemeClr>';
          } else {
            xml += '<a:schemeClr val="' + scv + '"/>';
          }
        } else {
          xml += '<a:srgbClr val="' + colorToHex(sc) + '"/>';
        }
        xml += '</a:gs>';
      }
      xml += '</a:gsLst>';
      if (fill.angle != null) {
        xml += '<a:lin ang="' + Math.round(fill.angle * 60000) + '" scaled="1"/>';
      }
      xml += '</a:gradFill>';
      return xml;
    }
    return '<a:noFill/>';
  }

  // ---- Line ----
  _lineXml(line: any): string {
    if (!line) return '';
    let xml = '<a:ln w="' + Math.round(line.width) + '"';
    if (line.cap && line.cap !== 'flat') xml += ' cap="' + line.cap + '"';
    xml += '>';
    const clr = line.color || '#000000';
    const lcv = colorVal(clr), lct = colorTransforms(clr);
    xml += '<a:solidFill>';
    if (isSchemeColor(clr)) {
      if (lct.length > 0) {
        xml += '<a:schemeClr val="' + lcv + '">' + colorTransformsXml(lct) + '</a:schemeClr>';
      } else {
        xml += '<a:schemeClr val="' + lcv + '"/>';
      }
    } else {
      xml += '<a:srgbClr val="' + colorToHex(clr) + '"/>';
    }
    xml += '</a:solidFill>';
    if (line.dash && line.dash !== 'solid') {
      xml += '<a:prstDash val="' + line.dash + '"/>';
    }
    xml += '</a:ln>';
    return xml;
  }

  // ---- Text Body ----
  _textBodyXml(tb: any): string {
    const bp = tb.bodyPr;
    let xml = '<p:txBody>';
    xml += '<a:bodyPr wrap="' + (bp.wrap || 'square') + '"';
    xml += ' lIns="' + bp.lIns + '" tIns="' + bp.tIns + '" rIns="' + bp.rIns + '" bIns="' + bp.bIns + '"';
    xml += ' anchor="' + (bp.anchor || 't') + '"';
    if (bp.anchorCtr) xml += ' anchorCtr="1"';
    if (bp.vert && bp.vert !== 'horz') xml += ' vert="' + bp.vert + '"';
    if (bp.rot) xml += ' rot="' + bp.rot + '"';
    xml += '>';
    if (bp.autoFit === 'spAutoFit') xml += '<a:spAutoFit/>';
    else if (bp.autoFit === 'shrink') xml += '<a:normAutofit/>';
    xml += '</a:bodyPr>';
    xml += '<a:lstStyle/>';

    for (let pi = 0; pi < tb.paragraphs.length; pi++) {
      xml += this._paragraphXml(tb.paragraphs[pi]);
    }
    if (tb.paragraphs.length === 0) {
      xml += '<a:p><a:endParaRPr lang="en-US"/></a:p>';
    }

    xml += '</p:txBody>';
    return xml;
  }

  // ---- Paragraph ----
  _paragraphXml(para: any): string {
    const pp = para.pPr;
    let xml = '<a:p>';

    xml += '<a:pPr';
    if (pp.algn && pp.algn !== 'l') xml += ' algn="' + pp.algn + '"';
    if (pp.lvl) xml += ' lvl="' + pp.lvl + '"';
    if (pp.indent) xml += ' indent="' + pp.indent + '"';
    if (pp.marL) xml += ' marL="' + pp.marL + '"';
    xml += '>';
    if (pp.spcBefPct != null) {
      xml += '<a:spcBef><a:spcPct val="' + pp.spcBefPct + '"/></a:spcBef>';
    } else if (pp.spcBef) {
      xml += '<a:spcBef><a:spcPts val="' + Math.round(pp.spcBef * 100 / 12700) + '"/></a:spcBef>';
    }
    if (pp.spcAftPct != null) {
      xml += '<a:spcAft><a:spcPct val="' + pp.spcAftPct + '"/></a:spcAft>';
    } else if (pp.spcAft) {
      xml += '<a:spcAft><a:spcPts val="' + Math.round(pp.spcAft * 100 / 12700) + '"/></a:spcAft>';
    }
    if (pp.lnSpcType === 'pts' && pp.lnSpc) {
      xml += '<a:lnSpc><a:spcPts val="' + pp.lnSpc + '"/></a:lnSpc>';
    } else if (pp.lnSpc && pp.lnSpc !== 100) {
      xml += '<a:lnSpc><a:spcPct val="' + (pp.lnSpc * 1000) + '"/></a:lnSpc>';
    }
    if (pp.buClr) {
      const buClrVal = colorVal(pp.buClr);
      const buClrTr = colorTransforms(pp.buClr);
      if (typeof buClrVal === 'string' && buClrVal.charAt(0) === '#') {
        xml += '<a:buClr><a:srgbClr val="' + buClrVal.replace('#','') + '"' + (buClrTr.length ? '>' + colorTransformsXml(buClrTr) + '</a:srgbClr>' : '/>') + '</a:buClr>';
      } else if (buClrVal) {
        xml += '<a:buClr><a:schemeClr val="' + buClrVal + '"' + (buClrTr.length ? '>' + colorTransformsXml(buClrTr) + '</a:schemeClr>' : '/>') + '</a:buClr>';
      }
    }
    if (pp.buSzPct) {
      xml += '<a:buSzPct val="' + pp.buSzPct + '"/>';
    }
    if (pp.buFont) {
      xml += '<a:buFont typeface="' + esc(pp.buFont) + '"/>';
    }
    if (pp.buNone) {
      xml += '<a:buNone/>';
    } else if (pp.buChar) {
      xml += '<a:buChar char="' + esc(pp.buChar) + '"/>';
    } else if (pp.buAutoNum) {
      xml += '<a:buAutoNum type="' + esc(pp.buAutoNum) + '"/>';
    }
    xml += '</a:pPr>';

    for (let ri = 0; ri < para.runs.length; ri++) {
      const run = para.runs[ri];
      if (run.text === '\n') {
        xml += '<a:br><a:rPr lang="en-US"/></a:br>';
        continue;
      }
      xml += this._runXml(run);
    }

    // endParaRPr determines empty paragraph height — use the last run's font size
    // so blank lines match the surrounding text size, not the theme default.
    const lastRun = para.runs.length > 0 ? para.runs[para.runs.length - 1] : null;
    const endSz = lastRun && lastRun.rPr && lastRun.rPr.sz ? lastRun.rPr.sz : null;
    if (endSz) {
      xml += '<a:endParaRPr lang="en-US" sz="' + endSz + '"/>';
    } else {
      xml += '<a:endParaRPr lang="en-US"/>';
    }
    xml += '</a:p>';
    return xml;
  }

  // ---- Run ----
  _runXml(run: any): string {
    const rp = run.rPr;
    let xml = '<a:r>';
    xml += '<a:rPr lang="en-US"';
    if (rp.sz) xml += ' sz="' + rp.sz + '"';
    if (rp.b) xml += ' b="1"';
    if (rp.i) xml += ' i="1"';
    if (rp.u && rp.u !== 'none') xml += ' u="' + rp.u + '"';
    if (rp.strike && rp.strike !== 'noStrike') xml += ' strike="' + rp.strike + '"';
    if (rp.cap && rp.cap !== 'none') xml += ' cap="' + rp.cap + '"';
    if (rp.baseline) xml += ' baseline="' + rp.baseline + '"';
    xml += ' dirty="0">';
    if (rp.color) {
      const rcv = colorVal(rp.color), rct = colorTransforms(rp.color);
      xml += '<a:solidFill>';
      if (isSchemeColor(rp.color)) {
        if (rct.length > 0) {
          xml += '<a:schemeClr val="' + rcv + '">' + colorTransformsXml(rct) + '</a:schemeClr>';
        } else {
          xml += '<a:schemeClr val="' + rcv + '"/>';
        }
      } else {
        xml += '<a:srgbClr val="' + colorToHex(rp.color) + '"/>';
      }
      xml += '</a:solidFill>';
    }
    if (rp.highlight) {
      xml += '<a:highlight>';
      xml += '<a:srgbClr val="' + colorToHex(rp.highlight) + '"/>';
      xml += '</a:highlight>';
    }
    if (rp.fontFamily) {
      xml += '<a:latin typeface="' + esc(rp.fontFamily) + '"/>';
      xml += '<a:cs typeface="' + esc(rp.fontFamily) + '"/>';
    }
    xml += '</a:rPr>';
    xml += '<a:t>' + esc(run.text) + '</a:t>';
    xml += '</a:r>';
    return xml;
  }

  // ---- docProps ----
  _appXml(pres: any): string {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">';
    xml += '<Application>BankOps</Application>';
    xml += '<Slides>' + pres.slides.length + '</Slides>';
    xml += '</Properties>';
    return xml;
  }

  _coreXml(): string {
    const now = new Date().toISOString();
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">';
    xml += '<dc:creator>BankOps</dc:creator>';
    xml += '<dcterms:created xsi:type="dcterms:W3CDTF">' + now + '</dcterms:created>';
    xml += '<dcterms:modified xsi:type="dcterms:W3CDTF">' + now + '</dcterms:modified>';
    xml += '</cp:coreProperties>';
    return xml;
  }
}

export default PptxWriter;
