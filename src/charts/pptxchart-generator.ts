/**
 * Generates a throwaway single-slide PPTX via pptxgenjs, then extracts the
 * native chart parts (chart XML, embedded Excel workbook, rels) from the ZIP.
 * These parts are later injected into our custom PPTX writer output.
 */

import JSZip from "jszip";
import { mapToPptxgenjs } from "./pptxchart-mapper.js";
import { MissingDependencyError } from "../raster.js";
import type { Bounds, PptxChartContent } from "../types.js";

export interface ChartParts {
  /** path → Buffer, e.g. "ppt/charts/chart1.xml", "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx" */
  parts: Record<string, Buffer>;
  contentTypeOverrides: { partName: string; contentType: string }[];
  /** <p:graphicFrame> XML with r:id replaced by "__CHART_RID_{chartIdx}__", or null */
  graphicFrameXml: string | null;
  chartIdx: number;
}

/** pptxgenjs is an optional peer dependency — load it lazily. */
async function loadPptxGenJS(): Promise<any> {
  let mod: any;
  try { mod = await import("pptxgenjs"); } catch { throw new MissingDependencyError("pptxgenjs", "native (editable) PPTX charts"); }
  return typeof mod === "function" ? mod : (mod.default || mod);
}

/**
 * Generate native PowerPoint chart parts from a PptxChartContent spec.
 *
 * @param content  - PptxChartContent spec
 * @param bounds   - EMU bounds { x, y, cx, cy }
 * @param chartIdx - Unique chart index for this presentation (1-based)
 * @returns { parts: {path → Buffer}, contentTypeOverrides: [], graphicFrameXml: string, chartIdx }
 */
export async function generateChartParts(content: PptxChartContent, bounds: Bounds, chartIdx: number): Promise<ChartParts> {
  const PptxGenJS = await loadPptxGenJS();

  // Convert EMU bounds to inches for pptxgenjs
  const boundsInches = {
    x: bounds.x / 914400,
    y: bounds.y / 914400,
    w: bounds.cx / 914400,
    h: bounds.cy / 914400,
  };

  // Create a throwaway presentation + slide
  const pptx = new PptxGenJS();
  // Match our slide size (standard 13.333 x 7.5 = widescreen)
  pptx.defineLayout({ name: 'CUSTOM', width: 13.333, height: 7.5 });
  pptx.layout = 'CUSTOM';

  const slide = pptx.addSlide();

  // Map our content to pptxgenjs args
  const mapped = mapToPptxgenjs(content, boundsInches);

  if (mapped.isCombo) {
    slide.addChart(mapped.chartType, mapped.options);
  } else {
    slide.addChart(mapped.chartType, mapped.data, mapped.options);
  }

  // Generate PPTX to buffer
  const buffer: Buffer = await pptx.write({ outputType: 'nodebuffer' });

  // Open the generated PPTX with JSZip and extract chart parts
  const zip = await JSZip.loadAsync(buffer);
  const parts: Record<string, Buffer> = {};
  const contentTypeOverrides: { partName: string; contentType: string }[] = [];

  // Detect the chart number pptxgenjs actually used (it has a global counter)
  const srcChartNum = detectChartNum(zip);

  // Extract chart XML files — pptxgenjs puts them in ppt/charts/
  const chartFiles = Object.keys(zip.files).filter((f) => f.startsWith('ppt/charts/') || f.startsWith('ppt/embeddings/'));

  for (const filePath of chartFiles) {
    if (zip.files[filePath].dir) continue;
    const fileBuffer = await zip.files[filePath].async('nodebuffer');

    // Rename chartN → chart{chartIdx} to normalize to our index
    const renamedPath = renameChartPath(filePath, srcChartNum, chartIdx);
    parts[renamedPath] = fileBuffer;
  }

  // Extract chart rels
  const chartRelsFiles = Object.keys(zip.files).filter((f) => f.startsWith('ppt/charts/_rels/'));
  for (const relsPath of chartRelsFiles) {
    if (zip.files[relsPath].dir) continue;
    const relsBuffer = await zip.files[relsPath].async('nodebuffer');
    // Update internal references from worksheetN to worksheet{chartIdx}
    let relsStr = relsBuffer.toString('utf-8');
    relsStr = relsStr.replace(
      new RegExp('Microsoft_Excel_Worksheet' + srcChartNum, 'g'),
      'Microsoft_Excel_Worksheet' + chartIdx
    );
    const renamedRelsPath = renameChartPath(relsPath, srcChartNum, chartIdx);
    parts[renamedRelsPath] = Buffer.from(relsStr, 'utf-8');
  }

  // Extract the graphicFrame XML from the slide
  const slideXml = await zip.files['ppt/slides/slide1.xml'].async('string');
  const graphicFrameXml = extractGraphicFrame(slideXml, chartIdx);

  // Build content type overrides
  contentTypeOverrides.push({
    partName: '/ppt/charts/chart' + chartIdx + '.xml',
    contentType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
  });

  return {
    parts: parts,
    contentTypeOverrides: contentTypeOverrides,
    graphicFrameXml: graphicFrameXml,
    chartIdx: chartIdx,
  };
}

/**
 * Detect the chart number pptxgenjs used inside the throwaway PPTX.
 * pptxgenjs has a global counter, so it may generate chart2, chart3, etc.
 */
function detectChartNum(zip: JSZip): number {
  const files = Object.keys(zip.files);
  for (const f of files) {
    const m = f.match(/ppt\/charts\/chart(\d+)\.xml$/);
    if (m) return parseInt(m[1], 10);
  }
  return 1; // fallback
}

/**
 * Rename chart file paths from chart{srcNum} → chart{destNum}
 */
function renameChartPath(filePath: string, srcNum: number, destNum: number): string {
  return filePath
    .replace(new RegExp('chart' + srcNum + '\\.xml', 'g'), 'chart' + destNum + '.xml')
    .replace(new RegExp('style' + srcNum + '\\.xml', 'g'), 'style' + destNum + '.xml')
    .replace(new RegExp('colors' + srcNum + '\\.xml', 'g'), 'colors' + destNum + '.xml')
    .replace(new RegExp('Microsoft_Excel_Worksheet' + srcNum, 'g'), 'Microsoft_Excel_Worksheet' + destNum);
}

/**
 * Extract the <p:graphicFrame> from the generated slide XML and rewrite
 * the rId to a placeholder that our writer will replace.
 */
function extractGraphicFrame(slideXml: string, chartIdx: number): string | null {
  // Find the graphicFrame containing the chart
  const gfMatch = slideXml.match(/<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/);
  if (!gfMatch) {
    // Fallback: build a minimal graphicFrame
    return null;
  }

  let gfXml = gfMatch[0];

  // Replace whatever rId was used (e.g. rId2) with a placeholder
  // Our writer will assign the actual rId when building slide rels
  gfXml = gfXml.replace(/r:id="rId\d+"/g, 'r:id="__CHART_RID_' + chartIdx + '__"');

  return gfXml;
}

export default { generateChartParts };
