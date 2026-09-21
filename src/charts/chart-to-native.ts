/**
 * Converts ChartContent to PptxChartContent format for native Excel export.
 * Supports bar, line, pie, combo, and scatter chart types.
 */

import type { ChartContent, PptxChartContent, StyleTokens } from "../types.js";
import { DEFAULT_COLORS } from "./chart-options.js";

/** style.chart tokens (all optional — callers pass `style.chart || {}`). */
export type ChartStyle = Partial<StyleTokens["chart"]>;

// Fallback palette — the same `DEFAULT_COLORS` the ECharts option builder
// uses, so the native PPTX render uses the same per-series colors as the
// editor when no theme palette is configured.
const EDITOR_DEFAULT_COLORS = DEFAULT_COLORS;

/**
 * Build an Excel number format code from chart formatting options.
 * @param prefix - e.g. "$"
 * @param suffix - e.g. "%"
 * @param decimals - decimal places (0+)
 * @param thousands - use thousands separator
 * @param negFmt - "minus" or "parens"
 * @returns Excel format code or null if default
 */
export function buildExcelFormatCode(prefix?: string, suffix?: string, decimals?: number | null, thousands?: boolean, negFmt?: string): string | null {
  if (!prefix && !suffix && decimals == null && !thousands && negFmt !== 'parens') {
    return null;
  }

  const dec = decimals != null ? decimals : 0;
  const sep = thousands ? '#,##' : '';

  const decPart = dec > 0 ? '.' + new Array(dec + 1).join('0') : '';
  const numPart = sep + '0' + decPart;

  // Quote prefix/suffix as Excel literals. Bare letters (e.g. "B") are
  // technically literals in Excel format syntax, but PowerPoint chart number
  // formats are stricter and frequently drop the unquoted text. Wrapping in
  // double quotes is the universally-portable form.
  function quoteLiteral(s?: string): string {
    if (!s) return '';
    return '"' + String(s).replace(/"/g, '""') + '"';
  }

  const positive = quoteLiteral(prefix) + numPart + quoteLiteral(suffix);

  if (negFmt === 'parens') {
    return positive + ';(' + positive + ')';
  }

  return positive;
}

/** Map legendPosition string to pptxgenjs legend pos code */
function mapLegendPos(pos: string): string {
  const posMap: Record<string, string> = { bottom: 'b', top: 't', left: 'l', right: 'r' };
  return posMap[pos] || 'b';
}

/** Apply common properties: legend, title, axis labels/bounds, fonts */
function applyCommon(spec: any, content: ChartContent, chartStyle: ChartStyle): void {
  // Legend
  spec.showLegend = content.showLegend !== false;
  if (content.legendPosition) {
    spec.legendPos = mapLegendPos(content.legendPosition);
  }

  // Title
  if (content.title) {
    spec.title = content.title;
    spec.showTitle = true;
  }

  // Axis bounds
  if (content.yMin != null) spec.valAxisMinVal = content.yMin;
  if (content.yMax != null) spec.valAxisMaxVal = content.yMax;

  // Axis labels
  if (content.xAxisLabel) spec.catAxisTitle = content.xAxisLabel;
  if (content.yAxisLabel) spec.valAxisTitle = content.yAxisLabel;

  // Colors — prefer theme palette, then fall back to the editor's default
  // palette so per-series colors match the on-screen render.
  if (chartStyle.colors && chartStyle.colors.length > 0) {
    spec.chartColors = chartStyle.colors.slice();
  } else {
    spec.chartColors = EDITOR_DEFAULT_COLORS.slice();
  }

  // Font pass-through (consumed by pptxchart-mapper). The chart-wide font is
  // applied to category-axis, value-axis, and data labels. The editor uses 11
  // for axes/legend, 10 for on-bar/on-line value labels (applyDataLabelStyle
  // overrides dataLabelFontSize), and 14 for the chart title.
  if (chartStyle.fontFamily) spec.fontFace = chartStyle.fontFamily;
  spec.fontSize = chartStyle.fontSize || 11;
  spec.legendFontSize = chartStyle.fontSize || 11;
  if (content.title) spec.titleFontSize = 14;

  // Data label + value-axis tick format code — the editor formats both, so
  // mirror that here so the embedded chart's Y-axis ticks show e.g. "$14B"
  // instead of raw "14".
  const fmtCode = buildExcelFormatCode(
    content.valuePrefix,
    content.valueSuffix,
    content.valueDecimals,
    content.valueThousands,
    content.negativeFormat
  );
  if (fmtCode) {
    spec.dataLabelFormatCode = fmtCode;
    spec.valAxisLabelFormatCode = fmtCode;
  }
}

/**
 * Map showXAxis/showYAxis from editor semantics onto pptxgenjs catAxisHidden/valAxisHidden.
 * For bar charts in `horizontal` mode the axes are swapped (cat is the Y-axis, val is X).
 */
function applyAxisVisibility(spec: any, content: ChartContent): void {
  if (content.showXAxis !== false && content.showYAxis !== false) return;
  const horizontal = content.horizontal === true;
  if (content.showXAxis === false) {
    if (horizontal) spec.valAxisHidden = true; else spec.catAxisHidden = true;
  }
  if (content.showYAxis === false) {
    if (horizontal) spec.catAxisHidden = true; else spec.valAxisHidden = true;
  }
}

/** Build a dataBorder spec from an editor border color/width, or null when unset. */
function buildDataBorder(color?: string, width?: number | null): { type: string; pt: number; color: string } | null {
  if (!color && width == null) return null;
  return {
    type: 'solid',
    pt: width != null ? width : 1,
    color: (color || '#000000').replace('#', ''),
  };
}

/** Map editor pointShape ("circle"|"rect"|"triangle"|"diamond") to pptxgenjs lineDataSymbol. */
function mapPointShape(shape?: string): string | null {
  switch (shape) {
    case 'rect': return 'square';
    case 'triangle': return 'triangle';
    case 'diamond': return 'diamond';
    case 'circle': return 'circle';
    default: return null;
  }
}

// Editor uses a fixed 10pt size for data labels — match it so the PPTX render
// looks like the on-screen render. pptxgenjs default is ~18pt.
const DATA_LABEL_FONT_SIZE = 10;

/**
 * Set data-label position + font size to match the editor's behavior:
 *   - Vertical bar: above the bar ("outEnd")
 *   - Horizontal bar: to the right of the bar ("outEnd")
 *   - Stacked bar with series labels: centered ("ctr")
 *   - Line: above the point ("t")
 *   - Pie: from labelPosition / pieValuePosition (outside / inside / center)
 */
function applyDataLabelStyle(spec: any, content: ChartContent): void {
  if (content.chartType === 'bar') {
    spec.dataLabelFontSize = DATA_LABEL_FONT_SIZE;
    const stackedWithSeriesLabels = content.stacked && content.showSeriesLabels;
    spec.dataLabelPosition = stackedWithSeriesLabels ? 'ctr' : 'outEnd';
  } else if (content.chartType === 'line') {
    spec.dataLabelFontSize = DATA_LABEL_FONT_SIZE;
    spec.dataLabelPosition = 't';
  } else if (content.chartType === 'pie') {
    // Editor uses fs(config) for pie labels — default 11pt — not the
    // bar/line 10pt.
    spec.dataLabelFontSize = spec.fontSize || 11;
    const pos: string | undefined = content.pieValuePosition || content.labelPosition;
    if (pos === 'inside') spec.dataLabelPosition = 'inEnd';
    else if (pos === 'center') spec.dataLabelPosition = 'ctr';
    else spec.dataLabelPosition = 'outEnd';
  }
}

/**
 * Match the editor's value-axis split-line behavior. ECharts draws very light
 * (#E8E8E8) split-lines by default; PowerPoint's default major gridlines are
 * much darker. Mirror the editor so the PPTX render doesn't suddenly grow
 * heavy gridlines. Hide entirely when showSplitLine is explicitly false.
 */
function applyGridlines(spec: any, content: ChartContent): void {
  if (content.showSplitLine === false) {
    spec.valGridLine = { style: 'none' };
  } else {
    spec.valGridLine = { style: 'solid', color: 'E8E8E8', size: 0.5 };
  }
}

/** Apply line-marker styling (pointShape/Size/BorderColor/BorderWidth) to the spec. */
function applyLineMarkers(spec: any, content: ChartContent): void {
  const sym = mapPointShape(content.pointShape);
  if (sym) spec.lineDataSymbol = sym;
  if (content.pointSize != null) spec.lineDataSymbolSize = content.pointSize;
  if (content.pointBorderColor) spec.lineDataSymbolLineColor = content.pointBorderColor;
  if (content.pointBorderWidth != null) spec.lineDataSymbolLineSize = content.pointBorderWidth;
}

/** Build standard series from categories + series arrays */
function buildCategorySeries(content: ChartContent): { name: string; labels: string[]; values: number[] }[] {
  const categories = content.categories || [];
  return (content.series || []).map((s) => ({
    name: s.name || 'Series',
    labels: categories.slice(),
    values: (s.data || []).map((v) => (Array.isArray(v) ? v[1] : v)),
  }));
}

// ─── Bar ───

export function convertBarToNative(content: ChartContent, chartStyle: ChartStyle): PptxChartContent | null {
  if (!content || content.chartType !== 'bar') return null;

  const series = buildCategorySeries(content);
  if (series.length === 0) return null;

  const spec: any = {
    type: 'pptxChart',
    chartType: 'bar',
    series: series,
    barDir: content.horizontal ? 'bar' : 'col',
    barGrouping: content.stacked ? 'stacked' : 'clustered',
    showValue: content.showBarValues !== false,
  };

  applyCommon(spec, content, chartStyle);
  applyAxisVisibility(spec, content);
  applyGridlines(spec, content);
  applyDataLabelStyle(spec, content);

  // barColor overrides theme colors
  if (content.barColor) {
    spec.chartColors = [content.barColor];
  }

  // categoryColors → per-data-point color array (single-series only — pptxgenjs
  // emits per-bar <c:dPt> fills when data.length === 1 and chartColors.length > 1).
  if (content.categoryColors && series.length === 1) {
    const cats = content.categories || [];
    const perBar: string[] = [];
    let anyOverride = false;
    for (let ci = 0; ci < cats.length; ci++) {
      const col = content.categoryColors[ci];
      if (col) { perBar.push(col); anyOverride = true; }
      else perBar.push(content.barColor || (spec.chartColors && spec.chartColors[0]) || '4472C4');
    }
    if (anyOverride) spec.chartColors = perBar;
  }

  // Bar borders
  const barBorder = buildDataBorder(content.barBorderColor, content.barBorderWidth);
  if (barBorder) spec.dataBorder = barBorder;

  return spec as PptxChartContent;
}

// ─── Line ───

export function convertLineToNative(content: ChartContent, chartStyle: ChartStyle): PptxChartContent | null {
  if (!content || content.chartType !== 'line') return null;

  const series = buildCategorySeries(content);
  if (series.length === 0) return null;

  const spec: any = {
    type: 'pptxChart',
    chartType: 'line',
    series: series,
    showValue: !!content.showLineValues,
    lineSmooth: !!content.smooth,
  };

  applyCommon(spec, content, chartStyle);
  applyAxisVisibility(spec, content);
  applyGridlines(spec, content);
  applyLineMarkers(spec, content);
  applyDataLabelStyle(spec, content);

  // lineColor overrides theme colors
  if (content.lineColor) {
    spec.chartColors = [content.lineColor];
  }

  return spec as PptxChartContent;
}

// ─── Pie ───

export function convertPieToNative(content: ChartContent, chartStyle: ChartStyle): PptxChartContent | null {
  if (!content || content.chartType !== 'pie') return null;

  const items = content.items || [];
  if (items.length === 0) return null;

  // Pie uses items [{name, value}] → single series with labels/values
  const labels = items.map((it) => it.name || '');
  const values = items.map((it) => it.value || 0);

  // Editor pie label model: `showLabel` controls slice-name visibility (default
  // true, unless explicitly false or `labelPosition === "none"`); `showPieValues`
  // controls whether the formatted value is shown. There is no percent toggle.
  const showName = content.showLabel !== false && content.labelPosition !== 'none';
  const showVal = !!content.showPieValues;

  const spec: any = {
    type: 'pptxChart',
    chartType: 'pie',
    series: [{ name: 'Data', labels: labels, values: values }],
    doughnut: !!content.doughnut,
    showValue: showVal,
    showPercent: false,
    showCatName: showName,
  };

  applyCommon(spec, content, chartStyle);
  applyDataLabelStyle(spec, content);

  // sliceColors (index → hex) → chartColors array, applied AFTER applyCommon so
  // it overrides the theme/default palette. Missing slices fall back to the
  // palette color at that index so every slice has a valid fill.
  if (content.sliceColors) {
    const palette: string[] = (spec.chartColors && spec.chartColors.length > 0) ? spec.chartColors : EDITOR_DEFAULT_COLORS;
    const colors: string[] = [];
    for (let i = 0; i < items.length; i++) {
      const sc = content.sliceColors[i];
      colors.push(sc || palette[i % palette.length]);
    }
    spec.chartColors = colors;
  }

  // Pie slice borders
  const pieBorder = buildDataBorder(content.pieBorderColor, content.pieBorderWidth);
  if (pieBorder) spec.dataBorder = pieBorder;

  return spec as PptxChartContent;
}

// ─── Combo ───

export function convertComboToNative(content: ChartContent, chartStyle: ChartStyle): PptxChartContent | null {
  if (!content || content.chartType !== 'combo') return null;

  const categories = content.categories || [];
  const bars = content.bars || [];
  const lines = content.lines || [];

  if (bars.length === 0 && lines.length === 0) return null;

  const comboCharts: any[] = [];

  // Bar sub-charts
  if (bars.length > 0) {
    const barSeries = bars.map((s) => ({
      name: s.name || 'Bar',
      labels: categories.slice(),
      values: (s.data || []).slice(),
    }));
    const barOpts: any = {};
    if (content.stacked) barOpts.barGrouping = 'stacked';
    comboCharts.push({
      chartType: 'bar',
      series: barSeries,
      options: barOpts,
    });
  }

  // Line sub-charts
  if (lines.length > 0) {
    const lineSeries = lines.map((s) => ({
      name: s.name || 'Line',
      labels: categories.slice(),
      values: (s.data || []).slice(),
    }));
    const lineOpts: any = {};
    if (content.dualAxis) lineOpts.secondaryValAxis = true;
    comboCharts.push({
      chartType: 'line',
      series: lineSeries,
      options: lineOpts,
    });
  }

  const spec: any = {
    type: 'pptxChart',
    chartType: 'combo',
    series: [], // combo uses comboCharts instead
    comboCharts: comboCharts,
    showValue: !!content.showBarValues || !!content.showLineValues,
  };

  applyCommon(spec, content, chartStyle);
  applyAxisVisibility(spec, content);
  applyGridlines(spec, content);
  // Match the editor's data-label font size for combo bars/lines.
  spec.dataLabelFontSize = DATA_LABEL_FONT_SIZE;
  spec.dataLabelPosition = 'outEnd';

  // Secondary value axis format code (for dual-axis combo). The editor uses
  // y2* fields with a fall-through to the primary value-format props.
  if (content.dualAxis) {
    const secFmt = buildExcelFormatCode(
      content.y2ValuePrefix != null ? content.y2ValuePrefix : content.valuePrefix,
      content.y2ValueSuffix != null ? content.y2ValueSuffix : content.valueSuffix,
      content.y2ValueDecimals != null ? content.y2ValueDecimals : content.valueDecimals,
      content.y2ValueThousands != null ? content.y2ValueThousands : content.valueThousands,
      content.y2NegativeFormat != null ? content.y2NegativeFormat : content.negativeFormat
    );
    if (secFmt) spec.secondaryValAxisLabelFormatCode = secFmt;
    if (content.y2Min != null) spec.secondaryValAxisMinVal = content.y2Min;
    if (content.y2Max != null) spec.secondaryValAxisMaxVal = content.y2Max;
  }

  return spec as PptxChartContent;
}

// ─── Scatter ───

export function convertScatterToNative(content: ChartContent, chartStyle: ChartStyle): PptxChartContent | null {
  if (!content || content.chartType !== 'scatter') return null;

  const rawSeries: { name?: string; data: any[] }[] = content.series || [];
  if (rawSeries.length === 0) return null;

  // Scatter data is series[{name, data: [[x,y], ...]}]
  // pptxgenjs scatter expects: series[{name, values: [{x, y}]}]
  // But our PptxChartContent uses labels/values format.
  // For scatter, we'll use the standard series format and the mapper handles it.
  const series = rawSeries.map((s) => {
    const pts: any[] = s.data || [];
    // Extract X values as labels and Y values as values
    const labels = pts.map((pt) => (Array.isArray(pt) ? String(pt[0]) : String(pt)));
    const values = pts.map((pt) => (Array.isArray(pt) ? pt[1] : pt));
    return {
      name: s.name || 'Series',
      labels: labels,
      values: values,
    };
  });

  const spec: any = {
    type: 'pptxChart',
    chartType: 'scatter',
    series: series,
    showValue: false,
  };

  applyCommon(spec, content, chartStyle);
  applyAxisVisibility(spec, content);
  applyGridlines(spec, content);
  applyLineMarkers(spec, content);

  return spec as PptxChartContent;
}

// ─── Dispatcher ───

/**
 * Convert any supported ChartContent to PptxChartContent.
 * @param content - ChartContent
 * @param chartStyle - style.chart tokens
 * @returns PptxChartContent spec, or null if unsupported/not convertible
 */
export function convertToNative(content: ChartContent, chartStyle: ChartStyle = {}): PptxChartContent | null {
  if (!content) return null;
  switch (content.chartType) {
    case 'bar': return convertBarToNative(content, chartStyle);
    case 'line': return convertLineToNative(content, chartStyle);
    case 'pie': return convertPieToNative(content, chartStyle);
    case 'combo': return convertComboToNative(content, chartStyle);
    case 'scatter': return convertScatterToNative(content, chartStyle);
    default: return null;
  }
}

export default {
  convertToNative,
  convertBarToNative,
  convertLineToNative,
  convertPieToNative,
  convertComboToNative,
  convertScatterToNative,
  buildExcelFormatCode,
};
