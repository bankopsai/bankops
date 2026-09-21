/**
 * Maps PptxChartContent → pptxgenjs addChart() arguments.
 * Returns { chartType, data, options } ready for slide.addChart().
 *
 * pptxgenjs' `ChartType` enum values are plain strings ("bar", "line", "pie",
 * "scatter", "doughnut", "area"), so this module needs no pptxgenjs import
 * and stays free of the optional dependency.
 */

import type { PptxChartContent } from "../types.js";

export interface BoundsInches { x: number; y: number; w: number; h: number }

export interface MappedChart {
  /** pptxgenjs chart type string, or (for combo) an array of { type, data, options } */
  chartType: any;
  data: any[] | null;
  options: any;
  isCombo: boolean;
}

// Values of pptxgenjs `ChartType` (string enum)
const chartTypeMap: Record<string, string> = {
  bar: 'bar',
  line: 'line',
  pie: 'pie',
  scatter: 'scatter',
  doughnut: 'doughnut',
  area: 'area',
};

export function mapToPptxgenjs(content: PptxChartContent, boundsInches: BoundsInches): MappedChart {
  // chart-to-native adds a few pass-through fields (fontFace, ...) that are
  // not part of the public PptxChartContent type.
  const c: any = content;

  // Build common options
  const opts: any = {
    x: boundsInches.x,
    y: boundsInches.y,
    w: boundsInches.w,
    h: boundsInches.h,
  };

  if (c.title && c.showTitle !== false) { opts.showTitle = true; opts.title = c.title; }
  if (c.showLegend === false) opts.showLegend = false;
  else { opts.showLegend = true; opts.legendPos = c.legendPos || 'b'; }
  if (c.chartColors) opts.chartColors = c.chartColors.map((col: string) => col.replace('#', ''));
  if (c.showValue) opts.showValue = true;
  else opts.showValue = false;
  if (c.showPercent) opts.showPercent = true;
  else opts.showPercent = false;
  if (c.showCatName) opts.showCatName = true;
  if (c.showSerName) opts.showSerName = true;
  if (c.dataLabelFormatCode) opts.dataLabelFormatCode = c.dataLabelFormatCode;
  if (c.catAxisTitle) opts.catAxisTitle = c.catAxisTitle;
  if (c.valAxisTitle) opts.valAxisTitle = c.valAxisTitle;
  if (c.valAxisMinVal != null) opts.valAxisMinVal = c.valAxisMinVal;
  if (c.valAxisMaxVal != null) opts.valAxisMaxVal = c.valAxisMaxVal;
  if (c.valAxisLabelFormatCode) opts.valAxisLabelFormatCode = c.valAxisLabelFormatCode;
  if (c.catAxisHidden) opts.catAxisHidden = true;
  if (c.valAxisHidden) opts.valAxisHidden = true;
  if (c.dataBorder) opts.dataBorder = c.dataBorder;
  if (c.valGridLine) opts.valGridLine = c.valGridLine;
  if (c.catGridLine) opts.catGridLine = c.catGridLine;

  // Handle combo chart
  if (c.chartType === 'combo' && c.comboCharts && c.comboCharts.length > 0) {
    let hasSecondary = false;
    const chartTypes = c.comboCharts.map((cc: any) => {
      const type = cc.chartType === 'area' ? chartTypeMap.area : (cc.chartType === 'line' ? chartTypeMap.line : chartTypeMap.bar);
      const data = cc.series.map((s: any) => ({ name: s.name, labels: s.labels, values: s.values }));
      const subOpts: any = {};
      if (cc.options && cc.options.secondaryValAxis) { subOpts.secondaryValAxis = true; hasSecondary = true; }
      if (cc.options && cc.options.barGrouping) subOpts.barGrouping = cc.options.barGrouping;
      return { type: type, data: data, options: subOpts };
    });
    // Build primary/secondary value-axis configs so each axis can have its own
    // tick formatting (the editor formats them independently via y2* fields).
    if (hasSecondary) {
      const primaryAx: any = {};
      if (c.valAxisTitle) { primaryAx.valAxisTitle = c.valAxisTitle; primaryAx.showValAxisTitle = true; }
      if (c.valAxisMinVal != null) primaryAx.valAxisMinVal = c.valAxisMinVal;
      if (c.valAxisMaxVal != null) primaryAx.valAxisMaxVal = c.valAxisMaxVal;
      if (c.valAxisLabelFormatCode) primaryAx.valAxisLabelFormatCode = c.valAxisLabelFormatCode;
      const secondaryAx: any = { secondaryValAxis: true };
      if (c.secondaryValAxisLabelFormatCode) secondaryAx.valAxisLabelFormatCode = c.secondaryValAxisLabelFormatCode;
      if (c.secondaryValAxisMinVal != null) secondaryAx.valAxisMinVal = c.secondaryValAxisMinVal;
      if (c.secondaryValAxisMaxVal != null) secondaryAx.valAxisMaxVal = c.secondaryValAxisMaxVal;
      opts.valAxes = [primaryAx, secondaryAx];
    }
    return { chartType: chartTypes, data: null, options: opts, isCombo: true };
  }

  // Resolve chart type
  let resolvedType: string;
  if (c.chartType === 'pie' && c.doughnut) {
    resolvedType = chartTypeMap.doughnut;
  } else {
    resolvedType = chartTypeMap[c.chartType] || chartTypeMap.bar;
  }

  // Bar-specific options
  if (c.chartType === 'bar') {
    if (c.barDir) opts.barDir = c.barDir;
    if (c.barGrouping) opts.barGrouping = c.barGrouping;
    if (c.gapWidthPct != null) opts.barGapWidthPct = c.gapWidthPct;
    if (c.barFillColor) opts.chartColors = [c.barFillColor.replace('#', '')];
  }

  // Line-specific
  if (c.chartType === 'line' || c.chartType === 'scatter') {
    if (c.lineSmooth) opts.lineSmooth = true;
    if (c.lineDataSymbol) opts.lineDataSymbol = c.lineDataSymbol;
    if (c.lineDataSymbolSize != null) opts.lineDataSymbolSize = c.lineDataSymbolSize;
    if (c.lineDataSymbolLineColor) opts.lineDataSymbolLineColor = c.lineDataSymbolLineColor.replace('#', '');
    if (c.lineDataSymbolLineSize != null) opts.lineDataSymbolLineSize = c.lineDataSymbolLineSize;
  }

  // Font pass-through
  if (c.fontFace) {
    opts.catAxisLabelFontFace = c.fontFace;
    opts.valAxisLabelFontFace = c.fontFace;
    opts.dataLabelFontFace = c.fontFace;
  }
  if (c.fontSize) {
    opts.catAxisLabelFontSize = c.fontSize;
    opts.valAxisLabelFontSize = c.fontSize;
    opts.dataLabelFontSize = c.fontSize;
  }
  if (c.legendFontSize != null) opts.legendFontSize = c.legendFontSize;
  if (c.titleFontSize != null) opts.titleFontSize = c.titleFontSize;

  // Data-label-specific style (set by chart-to-native to mirror the editor's
  // label position + font size; falls back to the chart-wide values above).
  if (c.dataLabelFontSize != null) opts.dataLabelFontSize = c.dataLabelFontSize;
  if (c.dataLabelColor) opts.dataLabelColor = c.dataLabelColor.replace('#', '');
  if (c.dataLabelPosition) opts.dataLabelPosition = c.dataLabelPosition;

  // Map series data (pptxgenjs format is identical)
  const data = (c.series || []).map((s: any) => ({ name: s.name, labels: s.labels, values: s.values }));

  return { chartType: resolvedType, data: data, options: opts, isCombo: false };
}

export default { mapToPptxgenjs };
