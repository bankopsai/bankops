/* Converts PptxChartContent → ECharts option for browser preview.
   All defaults match pptxgenjs PPTX output so the web preview looks
   like what PowerPoint renders.  PPTX is the source of truth. */

/* eslint-disable @typescript-eslint/no-explicit-any */

import type { PptxChartContent } from "../types.js";

// pptxgenjs default series palette (rgb values from OOXML)
const PPTX_DEFAULT_COLORS = [
  "#C0504D", // rgb(192,80,77)
  "#4F81BD", // rgb(79,129,189)
  "#9BBB59", // rgb(155,187,89)
  "#8064A2", // rgb(128,100,162)
  "#4BACC6", // rgb(75,172,198)
  "#F79646", // rgb(247,150,70)
];

// pptxgenjs axis / gridline color
const PPTX_AXIS_COLOR = "#888888";

interface PptxChartConfig {
  colors?: string[];
  fontFamily?: string;
  fontSize?: number;
}

function ff(config: PptxChartConfig): string { return config.fontFamily || "Arial"; }
function fs(config: PptxChartConfig): number { return config.fontSize || 12; }

function applyDefaults(option: any, content: PptxChartContent, config: PptxChartConfig): any {
  option.backgroundColor = "#FFFFFF";
  option.color = content.chartColors || config.colors || PPTX_DEFAULT_COLORS;
  if (!option.textStyle) option.textStyle = {};
  option.textStyle.fontFamily = ff(config);
  option.textStyle.fontSize = fs(config);
  option.animation = false;
  return option;
}

function buildLegend(content: PptxChartContent, config: PptxChartConfig): any {
  if (content.showLegend === false) return undefined;
  const textStyle = { fontSize: fs(config), fontFamily: ff(config) };
  switch (content.legendPos) {
    case "t": return { top: 5, textStyle };
    case "l": return { left: 5, orient: "vertical", textStyle };
    case "r": return { right: 5, orient: "vertical", textStyle };
    case "tr": return { right: 5, top: 5, textStyle };
    default: return { bottom: 5, textStyle };
  }
}

/** PPTX gap width is a percentage of bar width.  150% gap means the space
 *  between groups equals 1.5× the bar width.  ECharts `barCategoryGap`
 *  expresses the same concept as a percentage of the category slot. The
 *  mapping:  barCategoryGap ≈ gapWidth / (gapWidth + 100) × 100%.
 *  150 → 60%, 100 → 50%, 50 → 33%.  */
function gapWidthToCategoryGap(gapWidth: number): string {
  return Math.round(gapWidth / (gapWidth + 100) * 100) + "%";
}

function buildBarOption(content: PptxChartContent, config: PptxChartConfig): any {
  const isHorizontal = content.barDir === "bar";
  const categories = content.series[0]?.labels || [];
  const gapWidth = content.gapWidthPct ?? 150;

  const seriesArr = content.series.map((s) => {
    const item: any = {
      type: "bar", name: s.name, data: s.values,
      barGap: "0%",                             // overlap 0% — bars touch within group
      barCategoryGap: gapWidthToCategoryGap(gapWidth),
    };
    if (content.barFillColor) {
      item.itemStyle = { color: content.barFillColor };
    }
    if (content.barGrouping === "stacked" || content.barGrouping === "percentStacked") {
      item.stack = "total";
    }
    if (content.showValue) {
      item.label = { show: true, position: isHorizontal ? "right" : "top", fontSize: fs(config) };
    }
    return item;
  });

  const axisLineStyle = { lineStyle: { color: PPTX_AXIS_COLOR } };
  const catAxis: any = {
    type: "category", data: categories,
    axisLabel: { fontFamily: ff(config), fontSize: fs(config), color: "#000" },
    axisLine: axisLineStyle,
    axisTick: { alignWithLabel: true, lineStyle: { color: PPTX_AXIS_COLOR } },
  };
  if (content.catAxisTitle) { catAxis.name = content.catAxisTitle; catAxis.nameGap = 25; }

  const valAxis: any = {
    type: "value",
    axisLabel: { fontFamily: ff(config), fontSize: fs(config), color: "#000" },
    axisLine: axisLineStyle,
    axisTick: { lineStyle: { color: PPTX_AXIS_COLOR } },
    splitLine: { lineStyle: { color: PPTX_AXIS_COLOR, width: 1 } },
  };
  if (content.valAxisTitle) { valAxis.name = content.valAxisTitle; valAxis.nameGap = 40; }
  if (content.valAxisMinVal != null) valAxis.min = content.valAxisMinVal;
  if (content.valAxisMaxVal != null) valAxis.max = content.valAxisMaxVal;

  // Tight grid matching PPTX minimal margins
  const hasLegend = content.showLegend !== false;
  const grid = {
    left: isHorizontal ? 8 : 45,
    right: 12,
    top: content.title && content.showTitle !== false ? 35 : 8,
    bottom: hasLegend ? 30 : 8,
    containLabel: true,
  };

  return applyDefaults({
    title: content.showTitle !== false && content.title ? { text: content.title, left: "center", top: 4, textStyle: { fontSize: 14, fontWeight: "bold" } } : undefined,
    tooltip: { trigger: "axis" },
    legend: hasLegend ? buildLegend(content, config) : undefined,
    grid,
    xAxis: isHorizontal ? valAxis : catAxis,
    yAxis: isHorizontal ? catAxis : valAxis,
    series: seriesArr,
  }, content, config);
}

function buildLineOption(content: PptxChartContent, config: PptxChartConfig): any {
  const categories = content.series[0]?.labels || [];
  const axisLineStyle = { lineStyle: { color: PPTX_AXIS_COLOR } };

  const seriesArr = content.series.map((s) => ({
    type: "line", name: s.name, data: s.values,
    smooth: content.lineSmooth === true,
    symbol: "circle", symbolSize: 6,
    ...(content.showValue ? { label: { show: true, position: "top", fontSize: fs(config) } } : {}),
  }));

  const xAxis: any = {
    type: "category", data: categories,
    axisLabel: { fontFamily: ff(config), fontSize: fs(config), color: "#000" },
    axisLine: axisLineStyle,
    boundaryGap: false,
  };
  if (content.catAxisTitle) { xAxis.name = content.catAxisTitle; xAxis.nameGap = 25; }

  const yAxis: any = {
    type: "value",
    axisLabel: { fontFamily: ff(config), fontSize: fs(config), color: "#000" },
    axisLine: axisLineStyle,
    splitLine: { lineStyle: { color: PPTX_AXIS_COLOR, width: 1 } },
  };
  if (content.valAxisTitle) { yAxis.name = content.valAxisTitle; yAxis.nameGap = 40; }
  if (content.valAxisMinVal != null) yAxis.min = content.valAxisMinVal;
  if (content.valAxisMaxVal != null) yAxis.max = content.valAxisMaxVal;

  const hasLegend = content.showLegend !== false;
  return applyDefaults({
    title: content.showTitle !== false && content.title ? { text: content.title, left: "center", top: 4, textStyle: { fontSize: 14, fontWeight: "bold" } } : undefined,
    tooltip: { trigger: "axis" },
    legend: hasLegend ? buildLegend(content, config) : undefined,
    grid: { left: 45, right: 12, top: content.title && content.showTitle !== false ? 35 : 8, bottom: hasLegend ? 30 : 8, containLabel: true },
    xAxis,
    yAxis,
    series: seriesArr,
  }, content, config);
}

function buildPieOption(content: PptxChartContent, config: PptxChartConfig): any {
  const s = content.series[0];
  const items = s ? s.labels.map((label, i) => ({ name: label, value: s.values[i] || 0 })) : [];
  const inner = content.doughnut ? "40%" : "0%";
  const hasLegend = content.showLegend !== false;

  return applyDefaults({
    title: content.showTitle !== false && content.title ? { text: content.title, left: "center", top: 4, textStyle: { fontSize: 14, fontWeight: "bold" } } : undefined,
    tooltip: { trigger: "item", formatter: "{b}: {c} ({d}%)" },
    legend: hasLegend ? buildLegend(content, config) : undefined,
    series: [{
      type: "pie", radius: [inner, "70%"],
      center: ["50%", hasLegend ? "48%" : "50%"],
      data: items,
      label: {
        fontSize: fs(config), fontFamily: ff(config), position: "outside",
        ...(content.showPercent ? { formatter: "{b}: {d}%" } : {}),
      },
      labelLine: { show: true },
      emphasis: { itemStyle: { shadowBlur: 10, shadowOffsetX: 0, shadowColor: "rgba(0,0,0,0.3)" } },
    }],
  }, content, config);
}

function buildScatterOption(content: PptxChartContent, config: PptxChartConfig): any {
  const axisLineStyle = { lineStyle: { color: PPTX_AXIS_COLOR } };
  const seriesArr = content.series.map((s) => ({
    type: "scatter", name: s.name,
    data: s.labels.map((label, i) => [parseFloat(label) || i, s.values[i] || 0]),
    symbolSize: 10,
  }));

  const hasLegend = content.showLegend !== false && seriesArr.length > 1;
  return applyDefaults({
    title: content.showTitle !== false && content.title ? { text: content.title, left: "center", top: 4, textStyle: { fontSize: 14, fontWeight: "bold" } } : undefined,
    tooltip: { trigger: "item" },
    legend: hasLegend ? buildLegend(content, config) : undefined,
    grid: { left: 45, right: 12, top: content.title && content.showTitle !== false ? 35 : 8, bottom: hasLegend ? 30 : 8, containLabel: true },
    xAxis: { type: "value", axisLabel: { fontFamily: ff(config), fontSize: fs(config), color: "#000" }, axisLine: axisLineStyle, splitLine: { lineStyle: { color: PPTX_AXIS_COLOR, width: 1 } } },
    yAxis: { type: "value", axisLabel: { fontFamily: ff(config), fontSize: fs(config), color: "#000" }, axisLine: axisLineStyle, splitLine: { lineStyle: { color: PPTX_AXIS_COLOR, width: 1 } } },
    series: seriesArr,
  }, content, config);
}

function buildComboOption(content: PptxChartContent, config: PptxChartConfig): any {
  if (!content.comboCharts || content.comboCharts.length === 0) {
    return buildBarOption(content, config);
  }

  const categories = content.comboCharts[0]?.series[0]?.labels || [];
  const seriesArr: any[] = [];
  let hasDualAxis = false;
  const axisLineStyle = { lineStyle: { color: PPTX_AXIS_COLOR } };

  for (const chart of content.comboCharts) {
    if (chart.options?.secondaryValAxis) hasDualAxis = true;
    for (const s of chart.series) {
      const item: any = {
        type: chart.chartType === "area" ? "line" : chart.chartType,
        name: s.name,
        data: s.values,
      };
      if (chart.chartType === "bar") { item.barGap = "0%"; item.barCategoryGap = "60%"; }
      if (chart.chartType === "area") item.areaStyle = { opacity: 0.15 };
      if (chart.options?.secondaryValAxis) item.yAxisIndex = 1;
      if (chart.options?.barGrouping === "stacked") item.stack = "total";
      seriesArr.push(item);
    }
  }

  const yAxes: any[] = [
    { type: "value", axisLabel: { fontFamily: ff(config), fontSize: fs(config), color: "#000" }, axisLine: axisLineStyle, splitLine: { lineStyle: { color: PPTX_AXIS_COLOR, width: 1 } } },
  ];
  if (hasDualAxis) {
    yAxes.push({ type: "value", axisLabel: { fontFamily: ff(config), fontSize: fs(config), color: "#000" }, axisLine: axisLineStyle, splitLine: { show: false } });
  }

  const hasLegend = content.showLegend !== false;
  return applyDefaults({
    title: content.showTitle !== false && content.title ? { text: content.title, left: "center", top: 4, textStyle: { fontSize: 14, fontWeight: "bold" } } : undefined,
    tooltip: { trigger: "axis" },
    legend: hasLegend ? buildLegend(content, config) : undefined,
    grid: { left: 45, right: hasDualAxis ? 45 : 12, top: content.title && content.showTitle !== false ? 35 : 8, bottom: hasLegend ? 30 : 8, containLabel: true },
    xAxis: { type: "category", data: categories, axisLabel: { fontFamily: ff(config), fontSize: fs(config), color: "#000" }, axisLine: axisLineStyle },
    yAxis: yAxes,
    series: seriesArr,
  }, content, config);
}

export function buildPptxChartOption(content: PptxChartContent, config: PptxChartConfig = {}): any {
  switch (content.chartType) {
    case "bar": return buildBarOption(content, config);
    case "line": return buildLineOption(content, config);
    case "pie": return buildPieOption(content, config);
    case "scatter": return buildScatterOption(content, config);
    case "combo": return buildComboOption(content, config);
    default: return applyDefaults({}, content, config);
  }
}
