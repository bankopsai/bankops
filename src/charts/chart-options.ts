/* ECharts option builders — the single isomorphic source of truth for both the
   browser preview and the PPTX export (chart-generator.ts feeds these options
   to ECharts SSR). Only builds plain option objects: no node builtins, no
   echarts import, no rendering. */

/* eslint-disable @typescript-eslint/no-explicit-any */

export const DEFAULT_COLORS = [
  "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5",
  "#70AD47", "#264478", "#9B57A0", "#636363", "#EB7E3A",
];

export interface ChartConfig {
  colors?: string[];
  /** Render surface size in px; used to place pie/gauge centers when grid
   *  margins are given. Defaults to 800×500 when absent. */
  width?: number;
  height?: number;
  fontFamily?: string;
  fontSize?: number;
  backgroundColor?: string;
  barBorderColor?: string;
  barBorderWidth?: number;
  bar?: { showXAxis?: boolean; showYAxis?: boolean; showSplitLine?: boolean; showSeriesLabels?: boolean; gridLeft?: number; gridRight?: number; gridTop?: number; gridBottom?: number };
  pie?: { borderColor?: string; borderWidth?: number };
  combo?: { showPrimaryYAxis?: boolean; showSecondaryYAxis?: boolean; showSplitLine?: boolean; showBarValues?: boolean; showLineValues?: boolean; barValueColor?: string; lineValueColor?: string };
}

function applyDefaults(option: any, config: ChartConfig): any {
  if (!option.backgroundColor) option.backgroundColor = config.backgroundColor || "#FFFFFF";
  if (!option.color) option.color = config.colors || DEFAULT_COLORS;
  if (!option.textStyle) option.textStyle = {};
  if (!option.textStyle.fontFamily) option.textStyle.fontFamily = config.fontFamily || "Arial, Helvetica, sans-serif";
  if (!option.textStyle.fontSize) option.textStyle.fontSize = config.fontSize || 11;
  option.animation = false;
  return option;
}

/** Resolved font family and size from config */
function ff(config: ChartConfig): string { return config.fontFamily || "Arial, Helvetica, sans-serif"; }
function fs(config: ChartConfig): number { return config.fontSize || 11; }

function applyGridOverrides(grid: any, spec: any): any {
  if (spec.gridTop != null) grid.top = spec.gridTop;
  if (spec.gridRight != null) grid.right = spec.gridRight;
  if (spec.gridBottom != null) grid.bottom = spec.gridBottom;
  if (spec.gridLeft != null) grid.left = spec.gridLeft;
  return grid;
}

/** Pie and gauge don't use ECharts `grid`. Translate grid spacing into a
 *  shifted `center` so the chart sits in the middle of the usable area.
 *  Uses config.width/height (the actual render size) to convert px margins
 *  to percentages, falling back to a reference size of 800×500. */
function pieGaugeCenter(s: any, config: ChartConfig = {}): [string, string] {
  if (s.gridTop == null && s.gridBottom == null && s.gridLeft == null && s.gridRight == null) {
    return ["50%", "50%"];
  }
  const width = config.width || 800;
  const height = config.height || 500;
  const t = s.gridTop ?? (s.title ? 45 : 20);
  const b = s.gridBottom ?? 25;
  const l = s.gridLeft ?? 20;
  const r = s.gridRight ?? 20;
  const cx = (l + (width - l - r) / 2) / width * 100;
  const cy = (t + (height - t - b) / 2) / height * 100;
  return [cx.toFixed(1) + "%", cy.toFixed(1) + "%"];
}

function buildValueFormatter(s: any): (v: number) => string {
  return makeFormatter(s.valuePrefix, s.valueSuffix, s.valueDecimals, s.valueThousands, s.negativeFormat);
}

function makeFormatter(prefix?: string, suffix?: string, decimals?: number, thousands?: boolean, negFmt?: string): (v: number) => string {
  return (v: number) => {
    const isNeg = v < 0;
    const abs = Math.abs(v);
    let str = abs.toFixed(decimals ?? 0);
    if (thousands) {
      const parts = str.split(".");
      parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
      str = parts.join(".");
    }
    str = (prefix || "") + str + (suffix || "");
    if (isNeg) {
      str = negFmt === "parens" ? `(${str})` : `-${str}`;
    }
    return str;
  };
}

function applyYAxisRange(axis: any, s: any, secondary?: boolean): any {
  const minKey = secondary ? "y2Min" : "yMin";
  const maxKey = secondary ? "y2Max" : "yMax";
  if (s[minKey] != null) axis.min = s[minKey];
  if (s[maxKey] != null) axis.max = s[maxKey];
  return axis;
}

function legendSpec(pos: string | undefined, fontSize: number): any {
  const textStyle = { fontSize };
  switch (pos) {
    case "left": return { left: 5, orient: "vertical", textStyle };
    case "right": return { right: 5, orient: "vertical", textStyle };
    case "top": return { top: 5, textStyle };
    default: return { bottom: 5, textStyle };
  }
}

function applyXAxisRange(axis: any, s: any): any {
  if (s.xMin != null) axis.min = s.xMin;
  if (s.xMax != null) axis.max = s.xMax;
  return axis;
}

export function buildBarOption(spec: any, config: ChartConfig = {}): any {
  const s = spec || {};
  const isHorizontal = s.horizontal === true;

  const bd = config.bar || {};
  const valFormatter = buildValueFormatter(s);
  const seriesLabels = s.stacked && (s.showSeriesLabels ?? bd.showSeriesLabels);
  const catColors: Record<number, string> = s.categoryColors || {};
  const seriesArr = (s.series || []).map((ser: any) => {
    const item: any = { type: "bar", name: ser.name || "", data: ser.data || [], barMaxWidth: 50 };
    if (s.stacked) item.stack = "total";
    if (seriesLabels) {
      item.label = { show: true, position: "inside", fontSize: 10, formatter: (p: any) => valFormatter(p.value) };
    } else if (s.showBarValues !== false) {
      item.label = { show: true, position: isHorizontal ? "right" : "top", fontSize: 10, formatter: (p: any) => valFormatter(p.value) };
    }
    const itemStyle: any = {};
    const bbc = s.barBorderColor || config.barBorderColor;
    const bbw = s.barBorderWidth ?? config.barBorderWidth;
    if (bbc) itemStyle.borderColor = bbc;
    if (bbw != null) itemStyle.borderWidth = bbw;
    if (ser.color) itemStyle.color = ser.color;
    else if (s.barColor && (s.series || []).length <= 1) itemStyle.color = s.barColor;
    if (Object.keys(itemStyle).length) item.itemStyle = itemStyle;
    // Per-category color overrides (only for single-series bar charts)
    if (Object.keys(catColors).length > 0 && (s.series || []).length <= 1) {
      item.data = (ser.data || []).map((val: number, ci: number) =>
        catColors[ci] ? { value: val, itemStyle: { color: catColors[ci] } } : val
      );
    }
    return item;
  });

  const catAxis = applyXAxisRange({
    type: "category", data: s.categories || [],
    axisLabel: { fontFamily: ff(config), fontSize: fs(config) },
    axisTick: { alignWithLabel: true },
  }, s);
  if (s.xAxisLabel) {
    catAxis.name = s.xAxisLabel;
    catAxis.nameLocation = s.xAxisLabelPosition === "center" ? "middle" : "end";
    catAxis.nameGap = s.xAxisLabelPosition === "center" ? 30 : 15;
  }
  const hasValueFmt = s.valuePrefix || s.valueSuffix || s.valueDecimals || s.valueThousands;
  const valAxisLabel: any = { fontFamily: ff(config), fontSize: fs(config) };
  if (hasValueFmt) valAxisLabel.formatter = (v: number) => valFormatter(v);
  const valAxis = applyYAxisRange({
    type: "value",
    axisLabel: valAxisLabel,
    splitLine: { lineStyle: { color: "#E8E8E8" } },
  }, s);
  if (s.yAxisLabel) {
    valAxis.name = s.yAxisLabel;
    valAxis.nameLocation = s.yAxisLabelPosition === "top" ? "end" : "middle";
    valAxis.nameRotate = s.yAxisLabelPosition === "side" ? 90 : 0;
    valAxis.nameGap = s.yAxisLabelPosition === "side" ? 45 : 15;
  }

  // Stacked bar totals: invisible series that shows total labels above each stack
  if (s.stacked && (s.showTotalValues || (Array.isArray(s.stackTotal) && s.stackTotal.length > 0))) {
    const totals = Array.isArray(s.stackTotal) && s.stackTotal.length > 0
      ? s.stackTotal
      : (s.categories || []).map((_: unknown, ci: number) => (s.series || []).reduce((sum: number, ser: { data: number[] }) => sum + (ser.data[ci] || 0), 0));
    seriesArr.push({
      type: "bar", stack: "total", data: totals.map(() => 0), barMaxWidth: 50,
      itemStyle: { color: "transparent" }, emphasis: { itemStyle: { color: "transparent" } },
      label: { show: true, position: isHorizontal ? "right" : "top", fontSize: 10,
        formatter: (p: any) => valFormatter(totals[p.dataIndex]) },
    });
  }

  const finalXAxis = isHorizontal ? valAxis : catAxis;
  const finalYAxis = isHorizontal ? catAxis : valAxis;
  // showXAxis/showYAxis refer to logical axes: X = category, Y = value (regardless of orientation)
  if ((s.showXAxis ?? bd.showXAxis) === false) catAxis.show = false;
  if ((s.showYAxis ?? bd.showYAxis) === false) valAxis.show = false;
  // splitLine on the value axis (Y for vertical, X for horizontal)
  if ((s.showSplitLine ?? bd.showSplitLine) === false) valAxis.splitLine = { show: false };

  // Merge theme bar grid defaults under per-chart overrides
  const gridSpec = { ...s };
  if (gridSpec.gridLeft == null && bd.gridLeft != null) gridSpec.gridLeft = bd.gridLeft;
  if (gridSpec.gridRight == null && bd.gridRight != null) gridSpec.gridRight = bd.gridRight;
  if (gridSpec.gridTop == null && bd.gridTop != null) gridSpec.gridTop = bd.gridTop;
  if (gridSpec.gridBottom == null && bd.gridBottom != null) gridSpec.gridBottom = bd.gridBottom;

  return applyDefaults({
    title: s.title ? { text: s.title, left: "center", top: 8, textStyle: { fontSize: 14, fontWeight: "bold" } } : undefined,
    tooltip: { trigger: "axis" },
    legend: seriesArr.length > 1 ? { bottom: 5, textStyle: { fontSize: fs(config) } } : undefined,
    grid: applyGridOverrides({ left: isHorizontal ? 10 : 60, right: 20, top: s.title ? 45 : 20, bottom: seriesArr.length > 1 ? 40 : 25, containLabel: isHorizontal }, gridSpec),
    xAxis: finalXAxis,
    yAxis: finalYAxis,
    series: seriesArr,
  }, config);
}

export function buildLineOption(spec: any, config: ChartConfig = {}): any {
  const s = spec || {};
  const valFormatter = buildValueFormatter(s);
  const lineStyle: any = {};
  if (s.lineColor) lineStyle.color = s.lineColor;
  if (s.lineWidth != null) lineStyle.width = s.lineWidth;
  const pointStyle: any = {};
  if (s.pointColor) pointStyle.color = s.pointColor;
  if (s.pointBorderColor) pointStyle.borderColor = s.pointBorderColor;
  if (s.pointBorderWidth != null) pointStyle.borderWidth = s.pointBorderWidth;

  const seriesArr = (s.series || []).map((ser: any) => ({
    type: "line", name: ser.name || "", data: ser.data || [],
    smooth: s.smooth === true,
    areaStyle: s.area ? { opacity: 0.15 } : undefined,
    symbol: s.pointShape || "circle",
    symbolSize: s.pointSize ?? 6,
    ...(Object.keys(pointStyle).length ? { itemStyle: pointStyle } : {}),
    ...(Object.keys(lineStyle).length ? { lineStyle } : {}),
    ...(s.showLineValues ? { label: { show: true, position: "top", fontSize: 10, formatter: (p: any) => valFormatter(p.value) } } : {}),
  }));

  const xAxis: any = applyXAxisRange({
    type: "category", data: s.categories || [],
    axisLabel: { fontFamily: ff(config), fontSize: fs(config) },
    boundaryGap: false,
  }, s);
  if (s.xAxisLabel) {
    xAxis.name = s.xAxisLabel;
    xAxis.nameLocation = s.xAxisLabelPosition === "center" ? "middle" : "end";
    xAxis.nameGap = s.xAxisLabelPosition === "center" ? 30 : 15;
  }

  const yAxis: any = applyYAxisRange({
    type: "value",
    axisLabel: { fontFamily: ff(config), fontSize: fs(config) },
    splitLine: { lineStyle: { color: "#E8E8E8" } },
  }, s);
  if (s.yAxisLabel) {
    yAxis.name = s.yAxisLabel;
    yAxis.nameLocation = s.yAxisLabelPosition === "top" ? "end" : "middle";
    yAxis.nameRotate = s.yAxisLabelPosition === "side" ? 90 : 0;
    yAxis.nameGap = s.yAxisLabelPosition === "side" ? 45 : 15;
  }

  return applyDefaults({
    title: s.title ? { text: s.title, left: "center", top: 8, textStyle: { fontSize: 14, fontWeight: "bold" } } : undefined,
    tooltip: { trigger: "axis" },
    legend: seriesArr.length > 1 ? { bottom: 5, textStyle: { fontSize: fs(config) } } : undefined,
    grid: applyGridOverrides({ left: 60, right: 20, top: s.title ? 45 : 20, bottom: seriesArr.length > 1 ? 40 : 25, containLabel: false }, s),
    xAxis,
    yAxis,
    series: seriesArr,
  }, config);
}

export function buildPieOption(spec: any, config: ChartConfig = {}): any {
  const s = spec || {};
  const inner = s.innerRadius || (s.doughnut ? "40%" : "0%");
  const outer = s.outerRadius || "70%";
  const radius = [inner, outer];
  const showLabel = s.showLabel !== false && s.labelPosition !== "none";
  const showValues = !!s.showPieValues;
  const showAnyLabel = showLabel || showValues;
  const labelPos = showValues && s.pieValuePosition
    ? s.pieValuePosition
    : (s.labelPosition && s.labelPosition !== "none") ? s.labelPosition : "outside";
  const center: [string, string] = [s.centerX || "50%", s.centerY || "50%"];

  const pd = config.pie || {};
  const pieItemStyle: any = {};
  const pbc = s.pieBorderColor || pd.borderColor;
  const pbw = s.pieBorderWidth ?? pd.borderWidth;
  if (pbc) pieItemStyle.borderColor = pbc;
  if (pbw != null) pieItemStyle.borderWidth = pbw;

  // Build label formatter based on what's enabled
  const hasValueFmt = s.valuePrefix || s.valueSuffix || s.valueDecimals || s.valueThousands;
  const valueFmt = hasValueFmt ? buildValueFormatter(s) : null;
  let formatter: string | ((p: any) => string) | undefined;
  if (showLabel && showValues && valueFmt) formatter = (p: any) => `${p.name}: ${valueFmt(p.value)}`;
  else if (showLabel && showValues) formatter = "{b}: {c}";
  else if (showValues && valueFmt) formatter = (p: any) => valueFmt(p.value);
  else if (showValues) formatter = "{c}";

  return applyDefaults({
    title: s.title ? { text: s.title, left: "center", top: 8, textStyle: { fontSize: 14, fontWeight: "bold" } } : undefined,
    tooltip: { trigger: "item", formatter: "{b}: {c} ({d}%)" },
    legend: s.showLegend === false ? undefined : legendSpec(s.legendPosition, fs(config)),
    series: [{
      type: "pie", radius, center,
      data: (s.items || []).map((item: any, i: number) => {
        const d: any = { name: item.name, value: item.value };
        if (s.sliceColors && s.sliceColors[i]) d.itemStyle = { color: s.sliceColors[i] };
        return d;
      }),
      roseType: s.roseType || undefined,
      startAngle: s.startAngle ?? 90,
      padAngle: s.padAngle ?? 0,
      label: showAnyLabel ? { fontSize: fs(config), fontFamily: ff(config), position: labelPos, ...(formatter ? { formatter } : {}) } : { show: false },
      labelLine: showAnyLabel && labelPos === "outside" ? { show: true } : { show: false },
      itemStyle: Object.keys(pieItemStyle).length ? pieItemStyle : undefined,
      emphasis: { itemStyle: { shadowBlur: 10, shadowOffsetX: 0, shadowColor: "rgba(0,0,0,0.3)" } },
    }],
  }, config);
}

export function buildScatterOption(spec: any, config: ChartConfig = {}): any {
  const s = spec || {};
  const seriesArr = (s.series || []).map((ser: any) => ({
    type: "scatter", name: ser.name || "", data: ser.data || [], symbolSize: 10,
  }));

  return applyDefaults({
    title: s.title ? { text: s.title, left: "center", top: 8, textStyle: { fontSize: 14, fontWeight: "bold" } } : undefined,
    tooltip: { trigger: "item" },
    legend: seriesArr.length > 1 ? legendSpec(s.legendPosition, fs(config)) : undefined,
    grid: applyGridOverrides({ left: 60, right: s.legendPosition === "right" ? 100 : 20, top: s.title ? 45 : (s.legendPosition === "top" ? 40 : 20), bottom: seriesArr.length > 1 ? (s.legendPosition === "bottom" || !s.legendPosition ? 40 : 25) : 25, containLabel: false }, s),
    xAxis: applyXAxisRange({ type: "value", axisLabel: { fontFamily: ff(config), fontSize: fs(config) }, splitLine: { lineStyle: { color: "#E8E8E8" } } }, s),
    yAxis: applyYAxisRange({ type: "value", axisLabel: { fontFamily: ff(config), fontSize: fs(config) }, splitLine: { lineStyle: { color: "#E8E8E8" } } }, s),
    series: seriesArr,
  }, config);
}

export function buildWaterfallOption(spec: any, config: ChartConfig = {}): any {
  const s = spec || {};
  const categories = s.categories || [];
  const data = s.data || [];

  const base: (number | string)[] = [];
  const positive: (number | string)[] = [];
  const negative: (number | string)[] = [];
  let running = 0;

  for (let i = 0; i < data.length; i++) {
    const val = data[i];
    if (i === 0 || i === data.length - 1) {
      base.push(0);
      positive.push(Math.abs(val));
      negative.push("-");
    } else if (val >= 0) {
      base.push(running);
      positive.push(val);
      negative.push("-");
    } else {
      base.push(running + val);
      positive.push("-");
      negative.push(Math.abs(val));
    }
    if (i < data.length - 1) running += val;
  }

  return applyDefaults({
    title: s.title ? { text: s.title, left: "center", top: 8, textStyle: { fontSize: 14, fontWeight: "bold" } } : undefined,
    tooltip: { trigger: "axis", axisPointer: { type: "shadow" } },
    grid: applyGridOverrides({ left: 60, right: 20, top: s.title ? 45 : 20, bottom: 25, containLabel: false }, s),
    xAxis: { type: "category", data: categories, axisLabel: { fontFamily: ff(config), fontSize: fs(config) } },
    yAxis: applyYAxisRange({ type: "value", axisLabel: { fontFamily: ff(config), fontSize: fs(config) }, splitLine: { lineStyle: { color: "#E8E8E8" } } }, s),
    series: [
      { type: "bar", stack: "waterfall", data: base, itemStyle: { color: "transparent" }, emphasis: { itemStyle: { color: "transparent" } } },
      { type: "bar", stack: "waterfall", name: "Increase", data: positive, itemStyle: { color: "#4472C4" }, barMaxWidth: 50,
        label: { show: true, position: "top", fontSize: 10, formatter: (params: any) => (params.value !== "-" ? params.value : "") } },
      { type: "bar", stack: "waterfall", name: "Decrease", data: negative, itemStyle: { color: "#ED7D31" }, barMaxWidth: 50,
        label: { show: true, position: "bottom", fontSize: 10, formatter: (params: any) => (params.value !== "-" ? params.value : "") } },
    ],
  }, config);
}

export function buildComboOption(spec: any, config: ChartConfig = {}): any {
  const s = spec || {};
  const cd = config.combo || {};
  const seriesArr: any[] = [];

  // Primary axis formatter (bars) — uses existing valuePrefix/valueSuffix fields
  const primaryFmt = buildValueFormatter(s);
  const hasPrimaryFmt = s.valuePrefix || s.valueSuffix || s.valueDecimals || s.valueThousands;
  // Secondary axis formatter (lines) — uses y2* fields, falls back to primary
  const hasY2Fields = s.y2ValuePrefix != null || s.y2ValueSuffix != null || s.y2ValueDecimals != null || s.y2ValueThousands != null || s.y2NegativeFormat != null;
  const secFmt = hasY2Fields
    ? makeFormatter(s.y2ValuePrefix, s.y2ValueSuffix, s.y2ValueDecimals, s.y2ValueThousands, s.y2NegativeFormat)
    : primaryFmt;
  const hasSecFmt = hasY2Fields || hasPrimaryFmt;

  const barValues = s.showBarValues ?? cd.showBarValues;
  (s.bars || []).forEach((ser: any) => {
    const item: any = { type: "bar", name: ser.name || "", data: ser.data || [], barMaxWidth: 50 };
    if (s.stacked) item.stack = "total";
    if (barValues) {
      item.label = { show: true, position: "top", fontSize: 10, formatter: (p: any) => primaryFmt(p.value), ...(cd.barValueColor ? { color: cd.barValueColor } : {}) };
    }
    // Bar fill color
    const barItemStyle: any = {};
    if (s.barColor && (s.bars || []).length <= 1) barItemStyle.color = s.barColor;
    const bbc = s.barBorderColor || config.barBorderColor;
    const bbw = s.barBorderWidth ?? config.barBorderWidth;
    if (bbc) barItemStyle.borderColor = bbc;
    if (bbw != null) barItemStyle.borderWidth = bbw;
    if (Object.keys(barItemStyle).length) item.itemStyle = barItemStyle;
    seriesArr.push(item);
  });
  // Stacked bar totals for combo charts
  if (s.stacked && (s.showTotalValues || (Array.isArray(s.stackTotal) && s.stackTotal.length > 0))) {
    const totals = Array.isArray(s.stackTotal) && s.stackTotal.length > 0
      ? s.stackTotal
      : (s.categories || []).map((_: unknown, ci: number) => (s.bars || []).reduce((sum: number, ser: { data: number[] }) => sum + (ser.data[ci] || 0), 0));
    seriesArr.push({
      type: "bar", stack: "total", data: totals.map(() => 0), barMaxWidth: 50,
      itemStyle: { color: "transparent" }, emphasis: { itemStyle: { color: "transparent" } },
      label: { show: true, position: "top", fontSize: 10,
        formatter: (p: any) => primaryFmt(totals[p.dataIndex]) },
    });
  }
  const lineValues = s.showLineValues ?? cd.showLineValues;
  // Line styling for combo charts
  const comboLineStyle: any = {};
  if (s.lineColor) comboLineStyle.color = s.lineColor;
  if (s.lineWidth != null) comboLineStyle.width = s.lineWidth;
  const comboPointStyle: any = {};
  if (s.pointColor) comboPointStyle.color = s.pointColor;
  if (s.pointBorderColor) comboPointStyle.borderColor = s.pointBorderColor;
  if (s.pointBorderWidth != null) comboPointStyle.borderWidth = s.pointBorderWidth;
  (s.lines || []).forEach((ser: any) => {
    const item: any = {
      type: "line", name: ser.name || "", data: ser.data || [],
      smooth: s.smooth !== false, symbol: s.pointShape || "circle", symbolSize: s.pointSize ?? 6,
      yAxisIndex: s.dualAxis ? 1 : 0,
    };
    if (Object.keys(comboLineStyle).length) item.lineStyle = comboLineStyle;
    if (Object.keys(comboPointStyle).length) item.itemStyle = comboPointStyle;
    if (lineValues) {
      item.label = { show: true, fontSize: 10, formatter: (p: any) => secFmt(p.value), ...(cd.lineValueColor ? { color: cd.lineValueColor } : {}) };
    }
    seriesArr.push(item);
  });

  const primaryAxisLabel: any = { fontFamily: ff(config), fontSize: fs(config) };
  if (hasPrimaryFmt) primaryAxisLabel.formatter = (v: number) => primaryFmt(v);
  const primaryAxis = applyYAxisRange({ type: "value", axisLabel: primaryAxisLabel, splitLine: { lineStyle: { color: "#E8E8E8" } } }, s);
  if ((s.showSplitLine ?? cd.showSplitLine) === false) primaryAxis.splitLine = { show: false };
  if ((s.showYAxis ?? cd.showPrimaryYAxis) === false) primaryAxis.show = false;
  const yAxes: any[] = [primaryAxis];
  if (s.dualAxis) {
    const secAxisLabel: any = { fontFamily: ff(config), fontSize: fs(config) };
    if (hasSecFmt) secAxisLabel.formatter = (v: number) => secFmt(v);
    const secAxis = applyYAxisRange({ type: "value", axisLabel: secAxisLabel, splitLine: { show: false } }, s, true);
    if ((s.showSecondaryYAxis ?? cd.showSecondaryYAxis) === false) secAxis.show = false;
    yAxes.push(secAxis);
  }

  return applyDefaults({
    title: s.title ? { text: s.title, left: "center", top: 8, textStyle: { fontSize: 14, fontWeight: "bold" } } : undefined,
    tooltip: { trigger: "axis" },
    legend: { bottom: 5, textStyle: { fontSize: fs(config) } },
    grid: applyGridOverrides({ left: 60, right: s.dualAxis ? 60 : 20, top: s.title ? 45 : 20, bottom: 40, containLabel: false }, s),
    xAxis: { type: "category", data: s.categories || [], axisLabel: { fontFamily: ff(config), fontSize: fs(config) } },
    yAxis: yAxes,
    series: seriesArr,
  }, config);
}

export function buildGaugeOption(spec: any, config: ChartConfig = {}): any {
  const s = spec || {};
  const center = pieGaugeCenter(s, config);
  return applyDefaults({
    title: s.title ? { text: s.title, left: "center", top: 8, textStyle: { fontSize: 14, fontWeight: "bold" } } : undefined,
    series: [{
      type: "gauge", center, startAngle: 200, endAngle: -20,
      min: s.min ?? 0, max: s.max ?? 100,
      detail: { formatter: s.format || "{value}%", fontSize: 20, fontWeight: "bold", offsetCenter: [0, "60%"] },
      data: [{ value: s.value ?? 0, name: s.label || "" }],
      axisLine: { lineStyle: { width: 20, color: [[0.3, "#ED7D31"], [0.7, "#FFC000"], [1, "#70AD47"]] } },
      pointer: { width: 5 },
      title: { fontSize: 12, offsetCenter: [0, "80%"] },
    }],
  }, config);
}

export function buildChartOption(chartType: string, spec: any, config: ChartConfig = {}): any {
  switch (chartType) {
    case "bar": return buildBarOption(spec, config);
    case "line": return buildLineOption(spec, config);
    case "pie": return buildPieOption(spec, config);
    case "scatter": return buildScatterOption(spec, config);
    case "waterfall": return buildWaterfallOption(spec, config);
    case "combo": return buildComboOption(spec, config);
    case "gauge": return buildGaugeOption(spec, config);
    default: return applyDefaults({}, config);
  }
}
