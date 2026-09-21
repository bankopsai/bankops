// ===== ChartGenerator =====
// Server-side chart generation using Apache ECharts SSR mode.
// Renders charts to SVG via ECharts, converts to PNG via raster.ts (sharp).
// Designed for embedding chart images in PPTX slides.
//
// The ECharts option objects themselves are built by the isomorphic
// `buildChartOption` in ./chart-options.ts (shared with the browser preview);
// this module only adds the render config and the SSR → SVG → PNG pipeline.

import * as echarts from "echarts";
import { svgToPng } from "../raster.js";
import { buildChartOption, DEFAULT_COLORS, type ChartConfig } from "./chart-options.js";

export interface ChartGeneratorConfig {
  width?: number;
  height?: number;
  colors?: string[];
  fontFamily?: string;
  fontSize?: number;
  backgroundColor?: string;
  /** Render at Nx resolution for crisp output (SVG density = 72 * dpiScale). */
  dpiScale?: number;
  barBorderColor?: string | null;
  barBorderWidth?: number | null;
  /** theme-level bar defaults: { showXAxis, showYAxis, showSplitLine, gridLeft, ... } */
  bar?: ChartConfig["bar"] | null;
  /** theme-level pie defaults: { borderColor, borderWidth } */
  pie?: ChartConfig["pie"] | null;
  /** theme-level combo defaults: { showPrimaryYAxis, showSecondaryYAxis, showSplitLine, showBarValues, showLineValues } */
  combo?: ChartConfig["combo"] | null;
}

export default class ChartGenerator {
  width: number;
  height: number;
  colors: string[];
  fontFamily: string;
  fontSize: number;
  backgroundColor: string;
  dpiScale: number;
  barBorderColor: string | null;
  barBorderWidth: number | null;
  barDefaults: ChartConfig["bar"] | null;
  pieDefaults: ChartConfig["pie"] | null;
  comboDefaults: ChartConfig["combo"] | null;

  constructor(config?: ChartGeneratorConfig) {
    const c = config || {};
    this.width = c.width || 800;
    this.height = c.height || 500;
    this.colors = c.colors || DEFAULT_COLORS;
    this.fontFamily = c.fontFamily || "Arial, Helvetica, sans-serif";
    this.fontSize = c.fontSize || 11;
    this.backgroundColor = c.backgroundColor || "#FFFFFF";
    this.dpiScale = c.dpiScale || 1;  // Render at Nx resolution for crisp output
    this.barBorderColor = c.barBorderColor || null;
    this.barBorderWidth = c.barBorderWidth != null ? c.barBorderWidth : null;
    this.barDefaults = c.bar || null;
    this.pieDefaults = c.pie || null;
    this.comboDefaults = c.combo || null;
  }

  /** Config handed to the isomorphic option builders. */
  _optionConfig(): ChartConfig {
    return {
      width: this.width,
      height: this.height,
      colors: this.colors,
      fontFamily: this.fontFamily,
      fontSize: this.fontSize,
      backgroundColor: this.backgroundColor,
      barBorderColor: this.barBorderColor || undefined,
      barBorderWidth: this.barBorderWidth != null ? this.barBorderWidth : undefined,
      bar: this.barDefaults || undefined,
      pie: this.pieDefaults || undefined,
      combo: this.comboDefaults || undefined,
    };
  }

  /** Build the ECharts option for a chart type without rendering it. */
  buildOption(chartType: string, spec: any): any {
    return buildChartOption(chartType, spec, this._optionConfig());
  }

  // ---- Core render: ECharts option → PNG Buffer ----
  async _render(option: any): Promise<Buffer> {
    // Apply defaults (no-ops when the option came from buildChartOption)
    if (!option.backgroundColor) option.backgroundColor = this.backgroundColor;
    if (!option.color) option.color = this.colors;
    if (!option.textStyle) option.textStyle = {};
    if (!option.textStyle.fontFamily) option.textStyle.fontFamily = this.fontFamily;
    option.animation = false;

    const chart = echarts.init(null, null, {
      renderer: "svg",
      ssr: true,
      width: this.width,
      height: this.height,
    });
    chart.setOption(option);
    const svgStr = chart.renderToSVGString();
    chart.dispose();

    // Render at higher density for crisp PPTX output (default SVG density = 72)
    const density = Math.round(72 * (this.dpiScale || 1));
    return svgToPng(svgStr, { density });
  }

  // ---- Bar chart (vertical or horizontal, grouped or stacked) ----
  async bar(spec: any): Promise<Buffer> {
    return this._render(this.buildOption("bar", spec));
  }

  // ---- Line chart (optionally smooth, optionally area-filled) ----
  async line(spec: any): Promise<Buffer> {
    return this._render(this.buildOption("line", spec));
  }

  // ---- Pie / Doughnut chart ----
  async pie(spec: any): Promise<Buffer> {
    return this._render(this.buildOption("pie", spec));
  }

  // ---- Scatter plot ----
  async scatter(spec: any): Promise<Buffer> {
    return this._render(this.buildOption("scatter", spec));
  }

  // ---- Waterfall chart ----
  async waterfall(spec: any): Promise<Buffer> {
    return this._render(this.buildOption("waterfall", spec));
  }

  // ---- Combo chart (bars + lines on same axes) ----
  async combo(spec: any): Promise<Buffer> {
    return this._render(this.buildOption("combo", spec));
  }

  // ---- Gauge chart ----
  async gauge(spec: any): Promise<Buffer> {
    return this._render(this.buildOption("gauge", spec));
  }

  // ---- Convenience: resize for a specific output ----
  withSize(width: number, height: number): ChartGenerator {
    return new ChartGenerator({
      width,
      height,
      colors: this.colors,
      fontFamily: this.fontFamily,
      backgroundColor: this.backgroundColor,
      dpiScale: this.dpiScale,
    });
  }
}

export { ChartGenerator };
