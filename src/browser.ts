/**
 * bankops/browser — the isomorphic subset: types, schema/validation, layout engine,
 * style cascade, text budgets and ECharts option builders. No node builtins, no PPTX writer.
 */
export * from "./types.js";
export * from "./schema.js";
export { SlideLayout, type SlideLayoutConfig, type LayoutPreset } from "./layout.js";
export { SlideStyle, DEFAULT_STYLE, type StyleOverrides, type StyleSheet } from "./style.js";
export { resolveStyle, resolveStyleWithTheme, mergeStyleWithTheme, resolveDeck, resolveSlide } from "./resolve-style.js";
export * from "./text-metrics.js";
export { annotateDeck, resolveStyleFonts, type StyleFontDefaults } from "./annotate.js";
export { Units, EMU_PER_INCH, EMU_PER_POINT } from "./units.js";
export { LINE_SPACING } from "./defaults.js";
export * from "./charts/chart-options.js";
export { buildPptxChartOption } from "./charts/pptxchart-options.js";
export { fitTitle, isNumericColumn, type TitleFit, type FitTitleOptions } from "./fit.js";
