/**
 * bankops — deck JSON → investment-banking-grade PPTX, for AI agents.
 *
 *   import { renderPptx, validateDeck, annotateDeck } from "bankops";
 *   const pptx = await renderPptx(deck);
 */
export * from "./browser.js";
export { renderDeck, renderPptx, renderPptxToFile, renderFile, type RenderOptions, type RenderResult, type ComponentFactory } from "./render.js";
export { type ImageRequest, type ImageResolver } from "./resolvers.js";
export { SlideBuilder } from "./slide-builder.js";
export { default as ChartGenerator } from "./charts/chart-generator.js";
export { convertToNative } from "./charts/chart-to-native.js";
export { svgToPng, toPng, hasSharp, MissingDependencyError } from "./raster.js";
export { default as PptxParser } from "./pptx/pptx-parser.js";
export { default as PptxWriter } from "./pptx/pptx-writer.js";
export { default as DocumentModel } from "./pptx/document-model.js";
export { createMcpServer, skillSections, budgetReport, type McpServerOptions } from "./mcp/server.js";
export { renderSlidePng, pptxToPng, findSoffice } from "./preview.js";
export { createHandler as createServiceHandler, startServer as startService, type ServeOptions } from "./server/server.js";
