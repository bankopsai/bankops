# bankops — source conventions (contract for the CJS→TS conversion)

This package is TypeScript, ESM, `module: NodeNext`. `noImplicitAny` is OFF so untyped
parameters are allowed, but public methods should use the types in `src/types.ts` where obvious.

## Layout
```
src/
  index.ts            public API (node)                     — owner: lead
  browser.ts          isomorphic subset (no node builtins)   — owner: lead
  types.ts            deck JSON types (from app types.ts)   — owner: lead
  schema.ts           zod DeckSchema + validateDeck          — owner: lead
  units.ts            EMU/pt/inch conversions + color        — DONE
  defaults.ts         LINE_SPACING                           — DONE
  raster.ts           optional sharp: svgToPng, toPng        — DONE
  text-metrics.ts     font metrics, chars/line               — owner: lead
  annotate.ts         annotateDeck text budgets              — owner: lead
  layout.ts           SlideLayout (12-col grid → EMU)        — owner: lead
  style.ts            SlideStyle presets + cascade           — owner: lead
  resolve-style.ts    resolveStyle / resolveStyleWithTheme   — owner: lead
  render.ts           renderPptx(deck, options) → Buffer     — owner: lead
  slide-builder.ts    SlideBuilder                           — owner: agent C
  pptx/
    document-model.ts pptx-writer.ts pptx-parser.ts
    layout-catalog.ts layout-analyzer.ts theme-resolver.ts   — owner: agent A
  charts/
    chart-options.ts  isomorphic ECharts option builders (already TS)
    chart-generator.ts ECharts SSR → SVG → PNG via raster.ts
    chart-to-native.ts pptxchart-mapper.ts pptxchart-generator.ts
    pptxchart-options.ts (already TS)                        — owner: agent B
```
Each `.ts` under `src/pptx`, `src/charts`, and `src/slide-builder.ts` currently contains the
ORIGINAL CommonJS source, copied verbatim. Convert it in place.

## Rules
1. **ESM only.** No `require`, no `module.exports`, no `'use strict'`, no `require.main`.
   Relative imports MUST include the `.js` extension: `import DM from "./document-model.js"`.
2. **Keep every export name the original had.** If the original did `module.exports = X`
   (a constructor or object), do `export default X`. If the original exported an object of
   functions, ALSO export each function by name. Other modules import by these names.
3. **Constructor + prototype → `class`.** `function X(o){}` + `X.prototype.m = function(){}`
   becomes `class X { constructor(o){} m(){} }`; `X.foo = function` becomes `static foo()`.
   Declare instance fields at the top of the class (`pres: any;` is fine). Keep behavior identical.
4. **Cross-module imports** (same names as the original requires):
   - `./units.js` → `import U from "../units.js"` (default) — object with the same methods.
   - `./defaults.js` → `import Defaults from "../defaults.js"`.
   - `./document-model.js` → `import DM from "./document-model.js"` (default export = object of classes).
   - `./slide-layout.js` → `import SlideLayout from "./layout.js"` (or `../layout.js`).
   - `./slide-style.js` → `import SlideStyle from "./style.js"`. Lazy requires inside functions
     become normal top-level imports (there are no circular deps once converted).
   - `./chart-to-native`, `./chart-generator`, `./pptxchart-generator` → `../charts/*.js` or `./charts/*.js`.
5. **sharp** is never imported directly. Use `import { svgToPng, toPng } from "../raster.js"`.
6. **pptxgenjs** is optional: load with
   ```ts
   let mod: any;
   try { mod = await import("pptxgenjs"); } catch { throw new MissingDependencyError("pptxgenjs", "native (editable) PPTX charts"); }
   const PptxGenJS = typeof mod === "function" ? mod : (mod.default || mod);
   ```
   (`MissingDependencyError` is exported from `raster.ts`.) Only inside async functions.
7. **Node builtins**: `import fs from "node:fs"; import path from "node:path";`. Modules that
   must stay isomorphic (layout, style, units, text-metrics, annotate, chart-options,
   pptxchart-options, types, schema) must not import node builtins or jszip/echarts/sharp.
8. `var` → `const`/`let`. Keep the logic. Don't "improve" behavior; this is a port.
   Small typing where cheap: `bounds: Bounds`, `content: ContentSpec`, `style: StyleTokens`
   from `../types.js`. Use `any` freely for internal document-model plumbing.
9. Buffers: `Buffer` is fine in node-only modules.
10. Check your own files with `npx tsc --noEmit -p tsconfig.json 2>&1 | grep -E "src/(pptx|charts|slide-builder)"`
    (errors in files owned by others may appear while work is in progress; ignore those).
11. Do not run `npm install`, do not edit `package.json`, do not touch files you don't own.
12. The open package must not contain client names or credentials. Don't copy anything from
    `sdk/CLAUDE.md`.
