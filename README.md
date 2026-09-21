# bankops

Deck JSON in, investment-banking-grade PPTX out. Built for AI agents.

```bash
npm install bankops
```

```ts
import { renderPptxToFile } from "bankops";

await renderPptxToFile({
  style: "corporate",
  slides: [{
    title: "Semiconductor Investment Opportunities",
    sectionLabel: "Executive Summary",
    footer: "Source: company filings, BankOps analysis",
    body: { direction: "row", children: [
      { span: 7, header: "Key Themes", content: { type: "text", runs: [
        { text: "AI accelerators drive 30%+ data-center capex growth through 2027", bullet: true },
        { text: "Advanced packaging is the new bottleneck", bullet: true },
      ] } },
      { span: 5, header: "Revenue ($bn)", content: { type: "chart", chartType: "bar",
        categories: ["FY24", "FY25E", "FY26E"], series: [{ name: "Revenue", data: [61, 96, 128] }] } },
    ] },
  }],
}, "semis.pptx");
```

## When to use it

| Use bankops when | Do not use it when |
| --- | --- |
| You need an editable .pptx a banker or investor will open in PowerPoint | You need a web page or a PDF |
| The content is text, tables, KPIs, charts, timelines and logos laid out on a 12-column grid | You need free-form drawing or animation |
| An agent writes the content and a human polishes it | A human will build every slide by hand from scratch |

## What you get

- **A 12-column grid layout engine.** Slides are trees of nodes with `span`, `direction`, `gap`, `padding`, header bars and subheaders. The same engine drives the PPTX and any preview.
- **Eleven content types.** `text` (rich runs, bullets), `statGrid`, `cardGrid`, `table`, `profile`, `image`, `icon`, `chart` (bar, line, pie, combo, waterfall, scatter, gauge), `pptxChart` (native, editable Excel-backed charts), `line`, `timeline`.
- **Style presets and a brand cascade.** `corporate` (the IB house look), `minimal`, `dark`, `warm`, or your own tokens. Theme tokens sit under `deck.style` so an organization's brand applies to every deck.
- **Text budgets.** `annotateDeck()` tells an agent how many characters fit in every zone before it writes copy, so nothing overflows.
- **Validation that teaches.** `validateDeck()` returns JSON paths, messages and a fix for each problem, and warns about typos and spans that do not add up.
- **A CLI, an HTTP service and a JSON Schema.** `bankops render`, `bankops validate`, `bankops annotate`, `bankops schema`, `bankops serve`.

## Deck format in one screen

```jsonc
{
  "style": "corporate",                 // preset name or overrides
  "slides": [{
    "title": "Company Overview",        // optional
    "sectionLabel": "Section 1",        // optional eyebrow
    "footer": "Source: ...",            // optional
    "body": {                           // a grid node
      "direction": "row",               // "row" | "col"
      "children": [                     // a node has children OR content, never both
        { "span": 6, "header": "Overview", "content": { "type": "text", "text": "..." } },
        { "span": 6, "direction": "col", "children": [
          { "span": 6, "content": { "type": "statGrid", "items": [{ "value": "$1.2bn", "label": "Revenue" }] } },
          { "span": 6, "content": { "type": "table", "data": { "headers": ["Metric", "FY25"], "rows": [["EBITDA", "$310m"]], "colWidths": [8, 4] } } }
        ] }
      ]
    }
  }]
}
```

Rules that matter: sibling `span`s add up to 12; `gap`, `padding` and `margin` are inches; `fontSize` is hundredths of a point (`1400` = 14pt); colors are `#RRGGBB`.

Print the full JSON Schema with `npx bankops schema`. The agent skill (`npx bankops skill`) carries the layout recipes for company profiles, comps tables, market overviews and credentials pages.

## API

```ts
import {
  renderPptx, renderPptxToFile, renderDeck, renderFile,   // node: deck → PPTX
  validateDeck, parseDeck, deckJsonSchema, DeckSchema,    // validation, JSON Schema, zod
  annotateDeck,                                           // text budgets per zone
  SlideLayout, SlideStyle, resolveStyleWithTheme,         // layout engine, style cascade
  buildChartOption,                                       // ECharts options (isomorphic)
} from "bankops";
import { ... } from "bankops/browser";                    // the isomorphic subset, no node builtins
```

`renderDeck(deck, options)` returns `{ buffer, slides, warnings, lint }`. Options:

| Option | Meaning |
| --- | --- |
| `style` | Overrides `deck.style` (preset name or overrides) |
| `themeTokens` | Brand tokens used as the base layer under `deck.style` |
| `basePath` | Directory that relative image `path`s resolve against |
| `network` | Fetch image `url`s and Iconify icons (default `true`) |
| `resolveImage` | `(req) => Promise<Buffer \| null>` for images that need generation or search |
| `components` | Named nodes referenced by `$component` |
| `referencePptx` | A .pptx whose theme, fonts and colors are analyzed and reused |

The package makes no AI or search calls. Images marked `source: "ai"` or `source: "search"` are handed to your `resolveImage` function; without one they are skipped with a warning.

## Optional dependencies

- `sharp` renders chart images, icons and SVG or WebP images. Without it, `chart` content falls back to nothing and a warning is emitted; use `pptxChart` (native) instead.
- `pptxgenjs` produces native, editable charts for `pptxChart` content and `chart` content with `pptxExportMode: "nativeExcel"`.

## Render service

```bash
npx bankops serve --port 5490
curl -X POST localhost:5490/render -H 'content-type: application/json' -d @deck.json -o deck.pptx
```

`POST /render`, `POST /validate`, `POST /annotate`, `GET /schema`, `GET /health`. Stateless; the hosted service at bankops.ai runs this same binary.

## License

Free for personal, educational, research and other noncommercial use under the PolyForm Noncommercial License 1.0.0. Commercial use, including internal use at a bank, fund or advisory firm, requires a commercial license. See `COMMERCIAL-LICENSE.md` or contact licensing@bankops.ai.
