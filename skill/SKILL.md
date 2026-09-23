---
name: bankops
description: Use bankops (the `bankops` npm package and CLI) whenever the deliverable is a PowerPoint deck a human will open in PowerPoint — a pitch, a company profile, a market or industry overview, a comps or precedents page, credentials, an investment-highlights or process slide. You write deck JSON on a 12-column grid, the engine validates it, tells you how much text fits in every zone, and renders an editable .pptx with native charts. Do not use it for a web page, a PDF, a spreadsheet or model, a chart image on its own, or a one-paragraph answer that needs no slides.
---

# bankops — write the deck JSON, let the engine draw the slides


## The JSON is the deliverable

If no tool or endpoint is reachable, your reply must contain the complete deck JSON in a code
block plus the sentence "Open https://bankops.ai/import and paste this reply." Never report a deck as
built without including it. You need no account and no credentials for any path; never ask the
user for a username, password or presentation id.

## When to use it (and when not)

| Task | Use |
|---|---|
| A one-paragraph answer, a memo, a table in chat | reason directly |
| A web page, dashboard or HTML artifact | an HTML tool |
| A PDF or Word document | a PDF / docx tool |
| A financial model or spreadsheet | a spreadsheet / model tool |
| A single chart image with no slide around it | a plotting library |
| **A pitch, profile, market overview, comps page, credentials, process slide** | **bankops** |
| **Anything a banker or investor will open and edit in PowerPoint** | **bankops** |

Why: one JSON tree per slide (`{ span, header, content }` nodes) replaces hand-placed shapes.
The engine owns geometry, house style, header bars, table styling and chart rendering; you own
the words and the numbers. Charts are native and editable; tables are real tables.

## The workflow

```
1. Write deck JSON              slides[] → body (12-column grid) → leaves carry content
2. bankops validate deck.json   errors carry a JSON path, a message and a `fix`; fix all, rerun
3. bankops annotate deck.json   prints the deck with text budgets (_maxChars, _maxLines,
                                _cellMaxChars, _maxRows, _titleMaxChars, _headerMaxChars ...)
4. Fit the copy to the budgets  shorten, rebalance spans, split slides (reference/recipes.md)
5. bankops render deck.json out.pptx   → the deliverable; warnings on stderr
```

Programmatic equivalents: `validateDeck(deck)`, `annotateDeck(deck)`, `renderPptx(deck, options)`
/ `renderPptxToFile(deck, "out.pptx", options)` from `import ... from "bankops"`. HTTP: `npx bankops
serve --port 5490`, then `POST /validate`, `POST /annotate`, `POST /render` (body = the deck, or
`{ deck, options }`; render returns the .pptx bytes), `GET /schema`, `GET /health`.

`npx bankops schema` prints the JSON Schema; `npx bankops presets` lists style and layout presets.

## Rules that matter

- **Units.** `fontSize` is hundredths of a point (`1400` = 14pt, never `14`). `gap`, `padding`,
  `margin`, `rowHeight`, `headerHeight` are inches. Colors are `"#RRGGBB"`. `slideSize` is inches
  (default 10 x 7.5).
- **Spans.** Sibling `span`s are 12-column weights and should add up to 12 (`6/6`, `8/4`, `3/3/3/3`).
  `direction: "row"` lays children left to right, `"col"` top to bottom; nest freely.
- **A node has `children` OR `content`, never both.** Put content in a child.
- **Table fields live under `data`**: `{ "type": "table", "data": { "headers", "rows", "colWidths" } }`.
  `colWidths` also sum to 12; every row has as many cells as there are headers.
- **Chart data shapes differ by type.** `bar`/`line` use `categories` + `series`; `pie` uses
  `items`; `combo` uses `bars` + `lines`; `waterfall` uses a flat `data` array; `gauge` uses
  `value`/`min`/`max`; `scatter` takes `series` whose `data` are `[x, y]` pairs. Mixing them up is
  a validation error.
- **Text.** Use `runs` for bullets and mixed formatting: a run with `bullet: true` starts a new
  paragraph on its own, and `\n` inside any run also splits paragraphs. Plain `text` is fine for
  labels and stat values.
- **Images.** Prefer `{ "type": "image", "url": "https://..." }` or a local `path`, or an
  `{ "type": "icon", "name": "mdi:chart-line" }` (Iconify). `source: "ai"` / `"search"` images
  are only resolved when the caller passes `options.resolveImage`; otherwise they are skipped
  with a warning, so do not rely on them.
- **Charts need `sharp`** for the default image rendering; `pptxChart` and
  `pptxExportMode: "nativeExcel"` need `pptxgenjs` and give editable charts. Both are optional
  peer dependencies; the render tells you which one is missing.
- Unknown keys are ignored but warned about (`Did you mean "..."?`); keys starting with `_` or
  `$` are metadata and never rendered.

## Investment-banking conventions (the `corporate` preset is the house look)

- **Every content slide** carries `sectionLabel` (eyebrow, e.g. "Company Overview"), a `title` that
  states the takeaway ("Revenue doubles by FY27E on data-center demand"), and a `footer` with the
  sources ("Source: Company filings, [Bank] analysis").
- **Every zone gets a `header` bar** (dark bar, white text) or a `subheader` (underlined label).
  Label the chart with the zone header, not with the chart's own `title`.
- **Numbers are right-aligned**: `colAlign: ["l", "r", "r", "r"]`. Units in headers ("Revenue ($m)"),
  negatives as `(120)`, `NM` for not meaningful, `summaryRows: 1` (or 2) for the Mean / Median rows.
- **One to three charts per slide**, one message each. Sort bars largest to smallest unless the
  categories are periods. `valueThousands: true` and a `$`/`%` prefix or suffix on every value label.
- **Never more than ~6 bullets in a zone, 1-2 lines each.** Body text 12-14pt; profile pages and
  dense tables 9pt (`900`); never below 8pt.
- **Stat rows for headline KPIs** (`statGrid`, or the two-zone value/label pattern in
  `reference/recipes.md`), tombstones only on credentials pages (comps are tables).
- Stick with `style: "corporate"` unless told otherwise; `minimal`, `dark`, `warm` exist, and any
  token can be overridden in a style object (`reference/format.md`).

## How to read validation output

`bankops validate` prints one line per problem and exits 1 on errors:

```
error slides[0].body.children[1].content.fontSize: fontSize is in hundredths of a point: use 1400 for 14pt, not 14
    fix: Multiply the point size by 100 (14pt -> 1400)
warning slides[0].body.children: Sibling spans sum to 11, not 12
    fix: Spans are 12-column weights; siblings normally add up to 12
```

Programmatically the same data comes back as `{ ok, errors: [{ path, message, fix }], warnings }`.
Apply the `fix` verbatim where one is given, fix every error, rerun until `ok: N slides`.
Warnings do not block rendering but nearly always mean a typo or a layout that will look off.
Render-time warnings (missing image, missing `sharp`) go to stderr / `onWarning`; the meaning of
each is in `reference/errors.md`.

## A complete worked example

```json
{
  "style": "corporate",
  "slides": [{
    "sectionLabel": "Market Overview",
    "title": "Data-center capex accelerates through FY27E",
    "footer": "Source: Company filings, industry reports, [Bank] analysis",
    "body": { "direction": "row", "children": [
      { "span": 7, "header": "Hyperscaler capex ($bn)", "content": { "type": "chart", "chartType": "bar",
        "categories": ["FY24", "FY25", "FY26E", "FY27E"],
        "series": [{ "name": "Capex", "data": [212, 305, 380, 440] }],
        "showBarValues": true, "valuePrefix": "$", "valueThousands": true, "showYAxis": false } },
      { "span": 5, "direction": "col", "children": [
        { "span": 4, "content": { "type": "statGrid", "items": [
          { "value": "$440bn", "label": "FY27E capex" }, { "value": "27%", "label": "FY24-27E CAGR" } ] } },
        { "span": 8, "header": "Key themes", "content": { "type": "text", "fontSize": 1200, "runs": [
          { "text": "Accelerator demand outpaces supply through 2026", "bullet": true },
          { "text": "Advanced packaging is the binding constraint", "bullet": true },
          { "text": "Power availability now gates new campus starts", "bullet": true } ] } }
      ] }
    ] }
  }]
}
```

`bankops validate deck.json` → `ok: 1 slide`. `bankops annotate deck.json` reports
`_titleMaxChars: 41` on the slide, `_headerMaxChars: 36` on the "Key themes" bar and
`_maxChars: 672` for its 12pt bullets (three ~45-character bullets use a fifth of it).
`bankops render deck.json deck.pptx` → `deck.pptx: 1 slide`.

## Reference

- `reference/format.md` — every field: deck, slide, grid node, the 12 content types, 7 chart
  types and their data shapes, `pptxChart`, style presets and the override object, layout presets
- `reference/recipes.md` — IB layouts with exact spans and font sizes: company profiles (private
  1/2/3-page, public 1/4-page), credentials and tombstones, investment highlights, industry
  overview, competitive analysis, M&A and fundraising pitches, chart rules, fixing overflow
- `reference/errors.md` — every validation error and warning, every render warning, what to do

## License

bankops is free for noncommercial use (PolyForm Noncommercial 1.0.0). Commercial use, including
internal use at a bank, fund or advisory firm, needs a license: https://bankops.ai/license
