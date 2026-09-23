# Deck JSON format reference

Every field below is taken from the zod schema (`DeckSchema`) that `bankops validate` enforces.
Unknown keys are stripped with a warning; keys beginning with `_` or `$` are metadata and are
neither validated nor rendered. Units: `fontSize` in hundredths of a point (`1400` = 14pt),
lengths in inches, colors `"#RRGGBB"` (3-digit `#RGB` is accepted).

## 1. Deck

```jsonc
{
  "title": "Project Atlas — Management Presentation",   // optional, metadata only
  "style": "corporate",                                  // preset name or a style override object (§6)
  "slideSize": { "width": 10, "height": 7.5 },           // inches, optional (default 10 x 7.5)
  "slides": [ /* 1 to 200 slides */ ]
}
```

| Field | Type | Notes |
|---|---|---|
| `title` | string | optional |
| `style` | `"corporate"` \| `"minimal"` \| `"dark"` \| `"warm"` \| object | default `corporate`; object = overrides of §6 |
| `slideSize` | `{ width?, height? }` | positive numbers, inches; 13.333 x 7.5 for 16:9 |
| `slides` | slide[] | required, 1–200 |

## 2. Slide

```jsonc
{
  "id": "s1",                          // optional stable id
  "type": "grid",                      // optional; the only slide type
  "sectionLabel": "COMPANY OVERVIEW",  // eyebrow above the title (0.30" tall)
  "title": "Takeaway-style title",     // 0.55" tall, 28pt bold by default
  "subtitle": "Optional subtitle",
  "footer": "Source: Company filings", // 0.35" at the bottom, 8pt
  "margin": 0.5,                       // slide margin override, inches
  "body": { /* grid node */ }
}
```

Body height on a 7.5" slide with 0.5" margins = 6.5" minus 0.3 (sectionLabel) minus 0.65 (title)
minus 0.4 (footer) ≈ 5.15" when all three are present. Omitting the section label or footer gives
the body that space back.

## 3. Grid node

A node is a **container** (`children`) or a **leaf** (`content`), never both.

| Field | Type | Default | Meaning |
|---|---|---|---|
| `id` | string | — | stable id |
| `span` | number 0.5–12 | equal share | weight in the parent's 12-column grid (row: width, col: height) |
| `direction` | `"row"` \| `"col"` | `"row"` | how children are laid out |
| `gap` | inches | 0.15 | space between children |
| `padding`, `paddingTop/Right/Bottom/Left` | inches | 0 | inset applied to the children area |
| `margin`, `marginTop/Right/Bottom/Left` | inches | 0 | inset applied to the node's own box (background, border) |
| `header` | string | — | dark header bar (0.35") with white text above the content |
| `subheader` | string | — | bold label with an underline (0.22") above the content |
| `subfooter` | string | — | small italic note (0.22") below the content |
| `background` | color | — | zone fill |
| `backgroundOpacity` | 0–100 | 100 | fill opacity |
| `border` | color \| `{ color, width?, style? }` | — | all four sides; `width` in pt, `style` one of `solid`, `dash`, `dot`, `dashDot`, `lgDash`, `lgDashDot` |
| `borderTop/Right/Bottom/Left` | `{ color, width?, style? }` | — | per-side border (object form only) |
| `shadow` | `{ color?, alpha? 0–1, blur?, dist?, dir?, size? }` | — | drop shadow; blur/dist in pt, dir in degrees, size in % |
| `childShape` | `"chevron"` \| `"pyramid"` | — | draws each child as a chevron (process flows) or pyramid tier |
| `childShapeColor` | color | — | fill for the child shapes |
| `children` | node[] | — | at most 50 |
| `content` | content | — | one of the 12 content types (§4) |
| `$component` | string | — | expands a node from `options.components[name]` at render time; render fails if the name is unknown |
| `_sources` | `{ url?, api?, label?, date? }[]` | — | provenance metadata, not rendered |

Spans that do not sum to 12 are only a warning: the engine normalizes weights, but the layout will
not be what you meant. Children without `span` share what the explicit spans leave over.

Layout presets you can copy (`SlideLayout.preset(name)` returns the body skeleton):

| Name | Body |
|---|---|
| `twoColumn` | row: 6 / 6 |
| `threeColumn` | row: 4 / 4 / 4 |
| `quadrant` | row: two `col` children of 6, each split 6 / 6 |
| `sidebar` | row: 8 / 4 |
| `topBottom` | col: 4 / 8 |
| `wideTop` | col: 3 / 9, the 9 split row 6 / 6 |
| `dashboard` | col: 3 (row 3/3/3/3) / 9 (row: col 6 (6/6) / 6) |

## 4. Content types

`content.type` is one of `text`, `statGrid`, `cardGrid`, `table`, `profile`, `image`, `icon`,
`chart`, `pptxChart`, `line`, `timeline`.

### 4.1 `text`

```jsonc
{ "type": "text",
  "text": "Plain text.\nSecond paragraph.",     // used when runs is absent
  "runs": [ /* rich runs, take precedence */ ],
  "font": "Calibri", "fontSize": 1200, "color": "#333333",
  "bold": false, "italic": false, "underline": false,
  "textTransform": "uppercase",                  // only value: uppercase
  "align": "l",                                  // l | ctr | r
  "anchor": "t",                                 // t | ctr | b   (vertical)
  "lineSpacing": 1.0,                            // 0.5–3
  "bulletStyle": { "char": "•", "color": "#4472C4", "indent": 0.375, "marginLeft": 0.375, "spaceBefore": 2 },
  "paddingTop": 0.05, "paddingRight": 0.1, "paddingBottom": 0.05, "paddingLeft": 0.1 }   // text insets, inches
```

Run: `{ text, bold?, italic?, underline?, color?, fontSize?, font?, bullet?, bulletLevel? }`.
`bulletLevel` is 0–4 (0 = `•`, 1 = `–` sub-bullet).

Paragraph rules:
- `\n` inside `text` or inside any run's `text` starts a new paragraph.
- A run with `bullet: true` starts a new paragraph by itself when the current paragraph already
  has text, so one run per bullet needs no `\n`.
- The first run of a paragraph decides whether it is bulleted and at which level; later runs in
  the same paragraph (bold fragments, colored numbers) inherit it.

```jsonc
{ "type": "text", "fontSize": 1200, "runs": [
  { "text": "Revenue ", "bullet": true }, { "text": "+42% YoY", "bold": true, "color": "#16A34A" }, { "text": " on data-center demand" },
  { "text": "Gross margin expands 300bp", "bullet": true },
  { "text": "Packaging capacity is the constraint", "bullet": true, "bulletLevel": 1 } ] }
```

Defaults from the style: body 14pt Calibri, bullets 13pt; `align` defaults to the style's `body.align`.

### 4.2 `statGrid`

```jsonc
{ "type": "statGrid", "anchor": "ctr",
  "items": [ { "value": "$1.2bn", "label": "Revenue", "sublabel": "FY25A" },
             { "value": "34%",    "label": "EBITDA margin" } ] }   // 1+ items, laid out in a row
```

Values render 16pt bold in the accent color, labels at 85% of body size, sublabels at footnote size.

### 4.3 `cardGrid`

```jsonc
{ "type": "cardGrid", "items": [
  { "title": "Cloud platform", "lines": ["Multi-tenant", "99.9% SLA"] },
  { "title": "Mobile SDK",     "lines": ["iOS and Android"] } ] }   // 1+ items; bordered cards, accent titles
```

### 4.4 `table`

All fields live under `data`.

```jsonc
{ "type": "table", "data": {
  "headers": ["Company", "EV ($m)", "EV/Revenue", "EV/EBITDA"],   // optional; omit or [] for no header row
  "rows": [ ["Peer A", "4,200", "6.1x", "18.2x"],
            ["Peer B", "2,850", "4.8x", "15.0x"],
            ["Median", "",      "5.5x", "16.6x"] ],
  "colWidths": [5, 3, 2, 2],           // 12-column spans, sum to 12; default equal
  "colAlign": ["l", "r", "r", "r"],     // per column: l | ctr | r   (or align: { "1": "r", ... })
  "rowHeight": 0.4, "headerHeight": 0.4, // inches; one text line fits in ~0.3" at 9–11pt
  "fontSize": 1000, "headerFontSize": 1000,
  "summaryRows": 1,                     // bottom N rows styled bold on the header fill (Mean / Median)
  "verticalHeaders": false,             // true: first column styled as a header column
  "overflowWrap": "normal" } }          // normal | anywhere | break-word
```

A cell is a string or an image cell `{ "type": "image", "url"?, "path"?, "source"?, "searchTerm"?, "prompt"? }`
(logos in a comps or buyers table). Headers must be strings. The lint warns when a row's cell
count differs from the header count and when `colWidths` do not sum to 12.

### 4.5 `profile`

```jsonc
{ "type": "profile", "name": "[Executive name]", "items": ["Chief Executive Officer", "Joined 2019", "Previously CFO, [Company]"] }
```

A name line with bullet items beneath. For a management team page, one profile per zone in a
2 x 3 or 3 x 2 grid.

### 4.6 `image`

```jsonc
{ "type": "image",
  "url": "https://example.com/logo.png",   // fetched at render (options.network, default true)
  "path": "images/logo.png",               // local, relative to options.basePath (CLI: the deck's directory)
  "source": "search", "searchTerm": "[Company] logo",   // or source: "ai" + prompt — handed to options.resolveImage;
  "prompt": "...",                                        //   skipped with a warning when no resolver is passed
  "objectFit": "contain",                  // contain (logos) | cover (photos, crops) | fill (stretch)
  "anchor": "ctr",                         // t | ctr | b
  "width": "80%", "height": "80%",         // number (px) or "N%" of the zone
  "imagePadding": 0.1,                     // inset before the fit, inches
  "crop": { "l": 0, "t": 0, "r": 0, "b": 0 },   // percent per side
  "ext": "png" }
```

Give exactly one of `url`, `path`, or `source`+`searchTerm`/`prompt`. PNG and JPEG embed directly;
SVG and WebP are converted with `sharp` (skipped with a warning when it is missing).

### 4.7 `icon`

```jsonc
{ "type": "icon", "name": "mdi:chart-line",   // required, an Iconify id "prefix:name"
  "svg": "<svg ...>",                          // optional inline markup; avoids the network fetch
  "color": "#4472C4", "size": 60,              // size = % of the zone, 1–100 (default 80)
  "align": "middle",                           // top | middle | bottom
  "objectFit": "contain" }                     // contain | cover | fill
```

Needs network access to api.iconify.design (or `svg`) and `sharp`. Good for 2 x 2 "market trend"
tiles: an icon zone of span 2 beside a text zone of span 10.

### 4.8 `chart` (rendered image by default, native when asked)

Common fields: `chartType` (required), `title`, `showLegend`, `legendPosition` (`bottom` | `left` |
`right` | `top`), `yMin`/`yMax`, `xMin`/`xMax`, `gridTop/Right/Bottom/Left` (px inside the zone),
`showXAxis`, `showYAxis`, `showSplitLine`, `xAxisLabel`, `xAxisLabelPosition` (`end` | `center`),
`yAxisLabel`, `yAxisLabelPosition` (`top` | `side`), value-label formatting `valuePrefix`,
`valueSuffix`, `valueDecimals` (0–6), `valueThousands`, `negativeFormat` (`minus` | `parens`),
`pptxExportMode` (`image` | `nativeExcel`).

Data shape per type (validation enforces these):

| `chartType` | Required data | Type-specific fields |
|---|---|---|
| `bar` | `categories: string[]`, `series: [{ name?, data: number[] }]` | `horizontal`, `stacked`, `stackTotal: number[]`, `showBarValues`, `showSeriesLabels`, `showTotalValues`, `barColor`, `categoryColors: { "0": "#..." }`, `barBorderColor`, `barBorderWidth` |
| `line` | `categories`, `series` | `smooth`, `area`, `showLineValues`, `lineColor`, `lineWidth`, `pointShape` (`circle` \| `rect` \| `triangle` \| `diamond`), `pointSize`, `pointColor`, `pointBorderColor`, `pointBorderWidth` |
| `pie` | `items: [{ name, value }]` | `doughnut`, `innerRadius` `"40%"`, `outerRadius` `"70%"`, `centerX`/`centerY` `"50%"`, `startAngle`, `roseType` (`radius` \| `area`), `showLabel`, `labelPosition` (`outside` \| `inside` \| `center` \| `none`), `showPieValues`, `pieValuePosition`, `pieBorderColor`, `pieBorderWidth`, `padAngle`, `sliceColors: { "0": "#..." }` |
| `combo` | `categories`, `bars: [{ name?, data }]` and/or `lines: [{ name?, data }]` | `dualAxis`, `y2Min`/`y2Max`, `y2ValuePrefix`, `y2ValueSuffix`, `y2ValueDecimals`, `y2ValueThousands`, `y2NegativeFormat`, `showSecondaryYAxis`, `stacked`, `smooth`, `showBarValues`, `showLineValues` |
| `waterfall` | `categories`, `data: number[]` (2+; first and last are totals, the middle are deltas) | value formatting fields |
| `scatter` | `series: [{ name?, data: [[x, y], ...] }]` — each point is an `[x, y]` pair | point fields, `xMin`/`xMax`, `yMin`/`yMax` |
| `gauge` | `value` | `min` (0), `max` (100), `label`, `format` (`"{value}%"`) |

```jsonc
{ "type": "chart", "chartType": "combo", "categories": ["FY23", "FY24", "FY25"],
  "bars": [{ "name": "Revenue ($m)", "data": [410, 520, 680] }],
  "lines": [{ "name": "EBITDA margin", "data": [22, 25, 28] }],
  "dualAxis": true, "valuePrefix": "$", "valueThousands": true, "y2ValueSuffix": "%", "gridBottom": 45 }

{ "type": "chart", "chartType": "waterfall", "categories": ["FY24", "Volume", "Price", "FX", "FY25"],
  "data": [680, 85, 45, -30, 780], "valuePrefix": "$", "valueSuffix": "m", "negativeFormat": "parens" }

{ "type": "chart", "chartType": "pie", "items": [{ "name": "Software", "value": 65 }, { "name": "Services", "value": 35 }],
  "doughnut": true, "showLegend": true, "labelPosition": "none", "centerY": "45%" }
```

Rendering: by default the chart is drawn (ECharts) at the zone's exact size and embedded as a PNG,
which needs `sharp`. `pptxExportMode: "nativeExcel"` converts `bar`, `line`, `pie`, `combo` and
`scatter` charts to native PowerPoint charts (needs `pptxgenjs`); `waterfall` and `gauge` stay images.

### 4.9 `pptxChart` (native, editable, Excel-backed)

Prefer this when the reader will restyle or update the numbers in PowerPoint, or when `sharp` is
unavailable. Series carry their own labels.

```jsonc
{ "type": "pptxChart", "chartType": "bar",            // bar | line | pie | scatter | combo
  "series": [ { "name": "Revenue", "labels": ["FY23", "FY24", "FY25"], "values": [410, 520, 680] } ],
  "barDir": "col", "barGrouping": "clustered",         // col | bar; clustered | stacked | percentStacked
  "showValue": true, "dataLabelFormatCode": "$#,##0", "valAxisLabelFormatCode": "$#,##0",
  "showLegend": false, "legendPos": "b",               // b | t | l | r | tr
  "chartColors": ["#4472C4", "#A5A5A5"], "gapWidthPct": 80,
  "catAxisTitle": "", "valAxisTitle": "$m", "valAxisMinVal": 0, "valAxisMaxVal": 800,
  "catAxisHidden": false, "valAxisHidden": true,
  "valGridLine": { "style": "none" }, "catGridLine": { "style": "none" },   // style: solid | dash | dot | none, color, size
  "dataLabelPosition": "outEnd",   // b | bestFit | ctr | inBase | inEnd | l | outEnd | r | t
  "dataLabelFontSize": 9, "dataLabelColor": "#333333", "showCatName": false, "showSerName": false, "showPercent": false,
  "fontSize": 10, "legendFontSize": 9, "titleFontSize": 12, "title": "", "showTitle": false,
  "dataBorder": { "type": "solid", "pt": 0.5, "color": "#FFFFFF" },
  "lineSmooth": false, "lineDataSymbol": "circle", "lineDataSymbolSize": 6,   // circle | dash | diamond | dot | none | square | triangle
  "lineDataSymbolLineColor": "#4472C4", "lineDataSymbolLineSize": 1,
  "doughnut": false, "barFillColor": "#4472C4" }
```

Combo: `chartType: "combo"` with `comboCharts: [{ chartType: "bar" | "line" | "area", series: [...],
options?: { secondaryValAxis?: boolean, barGrouping?: string } }]`, plus `secondaryValAxisLabelFormatCode`,
`secondaryValAxisMinVal`, `secondaryValAxisMaxVal`. Font sizes here are plain points (`10`), not
hundredths.

### 4.10 `callout`

A key-takeaway box: a tinted panel with an accent bar on the left, an optional bold title in the
accent color, and text or runs. Use it for "bottom line", "so what", "key risk" and similar boxes,
usually in a span-4 or span-5 column beside a table or chart.

```jsonc
{ "type": "callout", "title": "Bottom line", "text": "Risk scales with what an agent can reach, change and persist.",
  "fontSize": 1200, "align": "l", "anchor": "ctr" }
```

| Field | Type | Notes |
|---|---|---|
| `title` | string | optional, bold, accent color |
| `text` | string | plain text (`\n` splits paragraphs) |
| `runs` | run[] | rich runs as in `text`; takes precedence over `text` |
| `font`, `fontSize`, `color` | | overrides; `fontSize` in hundredths of a point |
| `background`, `accentColor` | hex | override the panel fill and the bar/title color |
| `align`, `anchor` | | `"l" \| "ctr" \| "r"`, `"t" \| "ctr" \| "b"` (default centered vertically) |

Budgets: `_maxChars` as for `text`, minus the bar and padding, minus one line when a title is set.

### 4.11 `line`

```jsonc
{ "type": "line", "color": "#BFBFBF", "width": 9525 }   // horizontal rule; width in EMU (9525 = 0.75pt)
```

### 4.12 `timeline`

```jsonc
{ "type": "timeline", "title": "Company history",
  "events": [ { "date": "2019-03-01", "name": "Founded", "detail": "Incorporated in Delaware" },
              { "date": "2021-06-15", "name": "Series A", "detail": "$15m led by [Investor]" },
              { "date": "2024-09-01", "name": "Series C", "detail": "$120m at $1.1bn" } ],   // 1+ events, ISO dates
  "placement": "split",            // split | top | bottom
  "topCount": 2,                   // events above the line when split (default ceil(n/2))
  "dateFormat": "month-year",      // month | year | month-year | quarter | quarter-year
  "datePlacement": "auto",         // auto | top | bottom
  "textMode": "plain",             // plain | bullet
  "boxHeightPct": 28, "boxGap": 0.1, "boxOffset": 0.3, "boxWidth": 0, "boxHeight": 0,   // boxWidth/Height in EMU
  "boxBackground": "#F0F4F8", "boxBorderColor": "#CBD5E1", "boxBorderWidth": 0.5,
  "nameFont": "Calibri", "nameFontSize": 10, "nameBold": true, "nameItalic": false, "nameUnderline": false, "nameColor": "#1F2937", "nameAlign": "ctr",
  "detailFont": "Calibri", "detailFontSize": 8, "detailBold": false, "detailItalic": false, "detailUnderline": false, "detailColor": "#6B7280",
  "dateFont": "Calibri", "dateFontSize": 7, "dateColor": "#4472C4", "dateBold": true, "dateItalic": false,
  "lineColor": "#4472C4", "lineWidth": 25400, "nodeColor": "#FFFFFF", "nodeBorderColor": "#4472C4", "nodeBorderWidth": 19050,
  "connectorColor": "#9DB6DA", "connectorWidth": 9525 }
```

Timeline font sizes are plain points; line, node and connector widths are EMU (12700 = 1pt).
Four to eight events read well on a full-width zone.

## 5. Text budgets (`bankops annotate`)

`annotateDeck(deck, style?)` returns a deep copy with capacity fields added. They are estimates
from font metrics and the layout geometry (zone width and height after headers, gaps, padding).

| Where | Field | Meaning |
|---|---|---|
| slide | `_titleMaxChars`, `_titleCharsPerLine`, `_titleMaxLines` | title capacity (28pt bold, 0.62" tall) |
| slide | `_sectionLabelMaxChars`, `_footerMaxChars` | single-line capacities |
| node | `_headerMaxChars`, `_subheaderMaxChars` | header bar / subheader are single lines |
| text content | `_charsPerLine`, `_maxLines`, `_maxChars` | at the content's `fontSize` (or 14pt body default); bullets consume indent, so keep bullet zones ~15% under |
| table `data` | `_cellMaxChars[]`, `_headersMaxChars[]` | per column, at `fontSize` / `headerFontSize`, for one `rowHeight` / `headerHeight` |
| table `data` | `_maxRows` | data rows that fit below the header in the zone |

Compare each string's length against its budget before rendering; the annotated deck renders as-is
because `_` keys are ignored.

## 6. Style presets and the override object

Presets: `corporate` (white, slate titles, `#4472C4` accent, `#164556` header bars, Arial Bold /
Calibri — the IB house look), `minimal` (Helvetica Neue, `#2980b9`), `dark` (`#1a1a2e` background,
`#E94560` accent, Arial), `warm` (Georgia / Garamond, `#D57F5B`).

`style` may instead be an object; any subset of the defaults below overrides them (two levels
deep). Units here are human: points and inches, not hundredths.

```jsonc
{
  "slide":        { "background": "#FFFFFF", "margin": 0.5 },
  "colors":       { "primary": "#333333", "secondary": "#666666", "accent": "#4472C4", "darkFill": "#164556",
                    "border": "#BFBFBF", "headerBarText": "#FFFFFF", "cardFill": "#FFFFFF" },
  "fonts":        { "title": "Arial Bold", "body": "Calibri", "label": "Calibri", "dividerNumber": "Calibri Light" },
  "title":        { "size": 28, "bold": true, "color": "#405363", "y": 0.9 },
  "sectionLabel": { "size": 12, "bold": true, "italic": false, "color": null, "y": 0.43 },        // null = colors.accent
  "body":         { "size": 14, "color": null, "align": "l", "fontFamily": null, "lineSpacing": null, "bold": null, "italic": null },
  "bullet":       { "char": "•", "indent": 0.375, "marginLeft": 0.375, "size": 13, "font": null, "bulletFont": null,
                    "bulletColor": null, "color": null, "fontFamily": null, "lineSpacing": null },
  "footnote":     { "size": 8, "color": null, "y": null },
  "header":       { "size": null, "color": null, "background": null, "align": "ctr", "borderRadius": null,   // 0 = square bars
                    "fontFamily": null, "bold": null, "italic": null },
  "subheader":    { "size": null, "bold": true, "color": null, "align": "ctr", "lineColor": null, "lineWidth": 0.75, "fontFamily": null, "italic": null },
  "subfooter":    { "size": null, "italic": true, "color": null, "fontFamily": null },
  "table":        { "headerFill": null,            // null = colors.darkFill (same navy as zone header bars)
                    "headerTextColor": "#FFFFFF", "headerSize": 12, "headerBold": true, "headerFont": null,
                    "headerRuleColor": null, "headerRuleWidth": 1.25,   // rule under the header; null color = colors.accent
                    "borderMode": "horizontal",   // "horizontal" = hairlines between rows only (the IB look) | "grid" | "none"
                    "verticalHeaderFill": null, "verticalHeaderTextColor": null, "verticalHeaderSize": null, "verticalHeaderBold": null, "verticalHeaderFont": null,
                    "bodySize": 11, "bodyTextColor": null, "bodyFont": null,
                    "alternateRows": true, "evenRowFill": "#F4F6F9", "oddRowFill": "#FFFFFF", "borderColor": "#D9DEE5", "borderWidth": 0.5,
                    "summaryRowFill": "#E9EEF6", "summaryRowTextColor": null, "summaryRowSize": null, "summaryRowBold": true, "summaryRowFont": null,
                    "summaryRowBorderTopColor": null, "summaryRowBorderTopWidth": 1, "summaryRowBorderBottomColor": null, "summaryRowBorderBottomWidth": 1,
                    "summaryRowBorderLeftColor": null, "summaryRowBorderLeftWidth": 0, "summaryRowBorderRightColor": null, "summaryRowBorderRightWidth": 0 },
  "statGrid":     { "valueFontFamily": null, "valueSize": 16, "valueColor": null, "valueBold": true, "valueItalic": false,
                    "labelFontFamily": null, "labelSize": null, "labelColor": null, "labelBold": false, "labelItalic": false,
                    "sublabelFontFamily": null, "sublabelSize": null, "sublabelColor": null, "sublabelBold": false, "sublabelItalic": false },
  "cardGrid":     { "titleFontFamily": null, "titleSize": null, "titleColor": null, "titleBold": true, "titleItalic": false,
                    "lineFontFamily": null, "lineSize": null, "lineColor": null, "lineBold": false, "lineItalic": false,
                    "lineBullet": false, "lineBulletIndent": 0.05 },
  "chart":        { "colors": null, "fontFamily": null, "fontSize": null, "barBorderColor": null, "barBorderWidth": null,
                    "bar": null,    // { showXAxis, showYAxis, showSplitLine, showSeriesLabels, gridLeft, gridRight, gridTop, gridBottom }
                    "pie": null,    // { borderColor, borderWidth }
                    "combo": null },// { showPrimaryYAxis, showSecondaryYAxis, showSplitLine, showBarValues, showLineValues, barValueColor, lineValueColor }
  "divider":      { "numberSize": 115, "numberColor": null, "numberX": 0.5, "numberY": 2.4, "titleSize": 24, "titleColor": null, "titleY": null },
  "spacing":      { "columnGap": 0.25, "bulletSpacing": 0.125, "cardGap": 0.08, "cardBorderWidth": 0.75, "cardPaddingLeft": 0.075,
                    "cardPaddingTop": 0.05, "subheaderLineWidth": 0.75, "headerBarHeight": 0.35, "contentGap": 0.1, "zoneGap": 0.15 },
  "titleSlide":   { "titleSize": 36, "subtitleSize": 18 }
}
```

A brand: `{ "colors": { "accent": "#8A1538", "darkFill": "#1F2A44" }, "fonts": { "title": "Georgia", "body": "Arial" },
"chart": { "colors": ["#8A1538", "#1F2A44", "#A0A0A0"] } }`. With the API, `options.themeTokens`
supplies the same object as a base layer under `deck.style`; the CLI takes `--theme tokens.json`
and `--style corporate|minimal|dark|warm|'{...}'`.

## 7. Render options (API / CLI / HTTP)

| Option | CLI | HTTP `options` | Meaning |
|---|---|---|---|
| `style` | `--style` | yes | overrides `deck.style` |
| `themeTokens` | `--theme file` | yes | brand tokens under `deck.style` |
| `basePath` | `--base-path dir` | — | where relative image `path`s resolve (CLI default: the deck's directory) |
| `network` | `--no-network` | yes | fetch image URLs and Iconify icons (default true) |
| `timeoutMs` | — | — | per fetch, default 15000 |
| `resolveImage` | — | server-side only | `(req) => Promise<Buffer \| null>` for `source: "ai"` / `"search"` images and logos |
| `components` | — | yes | named nodes or factories for `$component` |
| `referencePptx` | — | — | a .pptx whose theme, fonts and colors are reused |
| `zoneBorders` | `--zone-borders` | yes | debug borders around every zone |
| `validate` | — | yes | default true; invalid decks throw `DeckValidationError` |
| `onWarning` | stderr | `x-bankops-warnings` header | non-fatal problems |

`renderDeck` returns `{ buffer, slides, warnings, lint }`; `renderPptx` returns the buffer;
`renderPptxToFile(deck, path)` and `renderFile(jsonPath, outPath)` write the file.
