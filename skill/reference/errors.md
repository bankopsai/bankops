# Errors and warnings

`bankops validate` / `validateDeck()` return `{ ok, errors: [{ path, message, fix? }], warnings }`.
Errors block rendering; warnings do not but almost always mean a typo or a layout that will look
wrong. `path` is a JSON path into the deck (`slides[2].body.children[0].content.fontSize`). When
`fix` is present, apply it as written. The CLI prints `error <path>: <message>` / `fix: ...` lines
and exits 1; the HTTP service answers `422 { error: "invalid deck", errors, warnings }`;
`renderPptx` throws `DeckValidationError` with the same `errors` and `warnings` arrays.

## 1. Validation errors

| Message (path) | Meaning | What to do |
|---|---|---|
| `Deck must be a JSON object` (`deck`) | Top level is an array, string or null. | Send `{ "slides": [ ... ] }`. |
| `A deck needs at least one slide` (`slides`) | `slides` missing or empty. | Add `slides: [{ title: "...", body: { ... } }]`. |
| `At most 200 slides per deck` | | Split the deck. |
| `Invalid input: expected array, received ...` / `expected object` / `expected string` / `expected number` / `expected boolean` | Wrong JSON type at that path (a common one: `"span": "6"`, `"data": [...]`). | Use the type from `format.md`; numbers unquoted, table fields under `data: { ... }`. |
| `Invalid discriminator value. Expected 'text' \| 'statGrid' \| ...` (`...content.type`) + `fix: content.type must be one of: ...` | Unknown content type (`bulletList`, `stat`, `diagram`, `html` ...). | Use one of `text`, `statGrid`, `cardGrid`, `table`, `profile`, `image`, `icon`, `chart`, `pptxChart`, `line`, `timeline`. Bullets are `text` with `runs`. |
| `Invalid option: expected one of ...` (any enum) | Value not in the enum, e.g. `align: "center"`, `direction: "column"`, `chartType: "column"`. | Enums are exact: `l` / `ctr` / `r`, `t` / `ctr` / `b`, `row` / `col`, the 7 chart types. |
| `Invalid input` (`style`) | `style` is neither a preset name nor an object (`"classic"`, `"Corporate"`). | `corporate`, `minimal`, `dark`, `warm` (lowercase) or an override object. |
| `Invalid input: expected object, received undefined` (`...content.data`) plus warnings `Unknown key "headers"` / `"rows"` | Table fields placed directly on the content instead of under `data`. | `{ "type": "table", "data": { "headers": [...], "rows": [...] } }`. |
| `fontSize is in hundredths of a point: use 1400 for 14pt, not 14` | `fontSize` below 100. | Multiply the point size by 100. |
| `fontSize must be an integer in hundredths of a point (1400 = 14pt)` | Fractional `fontSize`. | Round to an integer (`1250`). |
| `fontSize above 400pt is not supported` | `fontSize` > 40000. | You probably multiplied twice. |
| `Colors are hex strings like #4472C4` | Not `#RRGGBB` / `#RGB` (named colors, `rgb()`, missing `#`). | Use `"#RRGGBB"`. Applies to every color field including `border`, `shadow.color`, `chartColors`, `categoryColors` values. |
| `A node is either a container (children) or a leaf (content), never both. Move the content into a child node.` (`...content`) | A node has both `children` and `content`. | Wrap the content: `children: [{ content: {...} }, ...]` or move the header text into `header`. |
| `A node may have at most 50 children` | | Split the zone or paginate. |
| `Too small: expected number to be >=0.5` / `Too big: expected number to be <=12` (`span`) | `span` outside 0.5–12. | Spans are 12-column weights. |
| `Too small: expected number to be >=0` / `<=20` (`gap`, `padding`, `margin`, `rowHeight`, `headerHeight`, `imagePadding`, `indent` ...) | Lengths are inches. | You likely wrote points or EMU; 0.15" is a normal gap. |
| `Too small` / `Too big` on `backgroundOpacity` (0–100), `shadow.alpha` (0–1), `lineSpacing` (0.5–3), `bulletLevel` (0–4), `icon.size` (1–100), `valueDecimals` (0–6) | Out of range. | Use the range shown. |
| `Too small: expected number to be >0` (`colWidths[i]`) | Zero or negative column width. | Positive spans that sum to 12. |
| `Too small: expected array to have >=1 items` (`statGrid.items`, `cardGrid.items`, `timeline.events`) | Empty list. | Add at least one item, or remove the zone. |
| `Icon names are Iconify ids like "mdi:home"` (`icon.name`) | Not `prefix:name` in lowercase letters, digits and dashes. | e.g. `mdi:chart-line`, `lucide:building-2`. `name` is required even when `svg` is given. |
| `pie charts use items: [{ name, value }], not series` (`...content.items`) | Pie chart without `items`. | Convert `categories` + `series[0].data` into `items`. |
| `combo charts use bars: [...] and lines: [...] arrays plus categories` (`...content.bars`) | Combo without `bars` or `lines`. | Split the series into `bars` and `lines`. |
| `waterfall charts use a flat data: number[] (first and last are totals, middle values are deltas) plus categories` (`...content.data`) | Waterfall without `data` of 2+ numbers. | `data: [start, +delta, -delta, ..., end]`. |
| `gauge charts need value (and optionally min, max, label)` (`...content.value`) | Gauge without a numeric `value`. | Add `value`. |
| `scatter charts use series: [{ name, data: [[x, y], ...] }]` (`...content.series`) | Scatter without `series`. | Each series' `data` is an array of `[x, y]` pairs — but see the next row. |
| `bar charts use categories: string[] and series: [{ name, data: number[] }]` (also `line`) | Bar/line without `series`. | Add `categories` and `series`; `items` is for pie only. |
| `Invalid input: expected array, received ...` (`pptxChart.series`) | `pptxChart` needs `series: [{ name, labels, values }]`. | Each series carries its own `labels`; there is no `categories` field. |
| `Invalid input` on `slideSize.width` / `.height` | Not a positive number. | Inches, e.g. `13.333` x `7.5`. |

Zod reports every failing field, so one structural mistake (a table with `rows` outside `data`)
can produce several lines at the same path prefix; fix the outermost one first and rerun.

## 2. Validation warnings

| Message | Meaning | What to do |
|---|---|---|
| `Unknown key "x" is ignored` + `fix: Did you mean "y"?` | A key that is not in the schema at that level (deck, slide, node, or content of that type). Typos (`fontsize`, `colWidth`, `subHeader`) and keys from other content types (`items` on a bar chart, `text` on a table) show up here. | Rename to the suggestion, move the key to the right level, or delete it. The key is dropped at render, so the intended formatting is missing. |
| `Sibling spans sum to N, not 12` (`...children`) | All siblings have explicit spans and they do not add up to 12. | Rebalance (`7/5`, `3/3/3/3`); the engine normalizes weights, so the layout still renders but not as intended. |
| `colWidths sum to N, not 12` (`...content.data.colWidths`) | Table column spans off. | Make them sum to 12. |
| `Row has N cells but there are M headers` (`...content.data.rows[i]`) | Ragged table. | Pad or trim that row; a missing cell shifts every value to the left. |
| `Nesting deeper than 16 levels` + `fix: Flatten the layout` | Excessive nesting. | Flatten; nothing below that depth is linted. |

Keys that begin with `_` or `$` (`_maxChars`, `_sources`, `$component`) never trigger the
unknown-key warning.

## 3. Render-time warnings (non-fatal)

Emitted through `options.onWarning`, collected in `renderDeck().warnings`, printed by the CLI as
`warning: ...` on stderr, counted in the HTTP `x-bankops-warnings` header. The zone renders empty.

| Warning | Meaning | What to do |
|---|---|---|
| `<path>: image needs generation (<prompt>); pass options.resolveImage or supply url/path` / `image needs search (<term>)` | A `source: "ai"` / `"search"` image (or table logo cell) and no resolver was passed. The engine makes no AI or search calls itself. | Supply `url` or `path`, use an `icon`, or drop the image; in your own code pass `resolveImage`. |
| `<path>: resolver returned nothing for prompt "..."` / `search "..."` | Your resolver returned `null`. | Fix the resolver or fall back to `url`/`path`. |
| `image not found: <absolute path>` | Relative `path` did not resolve against `basePath` (CLI: the deck's directory). | Fix the path or pass `--base-path`. |
| `network disabled; skipped image URL <url>` | `--no-network` / `network: false`. | Use `path`, or enable network. |
| `failed to fetch image (<status>): <url>` / `failed to fetch image: <url> (<error>)` | HTTP error, timeout (`timeoutMs`, default 15s) or DNS failure. | Check the URL; download it and use `path`. |
| `sharp is not installed; skipped SVG image <hint> (npm install sharp)` / `... WebP image ...` | SVG and WebP need `sharp` to become PNG. | `npm install sharp`, or supply a PNG/JPEG. |
| `SVG to PNG failed for <hint>: ...` / `WebP to PNG failed ...` | `sharp` could not rasterize the file. | Convert it yourself; supply PNG. |
| `<path>: network disabled; icon "<name>" needs inline svg` | Icons are fetched from api.iconify.design. | Enable network or paste the SVG in `svg`. |
| `<path>: icon name "<name>" is not an Iconify id like "mdi:home"` | Malformed id. | `prefix:name`. |
| `<path>: icon "<name>" not found on Iconify (404)` | No such icon. | Check the id on iconify.design. |
| `<path>: icon fetch failed for "<name>": ...` | Network error. | Retry, or inline `svg`. |
| `<path>: sharp is not installed; skipped icon (npm install sharp)` / `icon SVG to PNG failed: ...` | Icons are rasterized with `sharp`. | Install it, or replace the icon with text. |
| `<path>: <lint message> (<fix>)` | Validation warnings are repeated at render. | See §2. |

## 4. Render-time failures (the render throws / CLI exits 1)

| Error | Meaning | What to do |
|---|---|---|
| `bankops: 'sharp' is required for rendering SVG (charts, icons) to PNG. Fix: npm install sharp` (`MissingDependencyError`) | A `chart` in the default image mode (or a `waterfall` / `gauge`, which are always images) needs `sharp`. | `npm install sharp`, or use `pptxChart` / `pptxExportMode: "nativeExcel"` for bar, line, pie, combo, scatter. |
| `bankops: 'pptxgenjs' is required for native (editable) PPTX charts. Fix: npm install pptxgenjs` | `pptxChart` content or `pptxExportMode: "nativeExcel"` without `pptxgenjs`. | Install it, or drop `pptxExportMode` and use `chart` (needs `sharp`). |
| `<path>: unknown $component "<name>"; known components: ...` / `; pass options.components` | `$component` names something not in `options.components`. | Register the component, or inline the node. Components are caller-defined; the package ships none. |
| `referencePptx not found: <path>` | `options.referencePptx` does not exist. | Fix the path. |
| `Unknown style preset: <name>. Known presets: corporate, minimal, dark, warm` | `--style` / `options.style` string is not a preset (validation catches `deck.style`). | Use a preset name or a JSON object. |
| `<file> is not valid JSON: ...` / `error: cannot read <file>` (CLI) | Bad or missing deck file. | Check the path and the JSON (trailing commas, comments). |
| `[SlideBuilder] Unknown chart type: <x>` (console) | Only reachable with `validate: false`. | Keep validation on. |

HTTP status codes: `400` invalid JSON body, `405` non-POST on a POST route, `404` unknown route
(the body lists the routes), `422` invalid deck, `500` any render failure (`{ error }`).

## 5. Things that are not errors but will look wrong

- Text longer than the zone's `_maxChars` is clipped silently. Run `bankops annotate` and fit
  the copy (`recipes.md` §12).
- A `chart.title` under a zone `header` doubles the label; use the header.
- `pptxChart` font sizes (`fontSize`, `dataLabelFontSize`) and `timeline` font sizes are plain
  points; everything else is hundredths.
- `headers: []` on a table means no header row; a row with fewer cells than headers shifts values.
- `span` on the root `body` is ignored; the body always fills the slide area.
