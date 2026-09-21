# Recipes — investment-banking layouts

Each recipe gives: when to use it, the grid (spans, headers), the content type per zone, copy
budgets, and a compact JSON skeleton you can fill. Spans are 12-column weights; `fontSize` is
hundredths of a point. Names in brackets (`[Company]`, `[Bank]`) are placeholders. Every slide
should carry `sectionLabel`, a takeaway `title` and a `footer` with sources; they are omitted from
some skeletons for brevity.

Contents

1. Building blocks: stat row, tombstone, numbered row, label/value rows, process chevrons
2. Private company profile — 1, 2 and 3 pages
3. Public company profile — 1 and 4 pages
4. Credentials section (overview, representative transactions, testimonials)
5. Investment highlights
6. Industry overview section
7. Competitive analysis (offerings matrix, dynamics, evolution, gap analysis)
8. M&A mandate pitch
9. Fundraising pitch
10. Common deck structures (IB pitch, product pitch)
11. Chart rules
12. Fixing overflow with the budgets

---

## 1. Building blocks

### 1.1 Stat / metric row

Headline KPIs. Either the built-in `statGrid`, or — for full control of size and color — one
`col` per stat with a value zone (span 7, 20pt bold accent, anchored bottom) over a label zone
(span 5, 9pt bold, anchored top). Four stats → span 3 each; three → 4; six → 2.

```json
{ "direction": "row", "children": [
  { "span": 3, "direction": "col", "children": [
    { "span": 7, "content": { "type": "text", "text": "$680m", "fontSize": 2000, "bold": true, "align": "ctr", "anchor": "b", "color": "#4472C4" } },
    { "span": 5, "content": { "type": "text", "text": "FY25 revenue", "fontSize": 900, "bold": true, "align": "ctr", "anchor": "t" } } ] },
  { "span": 3, "direction": "col", "children": [
    { "span": 7, "content": { "type": "text", "text": "32%", "fontSize": 2000, "bold": true, "align": "ctr", "anchor": "b", "color": "#4472C4" } },
    { "span": 5, "content": { "type": "text", "text": "YoY growth", "fontSize": 900, "bold": true, "align": "ctr", "anchor": "t" } } ] },
  { "span": 3, "direction": "col", "children": [
    { "span": 7, "content": { "type": "text", "text": "118%", "fontSize": 2000, "bold": true, "align": "ctr", "anchor": "b", "color": "#4472C4" } },
    { "span": 5, "content": { "type": "text", "text": "Net retention", "fontSize": 900, "bold": true, "align": "ctr", "anchor": "t" } } ] },
  { "span": 3, "direction": "col", "children": [
    { "span": 7, "content": { "type": "text", "text": "68%", "fontSize": 2000, "bold": true, "align": "ctr", "anchor": "b", "color": "#4472C4" } },
    { "span": 5, "content": { "type": "text", "text": "Gross margin", "fontSize": 900, "bold": true, "align": "ctr", "anchor": "t" } } ] } ] }
```

Sublabel variant: a third child (span 2, `fontSize: 700`, italic, `anchor: "t"`) for "+12% YoY".
Abbreviate values (`$1.2bn`, not `$1,200,000,000`); labels 2–4 words.

### 1.2 Transaction tombstone

Credentials pages only (comparables are tables). A `col` with 6 children, thin border, `gap: 0`;
parent grid `gap: 0.08`. Sort newest → oldest, top-left to bottom-right.

| Row | Span | Content |
|---|---|---|
| 1 | 2 | client name, `fontSize: 600`, bold, `align: "ctr"` |
| 2 | 1 | `{ "type": "line" }` |
| 3 | 3 | client logo `image` (`url` or `path`, `objectFit: "contain"`) |
| 4 | 1 | "acquired by" / "invested in", `fontSize: 600`, italic, centered |
| 5 | 3 | counterparty logo |
| 6 | 2 | date "Jan 2026", `fontSize: 800`, centered |

```json
{ "direction": "col", "gap": 0, "border": "#CCCCCC", "children": [
  { "span": 2, "content": { "type": "text", "text": "[Client]", "fontSize": 600, "bold": true, "align": "ctr" } },
  { "span": 1, "content": { "type": "line" } },
  { "span": 3, "content": { "type": "image", "path": "logos/client.png", "objectFit": "contain", "imagePadding": 0.05 } },
  { "span": 1, "content": { "type": "text", "text": "acquired by", "fontSize": 600, "italic": true, "align": "ctr" } },
  { "span": 3, "content": { "type": "image", "path": "logos/buyer.png", "objectFit": "contain", "imagePadding": 0.05 } },
  { "span": 2, "content": { "type": "text", "text": "Jan 2026", "fontSize": 800, "align": "ctr" } } ] }
```

Without logos, replace rows 3 and 5 with centered 9pt bold text zones naming the parties.
Grids: 6 x 3 (18 deals) or 8 x 3 (24) for a full slide; 3 x 3 or 4 x 3 for half a slide.

### 1.3 Label / value rows (instead of a small table)

Funding rounds, board members, announcements, initiatives: a `col` zone with one `row` child per
record. Text at `fontSize: 900`. Because each record is a row of text zones, the engine gives each
the same height; keep 4–8 rows per zone.

```json
{ "span": 6, "header": "Fundraising", "direction": "col", "gap": 0.05, "children": [
  { "direction": "row", "children": [
    { "span": 2, "content": { "type": "text", "text": "Jan 2022", "fontSize": 900 } },
    { "span": 3, "content": { "type": "text", "text": "Series B", "fontSize": 900, "bold": true } },
    { "span": 5, "content": { "type": "text", "text": "[Lead investor], [Investor]", "fontSize": 900 } },
    { "span": 2, "content": { "type": "text", "text": "$50m", "fontSize": 900, "align": "r" } } ] },
  { "direction": "row", "children": [
    { "span": 2, "content": { "type": "text", "text": "Sep 2024", "fontSize": 900 } },
    { "span": 3, "content": { "type": "text", "text": "Series C", "fontSize": 900, "bold": true } },
    { "span": 5, "content": { "type": "text", "text": "[Lead investor]", "fontSize": 900 } },
    { "span": 2, "content": { "type": "text", "text": "$120m", "fontSize": 900, "align": "r" } } ] } ] }
```

### 1.4 Numbered row (investment highlight, key theme)

Row: number zone (span 1.5, `fontSize: 2400`, bold, accent) | `col` (span 10.5): title zone
(bold, `1400`, one line) over a bullets zone (`1300`, `lineSpacing: 0.95`, 2 bullets of 1–2 lines).

### 1.5 Process chevrons

A `row` with `childShape: "chevron"` draws every child as a chevron; one text child per stage,
white bold centered text, 5–8 stages.

```json
{ "span": 3, "direction": "row", "gap": 0.05, "childShape": "chevron", "childShapeColor": "#164556", "children": [
  { "content": { "type": "text", "text": "Outreach", "fontSize": 900, "bold": true, "color": "#FFFFFF", "align": "ctr", "anchor": "ctr" } },
  { "content": { "type": "text", "text": "NDA", "fontSize": 900, "bold": true, "color": "#FFFFFF", "align": "ctr", "anchor": "ctr" } },
  { "content": { "type": "text", "text": "Mgmt. presentation", "fontSize": 900, "bold": true, "color": "#FFFFFF", "align": "ctr", "anchor": "ctr" } },
  { "content": { "type": "text", "text": "Data room", "fontSize": 900, "bold": true, "color": "#FFFFFF", "align": "ctr", "anchor": "ctr" } },
  { "content": { "type": "text", "text": "LOIs", "fontSize": 900, "bold": true, "color": "#FFFFFF", "align": "ctr", "anchor": "ctr" } },
  { "content": { "type": "text", "text": "Final bids", "fontSize": 900, "bold": true, "color": "#FFFFFF", "align": "ctr", "anchor": "ctr" } },
  { "content": { "type": "text", "text": "Closing", "fontSize": 900, "bold": true, "color": "#FFFFFF", "align": "ctr", "anchor": "ctr" } } ] }
```

---

## 2. Private company profile

Global rules for every profile page: `fontSize: 900` for all body text in profile zones; title
`"[Company] — Company Profile"` (add `(1/2)`, `(2/2)` for multi-page); `sectionLabel`
`"Company Profiles"`; logos as `image` with `url`/`path` and `objectFit: "contain"`, or omit the
logo zone and give the text the full width.

### 2.1 One page (2 x 2)

Body `col` of two `row`s, each with two span-6 zones, every quadrant with a `header`.

| Quadrant | Header | Layout | Content |
|---|---|---|---|
| top-left | Company Overview | row: text 9 / logo 3 | 2–4 sentences: founded, HQ, business model, differentiators |
| top-right | Products & Services | col | 4–6 bullet runs, one product each with a one-sentence description |
| bottom-left | Fundraising | col of rows (§1.3) | date 2 / round 3 / lead investors 5 / amount 2, oldest first |
| bottom-right | Management Team | col | one bullet per executive: bold name, title, one credential |

```json
{ "sectionLabel": "Company Profiles", "title": "[Company] — Company Profile", "footer": "Source: Company website, press releases, [Bank] analysis",
  "body": { "direction": "col", "children": [
    { "direction": "row", "children": [
      { "span": 6, "header": "Company Overview", "direction": "row", "children": [
        { "span": 9, "content": { "type": "text", "fontSize": 900, "text": "[Company] (founded 2016, [City]) sells ... to ... Its ... differentiates it by ..." } },
        { "span": 3, "content": { "type": "image", "path": "logos/company.png", "objectFit": "contain" } } ] },
      { "span": 6, "header": "Products & Services", "content": { "type": "text", "fontSize": 900, "runs": [
        { "text": "[Product A] — ", "bullet": true, "bold": true }, { "text": "one-sentence description" },
        { "text": "[Product B] — ", "bullet": true, "bold": true }, { "text": "one-sentence description" } ] } } ] },
    { "direction": "row", "children": [
      { "span": 6, "header": "Fundraising", "direction": "col", "gap": 0.05, "children": [
        { "direction": "row", "children": [
          { "span": 2, "content": { "type": "text", "text": "Jan 2022", "fontSize": 900 } },
          { "span": 3, "content": { "type": "text", "text": "Series B", "fontSize": 900, "bold": true } },
          { "span": 5, "content": { "type": "text", "text": "[Lead investor]", "fontSize": 900 } },
          { "span": 2, "content": { "type": "text", "text": "$50m", "fontSize": 900, "align": "r" } } ] } ] },
      { "span": 6, "header": "Management Team", "content": { "type": "text", "fontSize": 900, "runs": [
        { "text": "[Name]", "bullet": true, "bold": true }, { "text": " — Co-Founder & CEO; previously [Role] at [Company]" },
        { "text": "[Name]", "bullet": true, "bold": true }, { "text": " — CFO; previously [Role] at [Company]" } ] } } ] } ] } }
```

Budgets: at 9pt a span-6 quadrant holds about 6 bullets of ~90 characters; run `bankops annotate`
and keep each zone under its `_maxChars`.

### 2.2 Two pages

Page 1 (`(1/2)`): body `row` of two span-6 `col`s.
- Left: **Company Overview** (row: text 9 / logo 3), then **Products & Services** — one `row` per
  product: logo 2 | `col` 10 (name zone bold 900, description zone 900, 1–2 sentences).
- Right: **Fundraising** (label/value rows, oldest first), then **Management Team** (bullets).

Page 2 (`(2/2)`): body `row` of two span-6 `col`s.
- Left: **Recent Announcements** — rows: date 3 | one-sentence event 9, newest first; **Board
  Members** — rows: bold name 4 | role and affiliation 8.
- Right: **Strategic Focus** — rows: bold 3–5-word initiative 4 | one sentence 8; **Competitive
  Positioning** — rows: competitor logo 2 | bold name 3 | how [Company] differs 7.

### 2.3 Three pages

Pages 1–2 as in §2.2. Page 3 (`(3/3)`), body `row` of two span-6 `col`s:
- Left: **Management Team** extended — rows: bold name 3 | title 4 | one-sentence bio 5;
  **Board Members** extended — rows: bold name 3 | role 4 | one-sentence note 5.
- Right: **Recent CEO Appearances & Press** — rows: date 3 | event or publication 9, newest first;
  **Additional Context** — bullets or rows for partnerships, regulatory status, notable customers,
  awards.

---

## 3. Public company profile

Same global rules as §2. Titles `"[Company] ([TICKER]) — Company Profile"`. Financial tables use
`fontSize: 900`, `rowHeight: 0.22`–`0.26`, `headerHeight: 0.26`, `colAlign` left then right.
Figures in $m or $bn (use $bn from $1bn); negatives as `(120)`; `NM` for not meaningful.

### 3.1 One page (2 x 2)

| Quadrant | Header | Content |
|---|---|---|
| top-left | Company Overview | row: text 9 / logo 3 |
| top-right | Products & Services | 4–6 bullets |
| bottom-left | Management Team | one bullet per executive |
| bottom-right | Valuation & Financials | `col` of two tables: capital structure (§3.2 P2) over trading multiples |

### 3.2 Four pages

**P1 — Overview, Products, Management, Board.** Body `row` of two span-6 `col`s. Left: Company
Overview (text 9 / logo 3) then Products & Services (bullets, or one row per product with logo).
Right: Management Team (bullets) then Board of Directors (rows: bold name 4 | role and affiliation 8).

**P2 — Valuation, Financials, Shareholders, Stock.** Body `row` of two span-6 `col`s.
- Left, **Valuation & Capital Structure** (`col`, two tables): capital structure — no header row,
  4 rows Market cap | Total debt | Cash & investments | Enterprise value, `colWidths: [7, 5]`,
  `colAlign: ["l", "r"]`; trading multiples — `headers: ["Metric", "LTM", "NTM"]`, rows EV/Revenue |
  EV/EBITDA | P/E, `colWidths: [6, 3, 3]`, `colAlign: ["l", "r", "r"]`.
- Left, **Shareholder Overview**: `headers: ["Shareholder", "Shares (m)", "Ownership"]`,
  `colAlign: ["l", "r", "r"]`, six rows (top three holders, total institutional, total retail,
  total 100%), `summaryRows: 1`.
- Right, **Income Statement Summary**: headers `["", "FY[N-1]A", "FY[N]A", "FY[N+1]E", "FY[N+2]E"]`,
  8 rows Revenue | Gross profit | Gross margin | EBITDA | EBITDA margin | Net income | Net margin |
  EPS (diluted), `colWidths: [4, 2, 2, 2, 2]`, `colAlign: ["l", "r", "r", "r", "r"]`.
- Right, **[TICKER] — 52-week share price**: `chart` `line`, ~52 weekly closes, `categories`
  `"Mar-25"` style, `series: [{ "name": "Close", "data": [...] }]`, `smooth: true`, `area: true`,
  `valuePrefix: "$"`, `yMin`/`yMax` framing the range with ~10% headroom, `showLegend: false`.

```json
{ "span": 6, "header": "Income Statement Summary", "content": { "type": "table", "data": {
  "headers": ["$m", "FY24A", "FY25A", "FY26E", "FY27E"],
  "rows": [ ["Revenue", "1,210", "1,480", "1,790", "2,120"], ["Gross profit", "790", "985", "1,200", "1,430"],
            ["Gross margin", "65.3%", "66.6%", "67.0%", "67.5%"], ["EBITDA", "245", "340", "450", "560"],
            ["EBITDA margin", "20.2%", "23.0%", "25.1%", "26.4%"], ["Net income", "120", "190", "270", "350"],
            ["Net margin", "9.9%", "12.8%", "15.1%", "16.5%"], ["EPS (diluted)", "$1.02", "$1.58", "$2.21", "$2.84"] ],
  "colWidths": [4, 2, 2, 2, 2], "colAlign": ["l", "r", "r", "r", "r"], "fontSize": 900, "rowHeight": 0.24, "headerHeight": 0.26 } } }
```

**P3 — Strategic Focus and Competitive Positioning.** Left span 6: rows of bold initiative 4 | one
sentence 8, 4–6 initiatives. Right span 6: rows of logo 2 | bold competitor 3 | differentiation 7,
4–6 competitors.

**P4 — Product deep dive.** Two columns: description and bullets (span 6) beside a product image
or screenshot supplied by `url`/`path` (span 6, `objectFit: "contain"`), or a `cardGrid` of
feature cards when no image is available.

---

## 4. Credentials section

### 4.1 Credentials overview (2 x 2)

1. **Highlights** — stat row (§1.1): total transaction value, deals closed, senior bankers, league
   table rank.
2. **Practice areas** — `cardGrid` (title + 1–2 lines per sector) or icon tiles (icon 2 | text 10).
3. **Representative transactions** — 3 x 2 or 4 x 2 tombstones (§1.2).
4. **Team** — rows of bold name 4 | title and years 8, or one `profile` per zone.

### 4.2 Representative transactions

Full-slide tombstone grid, `col` of 3 `row`s, 6–8 tombstones per row (15–24 deals), newest first.

### 4.3 Client testimonials

Two `col` zones of span 6, each: quote text (span 8, 12pt italic, `anchor: "ctr"`) over an
attribution zone (span 4, 9pt bold: name, role, company). Optional third row: press quotes.

---

## 5. Investment highlights

Body `col` of five `row`s (§1.4). Common highlights: Large and growing TAM; Best-in-class unit
economics; Differentiated product / technology; Mission-critical customer relationships; Multiple
expansion vectors; Strong management team.

```json
{ "sectionLabel": "Investment Highlights", "title": "[Company] combines scale, margin and a clear path to $1bn revenue",
  "body": { "direction": "col", "gap": 0.1, "children": [
    { "direction": "row", "children": [
      { "span": 1.5, "content": { "type": "text", "text": "01", "fontSize": 2400, "bold": true, "color": "#4472C4", "align": "ctr", "anchor": "ctr" } },
      { "span": 10.5, "direction": "col", "gap": 0, "children": [
        { "span": 4, "content": { "type": "text", "text": "Large and growing TAM", "fontSize": 1400, "bold": true, "anchor": "b" } },
        { "span": 8, "content": { "type": "text", "fontSize": 1300, "lineSpacing": 0.95, "runs": [
          { "text": "$48bn market growing 14% a year, under 20% penetrated", "bullet": true },
          { "text": "Regulation pulls spend forward in the two largest segments", "bullet": true } ] } } ] } ] }
  ] } }
```

Repeat the row for 02–05. At five rows each bullets zone holds two bullets of ~110 characters.

---

## 6. Industry overview section

| Slide | Layout | Content |
|---|---|---|
| Market update | `col`: chart 5 / grid 7 | top: `bar` market size over 5–8 years; bottom: 2 x 2 trend tiles, each `row` icon 2 \| text 10 (bold title run + 3–4 bullets) |
| Sub-sector deep dive (one per sub-sector) | `row` 6 / 6 or 7 / 5 | text bullets (trends, drivers) beside a chart; or three `col`s of span 4 with subheaders |
| M&A activity | `col`: table 8 / chart 4 or `row` 7 / 5 | table Date \| Acquirer \| Target \| Deal size \| EV/Revenue, `colAlign` l l l r r; `bar` of annual deal count or value |
| Fundraising activity | same | table Date \| Company \| Round \| Investors \| Amount; `bar` of funding by year or quarter |
| Most active investors / acquirers | table or `cardGrid` | rank, name (logo cell optional), deal count, notable targets |
| Competitive dynamics | `col`: table 7 / matrix 5 | chronological event table; offerings matrix (players as columns, offering types as rows) highlighting gaps |
| Strategy of major players | table or `cardGrid` | players as columns, strategic priorities as rows |
| Acquisition targets | 2 x 2 profiles | §2.1 quadrants for the top targets |
| Industry insights | `row` 7 / 5 | 4–6 specific, defensible observations with numbers; a supporting chart |

---

## 7. Competitive analysis (3–5 slides)

**Offerings matrix.** Full-width table, rows = 8–15 competitors ordered by revenue, columns = 5–8
offering categories. First column span 3 (name, or a logo cell), remaining spans equal. Ratings
`●●●` strong, `●●○` moderate, `●○○` weak, `—` no offering (a gap); rating columns `ctr`. Subfooter:
"Ratings reflect relative market positioning; — denotes no offering." Follow with a text zone or
the slide's next row calling out the notable gaps.

**Competitive dynamics profile.** `col` of two `row`s. Top: **Nature of Competition** (span 6,
text: one bold-labelled paragraph per dynamic — price/cost, innovation race, brand and marketing,
scale and distribution, consolidation, regulatory moat, platform lock-in — with 2–3 sentences of
evidence) | **Market Concentration** (span 6: `bar` of top-5 shares or a stat row; one-line
characterization: top-3 > 60% highly concentrated, 40–60% moderate, < 40% fragmented). Bottom:
**Intensity Indicators** (span 6, table of 4–6 metrics such as gross-margin trend, churn, R&D %,
S&M %, new entrants, with an "Intensifying / Stable / Easing" column) | **Barriers to Entry**
(span 6, 4–6 bullets rated High / Medium / Low, last bullet the overall assessment).

**Competitive evolution.** Full-width table, oldest first, `colWidths: [2, 4, 3, 3]`: Year | Event
| Players | Impact (one sentence). 8–15 events mixing M&A, launches, entries and exits, regulation,
pricing shifts; end with the most recent. Alternatively a `timeline` for 6–8 events.

**Gap analysis (optional).** `row` 7 / 5. Left **Portfolio Gap Analysis**: table Potential acquirer
| Missing capability | Strategic rationale (`colWidths: [4, 4, 4]`), 4–8 rows. Right **Potential
Targets**: 4–6 bullets, bold company name, one-sentence profile, which acquirers benefit.

---

## 8. M&A mandate pitch

| Slide | Layout |
|---|---|
| Process overview | full-width chevron row (§1.5): Outreach → NDA → Management presentation → Data room → Second presentation → LOI → Final bids → Closing; below it a `row` of span-equal text zones with 2 bullets per stage |
| Strategic value to acquirers | `cardGrid` or three `col`s (Strategic, Financial sponsors, Platform), each 3–4 bullets: outcompete rivals, cross-sell, gaps filled, position strengthened |
| Buyer universe | three header zones side by side (Strategic — large cap, Strategic — mid cap, Financial sponsors), each a grid of logo zones (`image`, `objectFit: "contain"`) or bold name zones |
| Process timeline | table with weeks as columns and stages as rows; active cells `"■"`, `colWidths` first column 4, the rest equal; or a horizontal `bar` per phase |
| Transaction comparables | table Date \| Acquirer \| Target \| Deal size ($m) \| EV/Revenue \| EV/EBITDA \| EV/EBIT, `colAlign` l l l r r r r, Mean and Median rows last with `summaryRows: 2`; never tombstones |

```json
{ "sectionLabel": "Valuation", "title": "Precedent transactions imply 5.5x–6.5x revenue", "footer": "Source: Company filings, press releases, [Bank] analysis",
  "body": { "header": "Selected precedent transactions", "content": { "type": "table", "data": {
    "headers": ["Date", "Acquirer", "Target", "Deal size ($m)", "EV/Revenue", "EV/EBITDA"],
    "rows": [ ["Mar-26", "[Acquirer A]", "[Target A]", "2,400", "6.2x", "19.5x"],
              ["Nov-25", "[Acquirer B]", "[Target B]", "1,150", "5.1x", "16.0x"],
              ["Jun-25", "[Acquirer C]", "[Target C]", "780", "6.8x", "NM"],
              ["Mean", "", "", "", "6.0x", "17.8x"], ["Median", "", "", "", "6.2x", "17.8x"] ],
    "colWidths": [2, 3, 3, 2, 1, 1], "colAlign": ["l", "l", "l", "r", "r", "r"], "fontSize": 1000, "rowHeight": 0.32, "summaryRows": 2 } } } }
```

---

## 9. Fundraising pitch

- **Performance** — stat row of ARR, YoY growth, NRR, gross margin, then a `line` or `bar` of the
  trend (`row` 5 / 7, or stats on top span 3 and chart below span 9).
- **Growth strategy** — four `cardGrid` items or a 2 x 2 of header zones: market expansion,
  product roadmap, M&A, international; 3 bullets each showing how the trajectory supports the raise.
- **Process** — chevron row: Outreach → NDA → Management presentation → Data room → Second
  presentation → Term sheets → Final diligence → Closing.
- **Timeline** — weeks-as-columns table as in §8.
- **Investor universe** — groups Growth equity | Late-stage VC | Crossover | Strategic, logo or
  name grids under header zones.

---

## 10. Common deck structures

**IB pitch (corporate):** title slide → executive summary (`col`: stat row span 3 / `row` chart 8
+ key-metrics table 4, span 9) → market overview (`row`: chart 6 / bullets 6) → competitive
landscape (`row`: scatter or matrix 7 / `col` 5: peers table 7 + valuation stats 5) → company
profile(s) (§2–3) → valuation (§8 comps) → credentials (§4) → appendix.

Title slide: no `title`; body `row` with a spacer `{ "span": 1 }`, a `col` of span 7 holding the
deck title (`fontSize: 2800`, bold, `anchor: "b"`) over the date and "[Bank] | Strictly private
and confidential" (`fontSize: 1400`, secondary color), and a span-4 zone with `background`
`#164556` or a `url`/`path` image (`objectFit: "cover"`).

**Product / tech pitch (minimal):** title → problem (`row`: bullets 7 / stats 5) → solution
(`row`: `cardGrid` 6 / chart 6) → customer results (`row`: quotes 6 / three-stat row 6) → team →
ask.

---

## 11. Chart rules

- Label the chart with the zone `header`, not the chart `title`.
- `bar`: sort categories by value (largest first, or top to bottom when `horizontal`) unless they
  are periods. `showBarValues: true`; on single-series bars also `showYAxis: false`
  (`showXAxis: false` when horizontal), `gridLeft: 20`, `gridRight: 20`, and no `yAxisLabel`.
  `stacked: true` always with `stackTotal` (one total per category).
- `line`: `smooth` for prices, `area` for a single series, `yAxisLabel` with the unit
  (`"Revenue ($m)"`), `showLineValues` only with ≤ 8 points.
- `pie`: 2–6 slices. Legend mode (labels > 10 characters): `showLegend: true`,
  `labelPosition: "none"`, `centerY: "45%"`. Label mode: `showLegend: false`,
  `labelPosition: "outside"`. Never both.
- `combo`: bars for the absolute metric, line for the ratio, `dualAxis: true` when scales differ,
  `valuePrefix: "$"` left, `y2ValueSuffix: "%"` right.
- `waterfall` for bridges (revenue, EBITDA, cost walks); first and last values are totals.
- Multi-series charts with a bottom legend: `gridBottom: 45` so the legend does not overlap.
- Money: `valuePrefix: "$"`, `valueSuffix: "m"` or `"bn"` with the data already scaled,
  `valueThousands: true` always, `valueDecimals: 1` for `$4.2bn` / `23.5%`; percentages
  `valueSuffix: "%"` and no prefix.
- Close values that look identical: set `yMin`/`yMax` (e.g. 80–100 for a margin range).
- Editable in PowerPoint: add `pptxExportMode: "nativeExcel"` to the existing `chart` node (bar,
  line, pie, combo, scatter) or write a `pptxChart`. Never change a chart's `type`.
- Placeholder or estimated data: bracket the header (`"[Market size ($bn)]"`) and add a subfooter
  "Placeholder data — verify before use."

---

## 12. Fixing overflow with the budgets

`bankops annotate` gives every zone its capacity (`_maxChars`, `_maxLines`, `_charsPerLine`;
tables `_cellMaxChars[]`, `_headersMaxChars[]`, `_maxRows`; `_headerMaxChars`, `_titleMaxChars`).
Overflowing text is clipped in the PPTX, silently. Apply the least disruptive fix first:

1. **Shorten the copy.** Fewer sentences, 6 bullets → 4, abbreviations in cells ("Mgmt", "Rev").
   Never truncate with "...".
2. **Rebalance sibling spans.** 6/6 → 7/5 or 8/4 toward the crowded zone; 4/4/4 → 5/4/3. In a
   `col` of rows, move vertical span from the emptier row. Keep sums at 12.
3. **Restructure.** Turn a row-major 2 x 2 into two `col`s so the crowded zone can take span 8 of
   its column; give a dense zone its own full-width row; split a 10-bullet zone into 2–3 children
   with `subheader`s.
4. **Tables.** Widen the tight column (take from a wide neighbour; minimum useful span 2), abbreviate
   headers, `rowHeight: 0.5` for two-line cells, then `fontSize: 900`; too many rows → drop rows,
   `rowHeight: 0.3`, or paginate with headers repeated and the title suffixed "(cont.)".
5. **Split across slides.** Same `sectionLabel`, title "(cont.)"; group bullets by theme.
6. **Font size last.** Body ≥ 1000 (prefer 1300), tables ≥ 900, nothing below 800. One step
   (1300 → 1100) usually resolves a marginal overflow.

Headers and subheaders are single lines: shorten them to `_headerMaxChars`, or move detail into a
`subheader`. Titles wrap to two lines at most (`_titleMaxLines`); rephrase past that.
