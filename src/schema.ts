/**
 * Zod schema for the deck JSON format, plus validateDeck() with errors that teach.
 * Isomorphic. `DeckSchema` is the source of truth for the JSON Schema shipped to agents.
 */

import { z } from "zod";
import type { DeckJson } from "./types.js";
import { isNumericColumn } from "./fit.js";

// ---- Primitives ----

const hexColor = z
  .string()
  .regex(/^#([0-9a-fA-F]{6}|[0-9a-fA-F]{3})$/, { message: "Colors are hex strings like #4472C4" });

/** Font sizes are hundredths of a point: 1400 = 14pt. */
const fontSizeHundredths = z
  .number()
  .int({ message: "fontSize must be an integer in hundredths of a point (1400 = 14pt)" })
  .min(100, { message: "fontSize is in hundredths of a point: use 1400 for 14pt, not 14" })
  .max(40000, { message: "fontSize above 400pt is not supported" });

const inches = z.number().min(0).max(20);
const align = z.enum(["l", "ctr", "r"]);
const anchor = z.enum(["t", "ctr", "b"]);
const borderStyle = z.enum(["solid", "dash", "dot", "dashDot", "lgDash", "lgDashDot"]);
const borderSpec = z.object({ color: hexColor, width: z.number().optional(), style: borderStyle.optional() });

// ---- Content types ----

export const TextRunSchema = z.object({
  text: z.string(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  color: hexColor.optional(),
  fontSize: fontSizeHundredths.optional(),
  font: z.string().optional(),
  bullet: z.boolean().optional(),
  bulletLevel: z.number().int().min(0).max(4).optional(),
});

export const BulletStyleSchema = z.object({
  char: z.string().optional(),
  color: hexColor.optional(),
  indent: inches.optional(),
  marginLeft: inches.optional(),
  spaceBefore: z.number().optional(),
});

export const TextContentSchema = z.object({
  type: z.literal("text"),
  text: z.string().optional(),
  runs: z.array(TextRunSchema).optional(),
  font: z.string().optional(),
  fontSize: fontSizeHundredths.optional(),
  color: hexColor.optional(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  textTransform: z.literal("uppercase").optional(),
  align: align.optional(),
  anchor: anchor.optional(),
  lineSpacing: z.number().min(0.5).max(3).optional(),
  bulletStyle: BulletStyleSchema.optional(),
  paddingTop: inches.optional(),
  paddingRight: inches.optional(),
  paddingBottom: inches.optional(),
  paddingLeft: inches.optional(),
});

export const StatGridContentSchema = z.object({
  type: z.literal("statGrid"),
  items: z.array(z.object({ value: z.string(), label: z.string(), sublabel: z.string().optional() })).min(1),
  anchor: anchor.optional(),
});

export const CardGridContentSchema = z.object({
  type: z.literal("cardGrid"),
  items: z.array(z.object({ title: z.string(), lines: z.array(z.string()) })).min(1),
});

export const TableCellImageSchema = z.object({
  type: z.literal("image"),
  source: z.enum(["search", "ai", "upload"]).optional(),
  searchTerm: z.string().optional(),
  prompt: z.string().optional(),
  url: z.string().optional(),
  path: z.string().optional(),
});

export const TableContentSchema = z.object({
  type: z.literal("table"),
  data: z.object({
    headers: z.array(z.string()).optional(),
    rows: z.array(z.array(z.union([z.string(), TableCellImageSchema]))),
    colWidths: z.array(z.number().positive()).optional(),
    rowHeight: inches.optional(),
    headerHeight: inches.optional(),
    fontSize: fontSizeHundredths.optional(),
    headerFontSize: fontSizeHundredths.optional(),
    align: z.record(z.string(), align).optional(),
    colAlign: z.array(align).optional(),
    overflowWrap: z.enum(["normal", "anywhere", "break-word"]).optional(),
    verticalHeaders: z.boolean().optional(),
    summaryRows: z.number().int().min(0).optional(),
  }),
});

export const ProfileContentSchema = z.object({
  type: z.literal("profile"),
  name: z.string(),
  items: z.array(z.string()),
});

export const ImageContentSchema = z.object({
  type: z.literal("image"),
  source: z.enum(["ai", "search", "upload"]).optional(),
  prompt: z.string().optional(),
  searchTerm: z.string().optional(),
  path: z.string().optional(),
  url: z.string().optional(),
  width: z.union([z.number(), z.string()]).optional(),
  height: z.union([z.number(), z.string()]).optional(),
  anchor: anchor.optional(),
  objectFit: z.enum(["cover", "contain", "fill"]).optional(),
  imagePadding: inches.optional(),
  crop: z.object({ l: z.number(), t: z.number(), r: z.number(), b: z.number() }).optional(),
  ext: z.string().optional(),
});

export const IconContentSchema = z.object({
  type: z.literal("icon"),
  name: z.string().regex(/^[a-z0-9-]+:[a-z0-9-]+$/, { message: 'Icon names are Iconify ids like "mdi:home"' }),
  svg: z.string().optional(),
  color: hexColor.optional(),
  size: z.number().min(1).max(100).optional(),
  align: z.enum(["top", "middle", "bottom"]).optional(),
  objectFit: z.enum(["contain", "cover", "fill"]).optional(),
});

/** Series values: numbers, or [x, y] pairs for scatter charts. */
const series = z.object({ name: z.string().optional(), data: z.array(z.union([z.number(), z.tuple([z.number(), z.number()])])) });
const indexColorMap = z.record(z.string(), hexColor);

export const ChartContentSchema = z.object({
  type: z.literal("chart"),
  chartType: z.enum(["bar", "line", "pie", "combo", "waterfall", "scatter", "gauge"]),
  title: z.string().optional(),
  categories: z.array(z.string()).optional(),
  series: z.array(series).optional(),
  items: z.array(z.object({ name: z.string(), value: z.number() })).optional(),
  bars: z.array(series).optional(),
  lines: z.array(series).optional(),
  horizontal: z.boolean().optional(),
  stacked: z.boolean().optional(),
  stackTotal: z.array(z.number()).optional(),
  smooth: z.boolean().optional(),
  area: z.boolean().optional(),
  doughnut: z.boolean().optional(),
  roseType: z.enum(["radius", "area"]).optional(),
  startAngle: z.number().optional(),
  showLabel: z.boolean().optional(),
  labelPosition: z.enum(["outside", "inside", "center", "none"]).optional(),
  showLegend: z.boolean().optional(),
  legendPosition: z.enum(["bottom", "left", "right", "top"]).optional(),
  innerRadius: z.string().optional(),
  outerRadius: z.string().optional(),
  centerX: z.string().optional(),
  centerY: z.string().optional(),
  dualAxis: z.boolean().optional(),
  xMin: z.number().optional(), xMax: z.number().optional(),
  yMin: z.number().optional(), yMax: z.number().optional(),
  y2Min: z.number().optional(), y2Max: z.number().optional(),
  gridTop: z.number().optional(), gridRight: z.number().optional(),
  gridBottom: z.number().optional(), gridLeft: z.number().optional(),
  data: z.array(z.number()).optional(),
  value: z.number().optional(),
  min: z.number().optional(),
  max: z.number().optional(),
  label: z.string().optional(),
  format: z.string().optional(),
  showSeriesLabels: z.boolean().optional(),
  showTotalValues: z.boolean().optional(),
  showBarValues: z.boolean().optional(),
  valuePrefix: z.string().optional(),
  valueSuffix: z.string().optional(),
  valueDecimals: z.number().int().min(0).max(6).optional(),
  valueThousands: z.boolean().optional(),
  negativeFormat: z.enum(["minus", "parens"]).optional(),
  barBorderColor: hexColor.optional(),
  barBorderWidth: z.number().optional(),
  barColor: hexColor.optional(),
  categoryColors: indexColorMap.optional(),
  showLineValues: z.boolean().optional(),
  pointShape: z.enum(["circle", "rect", "triangle", "diamond"]).optional(),
  pointSize: z.number().optional(),
  pointBorderColor: hexColor.optional(),
  pointBorderWidth: z.number().optional(),
  pointColor: hexColor.optional(),
  lineColor: hexColor.optional(),
  lineWidth: z.number().optional(),
  showXAxis: z.boolean().optional(),
  showYAxis: z.boolean().optional(),
  showSecondaryYAxis: z.boolean().optional(),
  showSplitLine: z.boolean().optional(),
  y2ValuePrefix: z.string().optional(),
  y2ValueSuffix: z.string().optional(),
  y2ValueDecimals: z.number().int().optional(),
  y2ValueThousands: z.boolean().optional(),
  y2NegativeFormat: z.enum(["minus", "parens"]).optional(),
  xAxisLabel: z.string().optional(),
  xAxisLabelPosition: z.enum(["end", "center"]).optional(),
  yAxisLabel: z.string().optional(),
  yAxisLabelPosition: z.enum(["top", "side"]).optional(),
  pieBorderColor: hexColor.optional(),
  pieBorderWidth: z.number().optional(),
  padAngle: z.number().optional(),
  sliceColors: indexColorMap.optional(),
  showPieValues: z.boolean().optional(),
  pieValuePosition: z.enum(["outside", "inside"]).optional(),
  pptxExportMode: z.enum(["image", "nativeExcel"]).optional(),
}).superRefine((c, ctx) => {
  const needs = (field: string, ok: boolean, hint: string) => {
    if (!ok) ctx.addIssue({ code: "custom", path: [field], message: hint });
  };
  switch (c.chartType) {
    case "pie":
      needs("items", Array.isArray(c.items) && c.items.length > 0, 'pie charts use items: [{ name, value }], not series');
      break;
    case "combo":
      needs("bars", Array.isArray(c.bars) || Array.isArray(c.lines), "combo charts use bars: [...] and lines: [...] arrays plus categories");
      break;
    case "waterfall":
      needs("data", Array.isArray(c.data) && c.data.length > 1, "waterfall charts use a flat data: number[] (first and last are totals, middle values are deltas) plus categories");
      break;
    case "gauge":
      needs("value", typeof c.value === "number", "gauge charts need value (and optionally min, max, label)");
      break;
    case "scatter":
      needs("series", Array.isArray(c.series) && c.series.length > 0, "scatter charts use series: [{ name, data: [[x, y], ...] }]");
      break;
    default:
      needs("series", Array.isArray(c.series) && c.series.length > 0, `${c.chartType} charts use categories: string[] and series: [{ name, data: number[] }]`);
  }
});

const pptxSeries = z.object({ name: z.string(), labels: z.array(z.string()), values: z.array(z.number()) });
const gridLine = z.object({ style: z.enum(["solid", "dash", "dot", "none"]).optional(), color: hexColor.optional(), size: z.number().optional() });

export const PptxChartContentSchema = z.object({
  type: z.literal("pptxChart"),
  chartType: z.enum(["bar", "line", "pie", "scatter", "combo"]),
  series: z.array(pptxSeries),
  title: z.string().optional(),
  showTitle: z.boolean().optional(),
  showLegend: z.boolean().optional(),
  legendPos: z.enum(["b", "t", "l", "r", "tr"]).optional(),
  chartColors: z.array(hexColor).optional(),
  barDir: z.enum(["col", "bar"]).optional(),
  barGrouping: z.enum(["clustered", "stacked", "percentStacked"]).optional(),
  barFillColor: hexColor.optional(),
  gapWidthPct: z.number().optional(),
  lineSmooth: z.boolean().optional(),
  doughnut: z.boolean().optional(),
  catAxisTitle: z.string().optional(),
  valAxisTitle: z.string().optional(),
  valAxisMinVal: z.number().optional(),
  valAxisMaxVal: z.number().optional(),
  valAxisLabelFormatCode: z.string().optional(),
  showValue: z.boolean().optional(),
  showPercent: z.boolean().optional(),
  dataLabelFormatCode: z.string().optional(),
  comboCharts: z.array(z.object({
    chartType: z.enum(["bar", "line", "area"]),
    series: z.array(pptxSeries),
    options: z.object({ secondaryValAxis: z.boolean().optional(), barGrouping: z.string().optional() }).optional(),
  })).optional(),
  catAxisHidden: z.boolean().optional(),
  valAxisHidden: z.boolean().optional(),
  valGridLine: gridLine.optional(),
  catGridLine: gridLine.optional(),
  secondaryValAxisLabelFormatCode: z.string().optional(),
  secondaryValAxisMinVal: z.number().optional(),
  secondaryValAxisMaxVal: z.number().optional(),
  dataLabelFontSize: z.number().optional(),
  dataLabelColor: hexColor.optional(),
  dataLabelPosition: z.enum(["b", "bestFit", "ctr", "inBase", "inEnd", "l", "outEnd", "r", "t"]).optional(),
  showCatName: z.boolean().optional(),
  showSerName: z.boolean().optional(),
  fontSize: z.number().optional(),
  legendFontSize: z.number().optional(),
  titleFontSize: z.number().optional(),
  dataBorder: z.object({ type: z.enum(["solid", "dash"]).optional(), pt: z.number().optional(), color: hexColor }).optional(),
  lineDataSymbol: z.enum(["circle", "dash", "diamond", "dot", "none", "square", "triangle"]).optional(),
  lineDataSymbolSize: z.number().optional(),
  lineDataSymbolLineColor: hexColor.optional(),
  lineDataSymbolLineSize: z.number().optional(),
});

export const LineContentSchema = z.object({
  type: z.literal("line"),
  color: hexColor.optional(),
  width: z.number().optional(),
});

export const TimelineContentSchema = z.object({
  type: z.literal("timeline"),
  events: z.array(z.object({ date: z.string(), name: z.string(), detail: z.string() })).min(1),
  title: z.string().optional(),
  topCount: z.number().int().optional(),
  placement: z.enum(["split", "top", "bottom"]).optional(),
  boxWidth: z.number().optional(),
  boxHeight: z.number().optional(),
  boxGap: z.number().optional(),
  boxOffset: z.number().optional(),
  boxHeightPct: z.number().optional(),
  boxBorderColor: hexColor.optional(),
  boxBorderWidth: z.number().optional(),
  boxBackground: hexColor.optional(),
  nameFont: z.string().optional(),
  nameFontSize: z.number().optional(),
  nameBold: z.boolean().optional(),
  nameItalic: z.boolean().optional(),
  nameUnderline: z.boolean().optional(),
  nameColor: hexColor.optional(),
  nameAlign: align.optional(),
  detailFont: z.string().optional(),
  detailFontSize: z.number().optional(),
  detailBold: z.boolean().optional(),
  detailItalic: z.boolean().optional(),
  detailUnderline: z.boolean().optional(),
  detailColor: hexColor.optional(),
  lineColor: hexColor.optional(),
  lineWidth: z.number().optional(),
  nodeColor: hexColor.optional(),
  nodeBorderColor: hexColor.optional(),
  nodeBorderWidth: z.number().optional(),
  connectorColor: hexColor.optional(),
  connectorWidth: z.number().optional(),
  dateFormat: z.enum(["month", "year", "month-year", "quarter", "quarter-year"]).optional(),
  datePlacement: z.enum(["auto", "top", "bottom"]).optional(),
  dateFont: z.string().optional(),
  dateFontSize: z.number().optional(),
  dateColor: hexColor.optional(),
  dateBold: z.boolean().optional(),
  dateItalic: z.boolean().optional(),
  textMode: z.enum(["plain", "bullet"]).optional(),
});

export const CONTENT_TYPES = ["text", "statGrid", "cardGrid", "table", "profile", "image", "icon", "chart", "pptxChart", "line", "timeline"] as const;

export const ContentSchema = z.discriminatedUnion("type", [
  TextContentSchema,
  StatGridContentSchema,
  CardGridContentSchema,
  TableContentSchema,
  ProfileContentSchema,
  ImageContentSchema,
  IconContentSchema,
  ChartContentSchema as unknown as typeof TextContentSchema, // superRefine wrapper; discriminator still "type"
  PptxChartContentSchema,
  LineContentSchema,
  TimelineContentSchema,
]);

// ---- Grid node (recursive) ----

const SourceSchema = z.object({
  url: z.string().optional(),
  api: z.string().optional(),
  label: z.string().optional(),
  date: z.string().optional(),
});

export const GridNodeSchema: z.ZodType<any> = z.lazy(() =>
  z.object({
    id: z.string().optional(),
    span: z.number().min(0.5).max(12).optional(),
    direction: z.enum(["row", "col"]).optional(),
    gap: inches.optional(),
    padding: inches.optional(),
    paddingTop: inches.optional(), paddingRight: inches.optional(),
    paddingBottom: inches.optional(), paddingLeft: inches.optional(),
    margin: inches.optional(),
    marginTop: inches.optional(), marginRight: inches.optional(),
    marginBottom: inches.optional(), marginLeft: inches.optional(),
    header: z.string().optional(),
    subheader: z.string().optional(),
    subfooter: z.string().optional(),
    background: hexColor.optional(),
    backgroundOpacity: z.number().min(0).max(100).optional(),
    border: z.union([hexColor, borderSpec]).optional(),
    borderTop: borderSpec.optional(), borderRight: borderSpec.optional(),
    borderBottom: borderSpec.optional(), borderLeft: borderSpec.optional(),
    shadow: z.object({
      color: hexColor.optional(), alpha: z.number().min(0).max(1).optional(),
      blur: z.number().optional(), dist: z.number().optional(),
      dir: z.number().optional(), size: z.number().optional(),
    }).optional(),
    childShape: z.enum(["chevron", "pyramid"]).optional(),
    childShapeColor: hexColor.optional(),
    children: z.array(GridNodeSchema).max(50, { message: "A node may have at most 50 children" }).optional(),
    content: ContentSchema.optional(),
    $component: z.string().optional(),
    _task: z.string().optional(),
    _sources: z.array(SourceSchema).optional(),
  }).superRefine((node, ctx) => {
    if (node.children && node.children.length > 0 && node.content) {
      ctx.addIssue({ code: "custom", path: ["content"], message: "A node is either a container (children) or a leaf (content), never both. Move the content into a child node." });
    }
  })
);

export const SlideSchema = z.object({
  id: z.string().optional(),
  type: z.literal("grid").optional(),
  title: z.string().optional(),
  subtitle: z.string().optional(),
  sectionLabel: z.string().optional(),
  footer: z.string().optional(),
  margin: inches.optional(),
  _task: z.string().optional(),
  body: GridNodeSchema.optional(),
});

export const STYLE_PRESETS = ["corporate", "minimal", "dark", "warm"] as const;

export const DeckSchema = z.object({
  title: z.string().optional(),
  style: z.union([
    z.enum(STYLE_PRESETS),
    z.record(z.string(), z.unknown()),
  ]).optional(),
  slideSize: z.object({ width: z.number().positive().optional(), height: z.number().positive().optional() }).optional(),
  slides: z.array(SlideSchema).min(1, { message: "A deck needs at least one slide" }).max(200, { message: "At most 200 slides per deck" }),
});

export type Deck = z.infer<typeof DeckSchema>;

// ---- Validation with teaching errors ----

export interface DeckIssue {
  /** JSON path, e.g. "slides[2].body.children[0].content.fontSize" */
  path: string;
  message: string;
  /** What to change. */
  fix?: string;
}

export interface ValidationResult {
  ok: boolean;
  /** The parsed deck (unknown keys stripped) when ok. */
  deck?: DeckJson;
  errors: DeckIssue[];
  /** Non-fatal: spans that don't sum to 12, unknown keys, etc. */
  warnings: DeckIssue[];
}

function pathString(path: PropertyKey[]): string {
  let out = "";
  for (const seg of path) {
    if (typeof seg === "number") out += `[${seg}]`;
    else out += out ? `.${String(seg)}` : String(seg);
  }
  return out || "deck";
}

const FIX_HINTS: { test: RegExp; fix: string }[] = [
  { test: /items\[\d+\]\.lines/, fix: "cardGrid items are { title, lines: string[] }" },
  { test: /content\.data.*expected object/i, fix: 'table content is { type: "table", data: { headers, rows, colWidths } }' },
  { test: /Invalid discriminator|Invalid input.*type/i, fix: `content.type must be one of: ${CONTENT_TYPES.join(", ")}` },
  { test: /hundredths of a point/i, fix: "Multiply the point size by 100 (14pt -> 1400)" },
  { test: /hex strings/i, fix: 'Use "#RRGGBB" form, e.g. "#405363"' },
  { test: /at least one slide/i, fix: 'Add slides: [{ title: "...", body: { ... } }]' },
  { test: /container .*leaf/i, fix: "Wrap the content in a child: children: [{ content: {...} }]" },
];

function fixFor(message: string): string | undefined {
  for (const h of FIX_HINTS) if (h.test.test(message)) return h.fix;
  return undefined;
}

/** Suggest the closest known key for a typo (simple edit distance). */
function closest(key: string, known: string[]): string | undefined {
  const lower = key.toLowerCase();
  let best: string | undefined;
  let bestD = 3;
  for (const k of known) {
    const d = editDistance(lower, k.toLowerCase());
    if (d < bestD) { bestD = d; best = k; }
  }
  return best;
}

function editDistance(a: string, b: string): number {
  const dp: number[] = Array.from({ length: b.length + 1 }, (_, j) => j);
  for (let i = 1; i <= a.length; i++) {
    let prev = dp[0]; dp[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const tmp = dp[j];
      dp[j] = Math.min(dp[j] + 1, dp[j - 1] + 1, prev + (a[i - 1] === b[j - 1] ? 0 : 1));
      prev = tmp;
    }
  }
  return dp[b.length];
}

function shapeKeys(schema: any): string[] {
  const inner = schema?.def?.innerType ?? schema?._def?.innerType ?? schema;
  const shape = inner?.shape ?? inner?.def?.shape ?? inner?._def?.shape;
  return shape ? Object.keys(shape) : [];
}

const GRID_NODE_KEYS = (() => {
  const inner: any = (GridNodeSchema as any).def?.getter?.() ?? (GridNodeSchema as any)._def?.getter?.();
  return shapeKeys(inner);
})();
const SLIDE_KEYS = shapeKeys(SlideSchema);
const DECK_KEYS = shapeKeys(DeckSchema);
const CONTENT_KEYS: Record<string, string[]> = {
  text: shapeKeys(TextContentSchema), statGrid: shapeKeys(StatGridContentSchema), cardGrid: shapeKeys(CardGridContentSchema),
  table: shapeKeys(TableContentSchema), profile: shapeKeys(ProfileContentSchema), image: shapeKeys(ImageContentSchema),
  icon: shapeKeys(IconContentSchema), chart: shapeKeys(ChartContentSchema), pptxChart: shapeKeys(PptxChartContentSchema),
  line: shapeKeys(LineContentSchema), timeline: shapeKeys(TimelineContentSchema),
};

function isMeta(key: string): boolean { return key.startsWith("_") || key.startsWith("$"); }

/** Structural lint on the raw input: unknown keys, span sums, table widths. */
function lint(raw: any, warnings: DeckIssue[], errors: DeckIssue[]): void {
  if (!raw || typeof raw !== "object") return;
  for (const k of Object.keys(raw)) {
    if (!DECK_KEYS.includes(k) && !isMeta(k)) warnings.push(unknownKey(k, DECK_KEYS, "deck"));
  }
  if (!Array.isArray(raw.slides)) return;
  raw.slides.forEach((slide: any, i: number) => {
    if (!slide || typeof slide !== "object") return;
    for (const k of Object.keys(slide)) {
      if (!SLIDE_KEYS.includes(k) && !isMeta(k)) warnings.push(unknownKey(k, SLIDE_KEYS, `slides[${i}]`));
    }
    if (slide.body) lintNode(slide.body, `slides[${i}].body`, 0, warnings, errors);
  });
}

function unknownKey(key: string, known: string[], path: string): DeckIssue {
  const near = closest(key, known);
  return {
    path: `${path}.${key}`,
    message: `Unknown key "${key}" is ignored`,
    fix: near ? `Did you mean "${near}"?` : undefined,
  };
}

function lintNode(node: any, path: string, depth: number, warnings: DeckIssue[], errors: DeckIssue[]): void {
  if (!node || typeof node !== "object") return;
  if (Array.isArray(node.children) && node.children.length > 0 && node.content) {
    errors.push({ path: `${path}.content`, message: "A node is either a container (children) or a leaf (content), never both", fix: "Wrap the content in a child: children: [{ content: {...} }]" });
  }
  if (depth > 16) { warnings.push({ path, message: "Nesting deeper than 16 levels", fix: "Flatten the layout" }); return; }
  for (const k of Object.keys(node)) {
    if (!GRID_NODE_KEYS.includes(k) && !isMeta(k)) warnings.push(unknownKey(k, GRID_NODE_KEYS, path));
  }
  const c = node.content;
  if (c && typeof c === "object" && typeof c.type === "string" && CONTENT_KEYS[c.type]) {
    for (const k of Object.keys(c)) {
      if (!CONTENT_KEYS[c.type].includes(k) && !isMeta(k)) warnings.push(unknownKey(k, CONTENT_KEYS[c.type], `${path}.content`));
    }
    if (c.type === "table" && c.data && Array.isArray(c.data.colWidths)) {
      const sum = c.data.colWidths.reduce((a: number, b: number) => a + (Number(b) || 0), 0);
      if (Math.abs(sum - 12) > 0.001) warnings.push({ path: `${path}.content.data.colWidths`, message: `colWidths sum to ${sum}, not 12`, fix: "colWidths are 12-column spans; make them add up to 12" });
    }
    if (c.type === "table" && c.data && Array.isArray(c.data.rows) && Array.isArray(c.data.headers) && c.data.headers.length) {
      const bad = c.data.rows.findIndex((r: any) => Array.isArray(r) && r.length !== c.data.headers.length);
      if (bad >= 0) warnings.push({ path: `${path}.content.data.rows[${bad}]`, message: `Row has ${c.data.rows[bad].length} cells but there are ${c.data.headers.length} headers` });
    }
  }
  if (Array.isArray(node.children)) {
    const spans = node.children.filter(Boolean).map((ch: any) => ch && typeof ch.span === "number" ? ch.span : null);
    if (spans.length > 1 && spans.every((s: number | null) => s != null)) {
      const sum = spans.reduce((a: number, b: number) => a + b, 0);
      if (Math.abs(sum - 12) > 0.001) warnings.push({ path: `${path}.children`, message: `Sibling spans sum to ${sum}, not 12`, fix: "Spans are 12-column weights; siblings normally add up to 12" });
    }
    node.children.forEach((ch: any, i: number) => lintNode(ch, `${path}.children[${i}]`, depth + 1, warnings, errors));
  }
}

// ---- Normalization: accept the shapes models naturally write ----

/**
 * Rewrite common near-miss shapes into the canonical format, recording what changed.
 * Applied before validation by validateDeck; the deck returned there is the normalized one.
 *
 *   cardGrid items { title, text | description | body }  -> { title, lines: [text] }
 *   cardGrid items { title, lines: "one string" }         -> lines: ["one string"]
 *   table { columns | headers, rows } at content level    -> { data: { headers, rows } }
 *   text { bullets: string[] } or { items: string[] }     -> runs with bullet: true
 *   text { text: ["a", "b"] }                             -> runs with bullet: true
 *   statGrid items { value: 12 }                          -> value: "12"
 *   chart series { values: [...] }                        -> { data: [...] }
 *   slide with `content` or `children` but no `body`      -> wrapped in body
 *   "bulletList" content type                             -> text with bullet runs
 */
/** Banking convention: numeric columns are right-aligned. Applied only when no alignment was given. */
function alignNumericColumns(data: any, path: string, note: (p: string, m: string) => void): void {
  if (data.colAlign || data.align) return;
  const rows: any[][] = data.rows.filter((r: any) => Array.isArray(r));
  const ncols = Math.max(Array.isArray(data.headers) ? data.headers.length : 0, ...rows.map((r) => r.length));
  if (ncols < 2) return;
  const colAlign = Array.from({ length: ncols }, (_, i) => (isNumericColumn(rows.map((r) => r[i])) ? "r" : "l"));
  if (!colAlign.includes("r")) return;
  data.colAlign = colAlign;
  note(path, `numeric columns right-aligned (colAlign: ${JSON.stringify(colAlign)})`);
}

export function normalizeDeck(input: unknown, notes: DeckIssue[] = []): unknown {
  if (!input || typeof input !== "object" || Array.isArray(input)) return input;
  const deck: any = JSON.parse(JSON.stringify(input));
  const note = (path: string, message: string) => notes.push({ path, message });

  const fixContent = (c: any, path: string) => {
    if (!c || typeof c !== "object") return c;
    if (c.type === "bulletList") {
      const items = Array.isArray(c.items) ? c.items : [];
      note(path, 'bulletList became text with bullet runs');
      const { items: _i, type: _t, ...rest } = c;
      return { ...rest, type: "text", runs: items.map((t: any) => ({ text: String(t), bullet: true })) };
    }
    if (c.type === "text") {
      const list = Array.isArray(c.bullets) ? c.bullets : Array.isArray(c.items) ? c.items : Array.isArray(c.text) ? c.text : null;
      if (list && !c.runs) {
        note(path, "bullet list became runs with bullet: true");
        const { bullets: _b, items: _i, ...rest } = c;
        return { ...rest, text: undefined, runs: list.map((t: any) => (typeof t === "object" && t && "text" in t ? { bullet: true, ...t } : { text: String(t), bullet: true })) };
      }
      return c;
    }
    if (c.type === "cardGrid" && Array.isArray(c.items)) {
      c.items = c.items.map((it: any, i: number) => {
        if (!it || typeof it !== "object") return it;
        if (Array.isArray(it.lines)) return it;
        const text = typeof it.lines === "string" ? it.lines : it.text ?? it.description ?? it.body ?? it.detail;
        if (text == null) return it;
        note(`${path}.items[${i}]`, "card text became lines: [text]");
        const { text: _t, description: _d, body: _b, detail: _e, lines: _l, ...rest } = it;
        return { ...rest, lines: Array.isArray(text) ? text.map(String) : [String(text)] };
      });
      return c;
    }
    if (c.type === "table" && !c.data && Array.isArray(c.rows)) {
      note(path, "table columns/rows moved under data");
      const { rows, columns, headers, colWidths, colAlign, align, rowHeight, headerHeight, fontSize, headerFontSize, summaryRows, verticalHeaders, ...rest } = c;
      const data: any = { rows: rows.map((r: any) => (Array.isArray(r) ? r.map((cell: any) => (cell != null && typeof cell !== "object" ? String(cell) : cell)) : r)) };
      const h = headers ?? columns;
      if (Array.isArray(h)) data.headers = h.map(String);
      for (const [k, v] of Object.entries({ colWidths, colAlign, align, rowHeight, headerHeight, fontSize, headerFontSize, summaryRows, verticalHeaders })) if (v !== undefined) data[k] = v;
      alignNumericColumns(data, `${path}.data`, note);
      return { ...rest, data };
    }
    if (c.type === "table" && c.data && Array.isArray(c.data.rows)) {
      if (Array.isArray(c.data.columns) && !c.data.headers) { c.data.headers = c.data.columns.map(String); delete c.data.columns; note(`${path}.data`, "columns became headers"); }
      c.data.rows = c.data.rows.map((r: any) => (Array.isArray(r) ? r.map((cell: any) => (cell != null && typeof cell !== "object" ? String(cell) : cell)) : r));
      alignNumericColumns(c.data, `${path}.data`, note);
      return c;
    }
    if (c.type === "statGrid" && Array.isArray(c.items)) {
      c.items = c.items.map((it: any) => (it && typeof it === "object" && typeof it.value === "number" ? { ...it, value: String(it.value) } : it));
      return c;
    }
    if (c.type === "chart") {
      for (const key of ["series", "bars", "lines"]) {
        if (Array.isArray(c[key])) c[key] = c[key].map((ser: any) => (ser && typeof ser === "object" && !ser.data && Array.isArray(ser.values) ? (note(`${path}.${key}`, "series values became data"), { ...ser, data: ser.values, values: undefined }) : ser));
      }
      return c;
    }
    return c;
  };

  const walk = (node: any, path: string) => {
    if (!node || typeof node !== "object") return;
    if (node.content) node.content = fixContent(node.content, `${path}.content`);
    if (Array.isArray(node.children)) node.children.forEach((ch: any, i: number) => walk(ch, `${path}.children[${i}]`));
  };

  if (Array.isArray(deck.slides)) {
    deck.slides = deck.slides.map((slide: any, i: number) => {
      if (!slide || typeof slide !== "object") return slide;
      if (!slide.body && (slide.content || slide.children)) {
        note(`slides[${i}]`, "content/children wrapped in body");
        const { content, children, direction, gap, ...rest } = slide;
        slide = { ...rest, body: { ...(direction ? { direction } : {}), ...(gap != null ? { gap } : {}), ...(content ? { content } : {}), ...(children ? { children } : {}) } };
      }
      if (slide.body) walk(slide.body, `slides[${i}].body`);
      return slide;
    });
  }
  return deck;
}

/**
 * Validate a deck. Never throws. Errors carry a JSON path, a message, and a fix hint
 * where one is known. Warnings are non-fatal lint findings. Common near-miss shapes
 * are normalized first (see normalizeDeck); `deck` in the result is the normalized deck.
 */
export function validateDeck(rawInput: unknown): ValidationResult {
  const errors: DeckIssue[] = [];
  const warnings: DeckIssue[] = [];
  if (!rawInput || typeof rawInput !== "object" || Array.isArray(rawInput)) {
    return { ok: false, errors: [{ path: "deck", message: "Deck must be a JSON object", fix: 'Send { "slides": [ ... ] }' }], warnings };
  }
  const normalized: DeckIssue[] = [];
  const input = normalizeDeck(rawInput, normalized);
  for (const n of normalized) warnings.push({ path: n.path, message: `Normalized: ${n.message}`, fix: "Accepted; write the canonical shape next time" });
  const parsed = DeckSchema.safeParse(input);
  if (!parsed.success) {
    for (const issue of parsed.error.issues) {
      const message = issue.message;
      errors.push({ path: pathString(issue.path as PropertyKey[]), message, fix: fixFor(message) });
    }
  }
  lint(input, warnings, errors);
  const uniq = dedupe(errors);
  if (!parsed.success || uniq.length) return { ok: false, errors: uniq, warnings };
  return { ok: true, deck: parsed.data as unknown as DeckJson, errors: [], warnings };
}

function dedupe(issues: DeckIssue[]): DeckIssue[] {
  const seen = new Set<string>();
  return issues.filter((i) => { const k = i.path + "|" + i.message; if (seen.has(k)) return false; seen.add(k); return true; });
}

/** Parse or throw a DeckValidationError listing every problem. */
export function parseDeck(input: unknown): DeckJson {
  const r = validateDeck(input);
  if (!r.ok) throw new DeckValidationError(r.errors, r.warnings);
  return r.deck!;
}

export class DeckValidationError extends Error {
  readonly errors: DeckIssue[];
  readonly warnings: DeckIssue[];
  constructor(errors: DeckIssue[], warnings: DeckIssue[] = []) {
    super("Invalid deck:\n" + errors.map((e) => `  ${e.path}: ${e.message}${e.fix ? ` (fix: ${e.fix})` : ""}`).join("\n"));
    this.name = "DeckValidationError";
    this.errors = errors;
    this.warnings = warnings;
  }
}

/** JSON Schema (draft 2020-12) for the deck format, for tool definitions and editors. */
export function deckJsonSchema(): Record<string, unknown> {
  const schema = z.toJSONSchema(DeckSchema, { target: "draft-2020-12", unrepresentable: "any", cycles: "ref" }) as Record<string, unknown>;
  schema.$id = "https://bankops.ai/schema/deck.json";
  schema.title = "BankOps deck";
  schema.description = "A slide deck: a list of slides, each with a 12-column nested grid body whose leaves carry typed content.";
  return schema;
}
