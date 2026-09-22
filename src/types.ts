/**
 * Deck JSON — the canonical wire format. A deck is a list of slides; each slide has a
 * 12-column nested grid body whose leaves carry typed content. Isomorphic, no imports.
 */

export interface DeckJson {
  /** Style preset name ("corporate" | "minimal" | "dark" | "warm") or style overrides */
  style?: string | Record<string, unknown>;
  /** Optional deck title (used for file naming and metadata) */
  title?: string;
  slideSize?: { width?: number; height?: number };
  slides: SlideSpec[];
}

export interface SlideSpec {
  /** Optional stable id for patch-based editing */
  id?: string;
  type?: "grid";
  title?: string;
  subtitle?: string;
  sectionLabel?: string;
  footer?: string;
  margin?: number;
  _task?: string;
  body?: GridNode;
}

export interface GridNode {
  /** Optional stable id for patch-based editing */
  id?: string;
  span?: number;
  direction?: "row" | "col";
  gap?: number;
  padding?: number;
  paddingTop?: number;
  paddingRight?: number;
  paddingBottom?: number;
  paddingLeft?: number;
  margin?: number;
  marginTop?: number;
  marginRight?: number;
  marginBottom?: number;
  marginLeft?: number;
  header?: string;
  subheader?: string;
  subfooter?: string;
  background?: string;
  backgroundOpacity?: number;  // 0-100, default 100
  border?: string | { color: string; width?: number; style?: "solid" | "dash" | "dot" | "dashDot" | "lgDash" | "lgDashDot" };
  borderTop?: { color: string; width?: number; style?: "solid" | "dash" | "dot" | "dashDot" | "lgDash" | "lgDashDot" };
  borderRight?: { color: string; width?: number; style?: "solid" | "dash" | "dot" | "dashDot" | "lgDash" | "lgDashDot" };
  borderBottom?: { color: string; width?: number; style?: "solid" | "dash" | "dot" | "dashDot" | "lgDash" | "lgDashDot" };
  borderLeft?: { color: string; width?: number; style?: "solid" | "dash" | "dot" | "dashDot" | "lgDash" | "lgDashDot" };
  shadow?: {
    color?: string;     // hex color, default #000000
    alpha?: number;     // 0-1 opacity, default 0.5
    blur?: number;      // points, default 4
    dist?: number;      // points, default 3
    dir?: number;       // degrees 0-360, default 45
    size?: number;      // percentage, default 100
  };
  childShape?: "chevron" | "pyramid";
  childShapeColor?: string;
  children?: GridNode[];
  content?: ContentSpec;
  $component?: string;
  _task?: string;
  _sources?: {
    url?: string;       // URL of the source (web page, API endpoint, document)
    api?: string;       // API service name (e.g. "FMP", "World Bank", "SEC EDGAR")
    label?: string;     // Human-readable citation label (e.g. "Company 10-K, FY2025")
    date?: string;      // Date the data was retrieved (ISO format)
  }[];
}

// ─── Content Types ───

export interface TextRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  fontSize?: number;   // hundredths-pt (1400 = 14pt)
  font?: string;
  bullet?: boolean;    // paragraph-level: when true on first run of a paragraph, that paragraph is bulleted
  bulletLevel?: number;  // 0 = top bullet (•), 1 = sub-bullet (–), etc.
}

export interface BulletStyleOverrides {
  char?: string;           // bullet character (e.g. "•", "–", "▸", "■")
  color?: string;          // hex color for bullet character
  indent?: number;         // hanging indent in inches (default ~0.375)
  marginLeft?: number;     // left margin in inches (default ~0.375)
  spaceBefore?: number;    // space between bullet items in points
}

export interface TextContent {
  type: "text";
  text?: string;         // Plain text fallback (used when runs absent)
  runs?: TextRun[];      // Rich text runs (takes precedence)
  font?: string;
  fontSize?: number;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  textTransform?: "uppercase";  // CSS text-transform; PPTX cap="all"
  align?: "l" | "ctr" | "r";
  anchor?: "t" | "ctr" | "b";
  lineSpacing?: number;  // e.g. 1.0 = single (default), 1.5 = 1.5×, 0.8 = tight
  bulletStyle?: BulletStyleOverrides;  // per-zone bullet formatting overrides
  paddingTop?: number;     // text inset in inches (default ~0.05)
  paddingRight?: number;   // text inset in inches (default ~0.1)
  paddingBottom?: number;  // text inset in inches (default ~0.05)
  paddingLeft?: number;    // text inset in inches (default ~0.1)
}

export interface StatGridContent {
  type: "statGrid";
  items: { value: string; label: string; sublabel?: string }[];
  anchor?: "t" | "ctr" | "b";
}

export interface CardGridContent {
  type: "cardGrid";
  items: { title: string; lines: string[] }[];
}

export interface TableCellImage {
  type: "image";
  source?: "search" | "ai" | "upload";
  searchTerm?: string;
  prompt?: string;
  url?: string;
  buffer?: unknown;
  $search?: string;
}

export type TableCell = string | TableCellImage;

export interface TableContentSpec {
  type: "table";
  data: {
    headers?: string[];
    rows: TableCell[][];
    colWidths?: number[];      // 12-column grid spans (must sum to 12)
    rowHeight?: number;        // Body row height in inches (default: 0.4)
    headerHeight?: number;     // Header row height in inches (default: 0.4)
    fontSize?: number;         // Body font size override (hundredths-pt, e.g. 1100)
    headerFontSize?: number;   // Header font size override (hundredths-pt)
    align?: Record<number, "l" | "ctr" | "r">;  // Per-column alignment overrides
    colAlign?: ("l" | "ctr" | "r")[];             // Array shorthand for align (normalized to align)
    overflowWrap?: "normal" | "anywhere" | "break-word";  // CSS overflow-wrap for cell text (default: "normal")
    verticalHeaders?: boolean;   // When true, first column is styled as a header column
    summaryRows?: number;         // Number of bottom rows styled as summary rows
  };
}

/** Normalize table data: convert colAlign array → align record, treat headers:[] as no headers */
export function normalizeTableData<T extends { headers?: string[]; align?: Record<number, "l" | "ctr" | "r">; colAlign?: ("l" | "ctr" | "r")[] }>(data: T): T {
  const d = { ...data };
  // headers: [] → treat as no headers
  if (d.headers && d.headers.length === 0) {
    d.headers = undefined;
  }
  // colAlign array → align record
  if (d.colAlign && !d.align) {
    const align: Record<number, "l" | "ctr" | "r"> = {};
    for (let i = 0; i < d.colAlign.length; i++) {
      align[i] = d.colAlign[i];
    }
    d.align = align;
  }
  delete d.colAlign;
  return d;
}

export interface ProfileContentSpec {
  type: "profile";
  name: string;
  items: string[];
}

export interface ImageContent {
  type: "image";
  source?: "ai" | "search" | "upload";
  prompt?: string;
  searchTerm?: string;
  path?: string;
  url?: string;
  buffer?: unknown;
  width?: number | string;
  height?: number | string;
  anchor?: "t" | "ctr" | "b";
  objectFit?: "cover" | "contain" | "fill";
  imagePadding?: number;  // inset in inches, applied before fit calculation
  crop?: { l: number; t: number; r: number; b: number };
  ext?: string;
}

export interface IconContent {
  type: "icon";
  name: string;       // Iconify ID, e.g. "mdi:home"
  svg: string;        // Full SVG markup (stored after selection)
  color?: string;     // Hex fill color (default: theme body text color)
  size?: number;      // Percentage of zone (1-100, default: 80)
  align?: "top" | "middle" | "bottom";  // Vertical alignment (default: middle)
  objectFit?: "contain" | "cover" | "fill";  // How SVG fits zone (default: contain)
}

export interface ChartContent {
  type: "chart";
  chartType: "bar" | "line" | "pie" | "combo" | "waterfall" | "scatter" | "gauge";
  title?: string;
  categories?: string[];
  series?: { name?: string; data: (number | [number, number])[] }[];  // scatter uses [x, y] pairs
  items?: { name: string; value: number }[];
  bars?: { name?: string; data: number[] }[];
  lines?: { name?: string; data: number[] }[];
  horizontal?: boolean;
  stacked?: boolean;
  stackTotal?: number[];
  smooth?: boolean;
  area?: boolean;
  doughnut?: boolean;
  roseType?: "radius" | "area";
  startAngle?: number;
  showLabel?: boolean;
  labelPosition?: "outside" | "inside" | "center" | "none";
  showLegend?: boolean;
  legendPosition?: "bottom" | "left" | "right" | "top";
  innerRadius?: string;
  outerRadius?: string;
  centerX?: string;
  centerY?: string;
  dualAxis?: boolean;
  xMin?: number;
  xMax?: number;
  yMin?: number;
  yMax?: number;
  y2Min?: number;
  y2Max?: number;
  gridTop?: number;
  gridRight?: number;
  gridBottom?: number;
  gridLeft?: number;
  data?: number[];
  value?: number;
  min?: number;
  max?: number;
  label?: string;
  format?: string;
  // ── Bar-specific ──
  showSeriesLabels?: boolean;
  showTotalValues?: boolean;
  showBarValues?: boolean;
  valuePrefix?: string;
  valueSuffix?: string;
  valueDecimals?: number;
  valueThousands?: boolean;
  negativeFormat?: "minus" | "parens";
  barBorderColor?: string;
  barBorderWidth?: number;
  barColor?: string;
  categoryColors?: Record<number, string>;  // per-category bar color overrides (index → hex)
  // ── Line-specific ──
  showLineValues?: boolean;
  pointShape?: "circle" | "rect" | "triangle" | "diamond";
  pointSize?: number;
  pointBorderColor?: string;
  pointBorderWidth?: number;
  pointColor?: string;
  lineColor?: string;
  lineWidth?: number;
  showXAxis?: boolean;
  showYAxis?: boolean;
  showSecondaryYAxis?: boolean;
  showSplitLine?: boolean;
  // ── Secondary Y-axis formatting (combo dual-axis) ──
  y2ValuePrefix?: string;
  y2ValueSuffix?: string;
  y2ValueDecimals?: number;
  y2ValueThousands?: boolean;
  y2NegativeFormat?: "minus" | "parens";
  xAxisLabel?: string;
  xAxisLabelPosition?: "end" | "center";
  yAxisLabel?: string;
  yAxisLabelPosition?: "top" | "side";
  // ── Pie-specific ──
  pieBorderColor?: string;
  pieBorderWidth?: number;
  padAngle?: number;
  sliceColors?: Record<number, string>;  // per-slice pie color overrides (index → hex)
  showPieValues?: boolean;
  pieValuePosition?: "outside" | "inside";
  // ── PPTX Export Mode ──
  pptxExportMode?: "image" | "nativeExcel";
}

export interface PptxChartSeries {
  name: string;
  labels: string[];
  values: number[];
}

export interface PptxChartContent {
  type: "pptxChart";
  chartType: "bar" | "line" | "pie" | "scatter" | "combo";
  series: PptxChartSeries[];
  title?: string;
  showTitle?: boolean;
  showLegend?: boolean;
  legendPos?: "b" | "t" | "l" | "r" | "tr";
  chartColors?: string[];
  barDir?: "col" | "bar";
  barGrouping?: "clustered" | "stacked" | "percentStacked";
  barFillColor?: string;       // Override bar fill color (hex, e.g. "#C0504D")
  gapWidthPct?: number;        // PPTX gap width % (default 150). Space between bar groups.
  lineSmooth?: boolean;
  doughnut?: boolean;
  catAxisTitle?: string;
  valAxisTitle?: string;
  valAxisMinVal?: number;
  valAxisMaxVal?: number;
  valAxisLabelFormatCode?: string;
  showValue?: boolean;
  showPercent?: boolean;
  dataLabelFormatCode?: string;
  comboCharts?: {
    chartType: "bar" | "line" | "area";
    series: PptxChartSeries[];
    options?: { secondaryValAxis?: boolean; barGrouping?: string };
  }[];
  // Axis visibility
  catAxisHidden?: boolean;
  valAxisHidden?: boolean;
  // Major gridlines (per axis)
  valGridLine?: { style?: "solid" | "dash" | "dot" | "none"; color?: string; size?: number };
  catGridLine?: { style?: "solid" | "dash" | "dot" | "none"; color?: string; size?: number };
  // Combo dual-axis: secondary value axis tick formatting
  secondaryValAxisLabelFormatCode?: string;
  secondaryValAxisMinVal?: number;
  secondaryValAxisMaxVal?: number;
  // Data label styling (matches editor's on-bar/on-slice label render)
  dataLabelFontSize?: number;
  dataLabelColor?: string;
  dataLabelPosition?: "b" | "bestFit" | "ctr" | "inBase" | "inEnd" | "l" | "outEnd" | "r" | "t";
  showCatName?: boolean;
  showSerName?: boolean;
  // Chart-wide font sizes (editor's defaults: legend/axes 11pt, title 14pt)
  fontSize?: number;
  legendFontSize?: number;
  titleFontSize?: number;
  // Data point border (bars/slices)
  dataBorder?: { type?: "solid" | "dash"; pt?: number; color: string };
  // Line/scatter markers
  lineDataSymbol?: "circle" | "dash" | "diamond" | "dot" | "none" | "square" | "triangle";
  lineDataSymbolSize?: number;
  lineDataSymbolLineColor?: string;
  lineDataSymbolLineSize?: number;
}

export interface LineContentSpec {
  type: "line";
  color?: string;
  width?: number;
}

export interface TimelineEvent {
  date: string;          // ISO date string e.g. "2024-01-15"
  name: string;          // Bold heading
  detail: string;        // Body text (plain or bullet)
}

export interface TimelineContent {
  type: "timeline";
  events: TimelineEvent[];
  title?: string;
  // Layout
  topCount?: number;              // How many events on top (rest go below). Default: ceil(n/2)
  placement?: "split" | "top" | "bottom";  // Default: "split"
  boxWidth?: number;              // EMU override for text box width
  boxHeight?: number;             // EMU override for text box height
  boxGap?: number;                // Minimum gap between boxes in inches (default: auto)
  boxOffset?: number;             // Distance from timeline line to boxes in inches (default: auto ~9% of height)
  boxHeightPct?: number;           // Event box height as % of zone height (default: 28)
  boxBorderColor?: string;        // Event box border color. Default: #CBD5E1
  boxBorderWidth?: number;        // Event box border width in pt. Default: 0.5
  boxBackground?: string;         // Event box fill color. Default: #F0F4F8
  nameFont?: string;              // Event name font family
  nameFontSize?: number;          // Event name font size in pt (default: auto)
  nameBold?: boolean;             // Event name bold. Default: true
  nameItalic?: boolean;           // Event name italic. Default: false
  nameUnderline?: boolean;        // Event name underline. Default: false
  nameColor?: string;             // Event name color. Default: #1F2937
  nameAlign?: "l" | "ctr" | "r"; // Event name alignment. Default: ctr
  detailFont?: string;            // Detail text font family
  detailFontSize?: number;        // Detail text font size in pt (default: auto)
  detailBold?: boolean;           // Detail text bold. Default: false
  detailItalic?: boolean;         // Detail text italic. Default: false
  detailUnderline?: boolean;      // Detail text underline. Default: false
  detailColor?: string;           // Detail text color. Default: #6B7280
  // Styling
  lineColor?: string;             // Timeline line color. Default: style accent
  lineWidth?: number;             // Timeline line thickness (EMU). Default: 25400
  nodeColor?: string;             // Node fill color. Default: #FFFFFF
  nodeBorderColor?: string;       // Node border color. Default: lineColor
  nodeBorderWidth?: number;       // Node border thickness (EMU). Default: 19050
  connectorColor?: string;        // Elbow connector color. Default: lighter lineColor
  connectorWidth?: number;        // Connector thickness (EMU). Default: 9525
  dateFormat?: "month" | "year" | "month-year" | "quarter" | "quarter-year";  // Default: "month-year"
  datePlacement?: "auto" | "top" | "bottom";  // Where date labels appear. Default: "auto" (opposite side from box)
  dateFont?: string;              // Date label font family. Default: inherit from body
  dateFontSize?: number;          // Date label font size in pt. Default: 5
  dateColor?: string;             // Date label color. Default: lineColor
  dateBold?: boolean;             // Date label bold. Default: true
  dateItalic?: boolean;           // Date label italic. Default: false
  textMode?: "plain" | "bullet";  // How detail text renders. Default: "plain"
}

export type ContentSpec =
  | TextContent
  | StatGridContent
  | CardGridContent
  | TableContentSpec
  | ProfileContentSpec
  | ImageContent
  | IconContent
  | ChartContent
  | PptxChartContent
  | LineContentSpec
  | TimelineContent;

// ─── Layout Engine Output ───

export interface Bounds {
  x: number;
  y: number;
  cx: number;
  cy: number;
}

export interface LayoutTree {
  sectionLabelBounds: Bounds | null;
  titleBounds: Bounds | null;
  footerBounds: Bounds | null;
  body: LayoutNodeResult | null;
}

export interface LayoutNodeResult {
  bounds: Bounds;
  background: string | null;
  backgroundOpacity: number | null;
  border: string | { color: string; width?: number; style?: "solid" | "dash" | "dot" | "dashDot" | "lgDash" | "lgDashDot" } | null;
  borderTop: { color: string; width?: number; style?: "solid" | "dash" | "dot" | "dashDot" | "lgDash" | "lgDashDot" } | null;
  borderRight: { color: string; width?: number; style?: "solid" | "dash" | "dot" | "dashDot" | "lgDash" | "lgDashDot" } | null;
  borderBottom: { color: string; width?: number; style?: "solid" | "dash" | "dot" | "dashDot" | "lgDash" | "lgDashDot" } | null;
  borderLeft: { color: string; width?: number; style?: "solid" | "dash" | "dot" | "dashDot" | "lgDash" | "lgDashDot" } | null;
  header: string | null;
  headerBounds: Bounds | null;
  subheader: string | null;
  subheaderBounds: Bounds | null;
  subheaderLineBounds: Bounds | null;
  subfooter: string | null;
  subfooterBounds: Bounds | null;
  contentBounds: Bounds | null;
  shadow: GridNode["shadow"] | null;
  content: ContentSpec | null;
  childShape: string | null;
  childShapeColor: string | null;
  children: LayoutNodeResult[] | null;
  _node?: GridNode;
}

// ─── Style Tokens (output of SlideStyle.toInternal()) ───

export interface StyleTokens {
  sectionLabel: { font: string; fontSize: number; bold: boolean; italic: boolean; color: string; x: number; y: number };
  mainTitle: { font: string; fontSize: number; bold: boolean; color: string; x: number; y: number };
  bodyText: { font: string; fontSize: number; color: string; align: string; lineSpacing?: number | null; bold?: boolean; italic?: boolean; fontFamily?: string };
  bullet: { char: string; indent: number; marginLeft: number; bulletFont: string | null; bulletColor: string | null; font: string | null; fontSize: number; color: string; lineSpacing?: number | null; fontFamily?: string };
  footnote: { font: string; fontSize: number; color: string; y: number };
  divider: { numberFont: string; numberFontSize: number; numberColor: string; numberX: number; numberY: number; titleFont: string; titleFontSize: number; titleColor: string; titleY: number };
  palette: { primaryText: string; secondaryText: string; accent: string; darkFill: string; border: string };
  tableHeader: { fillColor: string; textColor: string; font: string | null; fontSize: number; bold: boolean; borderBottom?: { color: string; width: number } };
  tableVerticalHeader: { fillColor: string; textColor: string; font: string | null; fontSize: number; bold: boolean };
  tableBody: { alternateRows: boolean; fillEven: string; fillOdd: string; textColor: string; font: string | null; fontSize: number };
  tableSummaryRow: {
    fillColor: string;
    textColor: string;
    font: string | null;
    fontSize: number;
    bold: boolean;
    borderTop: { color: string; width: number };
    borderBottom: { color: string; width: number };
    borderLeft: { color: string; width: number };
    borderRight: { color: string; width: number };
  };
  /** mode: "horizontal" = rules between rows only (default), "grid" = every cell edge, "none" */
  tableBorder: { color: string; width: number; mode?: "horizontal" | "grid" | "none" };
  table?: { overflowWrap?: "normal" | "anywhere" | "break-word" };
  slide: { background: string };
  headerBar: { textColor: string };
  header: { fontSize: number | null; color: string | null; background: string | null; align: string; borderRadius: number | null; fontFamily?: string | null; bold?: boolean | null; italic?: boolean | null };
  subheader: { fontSize: number | null; bold: boolean; color: string | null; align: string; lineColor: string | null; lineWidth: number; fontFamily?: string | null; italic?: boolean | null };
  subfooter: { fontSize: number | null; italic: boolean; color: string | null; fontFamily?: string | null };
  card: { fill: string; borderWidth: number; paddingLeft: number; paddingTop: number };
  spacing: { columnGap: number; bulletSpacing: number; cardGap: number; headerBarHeight: number; contentGap: number; zoneGap: number; subheaderLineWidth: number };
  titleSlide: { titleSize: number; subtitleSize: number };
  chart: { colors: string[] | null; fontFamily: string | null; fontSize: number | null; barBorderColor: string | null; barBorderWidth: number | null; bar?: { showXAxis?: boolean; showYAxis?: boolean; showSplitLine?: boolean; showSeriesLabels?: boolean; gridLeft?: number; gridRight?: number; gridTop?: number; gridBottom?: number }; pie?: { borderColor?: string; borderWidth?: number }; combo?: { showPrimaryYAxis?: boolean; showSecondaryYAxis?: boolean; showSplitLine?: boolean; showBarValues?: boolean; showLineValues?: boolean; barValueColor?: string; lineValueColor?: string } };
  statGrid?: {
    valueFont: string; valueSize: number; valueColor: string; valueBold: boolean; valueItalic: boolean;
    labelFont: string; labelSize: number | null; labelColor: string; labelBold: boolean; labelItalic: boolean;
    sublabelFont: string; sublabelSize: number | null; sublabelColor: string; sublabelBold: boolean; sublabelItalic: boolean;
    paddingTop?: number | null;
  };
  cardGrid?: {
    titleFont: string; titleSize: number | null; titleColor: string; titleBold: boolean; titleItalic: boolean;
    lineFont: string; lineSize: number | null; lineColor: string; lineBold: boolean; lineItalic: boolean;
    lineBullet: boolean; lineBulletIndent: number;
  };
}
