/**
 * Port of legacy/test-basic.js + legacy/test-grid.js (+ the SlideBuilder helper
 * tests from legacy/test-layout.js). Everything renders to a Buffer and is parsed
 * back with PptxParser; nothing is written to disk.
 */
import { test } from "node:test";
import assert from "node:assert/strict";
import sharp from "sharp";
import SlideBuilder from "../src/slide-builder.js";
import SlideLayout from "../src/layout.js";
import ChartGenerator from "../src/charts/chart-generator.js";
import PptxParser from "../src/pptx/pptx-parser.js";

const EMU = 914400;
const inchesOf = (emu: number) => +(emu / EMU).toFixed(2);

/** Legacy `bulletList` content → text with bullet runs. */
const bullets = (items: string[], extra: Record<string, unknown> = {}) => ({
  type: "text" as const,
  runs: items.map((text) => ({ text, bullet: true })),
  ...extra,
});

/** Tiny placeholder PNG (replaces the legacy demo logos/headshots). */
async function makePng(color: string, w = 60, h = 60): Promise<Buffer> {
  return sharp({ create: { width: w, height: h, channels: 3, background: color } }).png().toBuffer();
}

/** Count header bars (rounded rectangles filled with the darkFill color). */
function headerBarCount(slide: any, fills = ["#164556", "#2D2D2D"]): number {
  let n = 0;
  for (const sh of slide.shapes) {
    if (sh.geometry && sh.geometry.type === "roundRect" && sh.fill && sh.fill.type === "solid" && fills.includes(sh.fill.color || "")) n++;
  }
  return n;
}

async function parse(buf: Buffer): Promise<any> {
  return new PptxParser().parseBuffer(buf);
}

// ===== test-basic.js =====

test("builder: deck from scratch round-trips through the parser", async () => {
  const builder = new SlideBuilder();
  await builder.addGridSlide({
    title: "Test Presentation",
    body: { direction: "row", children: [
      { span: 1 },
      { span: 7, direction: "col", children: [
        { content: { type: "text", text: "Test Presentation", fontSize: 2800 } },
        { content: { type: "text", text: "Created by PPTX SDK", fontSize: 1400 } },
      ] },
      { span: 4 },
    ] },
  });
  builder
    .addContentSlide("Key Points", ["First bullet", "Second bullet", "Third bullet"])
    .addTwoColumnSlide("Comparison", ["Left point A", "Left point B"], ["Right point A", "Right point B"]);
  await builder.addGridSlide({
    title: "Data Summary",
    body: { direction: "row", children: [
      { span: 12, content: { type: "table", data: { headers: ["Name", "Value", "Status"],
        rows: [["Alpha", "100", "Active"], ["Beta", "200", "Pending"], ["Gamma", "300", "Done"]] } } },
    ] },
  });
  builder.addCompanyProfileSlide("Acme Corp", {
      Industry: "Technology",
      Founded: "2010",
      Revenue: "$500M",
      Employees: "2,500",
      HQ: "San Francisco, CA",
    })
    .addSectionDividerSlide(2, "Analysis")
    .addBlankSlide();

  const buf = await builder.toBuffer();
  assert.ok(buf.length > 5000, "output should be substantial");
  assert.equal(builder.pres.slides.length, 7);

  const pres = await parse(buf);
  assert.equal(pres.slides.length, 7, "round-trip should preserve 7 slides");
  assert.equal(pres.theme.name, "Office Theme");
});

test("builder: setTheme overrides theme colors and fonts", async () => {
  const builder = new SlideBuilder();
  builder.setTheme({ colors: { accent1: "#FF0000", dk1: "#1a1a2e" }, fonts: { title: "Georgia", body: "Arial" } });
  await builder.addGridSlide({
    title: "Custom Theme",
    body: { direction: "row", children: [
      { span: 1 },
      { span: 7, direction: "col", children: [
        { content: { type: "text", text: "Custom Theme", fontSize: 2800 } },
        { content: { type: "text", text: "Red accent", fontSize: 1400 } },
      ] },
      { span: 4 },
    ] },
  });
  assert.equal(builder.pres.theme.clrScheme.accent1, "#FF0000");
  assert.equal(builder.pres.theme.fontScheme.majorFont.latin, "Georgia");
  assert.equal(builder.pres.theme.fontScheme.minorFont.latin, "Arial");
  const buf = await builder.toBuffer();
  assert.ok(buf.length > 1000);
});

// ===== test-grid.js: layout engine =====

test("grid: two-column layout", () => {
  const tree = new SlideLayout().compute({ body: { direction: "row", children: [{ span: 6 }, { span: 6 }] } });
  assert.ok(tree.body!.children !== null);
  assert.equal(tree.body!.children!.length, 2);
  const [left, right] = tree.body!.children!;
  assert.ok(Math.abs(left.bounds.cx - right.bounds.cx) < EMU * 0.1, "columns should be equal width");
  assert.ok(left.bounds.x < right.bounds.x, "left column should be to the left");
  const ratio = left.bounds.cx / tree.body!.bounds.cx;
  assert.ok(ratio > 0.4 && ratio < 0.55, `each column ~half width, got ${(ratio * 100).toFixed(1)}%`);
});

test("grid: three equal columns", () => {
  const tree = new SlideLayout().compute({ body: { direction: "row", children: [{ span: 4 }, { span: 4 }, { span: 4 }] } });
  assert.equal(tree.body!.children!.length, 3);
  const widths = tree.body!.children!.map((c) => c.bounds.cx);
  assert.ok(Math.max(...widths) - Math.min(...widths) < EMU * 0.1, "all columns should be equal width");
});

test("grid: nested 2x2 quadrant", () => {
  const layout = new SlideLayout();
  const tree = layout.compute({
    body: { direction: "row", children: [
      { span: 6, direction: "col", children: [{ span: 6 }, { span: 6 }] },
      { span: 6, direction: "col", children: [{ span: 6 }, { span: 6 }] },
    ] },
  });
  assert.equal(tree.body!.children!.length, 2);
  assert.equal(tree.body!.children![0].children!.length, 2);
  assert.equal(tree.body!.children![1].children!.length, 2);
  const leaves = layout.flattenLeaves(tree);
  assert.equal(leaves.length, 4);
  const tl = tree.body!.children![0].children![0];
  const bl = tree.body!.children![0].children![1];
  const tr = tree.body!.children![1].children![0];
  assert.ok(tl.bounds.y < bl.bounds.y, "top-left should be above bottom-left");
  assert.ok(tl.bounds.x < tr.bounds.x, "left should be to the left of right");
});

test("grid: unequal spans (8+4)", () => {
  const tree = new SlideLayout().compute({ body: { direction: "row", children: [{ span: 8 }, { span: 4 }] } });
  const [wide, narrow] = tree.body!.children!;
  const ratio = wide.bounds.cx / narrow.bounds.cx;
  assert.ok(ratio > 1.8 && ratio < 2.2, `wide should be ~2x narrow, got ${ratio.toFixed(2)}x`);
});

test("grid: column direction stacks vertically", () => {
  const tree = new SlideLayout().compute({ body: { direction: "col", children: [{ span: 4 }, { span: 8 }] } });
  const [top, bottom] = tree.body!.children!;
  assert.ok(top.bounds.y < bottom.bounds.y, "top should be above bottom");
  assert.equal(top.bounds.cx, bottom.bounds.cx, "both should have the same width");
  const ratio = bottom.bounds.cy / top.bounds.cy;
  assert.ok(ratio > 1.8 && ratio < 2.2, `bottom should be ~2x taller, got ${ratio.toFixed(2)}x`);
});

test("grid: title + section label + footer reduce the body area", () => {
  const layout = new SlideLayout();
  const noTitle = layout.compute({ body: { direction: "row", children: [{ span: 12 }] } });
  const withTitle = layout.compute({ title: "My Title", sectionLabel: "Section", footer: "Source: test", body: { direction: "row", children: [{ span: 12 }] } });
  assert.ok(withTitle.titleBounds !== null);
  assert.ok(withTitle.sectionLabelBounds !== null);
  assert.ok(withTitle.footerBounds !== null);
  assert.ok(withTitle.body!.bounds.cy < noTitle.body!.bounds.cy, "body should be smaller with title+footer");
});

test("grid: header bars push content down", () => {
  const tree = new SlideLayout().compute({ body: { direction: "row", children: [{ span: 6, header: "Left Zone" }, { span: 6, header: "Right Zone" }] } });
  const [left, right] = tree.body!.children!;
  assert.ok(left.headerBounds !== null);
  assert.ok(left.contentBounds !== null);
  assert.ok(left.contentBounds!.y > left.headerBounds!.y + left.headerBounds!.cy, "content should be below header");
  assert.ok(right.headerBounds !== null);
});

test("grid: auto span distributes evenly with no explicit spans", () => {
  const tree = new SlideLayout().compute({ body: { direction: "row", children: [{}, {}, {}] } });
  assert.equal(tree.body!.children!.length, 3);
  const widths = tree.body!.children!.map((c) => c.bounds.cx);
  assert.ok(Math.max(...widths) - Math.min(...widths) < EMU * 0.1, "auto columns should be equal width");
});

test("grid: larger gap produces narrower cells", () => {
  const layout = new SlideLayout();
  const small = layout.compute({ body: { direction: "row", gap: 0.05, children: [{ span: 6 }, { span: 6 }] } });
  const large = layout.compute({ body: { direction: "row", gap: 0.5, children: [{ span: 6 }, { span: 6 }] } });
  assert.ok(large.body!.children![0].bounds.cx < small.body!.children![0].bounds.cx);
  const gapBetween = large.body!.children![1].bounds.x - (large.body!.children![0].bounds.x + large.body!.children![0].bounds.cx);
  assert.equal(inchesOf(gapBetween), 0.5);
});

test("grid: subheader and subfooter ordering", () => {
  const layout = new SlideLayout();
  let tree = layout.compute({ body: { direction: "row", children: [{ span: 6, subheader: "Revenue Details" }, { span: 6, header: "Header Bar" }] } });
  const [left, right] = tree.body!.children!;
  assert.ok(left.subheaderBounds !== null, "left should have subheader bounds");
  assert.equal(left.headerBounds, null, "left should NOT have header bounds");
  assert.ok(right.headerBounds !== null, "right should have header bounds");
  assert.equal(right.subheaderBounds, null, "right should NOT have subheader bounds");
  assert.ok(left.contentBounds!.y > left.subheaderBounds!.y + left.subheaderBounds!.cy, "content should be below subheader");

  tree = layout.compute({ body: { direction: "row", children: [{
    span: 12, header: "Main Header", subheader: "Sub-heading", subfooter: "Note: Values are estimates",
    direction: "row", children: [{ span: 6 }, { span: 6 }],
  }] } });
  const c = tree.body!.children![0];
  assert.ok(c.headerBounds && c.subheaderBounds && c.subfooterBounds);
  assert.ok(c.subheaderBounds!.y > c.headerBounds!.y + c.headerBounds!.cy, "subheader below header");
  assert.ok(c.subfooterBounds!.y > c.subheaderBounds!.y, "subfooter below subheader");
  assert.ok(c.children![0].bounds.y >= c.subheaderBounds!.y + c.subheaderBounds!.cy, "children start at or below subheader");
  assert.ok(c.children![0].bounds.y + c.children![0].bounds.cy <= c.subfooterBounds!.y, "children end at or above subfooter");

  tree = layout.compute({ body: { direction: "row", children: [{ span: 12, subfooter: "Source: Internal data" }] } });
  const node = tree.body!.children![0];
  assert.ok(node.subfooterBounds !== null);
  assert.ok(node.contentBounds!.cy < node.bounds.cy, "content area smaller with subfooter");
});

test("grid: every layout preset computes a body with leaves", () => {
  const names = Object.keys(SlideLayout.presets);
  assert.ok(names.length >= 6, "should have 6+ presets");
  const layout = new SlideLayout();
  for (const name of names) {
    const spec = SlideLayout.preset(name, { title: "Test" });
    assert.ok(spec !== null, `preset ${name} should return a spec`);
    const tree = layout.compute(spec!);
    assert.ok(tree.body !== null, `preset ${name} should produce a body`);
    assert.ok(layout.flattenLeaves(tree).length >= 1, `preset ${name} should have at least 1 leaf`);
  }
});

// ===== test-grid.js: chart generation =====

test("charts: every chart type renders a PNG buffer", async () => {
  const gen = new ChartGenerator({ width: 600, height: 400 });

  const barBuf = await gen.bar({ title: "Test Bar", categories: ["A", "B", "C"], series: [{ name: "S1", data: [10, 20, 30] }] });
  assert.ok(Buffer.isBuffer(barBuf));
  assert.ok(barBuf.length > 1000, `bar PNG should be substantial, got ${barBuf.length}`);

  const lineBuf = await gen.line({ categories: ["Q1", "Q2", "Q3"], series: [{ name: "L1", data: [100, 120, 140] }], smooth: true, area: true });
  assert.ok(Buffer.isBuffer(lineBuf));
  assert.ok(lineBuf.length > 1000);

  assert.ok(Buffer.isBuffer(await gen.pie({ items: [{ name: "X", value: 60 }, { name: "Y", value: 40 }], doughnut: true })));
  assert.ok(Buffer.isBuffer(await gen.waterfall({ categories: ["Start", "Add", "Sub", "End"], data: [100, 30, -20, 110] })));
  assert.ok(Buffer.isBuffer(await gen.combo({ categories: ["A", "B", "C"], bars: [{ name: "Rev", data: [50, 60, 70] }], lines: [{ name: "Margin", data: [25, 28, 30] }] })));
  assert.ok(Buffer.isBuffer(await gen.scatter({ series: [{ name: "Pts", data: [[1, 2], [3, 4], [5, 1]] }] })));
  assert.ok(Buffer.isBuffer(await gen.gauge({ value: 75, label: "Score" })));

  const smallBuf = await gen.withSize(300, 200).bar({ categories: ["A"], series: [{ name: "S", data: [10] }] });
  assert.ok(Buffer.isBuffer(smallBuf));
  assert.ok(smallBuf.length < barBuf.length, "smaller chart should be a smaller PNG");
});

// ===== test-grid.js: the integration deck =====

test("builder: multi-slide grid deck with charts, images, tables and profiles", async () => {
  const builder = new SlideBuilder();
  const charts = new ChartGenerator({ width: 700, height: 450 });

  const barChart = await charts.bar({
    title: "Revenue by Segment",
    categories: ["Software", "Services", "Hardware", "Support"],
    series: [{ name: "FY2024", data: [320, 180, 95, 120] }, { name: "FY2025E", data: [380, 210, 88, 145] }],
  });
  const comboChart = await charts.combo({
    title: "Revenue & Margin Trend",
    categories: ["Q1'24", "Q2'24", "Q3'24", "Q4'24", "Q1'25", "Q2'25"],
    bars: [{ name: "Revenue ($M)", data: [145, 162, 178, 195, 205, 218] }],
    lines: [{ name: "Gross Margin %", data: [62, 64, 63, 65, 66, 68] }],
    dualAxis: true,
  });
  const scatterChart = await charts.scatter({
    title: "EV/Revenue vs Growth",
    series: [{ name: "Peers", data: [[12, 5.2], [18, 8.1], [25, 12.3], [8, 3.5], [32, 15.8], [15, 6.9], [22, 10.5], [28, 13.1]] }],
  });

  // ---- Slide 1: Two columns with subheaders ----
  await builder.addGridSlide({
    title: "Two Columns with Sub-headings",
    body: { direction: "row", children: [
      { span: 6, header: "Financial Overview", subheader: "Revenue Breakdown", subfooter: "Source: Company filings, FY2025",
        content: bullets(["Total Revenue: $823M", "Software: $480M (58%)", "Services: $220M (27%)", "Hardware: $123M (15%)"]) },
      { span: 6, header: "Growth Metrics", subheader: "Year-over-Year Performance", subfooter: "Note: 2025 figures are estimates",
        content: { type: "statGrid", items: [
          { value: "18.2%", label: "Revenue Growth" }, { value: "66%", label: "Gross Margin" },
          { value: "118%", label: "Net Retention" }, { value: "22%", label: "FCF Margin" },
        ] } },
    ] },
  }, { debug: true });

  // ---- Slide 2: Nested sub-containers ----
  await builder.addGridSlide({
    title: "Nested Sub-containers",
    sectionLabel: "Advanced Layout Demo",
    body: { direction: "row", children: [
      { span: 6, header: "Market Intelligence", direction: "col", children: [
        { span: 6, subheader: "Total Addressable Market", subfooter: "Industry research, 2025",
          content: { type: "statGrid", items: [{ value: "$15.2B", label: "TAM" }, { value: "14.3%", label: "CAGR" }] } },
        { span: 6, subheader: "Competitive Landscape",
          content: bullets(["Northbeam — 22% market share", "Ridgeline — 18% market share", "Cobalt — 12% market share", "Others — 48% fragmented"]) },
      ] },
      { span: 6, header: "Revenue Trend", subheader: "FY2022-FY2025E", subfooter: "Management estimates for FY2025",
        content: { type: "image", buffer: barChart, width: "100%", height: "100%" } },
    ] },
  }, { debug: true });

  // ---- Slide 3: Three-level nesting with charts rendered at container size ----
  await builder.addGridSlide({
    title: "Deep Nesting: Header + Sub-containers + Charts",
    footer: "Confidential — For Discussion Purposes Only",
    body: { direction: "row", children: [
      { span: 4, header: "Key Metrics", subheader: "Performance Summary", direction: "col", children: [
        { span: 4, subheader: "Revenue", content: { type: "chart", chartType: "bar", categories: ["2022", "2023", "2024", "2025E"], series: [{ name: "Revenue", data: [520, 640, 780, 920] }] } },
        { span: 4, subheader: "Margins", content: { type: "chart", chartType: "line", categories: ["2022", "2023", "2024", "2025E"], series: [{ name: "Margin %", data: [58, 62, 65, 68] }], smooth: true } },
        { span: 4, subheader: "Revenue Mix", content: { type: "chart", chartType: "pie", items: [{ name: "Subscription", value: 65 }, { name: "Services", value: 25 }, { name: "Other", value: 10 }] } },
      ] },
      { span: 8, header: "Detailed Analysis", direction: "col", children: [
        { span: 5, subheader: "EBITDA Bridge Analysis", subfooter: "Amounts in $M",
          content: { type: "chart", chartType: "waterfall", title: "EBITDA Bridge (FY24 to FY25E)",
            categories: ["FY2024", "Revenue Growth", "Cost Savings", "Investments", "FX Impact", "FY2025E"], data: [180, 65, 22, -35, -12, 220] } },
        { span: 7, direction: "row", children: [
          { span: 6, subheader: "Geographic Split",
            content: { type: "chart", chartType: "pie", title: "Revenue Mix", doughnut: true,
              items: [{ name: "North America", value: 42 }, { name: "Europe", value: 28 }, { name: "Asia Pacific", value: 18 }, { name: "Rest of World", value: 12 }] } },
          { span: 6, subheader: "Peer Valuation", subfooter: "Source: market data, Feb 2026",
            content: { type: "table", data: { headers: ["Company", "EV/Rev", "Growth"],
              rows: [["Northbeam", "12.3x", "18%"], ["Ridgeline", "8.5x", "12%"], ["Cobalt", "5.2x", "8%"], ["Stonegate", "10.5x", "22%"]] } } },
        ] },
      ] },
    ] },
  }, { debug: true });

  // ---- Slide 4: Dashboard ----
  await builder.addGridSlide({
    title: "Executive Dashboard",
    sectionLabel: "Q2 2025 Review",
    footer: "Prepared for Board of Directors",
    body: { direction: "col", children: [
      { span: 3, direction: "row", children: [
        { span: 3, subheader: "Revenue", content: { type: "statGrid", items: [{ value: "$218M", label: "+6.3% QoQ" }] } },
        { span: 3, subheader: "EBITDA", content: { type: "statGrid", items: [{ value: "$58M", label: "26.6% margin" }] } },
        { span: 3, subheader: "Customers", content: { type: "statGrid", items: [{ value: "4,820", label: "+1,247 net adds" }] } },
        { span: 3, subheader: "ARR", content: { type: "statGrid", items: [{ value: "$872M", label: "118% NRR" }] } },
      ] },
      { span: 9, direction: "row", children: [
        { span: 8, header: "Revenue & Margin", subheader: "Quarterly Trend", subfooter: "Q1-Q2 2025 are preliminary estimates",
          content: { type: "image", buffer: comboChart, width: "100%", height: "100%" } },
        { span: 4, header: "Highlights", subheader: "Operational KPIs",
          content: bullets(["Win Rate: 34%", "Avg Deal Size: $48K", "Sales Cycle: 42 days", "Churn: 1.2%", "NPS: 72"]) },
      ] },
    ] },
  }, { debug: true });

  // ---- Slide 5: Peer comparison ----
  await builder.addGridSlide({
    title: "Peer Comparison",
    sectionLabel: "Valuation",
    footer: "Source: market data, as of Feb 2026",
    body: { direction: "row", children: [
      { span: 7, header: "Valuation Overview", subheader: "EV/Revenue vs Revenue Growth", subfooter: "Bubble size indicates market cap",
        content: { type: "image", buffer: scatterChart, width: "100%", height: "100%" } },
      { span: 5, header: "Comparable Companies", direction: "col", children: [
        { span: 7, subheader: "Public Peers",
          content: { type: "table", data: { headers: ["Company", "EV/Rev", "Growth"],
            rows: [["Northbeam", "12.3x", "18%"], ["Ridgeline", "8.5x", "12%"], ["Cobalt", "5.2x", "8%"], ["Stonegate", "10.5x", "22%"]] } } },
        { span: 5, subheader: "Implied Valuation", subfooter: "Based on peer median multiples",
          content: { type: "statGrid", items: [{ value: "8.7x", label: "Median EV/Rev" }, { value: "$7.2B", label: "Implied EV" }] } },
      ] },
    ] },
  }, { debug: true });

  // ---- Slide 6: Company profiles (2x2 entity cards in two columns) ----
  const logoColors = ["#4472C4", "#ED7D31", "#70AD47", "#FFC000", "#5B9BD5", "#A5A5A5"];
  const logos: Buffer[] = [];
  for (const c of logoColors) logos.push(await makePng(c, 120, 120));
  const profileCard = (logo: Buffer, name: string, items: string[]) => ({
    span: 6, direction: "row" as const, gap: 0.08, children: [
      { span: 4, content: { type: "image" as const, buffer: logo, width: 120, height: 120 } },
      { span: 8, content: { type: "profile" as const, name, items } },
    ],
  });
  await builder.addGridSlide({
    title: "Key Players & Strategic Acquirers",
    sectionLabel: "Market Landscape",
    footer: "Source: Company filings, press releases",
    body: { direction: "row", children: [
      { span: 6, header: "Construction Software Leaders", direction: "col", gap: 0.1, children: [
        { span: 6, direction: "row", gap: 0.1, children: [
          profileCard(logos[0], "Northbeam", ["$1.2B revenue (FY24)", "Pure-play construction", "16,000+ customers"]),
          profileCard(logos[1], "Ridgeline", ["$5.8B revenue (FY24)", "BIM + Construction Cloud", "Integrated design-to-build"]),
        ] },
        { span: 6, direction: "row", gap: 0.1, children: [
          profileCard(logos[2], "Cobalt Build", ["Enterprise project controls", "Global infrastructure focus", "Part of a larger suite"]),
          profileCard(logos[3], "Stonegate", ["$3.7B revenue (FY24)", "ERP + field operations", "Recent tuck-in acquisition"]),
        ] },
      ] },
      { span: 6, header: "Strategic Acquirers", direction: "col", gap: 0.1, children: [
        { span: 6, direction: "row", gap: 0.1, children: [
          profileCard(logos[4], "Larkspur Systems", ["Infrastructure engineering", "$1.1B revenue (FY24)", "Digital twin platform"]),
          profileCard(logos[5], "Tidewater Software", ["Industrial group subsidiary", "Cost-planning platform", "European market leader"]),
        ] },
        { span: 6, direction: "row", gap: 0.1, children: [
          profileCard(logos[0], "Halcyon AB", ["Sensor + software conglomerate", "$5.4B revenue (FY24)", "Smart build solutions"]),
          profileCard(logos[1], "Brightwater", ["AEC software group", "Design and collaboration tools", "European HQ, global ops"]),
        ] },
      ] },
    ] },
  }, { debug: true });

  // ---- Slide 7: Table of contents ----
  const tocSections = ["Executive Summary", "Market Overview", "Competitive Landscape", "Financial Analysis", "Valuation", "Transaction Considerations", "Risk Factors", "Appendix"];
  const tocRows = tocSections.map((s, i) => ({
    direction: "row" as const, gap: 0.05, children: [
      { span: 2, content: { type: "text" as const, text: `${i < 9 ? "0" : ""}${i + 1}.`, font: "Arial Bold", fontSize: 2200, color: "#D57F5B", align: "ctr" as const, anchor: "ctr" as const } },
      { span: 10, content: { type: "text" as const, text: s, font: "Arial", fontSize: 2200, color: "#405363", anchor: "ctr" as const } },
    ],
  }));
  await builder.addGridSlide({
    title: "Table of Contents",
    sectionLabel: "Overview",
    body: { direction: "row", gap: 0, children: [
      { span: 8, direction: "col", gap: 0.05, padding: 0.15, children: tocRows },
      { span: 4, background: "#D9D9D9" },
    ] },
  });

  // ---- Slide 8: Section title ----
  await builder.addGridSlide({
    body: { direction: "row", gap: 0, children: [
      { span: 8, direction: "col", padding: 0.5, children: [
        { span: 6, content: { type: "text", text: "01.", font: "Arial Bold", fontSize: 6400, color: "#405363", anchor: "b" } },
        { span: 6, content: { type: "text", text: "EXECUTIVE SUMMARY", font: "Arial Bold", fontSize: 2000, color: "#D57F5B", anchor: "t" } },
      ] },
      { span: 4, background: "#8FA4B3" },
    ] },
  });

  // ---- Slide 9: Credentials overview ----
  const teamMembers = [
    { name: "James Sullivan", title: "Managing Director", exp: "22 years experience", color: "#405363" },
    { name: "Sarah Chen", title: "Managing Director", exp: "18 years experience", color: "#D57F5B" },
    { name: "Michael Torres", title: "Vice President", exp: "12 years experience", color: "#4472C4" },
    { name: "Emily Park", title: "Vice President", exp: "9 years experience", color: "#70AD47" },
  ];
  const teamCards: any[] = [];
  for (const tm of teamMembers) {
    teamCards.push({ direction: "row", gap: 0.08, children: [
      { span: 4, content: { type: "image", buffer: await makePng(tm.color, 80, 80), width: 80, height: 80 } },
      { span: 8, content: { type: "profile", name: tm.name, items: [tm.title, tm.exp] } },
    ] });
  }
  const practiceAreas = ["Enterprise\nSoftware", "Cloud &\nSaaS", "Data &\nAnalytics", "Cybersecurity", "AI / ML"];
  const paColors = ["#4472C4", "#ED7D31", "#70AD47", "#C00000", "#7030A0"];
  const practiceIcons: any[] = [];
  for (let i = 0; i < practiceAreas.length; i++) {
    practiceIcons.push({ direction: "col", gap: 0.02, children: [
      { span: 6, content: { type: "image", buffer: await makePng(paColors[i]), width: 60, height: 60 } },
      { span: 6, content: { type: "text", text: practiceAreas[i], font: "Arial", fontSize: 800, color: "#405363", align: "ctr", anchor: "t" } },
    ] });
  }
  await builder.addGridSlide({
    title: "Credentials Overview",
    sectionLabel: "Harbor Point Advisors",
    body: { direction: "col", children: [
      { span: 6, direction: "row", gap: 0.1, children: [
        { span: 6, header: "Harbor Point by the Numbers",
          content: { type: "statGrid", items: [
            { value: "$48B+", label: "Transaction Value", sublabel: "Since 2005" },
            { value: "320+", label: "Deals Closed", sublabel: "M&A + Capital Raises" },
            { value: "85", label: "Senior Bankers", sublabel: "Across 12 offices" },
            { value: "#3", label: "League Table Rank", sublabel: "Tech M&A (2024)" },
          ] } },
        { span: 6, header: "Harbor Point Technology Practice Areas", direction: "col", gap: 0.05, padding: 0.1, children: [
          { span: 6, direction: "row", gap: 0.05, children: practiceIcons.slice(0, 3) },
          { span: 6, direction: "row", gap: 0.05, children: practiceIcons.slice(3, 5) },
        ] },
      ] },
      { span: 6, direction: "row", gap: 0.1, children: [
        { span: 6, header: "Representative Transactions",
          content: { type: "cardGrid", items: [
            { title: "TechBuild Inc.", lines: ["Acquired by Northbeam", "September 2025"] },
            { title: "FieldOps Pro", lines: ["Investment from Crestline Equity", "June 2025"] },
            { title: "SiteSync Corp", lines: ["Merged with Ridgeline", "January 2025"] },
            { title: "SmartBuild Co", lines: ["Recapitalized by Bayview Capital", "In-market"] },
          ] } },
        { span: 6, header: "Technology Team", direction: "col", gap: 0.08, children: [
          { span: 6, direction: "row", gap: 0.08, children: [teamCards[0], teamCards[1]] },
          { span: 6, direction: "row", gap: 0.08, children: [teamCards[2], teamCards[3]] },
        ] },
      ] },
    ] },
  });

  // ---- Slide 10: Client quotes ----
  const quoteCell = (quote: string, author: string, role: string, company: string) => ({
    direction: "col" as const, gap: 0.05, padding: 0.15, border: "#999999", children: [
      { span: 2, content: { type: "text" as const, text: "“", font: "Georgia", fontSize: 4800, color: "#D57F5B", bold: true, align: "l" as const, anchor: "b" as const } },
      { span: 5, content: { type: "text" as const, text: quote, font: "Georgia", fontSize: 1200, color: "#405363", align: "l" as const, anchor: "t" as const } },
      { span: 1, content: { type: "line" as const, color: "#D57F5B", width: 12700 } },
      { span: 2, content: { type: "text" as const, text: `${author}\n${role}, ${company}`, font: "Arial", fontSize: 1000, color: "#666666", align: "l" as const, anchor: "t" as const } },
    ],
  });
  await builder.addGridSlide({
    title: "What Our Clients are Saying",
    sectionLabel: "Client Testimonials",
    body: { direction: "col", gap: 0.1, children: [
      { span: 7, direction: "row", gap: 0.2, children: [
        quoteCell("Harbor Point's deep sector expertise and relentless execution made all the difference in our strategic review process. They delivered results that exceeded our expectations.", "Sarah Chen", "CEO", "CloudSync Technologies"),
        quoteCell("The team brought an unparalleled understanding of the software landscape. Their guidance through our Series D was instrumental in achieving a valuation that reflected our true potential.", "James Morrison", "CFO", "DataVault Systems"),
      ] },
      { span: 3, background: "#F5F0EB", direction: "row", gap: 0.2, padding: 0.15, children: [
        { span: 4, content: { type: "text", text: "\"The most trusted advisor in enterprise software M&A.\"", font: "Georgia", fontSize: 1400, color: "#405363", align: "ctr", anchor: "ctr" } },
        { span: 4, content: { type: "text", text: "— Private Equity Quarterly, 2025", font: "Arial", fontSize: 1000, color: "#999999", align: "ctr", anchor: "ctr" } },
        { span: 4, content: { type: "text", text: "\"Harbor Point consistently delivers best-in-class outcomes.\"", font: "Georgia", fontSize: 1400, color: "#405363", align: "ctr", anchor: "ctr" } },
      ] },
    ] },
  });

  // ---- Slide 11: Section title ----
  await builder.addGridSlide({
    body: { direction: "row", gap: 0, children: [
      { span: 8, direction: "col", padding: 0.5, children: [
        { span: 6, content: { type: "text", text: "02.", font: "Arial Bold", fontSize: 6400, color: "#405363", anchor: "b" } },
        { span: 6, content: { type: "text", text: "INDUSTRY OVERVIEW", font: "Arial Bold", fontSize: 2000, color: "#D57F5B", anchor: "t" } },
      ] },
      { span: 4, background: "#8FA4B3" },
    ] },
  });

  // ---- Slide 12: Market update (chart + icon trend cells) ----
  const trendColors = ["#4472C4", "#70AD47", "#ED7D31", "#7030A0"];
  const trendIcons: Buffer[] = [];
  for (const c of trendColors) trendIcons.push(await makePng(c));
  const trendCell = (iconBuf: Buffer, title: string, items: string[]) => ({
    direction: "row" as const, gap: 0.06, children: [
      { span: 2, content: { type: "image" as const, buffer: iconBuf, width: 60, height: 60 } },
      { span: 10, direction: "col" as const, gap: 0, children: [
        { span: 3, content: { type: "text" as const, text: title, font: "Arial Bold", fontSize: 1100, color: "#405363", align: "l" as const, anchor: "b" as const } },
        { span: 9, content: bullets(items, { fontSize: 1200, lineSpacing: 1.0 }) },
      ] },
    ],
  });
  await builder.addGridSlide({
    title: "AdTech Market Update",
    sectionLabel: "Industry Overview",
    body: { direction: "col", gap: 0.1, children: [
      { span: 6, header: "Global AdTech Market", content: { type: "chart", chartType: "bar", title: "Global AdTech Market Size ($B)",
        categories: ["2020", "2021", "2022", "2023", "2024", "2025E", "2026E", "2027E"],
        series: [{ name: "Market Size", data: [378, 455, 520, 590, 680, 780, 890, 1020] }] } },
      { span: 6, header: "Select Market Trends", direction: "col", gap: 0.08, children: [
        { span: 6, direction: "row", gap: 0.1, children: [
          trendCell(trendIcons[0], "AI-Powered Optimization", ["Generative AI for ad creative at scale", "Real-time bidding optimization +35% ROAS", "Predictive audience modeling"]),
          trendCell(trendIcons[1], "Connected TV (CTV) Growth", ["CTV ad spend up 22% YoY to $31B", "Programmatic CTV now 68% of total", "Measurement standards maturing"]),
        ] },
        { span: 6, direction: "row", gap: 0.1, children: [
          trendCell(trendIcons[2], "Privacy-First Advertising", ["Cookie deprecation driving innovation", "Contextual targeting revival (+45%)", "Clean room adoption accelerating"]),
          trendCell(trendIcons[3], "Identity Resolution", ["First-party data strategies critical", "Universal ID frameworks consolidating", "Cross-device graph accuracy at 92%"]),
        ] },
      ] },
    ] },
  });

  // ---- Slide 13: Chart quadrant ----
  await builder.addGridSlide({
    title: "AdTech Industry Dynamics and Insights",
    sectionLabel: "Industry Overview",
    body: { direction: "col", gap: 0.1, children: [
      { span: 6, direction: "row", gap: 0.1, children: [
        { header: "AI Enhancing Efficiency in Advertising", direction: "row", gap: 0.08, children: [
          { content: { type: "chart", chartType: "bar", title: "GenAI Ad Spend ($B)", categories: ["2023", "2024", "2025E", "2026E", "2027E"], series: [{ name: "Spend", data: [2.1, 4.8, 8.5, 14.2, 22.0] }] } },
          { content: { type: "chart", chartType: "line", title: "Ad Production Cost Index", categories: ["2021", "2022", "2023", "2024", "2025E"], series: [{ name: "Cost Index", data: [100, 94, 82, 68, 55] }], smooth: true } },
        ] },
        { header: "Global Ad Spend Growth", direction: "row", gap: 0.08, children: [
          { content: { type: "chart", chartType: "bar", title: "Global Ad Spend ($B)", categories: ["2022", "2023", "2024", "2025E", "2026E"], series: [{ name: "Digital", data: [485, 550, 625, 710, 795] }, { name: "Traditional", data: [310, 295, 280, 270, 260] }] } },
          { content: { type: "chart", chartType: "line", title: "YoY Growth Rate (%)", categories: ["2022", "2023", "2024", "2025E", "2026E"], series: [{ name: "Digital", data: [12, 14, 13.5, 13.6, 12.0] }, { name: "Total", data: [8, 6, 7, 8, 7] }] } },
        ] },
      ] },
      { span: 6, direction: "row", gap: 0.1, children: [
        { header: "OTT Paving Way for U.S. Digital Ad Spend", direction: "row", gap: 0.08, children: [
          { content: { type: "chart", chartType: "bar", title: "U.S. CTV Ad Spend ($B)", categories: ["2022", "2023", "2024", "2025E", "2026E"], series: [{ name: "CTV", data: [18.9, 24.6, 28.8, 33.5, 38.2] }] } },
          { content: { type: "chart", chartType: "combo", title: "Digital Share of U.S. Ad Spend", categories: ["2022", "2023", "2024", "2025E", "2026E"],
            bars: [{ name: "Digital ($B)", data: [210, 238, 272, 310, 348] }], lines: [{ name: "% of Total", data: [64, 67, 71, 74, 77] }], dualAxis: true } },
        ] },
        { header: "Total U.S. Ad Spend Analysis", content: { type: "chart", chartType: "combo", title: "Total U.S. Ad Spend",
          categories: ["2020", "2021", "2022", "2023", "2024", "2025E", "2026E"],
          bars: [{ name: "Digital", data: [152, 189, 210, 238, 272, 310, 348] }, { name: "TV/Linear", data: [60, 65, 68, 64, 61, 58, 55] }, { name: "Other", data: [28, 24, 22, 20, 18, 16, 15] }],
          lines: [{ name: "Total ($B)", data: [240, 278, 300, 322, 351, 384, 418] }], dualAxis: true } },
      ] },
    ] },
  });

  // ---- Slide 14: Key trends (3 columns of bullets + chart) ----
  await builder.addGridSlide({
    title: "Media Planning & AdTech Infrastructure Key Trends",
    sectionLabel: "Industry Overview",
    body: { direction: "col", gap: 0.08, children: [
      { span: 2, content: { type: "text", text: "As demand for programmatic advertising grows and privacy regulations reshape the landscape, three infrastructure pillars are emerging as critical differentiators for scaled media platforms.", font: "Arial", fontSize: 1200, color: "#405363", align: "l", anchor: "ctr" } },
      { span: 10, direction: "row", gap: 0.15, children: [
        { header: "Programmatic Supply Chain", direction: "col", gap: 0.08, children: [
          { span: 8, content: bullets(["DSP/SSP consolidation accelerating — top 5 control 72% of spend", "Header bidding now standard; server-side adoption at 45%", "Supply-path optimization reducing intermediary fees by 18%", "CTV programmatic growing 3x faster than display"]) },
          { span: 4, content: { type: "chart", chartType: "bar", title: "U.S. Programmatic Spend ($B)", categories: ["2022", "2023", "2024", "2025E", "2026E"], series: [{ name: "Spend", data: [118, 141, 168, 196, 224] }] } },
        ] },
        { header: "Data Infrastructure & Identity", direction: "col", gap: 0.08, children: [
          { span: 8, content: bullets(["CDP market growing at 34% CAGR — $5.2B by 2026", "Data clean rooms adopted by 62% of top-100 advertisers", "First-party data strategies now #1 priority for CMOs", "Unified ID frameworks gaining publisher traction"]) },
          { span: 4, content: { type: "chart", chartType: "pie", title: "Data Strategy Adoption", items: [{ name: "First-Party", value: 42 }, { name: "Clean Rooms", value: 28 }, { name: "Contextual", value: 18 }, { name: "Universal ID", value: 12 }] } },
        ] },
        { header: "AI & Automation", direction: "col", gap: 0.08, children: [
          { span: 8, content: bullets(["ML-driven bid optimization improving ROAS by 25-40%", "Generative AI cutting creative production costs 55%", "Automated campaign management reducing setup time 70%", "Predictive audience modeling outperforming lookalikes 2:1"]) },
          { span: 4, content: { type: "chart", chartType: "line", title: "AI Adoption in AdTech (%)", categories: ["2021", "2022", "2023", "2024", "2025E"], series: [{ name: "Adoption", data: [22, 35, 51, 68, 82] }], smooth: true, area: true } },
        ] },
      ] },
    ] },
  });

  const totalSlides = 14;
  assert.equal(builder.pres.slides.length, totalSlides);

  const buf = await builder.toBuffer();
  assert.ok(buf.length > 50000, "output should be substantial (has embedded images)");

  const parsed = await parse(buf);
  assert.equal(parsed.slides.length, totalSlides, "round-trip should preserve every slide");
  const imgCount = Object.keys(parsed.media).filter((k) => k.includes("img_")).length;
  assert.ok(imgCount >= 5, `should have at least 5 embedded images, got ${imgCount}`);

  // Chart-only slides still produce picture shapes; the section-title slides produce none.
  const picCounts = parsed.slides.map((s: any) => s.shapes.filter((sh: any) => sh.type === "pic").length);
  assert.ok(picCounts[2] >= 5, "deep-nesting slide should embed its 5 charts");
  assert.equal(picCounts[7], 0, "section title slide has no images");
  assert.ok(picCounts[12] >= 7, "chart quadrant slide should embed 7 charts");
});

// ===== test-layout.js: SlideBuilder helper slides (no reference deck needed) =====

test("builder: addQuadrantSlide renders four header bars", async () => {
  const builder = new SlideBuilder();
  builder.addQuadrantSlide("Market Overview", {
    topLeft: { header: "Key Metrics", content: { type: "statGrid", items: [
      { value: "$15B", label: "Market Size", sublabel: "2028E" }, { value: "12%", label: "CAGR" }, { value: "400+", label: "Players" },
    ] } },
    topRight: { header: "Growth Drivers", content: bullets(["Cloud adoption accelerating", "AI/ML integration", "Consolidation wave"]) },
    bottomLeft: { header: "Recent Deals", content: { type: "cardGrid", items: [
      { title: "Acme Corp", lines: ["$500M", "Q1 2026"] }, { title: "Beta Inc", lines: ["$200M", "Q4 2025"] },
    ] } },
    bottomRight: { header: "Key Players", content: bullets(["Revenue Leader: Northbeam ($720M)", "Fastest Growing: Ridgeline (45% YoY)"]) },
  });
  assert.equal(builder.pres.slides.length, 1);
  const slide = builder.pres.slides[0];
  assert.equal(headerBarCount(slide), 4);
  assert.ok(slide.shapes.length > 10, `should have many shapes, got ${slide.shapes.length}`);
  const buf = await builder.toBuffer();
  assert.ok(buf.length > 5000);
});

test("builder: addThreeColumnSlide renders three header bars", async () => {
  const builder = new SlideBuilder();
  builder.addThreeColumnSlide("Industry Trends", [
    { header: "Market A", content: bullets(["Trend 1", "Trend 2", "Trend 3"]) },
    { header: "Market B", content: { type: "statGrid", items: [{ value: "50%", label: "Growth" }] } },
    { header: "Market C", content: { type: "text", text: "Emerging market with high potential." } },
  ]);
  assert.equal(builder.pres.slides.length, 1);
  assert.equal(headerBarCount(builder.pres.slides[0]), 3);
  assert.ok((await builder.toBuffer()).length > 5000);
});

test("builder: addTimelineSlide draws connectors and event nodes", async () => {
  const builder = new SlideBuilder();
  builder.addTimelineSlide("Company History", [
    { date: "2018", label: "Founded", detail: "Company established in Austin, TX" },
    { date: "2020", label: "Series A", detail: "$10M raised" },
    { date: "2022", label: "IPO", detail: "Listed on NYSE" },
    { date: "2025", label: "Expansion", detail: "Global offices" },
  ]);
  assert.equal(builder.pres.slides.length, 1);
  const slide = builder.pres.slides[0];
  const cxnCount = slide.shapes.filter((s: any) => s.type === "cxnSp").length;
  assert.ok(cxnCount >= 5, `should have connectors (1 main + 4 vertical), got ${cxnCount}`);
  const nodeCount = slide.shapes.filter((s: any) => s.geometry && s.geometry.type === "ellipse").length;
  assert.equal(nodeCount, 4);
  assert.ok((await builder.toBuffer()).length > 5000);
});

test("builder: addTwoPanelSlide horizontal and vertical", async () => {
  const builder = new SlideBuilder();
  builder.addTwoPanelSlide("Comparison", {
    orientation: "horizontal",
    first: { header: "Pros", content: bullets(["Fast", "Cheap"]) },
    second: { header: "Cons", content: bullets(["Complex", "Risky"]) },
  });
  builder.addTwoPanelSlide("Analysis", {
    orientation: "vertical",
    first: { header: "Overview", content: { type: "text", text: "Market analysis text here." } },
    second: { header: "Details", content: bullets(["Size: $10B", "Growth: 15% CAGR"]) },
  });
  assert.equal(builder.pres.slides.length, 2);
  assert.ok((await builder.toBuffer()).length > 5000);
  for (let si = 0; si < 2; si++) assert.equal(headerBarCount(builder.pres.slides[si]), 2, `slide ${si + 1} should have 2 header bars`);
});
