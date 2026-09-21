/**
 * Port of legacy/test-style.js. test/style.test.ts already covers the corporate
 * defaults, preset listing, two-level merge and the theme cascade; this file
 * carries the rest: the full internal token set, unit conversions, color cascade
 * rules, preset specifics, the SlideBuilder integration and rendered decks.
 */
import { test } from "node:test";
import assert from "node:assert/strict";
import SlideStyle from "../src/style.js";
import SlideBuilder from "../src/slide-builder.js";
import ChartGenerator from "../src/charts/chart-generator.js";
import PptxParser from "../src/pptx/pptx-parser.js";

const approxEq = (a: number, b: number, tolerance = 1) => Math.abs(a - b) <= tolerance;

test("default style produces the full internal token set", () => {
  const int: any = new SlideStyle().toInternal();
  const expectedKeys = [
    "sectionLabel", "mainTitle", "bodyText", "bullet", "footnote", "divider", "palette", "tableHeader", "tableBody",
    "slide", "headerBar", "header", "subheader", "subfooter", "card", "spacing", "titleSlide",
  ];
  for (const k of expectedKeys) assert.ok(int[k] != null, `missing key: ${k}`);

  // Unit conversions
  assert.equal(int.mainTitle.fontSize, 2800, "title 28pt -> 2800");
  assert.equal(int.bodyText.fontSize, 1400, "body 14pt -> 1400");
  assert.equal(int.footnote.fontSize, 800, "footnote 8pt -> 800");
  assert.ok(approxEq(int.sectionLabel.x, 457200), "margin 0.5in -> 457200 EMU");
  assert.ok(approxEq(int.mainTitle.y, 822960, 10), "title y 0.9in -> ~822960 EMU");
  assert.equal(int.bullet.indent, -342900, "bullet indent negated");
  assert.equal(int.bullet.marginLeft, 342900);
  assert.equal(int.divider.numberFontSize, 11500, "divider number 115pt -> 11500");

  // Cascade resolution
  assert.equal(int.mainTitle.color, "#405363");
  assert.equal(int.sectionLabel.color, "#4472C4", "section label color cascades from colors.accent");
  assert.equal(int.footnote.color, "#666666", "footnote color cascades from colors.secondary");
  assert.equal(int.palette.darkFill, "#164556");

  // Previously hardcoded categories
  assert.equal(int.slide.background, "#FFFFFF");
  assert.equal(int.headerBar.textColor, "#FFFFFF");
  assert.equal(int.spacing.columnGap, 228600);
  assert.ok(approxEq(int.spacing.bulletSpacing, 114300));
  assert.equal(int.titleSlide.titleSize, 3600);
  assert.equal(int.titleSlide.subtitleSize, 1800);
});

test("partial override keeps every other default", () => {
  const int = new SlideStyle({ colors: { accent: "#FF0000" } }).toInternal();
  assert.equal(int.palette.accent, "#FF0000");
  assert.equal(int.sectionLabel.color, "#FF0000", "section label cascades from the new accent");
  assert.equal(int.palette.primaryText, "#333333");
  assert.equal(int.bodyText.font, "Calibri");
  assert.equal(int.mainTitle.font, "Arial Bold");
});

test("colors cascade into every dependent token unless set explicitly", () => {
  const int = new SlideStyle({ colors: { primary: "#111111", secondary: "#222222", accent: "#333333" } }).toInternal();
  assert.equal(int.mainTitle.color, "#405363", "title.color is explicit, not cascaded");
  assert.equal(int.bodyText.color, "#111111", "body.color -> colors.primary");
  assert.equal(int.bullet.color, "#111111", "bullet.color -> colors.primary");
  assert.equal(int.footnote.color, "#222222", "footnote.color -> colors.secondary");
  assert.equal(int.sectionLabel.color, "#333333", "sectionLabel.color -> colors.accent");
  assert.equal(int.divider.titleColor, "#333333", "divider.titleColor -> colors.accent");
  assert.equal(int.tableHeader.fillColor, "#333333", "tableHeader.fillColor -> colors.accent");

  const int2 = new SlideStyle({ colors: { primary: "#111111" }, title: { color: "#AAAAAA" } }).toInternal();
  assert.equal(int2.mainTitle.color, "#AAAAAA", "explicit title.color overrides the cascade");
  assert.equal(int2.bodyText.color, "#111111", "body still uses the cascade");
});

test("points and inches convert to hundredths-pt and EMU", () => {
  const int = new SlideStyle({
    title: { size: 48, y: 1.5 },
    body: { size: 10 },
    spacing: { columnGap: 0.5, headerBarHeight: 0.4 },
    titleSlide: { titleSize: 44, subtitleSize: 20 },
  }).toInternal();
  assert.equal(int.mainTitle.fontSize, 4800);
  assert.ok(approxEq(int.mainTitle.y, 1371600), "1.5in -> 1371600 EMU");
  assert.equal(int.bodyText.fontSize, 1000);
  assert.equal(int.spacing.columnGap, 457200, "0.5in -> 457200 EMU");
  assert.ok(approxEq(int.spacing.headerBarHeight, 365760), "0.4in -> 365760 EMU");
  assert.equal(int.titleSlide.titleSize, 4400);
  assert.equal(int.titleSlide.subtitleSize, 2000);
});

test("every preset yields a complete token set; dark and warm specifics", () => {
  const names = Object.keys(SlideStyle.presets);
  assert.ok(names.length >= 4);
  for (const name of names) {
    const int: any = SlideStyle.from(name).toInternal();
    assert.ok(int.mainTitle && int.bodyText && int.palette, `${name} has required sections`);
    assert.ok(int.slide && int.spacing && int.card, `${name} has slide/spacing/card sections`);
  }
  const dark = SlideStyle.from("dark").toInternal();
  assert.equal(dark.slide.background, "#1a1a2e");
  assert.equal(dark.palette.primaryText, "#EAEAEA");
  assert.equal(dark.palette.accent, "#E94560");
  const warm = SlideStyle.from("warm").toInternal();
  assert.equal(warm.mainTitle.font, "Georgia");
  assert.equal(warm.bodyText.font, "Garamond");
});

test("merge chains preserve earlier overrides", () => {
  const style = SlideStyle.from("corporate").merge({ colors: { accent: "#FF0000" } });
  const int = style.toInternal();
  assert.equal(int.palette.accent, "#FF0000");
  assert.equal(int.palette.primaryText, "#333333");
  const int2 = style.merge({ title: { size: 40 } }).toInternal();
  assert.equal(int2.palette.accent, "#FF0000", "accent survives a double merge");
  assert.equal(int2.mainTitle.fontSize, 4000);
});

test("SlideBuilder accepts style objects, SlideStyle instances and slide sizes", () => {
  const b1 = new SlideBuilder({ style: { colors: { accent: "#FF0000" } } });
  assert.equal(b1._style.palette.accent, "#FF0000");
  assert.equal(b1._style.slide.background, "#FFFFFF");

  const b2 = new SlideBuilder({ style: SlideStyle.from("dark") });
  assert.equal(b2._style.slide.background, "#1a1a2e");
  assert.equal(b2._style.palette.primaryText, "#EAEAEA");
  assert.equal(b2._style.headerBar.textColor, "#FFFFFF");

  const b3 = new SlideBuilder({ slideSize: { width: 13.333, height: 7.5 }, style: { title: { size: 28 } } });
  assert.equal(b3._style.mainTitle.fontSize, 2800);
  assert.ok(b3.pres.slideSize.cx > 12000000, "wide slide");
});

test("SlideBuilder default tokens match the corporate stylesheet", () => {
  const s = new SlideBuilder()._style;
  assert.equal(s.sectionLabel.font, "Calibri");
  assert.equal(s.sectionLabel.fontSize, 1200);
  assert.equal(s.mainTitle.font, "Arial Bold");
  assert.equal(s.mainTitle.fontSize, 2800);
  assert.equal(s.bodyText.fontSize, 1400);
  assert.equal(s.palette.darkFill, "#164556");
  assert.equal(s.palette.accent, "#4472C4");
  assert.equal(s.slide.background, "#FFFFFF");
  assert.equal(s.headerBar.textColor, "#FFFFFF");
  assert.equal(s.spacing.columnGap, 228600);
  assert.equal(s.spacing.bulletSpacing, 114300);
  assert.equal(s.spacing.zoneGap, 137160);
  assert.equal(s.card.fill, "#FFFFFF");
  assert.equal(s.card.borderWidth, 9525);
  assert.equal(s.titleSlide.titleSize, 3600);
});

test("SlideStyle.defaults exposes the human-unit stylesheet", () => {
  assert.ok(SlideStyle.defaults != null);
  assert.equal(SlideStyle.defaults.colors.primary, "#333333");
  assert.equal(SlideStyle.defaults.title.size, 28);
  assert.equal(SlideStyle.defaults.fonts.body, "Calibri");
});

test("styled decks render: dark, warm and custom", async () => {
  const titleBody = (title: string, subtitle: string) => ({
    direction: "row" as const, children: [
      { span: 1 },
      { span: 7, direction: "col" as const, children: [
        { content: { type: "text" as const, text: title, fontSize: 2800 } },
        { content: { type: "text" as const, text: subtitle, fontSize: 1400 } },
      ] },
      { span: 4 },
    ],
  });

  const dark = new SlideBuilder({ style: SlideStyle.from("dark") });
  await dark.addGridSlide({ title: "Dark Theme Demo", body: titleBody("Dark Theme Demo", "Generated with SlideStyle") });
  dark.addContentSlide("Key Features", ["Custom color palettes", "Human-friendly units (pt, inches)", "Cascading defaults", "Built-in presets"], { sectionLabel: "Style System" });
  const barChart = await new ChartGenerator({ width: 600, height: 400 }).bar({
    title: "Quarterly Revenue", categories: ["Q1", "Q2", "Q3", "Q4"], series: [{ name: "Revenue", data: [120, 150, 180, 210] }],
  });
  await dark.addGridSlide({
    title: "Dashboard",
    sectionLabel: "Dark Theme",
    body: { direction: "row", children: [
      { span: 7, header: "Revenue Trend", content: { type: "image", buffer: barChart, width: 600, height: 400 } },
      { span: 5, header: "Key Metrics", content: { type: "statGrid", items: [
        { value: "$660M", label: "Total Revenue" }, { value: "23%", label: "YoY Growth" }, { value: "4,200", label: "Customers" },
      ] } },
    ] },
  });
  const darkBuf = await dark.toBuffer();
  assert.ok(darkBuf.length > 5000, "dark deck generated");
  assert.equal(dark.pres.slides.length, 3);
  assert.equal(dark.pres.slides[0].background.color, "#1a1a2e", "dark slide background applied");
  const parsed = await new PptxParser().parseBuffer(darkBuf);
  assert.equal(parsed.slides.length, 3);

  const warm = new SlideBuilder({ style: SlideStyle.from("warm") });
  await warm.addGridSlide({ title: "Warm Theme Demo", body: titleBody("Warm Theme Demo", "Georgia & Garamond") });
  warm.addContentSlide("Earthy Tones", ["Rich warm palette", "Serif fonts for elegance", "Professional and inviting"]);
  assert.ok((await warm.toBuffer()).length > 3000, "warm deck generated");

  const custom = new SlideBuilder({ style: {
    colors: { primary: "#1B2838", accent: "#66C0F4", darkFill: "#171A21", secondary: "#8F98A0" },
    fonts: { title: "Arial Bold", body: "Arial" },
    slide: { background: "#1B2838" },
    title: { size: 28 },
  } });
  await custom.addGridSlide({ title: "Custom Style", body: titleBody("Custom Style", "Navy and sky-blue color scheme") });
  assert.ok((await custom.toBuffer()).length > 3000, "custom deck generated");
  assert.equal(custom.pres.slides[0].background.color, "#1B2838");
});

test("dark preset overrides the corporate title color so titles are legible", () => {
  const dark = SlideStyle.from("dark").toInternal();
  assert.notEqual(dark.mainTitle.color, "#405363", "corporate slate title on a #1a1a2e background is unreadable");
  assert.equal(dark.mainTitle.color, "#EAEAEA");
});
