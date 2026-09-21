import { test } from "node:test";
import assert from "node:assert/strict";
import SlideLayout from "../src/layout.js";

const EMU = 914400;
const inchesOf = (emu: number) => +(emu / EMU).toFixed(2);

test("two-column layout splits the body evenly with a gap", () => {
  const tree = new SlideLayout().compute({ body: { direction: "row", children: [{ span: 6 }, { span: 6 }] } });
  const [a, b] = tree.body!.children!;
  assert.equal(inchesOf(a.bounds.x), 0.5);
  assert.equal(inchesOf(a.bounds.cx), inchesOf(b.bounds.cx));
  assert.equal(inchesOf(b.bounds.x), inchesOf(a.bounds.x + a.bounds.cx + 0.15 * EMU));
});

test("title, section label and footer reserve space", () => {
  const tree = new SlideLayout().compute({ title: "T", sectionLabel: "S", footer: "F", body: { span: 12 } });
  assert.ok(tree.sectionLabelBounds && tree.titleBounds && tree.footerBounds);
  assert.ok(tree.body!.bounds.y > tree.titleBounds!.y + tree.titleBounds!.cy);
  assert.ok(tree.body!.bounds.y + tree.body!.bounds.cy <= tree.footerBounds!.y);
});

test("auto spans fill the remaining columns", () => {
  const tree = new SlideLayout().compute({ body: { direction: "row", children: [{ span: 4 }, {}, {}] } });
  const [a, b, c] = tree.body!.children!;
  assert.equal(inchesOf(b.bounds.cx), inchesOf(c.bounds.cx));
  assert.equal(inchesOf(a.bounds.cx), inchesOf(b.bounds.cx));
});

test("header and subheader shrink the content area", () => {
  const plain = new SlideLayout().compute({ body: { span: 12 } }).body!;
  const withHeader = new SlideLayout().compute({ body: { span: 12, header: "H", subheader: "S" } }).body!;
  assert.ok(withHeader.headerBounds && withHeader.subheaderBounds && withHeader.subheaderLineBounds);
  assert.ok(withHeader.contentBounds!.cy < plain.contentBounds!.cy);
  assert.ok(withHeader.contentBounds!.y > withHeader.subheaderLineBounds!.y);
});

test("margin and padding inset the node", () => {
  const tree = new SlideLayout().compute({ body: { margin: 0.25, padding: 0.1, direction: "col", children: [{ span: 12 }] } });
  const body = tree.body!;
  assert.equal(inchesOf(body.bounds.x), 0.75);
  assert.equal(inchesOf(body.children![0].bounds.x), 0.85);
});

test("presets are deep-cloned and accept overrides", () => {
  const a = SlideLayout.preset("quadrant", { title: "Q" })!;
  const b = SlideLayout.preset("quadrant")!;
  assert.equal(a.title, "Q");
  assert.equal(b.title, undefined);
  assert.notEqual(a.body, b.body);
  assert.equal(SlideLayout.preset("nope"), null);
  assert.equal(Object.keys(SlideLayout.presets).length, 7);
});

test("flattenLeaves returns every leaf", () => {
  const layout = new SlideLayout();
  const tree = layout.compute(SlideLayout.preset("dashboard")!);
  assert.equal(layout.flattenLeaves(tree).length, 7);
});
