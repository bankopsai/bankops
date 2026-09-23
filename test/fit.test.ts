import { test } from "node:test";
import assert from "node:assert/strict";
import { fitTitle, isNumericColumn } from "../src/fit.js";
import { validateDeck } from "../src/schema.js";
import { renderDeck } from "../src/render.js";
import PptxParser from "../src/pptx/pptx-parser.js";

test("short titles keep their size; long ones shrink, then wrap", () => {
  const base = { fontSizePt: 28, font: "Arial Bold", widthInches: 9 };
  assert.deepEqual(fitTitle({ ...base, text: "Financial Summary" }).lines, 1);
  const mid = fitTitle({ ...base, text: "Leading semiconductor names trade at a wide valuation dispersion" });
  assert.ok(mid.fontSizePt < 28 && mid.fontSizePt >= 20, String(mid.fontSizePt));
  const long = fitTitle({ ...base, text: "An extremely long title that no reasonable shrinking could ever keep on a single line of a slide" });
  assert.equal(long.fontSizePt, 20);
  assert.equal(long.lines, 2);
});

test("numeric columns are detected and right-aligned when no alignment was given", () => {
  assert.equal(isNumericColumn(["$5.4tn", "12.3x", "(4.5)", "31%", "1,234", "-8bps", ""]), true);
  assert.equal(isNumericColumn(["NVIDIA", "TSMC"]), false);
  assert.equal(isNumericColumn(["n/a", "12"]), false);
  const r = validateDeck({ slides: [{ body: { content: { type: "table", data: { headers: ["Co", "EV/Rev", "Exposure"], rows: [["A", "9.2x", "AI"], ["B", "7.4x", "Memory"]] } } } }] });
  assert.equal(r.ok, true);
  assert.deepEqual((r.deck as any).slides[0].body.content.data.colAlign, ["l", "r", "l"]);
  const keep = validateDeck({ slides: [{ body: { content: { type: "table", data: { headers: ["Co", "EV"], rows: [["A", "9"]], colAlign: ["ctr", "ctr"] } } } }] });
  assert.deepEqual((keep.deck as any).slides[0].body.content.data.colAlign, ["ctr", "ctr"]);
});

test("a long title renders smaller and the body starts below it", async () => {
  const title = "Leading semiconductor names trade at a wide valuation dispersion";
  const r = await renderDeck({ slides: [{ title, body: { content: { type: "text", text: "body" } } }] });
  const pres = await new PptxParser().parse(r.buffer);
  const shapes = pres.slides[0].shapes;
  const t = shapes.find((s: any) => s.name === "Title");
  const sz = t.textBody.paragraphs[0].runs[0].rPr.sz;
  assert.ok(sz < 2800, `title size ${sz}`);
  const body = shapes.find((s: any) => s.name !== "Title" && s.textBody && s.textBody.paragraphs.some((p: any) => p.runs.some((x: any) => x.text === "body")));
  assert.ok(body.xfrm.off.y >= t.xfrm.off.y + t.xfrm.ext.cy, "body below title");
});

test("a callout renders as a panel, a bar and text", async () => {
  const r = await renderDeck({ slides: [{ body: { content: { type: "callout", title: "Bottom line", text: "Agents that can reach, change and persist are the risk." } } }] });
  const pres = await new PptxParser().parse(r.buffer);
  const names = pres.slides[0].shapes.map((s: any) => s.name);
  assert.ok(names.includes("Callout") && names.includes("Callout Bar"), names.join(","));
  const textShape = pres.slides[0].shapes.find((s: any) => s.textBody && s.textBody.paragraphs.some((p: any) => p.runs.some((x: any) => /Bottom line/.test(x.text))));
  assert.ok(textShape);
});

test("legacy bulletList content renders as bullets and string line spacing is coerced", async () => {
  const r = await renderDeck({ slides: [{ body: { content: { type: "bulletList", items: ["one", "two"], fontSize: 900, lineSpacing: "single" } } }] }, { validate: false });
  const pres = await new PptxParser().parse(r.buffer);
  const tb = pres.slides[0].shapes.find((s: any) => s.textBody)?.textBody;
  assert.ok(tb, "text shape rendered");
  assert.equal(tb.paragraphs.length, 2);
  assert.equal(tb.paragraphs[0].runs[0].text, "one");
  assert.ok(tb.paragraphs[0].pPr.buChar, "bulleted");
  assert.equal(tb.paragraphs[0].pPr.lnSpc, 100);
});
