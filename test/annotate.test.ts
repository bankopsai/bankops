import { test } from "node:test";
import assert from "node:assert/strict";
import { annotateDeck } from "../src/annotate.js";
import { tableCellCapacity, generateReferenceTable, calcCharsPerLine } from "../src/text-metrics.js";

test("annotates titles, text zones and tables without mutating the input", () => {
  const deck = { slides: [{ title: "Market Overview", body: { direction: "row", children: [
    { span: 8, header: "Trends", content: { type: "text", text: "x" } },
    { span: 4, content: { type: "table", data: { headers: ["A", "B"], rows: [["1", "2"]], colWidths: [6, 6] } } },
  ] } }] };
  const { annotated } = annotateDeck(deck as any);
  const s = annotated.slides[0];
  assert.ok(s._titleMaxChars > 0);
  const [text, table] = s.body.children;
  assert.ok(text._headerMaxChars > 0);
  assert.ok(text.content._maxChars > 0 && text.content._charsPerLine > 0 && text.content._maxLines > 0);
  assert.equal(table.content.data._cellMaxChars.length, 2);
  assert.ok(table.content.data._maxRows > 0);
  assert.equal((deck.slides[0] as any)._titleMaxChars, undefined);
});

test("bigger fonts mean smaller budgets", () => {
  const mk = (fontSize: number) => annotateDeck({ slides: [{ body: { content: { type: "text", text: "x", fontSize } } }] } as any).annotated.slides[0].body.content._maxChars;
  assert.ok(mk(1000) > mk(2000));
});

test("table cell capacity and the reference table", () => {
  const c = tableCellCapacity({ colSpan: 6, rowHeight: 0.4 });
  assert.equal(c.colWidthInches, 4.5);
  assert.ok(c.charsPerLine > 40 && c.maxLines >= 1);
  assert.ok(generateReferenceTable().includes("| Span |"));
  assert.ok(calcCharsPerLine(4.5, 11, "Arial") < calcCharsPerLine(4.5, 11, "Calibri"));
});
