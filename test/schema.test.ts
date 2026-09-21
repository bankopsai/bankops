import { test } from "node:test";
import assert from "node:assert/strict";
import { validateDeck, parseDeck, deckJsonSchema, DeckValidationError, CONTENT_TYPES } from "../src/schema.js";

const good = {
  style: "corporate",
  slides: [
    { title: "Overview", sectionLabel: "Section 1", footer: "Source: company filings",
      body: { direction: "row", children: [
        { span: 6, header: "Highlights", content: { type: "text", runs: [{ text: "Revenue up 12%", bullet: true }] } },
        { span: 6, content: { type: "chart", chartType: "bar", categories: ["FY24", "FY25"], series: [{ name: "Revenue", data: [100, 112] }] } },
      ] } },
  ],
};

test("accepts a valid deck and strips nothing it needs", () => {
  const r = validateDeck(good);
  assert.equal(r.ok, true);
  assert.equal(r.errors.length, 0);
  assert.equal(r.warnings.length, 0);
  assert.equal(r.deck!.slides.length, 1);
});

test("rejects a non-object and an empty slides array with fixes", () => {
  assert.equal(validateDeck(null).ok, false);
  const r = validateDeck({ slides: [] });
  assert.equal(r.ok, false);
  assert.match(r.errors[0].message, /at least one slide/);
  assert.ok(r.errors[0].fix);
});

test("teaches the fontSize convention", () => {
  const r = validateDeck({ slides: [{ body: { content: { type: "text", text: "x", fontSize: 14 } } }] });
  assert.equal(r.ok, false);
  assert.match(r.errors[0].message, /hundredths of a point/);
  assert.equal(r.errors[0].path, "slides[0].body.content.fontSize");
});

test("lists content types on an unknown type", () => {
  const r = validateDeck({ slides: [{ body: { content: { type: "bullets" } } }] });
  assert.equal(r.ok, false);
  for (const t of CONTENT_TYPES) assert.ok(r.errors[0].fix!.includes(t));
});

test("container and leaf are mutually exclusive", () => {
  const r = validateDeck({ slides: [{ body: { children: [{ span: 12 }], content: { type: "text", text: "x" } } }] });
  assert.equal(r.ok, false);
  assert.match(r.errors[0].message, /container/);
});

test("pie charts need items, combos need bars/lines, waterfalls need data", () => {
  const mk = (c: object) => validateDeck({ slides: [{ body: { content: { type: "chart", ...c } } }] });
  assert.match(mk({ chartType: "pie", series: [] }).errors[0].message, /items/);
  assert.match(mk({ chartType: "combo", categories: ["a"] }).errors[0].message, /bars/);
  assert.match(mk({ chartType: "waterfall", categories: ["a"] }).errors[0].message, /data/);
  assert.equal(mk({ chartType: "pie", items: [{ name: "a", value: 1 }] }).ok, true);
});

test("warns on typos with a suggestion, and on spans that do not sum to 12", () => {
  const r = validateDeck({ slides: [{ body: { children: [
    { span: 6, content: { type: "text", text: "x", colour: "#000000" } },
    { span: 5, content: { type: "text", text: "y" } },
  ] } }] });
  assert.equal(r.ok, true);
  const keys = r.warnings.map((w) => w.message);
  assert.ok(keys.some((m) => m.includes('"colour"')));
  assert.ok(r.warnings.find((w) => w.fix?.includes('"color"')));
  assert.ok(keys.some((m) => m.includes("sum to 11")));
});

test("ignores underscore and dollar metadata keys", () => {
  const r = validateDeck({ _meta: 1, slides: [{ _task: "todo", body: { _overflow: true, content: { type: "text", text: "x", _maxChars: 10 } } }] });
  assert.equal(r.ok, true);
  assert.equal(r.warnings.length, 0);
});

test("parseDeck throws a DeckValidationError with every problem", () => {
  assert.throws(() => parseDeck({ slides: [{ body: { content: { type: "nope" } } }] }), DeckValidationError);
});

test("emits a JSON Schema with definitions", () => {
  const s = deckJsonSchema() as any;
  assert.equal(s.$schema, "https://json-schema.org/draft/2020-12/schema");
  assert.ok(s.properties.slides);
  assert.ok(JSON.stringify(s).includes("statGrid"));
});

test("scatter series accept [x, y] pairs", () => {
  const r = validateDeck({ slides: [{ body: { content: { type: "chart", chartType: "scatter", series: [{ name: "Peers", data: [[1.2, 14], [2.5, 22]] }] } } }] });
  assert.equal(r.ok, true, JSON.stringify(r.errors));
});
