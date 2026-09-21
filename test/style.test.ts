import { test } from "node:test";
import assert from "node:assert/strict";
import SlideStyle from "../src/style.js";
import { resolveStyleWithTheme, resolveStyle } from "../src/resolve-style.js";

test("defaults produce the corporate token set", () => {
  const t = new SlideStyle().toInternal();
  assert.equal(t.mainTitle.fontSize, 2800);
  assert.equal(t.palette.accent, "#4472C4");
  assert.equal(t.bullet.indent, -Math.round(0.375 * 914400));
  assert.equal(t.tableHeader.fillColor, "#4472C4");
  assert.ok(t.tableSummaryRow.borderTop.width > 0);
});

test("presets exist and from() throws on unknown names", () => {
  assert.deepEqual(Object.keys(SlideStyle.presets), ["corporate", "minimal", "dark", "warm"]);
  assert.equal(SlideStyle.from("dark").toInternal().slide.background, "#1a1a2e");
  assert.throws(() => SlideStyle.from("neon"), /Unknown style preset/);
});

test("merge is two-level deep and non-destructive", () => {
  const base = new SlideStyle({ colors: { accent: "#FF0000" } });
  const merged = base.merge({ colors: { primary: "#000001" }, title: { size: 30 } });
  assert.equal(merged.toInternal().palette.accent, "#FF0000");
  assert.equal(merged.toInternal().palette.primaryText, "#000001");
  assert.equal(merged.toInternal().mainTitle.fontSize, 3000);
  assert.equal(base.toInternal().mainTitle.fontSize, 2800);
});

test("theme tokens sit under deck.style", () => {
  const theme = { colors: { accent: "#123456", primary: "#111111" }, fonts: { body: "Georgia" } };
  const withPreset = resolveStyleWithTheme("minimal", theme);
  assert.equal(withPreset.palette.accent, "#2980b9");
  assert.equal(withPreset.bodyText.font, "Helvetica Neue");
  const withOverrides = resolveStyleWithTheme({ title: { size: 20 } }, theme);
  assert.equal(withOverrides.palette.accent, "#123456");
  assert.equal(withOverrides.bodyText.font, "Georgia");
  assert.equal(withOverrides.mainTitle.fontSize, 2000);
  assert.equal(resolveStyle().palette.accent, "#4472C4");
});
