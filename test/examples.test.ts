/**
 * Every deck in examples/ must validate with no errors and no warnings, render
 * offline without warnings, and re-parse with the expected slide count.
 *
 * Rendering runs with `network: false` so the suite is hermetic. An `icon` zone
 * that carries only an Iconify name needs the network, so a deck with such a
 * zone is expected to produce exactly one warning per icon (the "network
 * disabled" message) and nothing else. Every other deck must produce zero.
 */
import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { validateDeck } from "../src/schema.js";
import { renderDeck } from "../src/render.js";
import PptxParser from "../src/pptx/pptx-parser.js";

const examplesDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../examples");
const files = fs.readdirSync(examplesDir).filter((f) => f.endsWith(".json")).sort();

/** Count `icon` content zones that rely on an Iconify name (no inline svg). */
function countNetworkIcons(node: any): number {
  if (!node || typeof node !== "object") return 0;
  if (Array.isArray(node)) return node.reduce((n, c) => n + countNetworkIcons(c), 0);
  let n = 0;
  if (node.content && node.content.type === "icon" && !node.content.svg) n++;
  for (const k of Object.keys(node)) if (typeof node[k] === "object") n += countNetworkIcons(node[k]);
  return n;
}

test("examples/ ships the documented decks", () => {
  assert.deepEqual(files, ["company-profile.json", "comps-table.json", "dark-theme.json", "market-overview.json", "pitch-deck.json"]);
});

for (const file of files) {
  test(`examples/${file} validates, renders offline and re-parses`, async () => {
    const deck = JSON.parse(fs.readFileSync(path.join(examplesDir, file), "utf8"));

    const v = validateDeck(deck);
    assert.equal(v.ok, true, JSON.stringify(v.errors));
    assert.equal(v.errors.length, 0);
    assert.deepEqual(v.warnings, [], "lint warnings must be fixed in the example, not tolerated");

    const iconCount = countNetworkIcons(deck);
    assert.ok(iconCount <= 1, "at most one Iconify icon per example deck");

    const r = await renderDeck(deck, { network: false });
    assert.equal(r.lint.length, 0);
    assert.equal(r.warnings.length, iconCount, `unexpected warnings: ${JSON.stringify(r.warnings)}`);
    for (const w of r.warnings) assert.match(w, /network disabled; icon ".+" needs inline svg/);

    assert.equal(r.slides, deck.slides.length);
    assert.ok(r.buffer.length > 10000, "PPTX should be substantial");
    const parsed = await new PptxParser().parseBuffer(r.buffer);
    assert.equal(parsed.slides.length, deck.slides.length);
    assert.equal(parsed.slides.length, v.deck!.slides.length);
  });
}
