/** Resolve deck.style (preset name or overrides) and optional theme tokens into StyleTokens. */

import SlideStyle, { type StyleOverrides } from "./style.js";
import type { DeckJson, GridNode, SlideSpec, StyleTokens } from "./types.js";

function resolveNode(node: GridNode): GridNode {
  if (!node || typeof node !== "object") return node;
  if (node.children) {
    return { ...node, children: node.children.filter(Boolean).map((child) => resolveNode(child)) };
  }
  return node;
}

/** Strip null children recursively. */
export function resolveSlide(slide: SlideSpec): SlideSpec {
  if (slide.body) return { ...slide, body: resolveNode(slide.body) };
  return slide;
}

/** StyleTokens for a deck.style value (undefined = corporate defaults). */
export function resolveStyle(style?: string | StyleOverrides): StyleTokens {
  let ss: SlideStyle;
  if (!style) ss = new SlideStyle();
  else if (typeof style === "string") ss = SlideStyle.from(style);
  else ss = new SlideStyle(style);
  return ss.toInternal();
}

/**
 * Theme tokens (an organization's brand) form the base layer; deck.style
 * (preset name or overrides) is merged on top.
 */
export function resolveStyleWithTheme(deckStyle?: string | StyleOverrides, themeTokens?: StyleOverrides): StyleTokens {
  if (!themeTokens) return resolveStyle(deckStyle);
  return mergeStyleWithTheme(deckStyle, themeTokens).toInternal();
}

/** Same cascade as resolveStyleWithTheme but returns the SlideStyle (human units). */
export function mergeStyleWithTheme(deckStyle?: string | StyleOverrides, themeTokens?: StyleOverrides): SlideStyle {
  let ss = new SlideStyle(themeTokens || {});
  if (typeof deckStyle === "string") {
    const preset = SlideStyle.presets[deckStyle];
    if (!preset) throw new Error(`Unknown style preset: ${deckStyle}`);
    ss = ss.merge(preset);
  } else if (deckStyle && typeof deckStyle === "object") {
    ss = ss.merge(deckStyle);
  }
  return ss;
}

/** Resolve all slides plus the style of a deck. */
export function resolveDeck(deck: DeckJson): { slides: SlideSpec[]; style: StyleTokens } {
  const style = resolveStyle(deck.style);
  const slides = (deck.slides || []).map((s) => resolveSlide(s));
  return { slides, style };
}
