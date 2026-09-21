/** Resolves scheme colors / theme fonts against a parsed Theme (with clrMap / clrMapOvr). */

import U from "../units.js";

/** DrawingML color transform as parsed from XML: `<a:lumMod val="75000"/>` -> { type: 'lumMod', val: 75000 } */
export interface ColorTransform { type: string; val: number; }

// Applies an ordered list of DrawingML transforms (tint, shade, lumMod/Off, satMod/Off, alpha)
// to a hex color. Mirrors the legacy units.applyColorTransforms(hex, transforms[]) contract.
function applyTransformList(hex: string, transforms?: ColorTransform[] | null): { color: string; alpha: number | null } {
  if (!transforms || transforms.length === 0) return { color: hex, alpha: null };
  let rgb = U.hexToRgb(hex);
  let alpha: number | null = null;
  for (let i = 0; i < transforms.length; i++) {
    const t = transforms[i], v = t.val / 100000;
    switch (t.type) {
      case 'tint':
        rgb.r = rgb.r + (255 - rgb.r) * v;
        rgb.g = rgb.g + (255 - rgb.g) * v;
        rgb.b = rgb.b + (255 - rgb.b) * v;
        break;
      case 'shade':
        rgb.r = rgb.r * v;
        rgb.g = rgb.g * v;
        rgb.b = rgb.b * v;
        break;
      case 'lumMod': case 'lumOff': case 'satMod': case 'satOff': {
        const hsl = U.rgbToHsl(rgb.r, rgb.g, rgb.b);
        if (t.type === 'lumMod') hsl.l = hsl.l * v;
        else if (t.type === 'lumOff') hsl.l = hsl.l + v;
        else if (t.type === 'satMod') hsl.s = hsl.s * v;
        else if (t.type === 'satOff') hsl.s = hsl.s + v;
        hsl.l = Math.max(0, Math.min(1, hsl.l));
        hsl.s = Math.max(0, Math.min(1, hsl.s));
        rgb = U.hslToRgb(hsl.h, hsl.s, hsl.l);
        break;
      }
      case 'alpha':
        alpha = v * 100;
        break;
    }
  }
  return { color: U.rgbToHex(rgb.r, rgb.g, rgb.b), alpha: alpha };
}

export class ThemeResolver {
  theme: any;

  constructor(theme: any) {
    this.theme = theme;
  }

  resolveColor(colorValue: any, slide?: any): string {
    if (!colorValue) return '#000000';
    let baseVal: any, transforms: ColorTransform[] | null = null;
    if (typeof colorValue === 'object' && colorValue.val !== undefined) {
      baseVal = colorValue.val;
      transforms = colorValue.transforms;
    } else {
      baseVal = colorValue;
    }
    let hex = '#000000';
    if (typeof baseVal === 'string') {
      if (baseVal.charAt(0) === '#') {
        hex = baseVal;
      } else if (this.theme && this.theme.clrScheme) {
        let mapped = baseVal;
        const aliasMap: Record<string, string> = { bg1: 'lt1', bg2: 'lt2', tx1: 'dk1', tx2: 'dk2' };
        if (aliasMap[mapped]) mapped = aliasMap[mapped];
        if (slide && slide.clrMapOvr && slide.clrMapOvr[baseVal]) {
          mapped = slide.clrMapOvr[baseVal];
        } else if (slide && slide.slideLayout && slide.slideLayout.slideMaster) {
          const master = slide.slideLayout.slideMaster;
          if (master.clrMap && master.clrMap[baseVal]) {
            mapped = master.clrMap[baseVal];
          }
        }
        const resolved = this.theme.clrScheme[mapped];
        if (resolved) hex = resolved;
        else hex = '#' + baseVal;
      } else {
        hex = '#' + baseVal;
      }
    }
    if (transforms && transforms.length > 0) {
      const result = applyTransformList(hex, transforms);
      hex = result.color;
    }
    return hex;
  }

  resolveColorAlpha(colorValue: any): number | null {
    if (!colorValue || typeof colorValue !== 'object' || !colorValue.transforms) return null;
    for (let i = 0; i < colorValue.transforms.length; i++) {
      if (colorValue.transforms[i].type === 'alpha') return colorValue.transforms[i].val / 1000;
    }
    return null;
  }

  resolveFont(fontRef: any): string {
    if (!fontRef || !this.theme) return fontRef || 'Calibri';
    if (fontRef === '+mj-lt' || fontRef === '+mj-ea' || fontRef === '+mj-cs') {
      return this.theme.fontScheme.majorFont.latin || 'Calibri Light';
    }
    if (fontRef === '+mn-lt' || fontRef === '+mn-ea' || fontRef === '+mn-cs') {
      return this.theme.fontScheme.minorFont.latin || 'Calibri';
    }
    return fontRef;
  }
}

export default ThemeResolver;
