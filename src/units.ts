/** EMU / point / inch / pixel conversions and color helpers. Pure, isomorphic. */

export const EMU_PER_INCH = 914400;
export const EMU_PER_POINT = 12700;
export const EMU_PER_CM = 360000;

export interface FontFamilyParts {
  family: string;
  weight: "bold" | "light" | "";
  style: "italic" | "";
}

export const Units = {
  EMU_PER_INCH,
  EMU_PER_POINT,
  EMU_PER_CM,
  emuToPoints: (emu: number): number => emu / EMU_PER_POINT,
  emuToPx: (emu: number, dpi?: number): number => (emu / EMU_PER_INCH) * (dpi || 96),
  emuToInches: (emu: number): number => emu / EMU_PER_INCH,
  emuToCm: (emu: number): number => emu / EMU_PER_CM,
  pointsToEmu: (pt: number): number => pt * EMU_PER_POINT,
  pxToEmu: (px: number, dpi?: number): number => (px / (dpi || 96)) * EMU_PER_INCH,
  inchesToEmu: (inches: number): number => inches * EMU_PER_INCH,
  cmToEmu: (cm: number): number => cm * EMU_PER_CM,
  /** PPTX font sizes are hundredths of a point: 1400 -> 14pt */
  fontSizeToPoints: (sz: number): number => sz / 100,
  pointsToFontSize: (pt: number): number => pt * 100,
  rotToDeg: (rot?: number): number => (rot || 0) / 60000,
  degToRot: (deg: number): number => deg * 60000,

  /** "Arial Bold" -> { family: "Arial", weight: "bold", style: "" } */
  decomposeFontFamily(name?: string | null): FontFamilyParts {
    if (!name) return { family: "", weight: "", style: "" };
    let family = name;
    let weight: FontFamilyParts["weight"] = "";
    let style: FontFamilyParts["style"] = "";
    if (/\bBold\b/i.test(family)) { weight = "bold"; family = family.replace(/\s*\bBold\b/i, ""); }
    if (/\bLight\b/i.test(family)) { weight = "light"; family = family.replace(/\s*\bLight\b/i, ""); }
    if (/\bItalic\b/i.test(family)) { style = "italic"; family = family.replace(/\s*\bItalic\b/i, ""); }
    return { family: family.trim(), weight, style };
  },

  hexToRgb(hex: string): { r: number; g: number; b: number } {
    const h = hex.replace("#", "");
    const n = parseInt(h.length === 3 ? h.split("").map((c) => c + c).join("") : h, 16);
    return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 };
  },
  rgbToHex(r: number, g: number, b: number): string {
    const c = (v: number) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, "0");
    return "#" + c(r) + c(g) + c(b);
  },
  rgbToHsl(r: number, g: number, b: number): { h: number; s: number; l: number } {
    r /= 255; g /= 255; b /= 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    let h = 0, s = 0;
    const l = (max + min) / 2;
    if (max !== min) {
      const d = max - min;
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
      else if (max === g) h = (b - r) / d + 2;
      else h = (r - g) / d + 4;
      h /= 6;
    }
    return { h, s, l };
  },
  hslToRgb(h: number, s: number, l: number): { r: number; g: number; b: number } {
    if (s === 0) { const v = l * 255; return { r: v, g: v, b: v }; }
    const hue2rgb = (p: number, q: number, t: number) => {
      if (t < 0) t += 1; if (t > 1) t -= 1;
      if (t < 1 / 6) return p + (q - p) * 6 * t;
      if (t < 1 / 2) return q;
      if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
      return p;
    };
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    return { r: hue2rgb(p, q, h + 1 / 3) * 255, g: hue2rgb(p, q, h) * 255, b: hue2rgb(p, q, h - 1 / 3) * 255 };
  },
  /** Apply DrawingML color transforms (lumMod, lumOff, tint, shade) to a hex color */
  applyColorTransforms(hex: string, transforms?: Record<string, number> | null): string {
    if (!transforms) return hex;
    const rgb = Units.hexToRgb(hex);
    const hsl = Units.rgbToHsl(rgb.r, rgb.g, rgb.b);
    let l = hsl.l;
    if (transforms.lumMod != null) l = l * (transforms.lumMod / 100000);
    if (transforms.lumOff != null) l = l + transforms.lumOff / 100000;
    l = Math.max(0, Math.min(1, l));
    let out = Units.hslToRgb(hsl.h, hsl.s, l);
    if (transforms.tint != null) {
      const t = transforms.tint / 100000;
      out = { r: 255 - (255 - out.r) * t, g: 255 - (255 - out.g) * t, b: 255 - (255 - out.b) * t };
    }
    if (transforms.shade != null) {
      const s = transforms.shade / 100000;
      out = { r: out.r * s, g: out.g * s, b: out.b * s };
    }
    return Units.rgbToHex(out.r, out.g, out.b);
  },
};

export default Units;
