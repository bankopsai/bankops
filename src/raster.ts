/**
 * Optional raster helpers. `sharp` is an optional peer dependency: it is loaded
 * lazily and a clear error is thrown when it is missing. Node-only.
 */

type SharpModule = typeof import("sharp");

let sharpPromise: Promise<SharpModule | null> | null = null;

async function loadSharp(): Promise<SharpModule | null> {
  if (!sharpPromise) {
    sharpPromise = import("sharp").then((m) => (m.default || m) as SharpModule).catch(() => null);
  }
  return sharpPromise;
}

export class MissingDependencyError extends Error {
  readonly dependency: string;
  readonly fix: string;
  constructor(dependency: string, purpose: string) {
    const fix = `npm install ${dependency}`;
    super(`bankops: '${dependency}' is required for ${purpose}. Fix: ${fix}`);
    this.name = "MissingDependencyError";
    this.dependency = dependency;
    this.fix = fix;
  }
}

/** True when sharp can be loaded. */
export async function hasSharp(): Promise<boolean> {
  return (await loadSharp()) !== null;
}

/** Strip constructs that make librsvg choke (non-absolute xmlns, <metadata>). */
export function sanitizeSvg(svg: string): string {
  let out = svg.replace(/\s+xmlns:[a-z0-9]+=["'][^"']*["']/gi, (m) => (/https?:\/\//.test(m) ? m : ""));
  out = out.replace(/<metadata[\s>][\s\S]*?<\/metadata>/gi, "");
  return out;
}

/** Render an SVG string or buffer to PNG. `density` controls raster resolution (default 72). */
export async function svgToPng(svg: string | Buffer, opts: { density?: number } = {}): Promise<Buffer> {
  const sharp = await loadSharp();
  if (!sharp) throw new MissingDependencyError("sharp", "rendering SVG (charts, icons) to PNG");
  const input = typeof svg === "string" ? Buffer.from(sanitizeSvg(svg), "utf8") : Buffer.from(sanitizeSvg(svg.toString("utf8")), "utf8");
  return sharp(input, opts.density ? { density: opts.density } : undefined).png().toBuffer();
}

/** Convert any sharp-readable image (WebP, GIF, TIFF...) to PNG. */
export async function toPng(image: Buffer): Promise<Buffer> {
  const sharp = await loadSharp();
  if (!sharp) throw new MissingDependencyError("sharp", "converting images to PNG");
  return sharp(image).png().toBuffer();
}

export function isSvg(buf: Buffer, hint?: string): boolean {
  if (hint && /\.svg(\?|$)/i.test(hint)) return true;
  if (buf.length > 200000) return false;
  return buf.subarray(0, 500).toString("utf8").includes("<svg");
}

export function isWebP(buf: Buffer): boolean {
  return buf.length >= 12 && buf[0] === 0x52 && buf[1] === 0x49 && buf[2] === 0x46 && buf[3] === 0x46 &&
    buf[8] === 0x57 && buf[9] === 0x45 && buf[10] === 0x42 && buf[11] === 0x50;
}
