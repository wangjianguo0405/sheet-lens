/**
 * The theme colours the canvas painter needs, resolved from the CSS custom
 * properties in `index.css`.
 *
 * They cannot be read from a stylesheet directly, and they are not literals to
 * copy: the tokens are `oklch()`, and one of them (`--border` under `.dark`)
 * carries its own alpha. Chromium's 2D context does parse `oklch()` now, but
 * whether a given WebView build does is not something to depend on, so each
 * token is resolved once by painting it and reading the pixel back. The result
 * is cached for the life of the page.
 *
 * Resolving at runtime rather than copying the literals is what keeps the grid
 * right if a theme toggle is ever added; today the app only ever applies the
 * light palette.
 */

export interface Rgba {
  r: number;
  g: number;
  b: number;
  /** 0 to 1. */
  a: number;
}

export interface CanvasTheme {
  /** The page behind the print area, and the default cell background. */
  background: string;
  /** The default cell text colour. */
  foreground: string;
  /** Gridlines — `border-border/60`, i.e. the border colour at 60% alpha. */
  gridline: string;
  /** The full-strength border colour, for an explicit cell border of no colour. */
  border: string;
  /** Gutter background. */
  muted: string;
  /** Gutter labels, booleans. */
  mutedForeground: string;
  /** Error cells. */
  destructive: string;
}

/**
 * The light palette from `index.css`, as a fallback for a document that cannot
 * be painted into (no DOM, or a context that refuses `getImageData`). These are
 * the tokens' values and must be updated with them.
 */
const FALLBACK: CanvasTheme = {
  background: "oklch(1 0 0)",
  foreground: "oklch(0.145 0 0)",
  gridline: "oklch(0.922 0 0 / 60%)",
  border: "oklch(0.922 0 0)",
  muted: "oklch(0.97 0 0)",
  mutedForeground: "oklch(0.556 0 0)",
  destructive: "oklch(0.577 0.245 27.325)",
};

const VARS: Record<keyof Omit<CanvasTheme, "gridline">, string> = {
  background: "--background",
  foreground: "--foreground",
  border: "--border",
  muted: "--muted",
  mutedForeground: "--muted-foreground",
  destructive: "--destructive",
};

/** Gridlines are the border ink at 60%, the way `border-border/60` reads. */
const GRIDLINE_ALPHA = 0.6;

let probeContext: CanvasRenderingContext2D | null | undefined;

/**
 * A 1×1 context kept only for turning a CSS colour string into RGBA.
 * `willReadFrequently` silences the readback cost warning; this is called a
 * handful of times per page, but the hint is free.
 */
function probe(): CanvasRenderingContext2D | null {
  if (probeContext === undefined) {
    probeContext = document
      .createElement("canvas")
      .getContext("2d", { willReadFrequently: true });
  }
  return probeContext;
}

/**
 * A CSS colour string as RGBA, by painting one pixel and reading it back.
 * Returns null if the renderer will not paint it or will not give the pixel
 * back — the caller falls back to the literals.
 */
export function toRgba(value: string): Rgba | null {
  const context = probe();
  if (!context) return null;
  try {
    context.clearRect(0, 0, 1, 1);
    // An unparseable value leaves `fillStyle` where it was, so paint it the
    // known black first and treat "still black" as a failure.
    context.fillStyle = "#000000";
    context.fillStyle = value;
    if (context.fillStyle === "#000000") return null;
    context.fillRect(0, 0, 1, 1);
    const [r, g, b, a] = context.getImageData(0, 0, 1, 1).data;
    return { r, g, b, a: a / 255 };
  } catch {
    return null;
  }
}

export function rgba(color: Rgba): string {
  return `rgba(${color.r}, ${color.g}, ${color.b}, ${Number(color.a.toFixed(4))})`;
}

/** The same colour at a given alpha, replacing whatever alpha it had. */
export function withAlpha(color: Rgba, alpha: number): Rgba {
  return { ...color, a: alpha };
}

/** Reads a custom property off the document root, trimmed. */
function readVar(name: string): string {
  return getComputedStyle(document.documentElement)
    .getPropertyValue(name)
    .trim();
}

let cached: CanvasTheme | null = null;

/**
 * The theme, resolved once and cached.
 *
 * A token that is declared but resolves to nothing (or that the renderer will
 * not paint) falls back to its literal, so one bad token cannot take the whole
 * palette with it.
 */
export function canvasTheme(): CanvasTheme {
  if (cached) return cached;

  const resolved: Partial<Record<keyof CanvasTheme, string>> = {};
  const rgbaByToken: Partial<Record<keyof Omit<CanvasTheme, "gridline">, Rgba>> = {};

  for (const [key, variable] of Object.entries(VARS) as [
    keyof Omit<CanvasTheme, "gridline">,
    string,
  ][]) {
    let value: string;
    try {
      value = readVar(variable);
    } catch {
      value = "";
    }
    // An undeclared property reads as the empty string.
    const source = value || FALLBACK[key];
    const parsed = toRgba(source);
    if (parsed) rgbaByToken[key] = parsed;
    resolved[key] = parsed ? rgba(parsed) : source;
  }

  const borderInk = rgbaByToken.border;
  resolved.gridline = borderInk
    ? rgba(withAlpha(borderInk, GRIDLINE_ALPHA))
    : FALLBACK.gridline;

  cached = resolved as CanvasTheme;
  return cached;
}

/** Drop the cache, for a theme change or a test. */
export function resetCanvasTheme(): void {
  cached = null;
}
