/**
 * Text measurement and line breaking for the canvas grid.
 *
 * Everything the DOM renderer got from the layout engine for free has to be
 * asked for here: the width of a run, where a line would break, where the
 * baseline sits inside a line box, and how much of a too-long value fits before
 * an ellipsis. The results are cached, because a scroll repaints every visible
 * cell and `measureText` is the one call in the paint loop that is not free.
 *
 * The font stack is the DOM renderer's own (`Grid.tsx`), deliberately
 * character for character: any divergence changes the metrics, and the metrics
 * decide column widths and text fitting.
 */

import { CJK_FALLBACK, LINE_HEIGHT, measureContext } from "@/lib/gridGeometry";
import type { RichFont, RichRun } from "@/lib/excel";

const STACKS = new Map<string, string>();

/**
 * The font stack a cell draws in. A cell that names a family uses it ahead of
 * the CJK fallbacks; one that does not inherits the workbook's family, and
 * behind that the same browsers-default chain the scroller declared.
 */
export function cellFontStack(
  cellFamily: string | null,
  workbookFamily: string | null,
): string {
  const key = `${cellFamily ?? ""}\u0000${workbookFamily ?? ""}`;
  const known = STACKS.get(key);
  if (known !== undefined) return known;

  const stack = cellFamily
    ? `"${cellFamily}", ${CJK_FALLBACK}`
    : workbookFamily
      ? `"${workbookFamily}", ${CJK_FALLBACK}, system-ui, sans-serif`
      : `${CJK_FALLBACK}, system-ui, sans-serif`;
  if (STACKS.size < CACHE_LIMIT) STACKS.set(key, stack);
  return stack;
}

const FONTS = new Map<string, string>();

/** The CSS `font` shorthand for a stack. Weight and slant are left out when normal. */
export function fontString(
  stack: string,
  sizePx: number,
  bold: boolean,
  italic = false,
): string {
  const key = `${bold ? 1 : 0}${italic ? 1 : 0}\u0000${sizePx}\u0000${stack}`;
  const known = FONTS.get(key);
  if (known !== undefined) return known;

  const font = `${italic ? "italic " : ""}${bold ? "600 " : ""}${sizePx}px ${stack}`;
  if (FONTS.size < CACHE_LIMIT) FONTS.set(key, font);
  return font;
}

/**
 * Caches are keyed by the whole font string and dropped wholesale if they grow
 * past this, which keeps a long browse of many sheets from pinning every string
 * it ever drew. A single sheet's visible text is far below it.
 */
const CACHE_LIMIT = 20000;

function bounded<K, V>(map: Map<K, V>, key: K, value: V): V {
  if (map.size >= CACHE_LIMIT) map.clear();
  map.set(key, value);
  return value;
}

const WIDTHS = new Map<string, number>();

/** The advance width of `text` in `font`, cached. */
export function textWidth(font: string, text: string): number {
  const key = `${font}\u0000${text}`;
  const known = WIDTHS.get(key);
  if (known !== undefined) return known;

  const context = measureContext();
  let width = 0;
  if (context) {
    context.font = font;
    width = context.measureText(text).width;
  }
  return bounded(WIDTHS, key, width);
}

export interface LineMetrics {
  /** Distance from the alphabetic baseline up to the font's own top. */
  ascent: number;
  descent: number;
  /** The line box height the DOM renderer used: `fontSize * LINE_HEIGHT`. */
  lineHeight: number;
}

const METRICS = new Map<string, LineMetrics>();

/**
 * Where the baseline sits inside a line box, from the font's own bounding box.
 *
 * CSS centres the leading: for a line box of height `H` and a font whose ascent
 * and descent total `A + D`, the space left over is split evenly top and
 * bottom, so the alphabetic baseline lands at `(H - (A + D)) / 2 + A` from the
 * top of the box. `Grid.tsx` got this from `line-height: 1.269`; a canvas has
 * to compute it.
 */
export function lineMetrics(font: string, sizePx: number): LineMetrics {
  const key = font;
  const known = METRICS.get(key);
  if (known !== undefined) return known;

  const lineHeight = sizePx * LINE_HEIGHT;
  let ascent = sizePx * 0.8;
  let descent = sizePx * 0.2;
  const context = measureContext();
  if (context) {
    context.font = font;
    const measured = context.measureText("Mg");
    const fontAscent = measured.fontBoundingBoxAscent;
    const fontDescent = measured.fontBoundingBoxDescent;
    if (Number.isFinite(fontAscent) && fontAscent > 0) ascent = fontAscent;
    if (Number.isFinite(fontDescent) && fontDescent >= 0) descent = fontDescent;
  }
  return bounded(METRICS, key, { ascent, descent, lineHeight });
}

/**
 * The value as Excel draws it in a cell that does not wrap: the line breaks are
 * not there at all — not a break, and not a space either, so the text runs on
 * across them. `measureText` charges a space for `\n`, so a value like this has
 * to lose them before anything measures it. Pinned against Excel by rendering
 * `ABC\nDEF` in an unwrapped cell: it comes out exactly as wide as `ABCDEF`,
 * where `ABC DEF` is a space wider (`MS PGothic` 11pt and 20pt alike).
 */
export function flatText(text: string): string {
  return text.replace(/\n/g, "");
}

/**
 * The lines a cell's text occupies. A wrapped cell (`whitespace-pre-wrap
 * break-words`) breaks at spaces and, where a single word cannot fit, inside
 * the word, and a break in the value is a line of its own; an unwrapped one is
 * a single line of the whole value, its breaks drawn as nothing.
 *
 * Spaces are preserved rather than collapsed, which is what `pre-wrap` means
 * and what these documents rely on for their indentation runs.
 *
 * `measure` overrides how a slice of the value is measured, for text whose runs
 * carry fonts of their own — a word in a larger font has to wrap where Excel
 * wraps it, not where the cell's own font would have.
 */
export function layoutLines(
  font: string,
  text: string,
  maxWidth: number,
  wrap: boolean,
  measure?: (from: number, to: number) => number,
): string[] {
  if (!wrap) return [flatText(text)];
  if (maxWidth <= 0) return [""];

  const width = measure ?? ((from: number, to: number) => textWidth(font, text.slice(from, to)));
  const lines: string[] = [];
  /** Where the line being built starts in `text`, which is what `measure` takes. */
  let pos = 0;
  for (const paragraph of text.split("\n")) {
    const end = pos + paragraph.length;
    if (paragraph === "") {
      lines.push("");
      pos = end + 1;
      continue;
    }
    let line = "";
    // Split keeps the separators so the spaces can be re-emitted verbatim.
    for (const chunk of paragraph.split(/(?<=\s)/)) {
      if (width(pos, pos + line.length + chunk.length) <= maxWidth) {
        line += chunk;
        continue;
      }
      if (line !== "") {
        lines.push(line.replace(/\s+$/, ""));
        pos += line.length;
        line = "";
      }
      // The chunk alone may still be too wide — break it mid-word, as
      // `break-words` does.
      let rest = chunk;
      while (rest !== "" && width(pos, pos + rest.length) > maxWidth) {
        let lo = 1;
        let hi = rest.length;
        while (lo < hi) {
          const mid = (lo + hi + 1) >> 1;
          if (width(pos, pos + mid) <= maxWidth) lo = mid;
          else hi = mid - 1;
        }
        lines.push(rest.slice(0, lo));
        rest = rest.slice(lo);
        pos += lo;
      }
      line = rest;
    }
    lines.push(line.replace(/\s+$/, ""));
    pos = end + 1;
  }
  return lines;
}

export interface RunSpan {
  /** Offsets into the string the runs spell out, which is the cell's text. */
  start: number;
  end: number;
  color: string | null;
  /** The run's own font, or null when its whole font comes from the cell. */
  font: RichFont | null;
}

/**
 * Where each run sits inside the string they concatenate to, and what it draws
 * with.
 *
 * Drawing a line in its runs' own fonts and colours means knowing which part of
 * the source each line came from, and the runs are laid out along that same
 * source, so this is the map from one to the other.
 */
export function runSpans(runs: RichRun[]): RunSpan[] {
  const spans: RunSpan[] = [];
  let at = 0;
  for (const run of runs) {
    spans.push({ start: at, end: at + run.text.length, color: run.color, font: run.font ?? null });
    at += run.text.length;
  }
  return spans;
}

/** Drops every measurement cache — call after `document.fonts.ready`. */
export function clearTextCaches(): void {
  WIDTHS.clear();
  METRICS.clear();
  // Font strings themselves are font-independent, so they survive, but the
  // stacks are cheap to rebuild and keeping them in step is simpler.
  STACKS.clear();
  FONTS.clear();
}
