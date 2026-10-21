import type { OutlineCol, OutlineRow } from "./excel";
import { COL_HEADER_H, ROW_HEADER_W } from "./gridGeometry";

/**
 * Excel's outline (Data > Group), turned from the levels the file records into
 * the brackets and buttons the grid draws.
 *
 * The file does not write a group as a range. Every *detail* line carries
 * `outlineLevel`, and the line beside the group — its summary — carries
 * `collapsed="1"` while the group is closed. So a group is the maximal run of
 * consecutive lines whose level reaches `L`, and its summary is the line just
 * outside that run, below by default and above when `<outlinePr summaryBelow>`
 * says so.
 *
 * The chrome is view decoration: `CopyPicture` and the PDF do not contain it,
 * so every number here was measured off Excel's own window at 100% zoom:
 *
 * - The gutter holds one 16px strip per level, plus one strip for "no level",
 *   with 2px before the first and 2px after the last.
 * - A bracket is a 1px `#ABABAB` line down the middle of its strip, with a
 *   4px cap at the end away from the summary.
 * - The +/- button is 16x16, 1px `#828790` border, `#E4E4E4` face, a
 *   `#646464` glyph, and sits on the summary line.
 */
export interface OutlineGroup {
  /** 1-based, as `<row outlineLevel>` writes it. */
  level: number;
  /** First and last detail line (inclusive). */
  from: number;
  to: number;
  /** The line carrying the button — outside the detail run. */
  summary: number;
}

export interface LineOutline {
  groups: OutlineGroup[];
  /** Line -> level, for the level buttons and for the dots. */
  levels: Map<number, number>;
  /** How many strips the gutter needs: the deepest level, plus one. */
  strips: number;
  /** Whether a group's summary line comes after its detail (below / right). */
  summaryAfter: boolean;
}

export const OUTLINE_STRIP_W = 16;
export const OUTLINE_PAD = 2;
export const OUTLINE_TAIL = 2;
/** The line Excel draws, measured 1px at 100%. */
export const OUTLINE_LINE_W = 1;
/** Length of the bracket's end cap, toward the numbers. */
export const OUTLINE_CAP = 4;
/** The +/- and level buttons; they tile their strip. */
export const OUTLINE_BUTTON = 16;

/** The +/- square on a group's summary line, centred in its strip. */
export function groupButtonBox(
  strip: number,
  lineStart: number,
  lineSize: number,
  vertical: boolean,
): { x: number; y: number; w: number; h: number } {
  const size = Math.min(OUTLINE_BUTTON, Math.max(0, lineSize));
  const offset = (OUTLINE_STRIP_W - size) / 2;
  if (vertical) {
    return {
      x: lineStart + (lineSize - size) / 2,
      y: stripStart(strip) + offset,
      w: size,
      h: size,
    };
  }
  return {
    x: stripStart(strip) + offset,
    y: lineStart + (lineSize - size) / 2,
    w: size,
    h: size,
  };
}

export const OUTLINE_LINE_COLOR = "#ABABAB";
export const OUTLINE_BUTTON_BORDER = "#828790";
export const OUTLINE_BUTTON_FACE = "#E4E4E4";
export const OUTLINE_BUTTON_INK = "#646464";

/** The room the gutter takes beside the row numbers or column letters. */
export function outlineGutter(strips: number): number {
  return strips <= 0
    ? 0
    : OUTLINE_PAD + strips * OUTLINE_STRIP_W + OUTLINE_TAIL;
}

/** Left edge of strip `index` (0-based), within the gutter. */
export function stripStart(index: number): number {
  return OUTLINE_PAD + index * OUTLINE_STRIP_W;
}

/** Left edge of the line itself: 1px, a shade left of the strip's centre. */
export function stripLine(index: number): number {
  return OUTLINE_PAD + 5 + index * OUTLINE_STRIP_W;
}

/**
 * The corner's level buttons, horizontal (rows) or vertical (columns).
 *
 * Excel centres the little cluster: the row buttons sit mid-height in the
 * column-letter band, and the column buttons mid-width over the row numbers,
 * with the same margin on both sides. When the sheet groups rows too, the
 * cluster moves right of the row gutter (a capture of a sheet with both shows
 * the columns' buttons at 55–78 of an 81px corner); a cluster wider than the
 * header band overflows evenly on both sides.
 */
export function rowLevelButtonRect(
  index: number,
  colGutter: number,
): { x: number; y: number; w: number; h: number } {
  return {
    x: stripStart(index),
    y: colGutter + Math.round((COL_HEADER_H - OUTLINE_BUTTON) / 2),
    w: OUTLINE_BUTTON,
    h: OUTLINE_BUTTON,
  };
}

export function colLevelButtonRect(
  index: number,
  rowGutter: number,
): { x: number; y: number; w: number; h: number } {
  // They stack down one column, so it is one button's width being centred,
  // not the run of strips.
  return {
    x: rowGutter + Math.round((ROW_HEADER_W - OUTLINE_BUTTON) / 2),
    y: stripStart(index),
    w: OUTLINE_BUTTON,
    h: OUTLINE_BUTTON,
  };
}

/**
 * Groups from the file's level facts.
 *
 * A level-2 group inside a level-1 group is found the same way as any other:
 * the level-1 run is every line at level >= 1, and the level-2 run the lines
 * inside it at level >= 2.
 */
export function buildLineOutline(
  facts: { line: number; level: number }[],
  summaryAfter: boolean,
): LineOutline {
  const levels = new Map<number, number>();
  for (const fact of facts) {
    if (fact.level > 0) levels.set(fact.line, fact.level);
  }
  let maxLevel = 0;
  for (const level of levels.values()) maxLevel = Math.max(maxLevel, level);
  const groups: OutlineGroup[] = [];
  if (maxLevel > 0) {
    const lines = Array.from(levels.entries())
      .filter(([, level]) => level >= 1)
      .sort((a, b) => a[0] - b[0]);
    for (let level = 1; level <= maxLevel; level += 1) {
      let from = -1;
      let to = -1;
      const close = () => {
        if (from < 0) return;
        groups.push({
          level,
          from,
          to,
          summary: summaryAfter ? to + 1 : from - 1,
        });
        from = -1;
      };
      for (const [line, own] of lines) {
        if (own < level) {
          close();
          continue;
        }
        if (from < 0) from = line;
        else if (line !== to + 1) {
          close();
          from = line;
        }
        to = line;
      }
      close();
    }
  }
  return {
    groups,
    levels,
    strips: maxLevel > 0 ? maxLevel + 1 : 0,
    summaryAfter,
  };
}

export function rowOutline(outline: {
  rows: OutlineRow[];
  summaryBelow: boolean;
}): LineOutline {
  return buildLineOutline(
    outline.rows.map((entry) => ({ line: entry.row, level: entry.level })),
    outline.summaryBelow,
  );
}

export function colOutline(outline: {
  cols: OutlineCol[];
  summaryRight: boolean;
}): LineOutline {
  const facts: { line: number; level: number }[] = [];
  for (const run of outline.cols) {
    for (let col = run.from; col <= run.to; col += 1) {
      facts.push({ line: col, level: run.level });
    }
  }
  return buildLineOutline(facts, outline.summaryRight);
}

/**
 * Whether a group is shown open. The file's `collapsed` flag is the saved
 * answer, but the reader can open and close a group without writing anything
 * back, so the line's own visibility is what decides: a group is open as soon
 * as any of its detail lines is on screen.
 */
export function groupOpen(
  group: OutlineGroup,
  hidden: Set<number>,
  displayOf: (line: number) => number,
): boolean {
  for (let line = group.from; line <= group.to; line += 1) {
    if (!hidden.has(line) && displayOf(line) >= 0) return true;
  }
  return false;
}

/** The detail lines a click on the group's button changes. */
export function groupToggleLines(
  group: OutlineGroup,
  levelOf: (line: number) => number,
  open: boolean,
): number[] {
  const lines: number[] = [];
  for (let line = group.from; line <= group.to; line += 1) {
    const level = levelOf(line);
    // Closing takes the whole run, nested lines included. Opening shows only
    // the group's own level and leaves anything deeper as it was — a nested
    // group stays closed, which is what Excel does when it opens an outer
    // group whose inner one was closed.
    if (open ? level >= group.level : level === group.level) lines.push(line);
  }
  return lines;
}
