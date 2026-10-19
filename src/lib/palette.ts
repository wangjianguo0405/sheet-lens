/**
 * The colours every colour menu offers, and the ones the reader has just used.
 *
 * One set for the whole app rather than one per menu: a cell's fill, a review
 * mark and a sheet tab's colour are all chosen from the same palette in Excel,
 * and the colours a reader reaches for are the same colours wherever they are
 * pointing. Only the colours they are *pointing at* differ, and that is the
 * state each menu keeps for itself.
 */

/**
 * Excel's own Standard Colours row, so the palette is the one a reader already
 * knows rather than a set invented here.
 */
export const STANDARD_COLORS = [
  "#C00000",
  "#FF0000",
  "#FFC000",
  "#FFFF00",
  "#92D050",
  "#00B050",
  "#00B0F0",
  "#0070C0",
  "#002060",
  "#7030A0",
];

/**
 * The Office theme's colours, laid out the way Excel lays them out: ten families
 * across, five rows down, each family's own colour in the top row with its
 * lighter tints under it and its darkest at the foot.
 *
 * Excel's theme grid holds whatever theme the workbook carries, and a file can
 * name its own accents — these are the ones Excel itself ships with, since the
 * parser does not read a workbook's theme out.
 */
export const THEME_COLORS = [
  [
    "#FFFFFF",
    "#000000",
    "#E7E6E6",
    "#44546A",
    "#4472C4",
    "#ED7D31",
    "#A5A5A5",
    "#FFC000",
    "#5B9BD5",
    "#70AD47",
  ],
  [
    "#F2F2F2",
    "#7F7F7F",
    "#D0CECE",
    "#D6DCE5",
    "#D9E2F3",
    "#FBE5D6",
    "#EDEDED",
    "#FFF2CC",
    "#DEEBF7",
    "#E2EFDA",
  ],
  [
    "#D9D9D9",
    "#595959",
    "#AEAAAA",
    "#ADB9CA",
    "#B4C6E7",
    "#F8CBAD",
    "#DBDBDB",
    "#FFE699",
    "#BDD7EE",
    "#C6E0B4",
  ],
  [
    "#BFBFBF",
    "#404040",
    "#757070",
    "#8496B0",
    "#8FAADC",
    "#F4B183",
    "#C9C9C9",
    "#FFD966",
    "#9DC3E6",
    "#A9D08E",
  ],
  [
    "#7F7F7F",
    "#262626",
    "#3A3838",
    "#323F4F",
    "#2F5496",
    "#C55A11",
    "#7B7B7B",
    "#BF8F00",
    "#2E75B6",
    "#548235",
  ],
];

/**
 * What the paint bucket starts out holding, and fills with on its own: the
 * theme's light gold, which is a shade a cell's text still reads through.
 */
export const DEFAULT_COLOR = "#FFE699";

const PALETTE = new Set(
  [...STANDARD_COLORS, ...THEME_COLORS.flat()].map((color) => color.toUpperCase()),
);

/** Whether a stored value is one of the palette's, rather than anything at all. */
export function isPaletteColor(value: unknown): value is string {
  return typeof value === "string" && PALETTE.has(value.toUpperCase());
}

/** How many of them a palette shows, which is one row of them. */
export const RECENT_SHOWN = 10;

const RECENT_KEY = "sheetlens.recentColors";

/**
 * The colours last picked anywhere, most recent first.
 *
 * Kept apart from the menus that show them, because the row is the same row in
 * all of them: a colour used to mark a cell is one the reader may well reach for
 * again on a sheet tab.
 */
export function readRecentColors(): string[] {
  try {
    const stored = JSON.parse(localStorage.getItem(RECENT_KEY) ?? "[]");
    if (!Array.isArray(stored)) return [];
    return stored.filter(isPaletteColor).slice(0, RECENT_SHOWN);
  } catch {
    // A webview that refuses storage, or a value it wrote badly, has no list.
    return [];
  }
}

/** Puts `fill` at the head of that list, where a palette will show it. */
export function rememberColor(fill: string): void {
  try {
    const next = [fill, ...readRecentColors().filter((other) => other !== fill)];
    localStorage.setItem(RECENT_KEY, JSON.stringify(next.slice(0, RECENT_SHOWN)));
  } catch {
    // As above: the colour is still used, it just does not survive the window.
  }
}
