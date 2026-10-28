/**
 * The colour the font-colour button paints with.
 *
 * The same arrangement the review bucket has: a colour kept between sessions,
 * because Excel's own button still shows the colour it was last given the next
 * day, and the cells it was laid over are forgotten with their tab.
 *
 * It is a font colour rather than a fill, so the default is Excel's red — the
 * one a reader reaches for to point at something — and it is chosen from the
 * same palette, with its picks remembered in the same recent row.
 */
import { isPaletteColor } from "@/lib/palette";

const STORAGE_KEY = "sheetlens.fontColor";

/** Excel's own "Automatic" red, which its font-colour button starts on. */
export const DEFAULT_FONT_COLOR = "#C00000";

export function initialFontColor(): string {
  try {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (isPaletteColor(stored)) return stored;
  } catch {
    // A webview that refuses storage still gets the default.
  }
  return DEFAULT_FONT_COLOR;
}

export function storeFontColor(color: string): void {
  try {
    localStorage.setItem(STORAGE_KEY, color);
  } catch {
    // As above: the colour changes, the choice just does not survive the window.
  }
}
