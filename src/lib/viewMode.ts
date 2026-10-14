import type { WorkbookInfo } from "@/lib/excel";

/**
 * Which of Excel's views the grid is drawing. `pageBreak` is Page Break
 * Preview — the print area shaded around and boxed, so where the page ends can
 * be seen. `normal` is the ordinary view, drawn with no page at all.
 *
 * It belongs to the workbook rather than to a sheet, because that is what
 * decides it: whether the file set a print area anywhere.
 */
export type ViewMode = "normal" | "pageBreak";

/**
 * What every sheet of a workbook starts in.
 *
 * The page view is for a workbook that means to be printed — somebody set a
 * print area on one of its sheets, and the sheets that were not given one are
 * still read beside the ones that were. Where nobody set one anywhere, the
 * shading would be drawn around a page the app worked out for itself and
 * nothing in the file asks for, so those open plainly, as Excel draws them.
 *
 * A default only: each sheet is switched on its own from there.
 */
export function initialView(info: WorkbookInfo): ViewMode {
  return info.sheets.some((sheet) => sheet.printArea) ? "pageBreak" : "normal";
}
