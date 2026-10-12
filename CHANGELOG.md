# Changelog

**English** | [简体中文](CHANGELOG.zh-CN.md) | [日本語](CHANGELOG.ja.md)

One section per release. The section for the version being published becomes
both the release notes on GitHub and the notes the app shows when it offers the
update — so write it for whoever is being asked to install it, not for whoever
wrote the code.

## [0.3.2] - 2026-10-12

Two things that went wrong on a long sheet.

- Borders no longer stop partway down a heavily formatted sheet. The parser
  gave up on reading styles past a certain point, and what it dropped read as
  missing borders to the right of the table
- The row numbers stay out of the corner where the row and column headers meet
  while the sheet is scrolled

## [0.3.1] - 2026-10-12

Updating replaces more of the folder now. It used to swap only `SheetLens.exe`,
leaving the README and the changelog beside it describing whatever version
happened to be unpacked there first.

- An update rewrites the README, the changelog and the licence in the folder too
- The README now describes the grid as it is — drawn on a canvas since 0.3.0,
  not rendered in CSS

## [0.3.0] - 2026-10-12

The grid is painted onto a canvas instead of being built out of an element per
cell, which is what used to make a long sheet slow to scroll and slow to click
in. The rest of this release is the keyboard catching up with it.

- Large sheets are responsive again: scrolling, selecting and clicking no longer
  stall the way they did
- Shift with the arrow keys, Page Up/Down or Home/End extends the selection, and
  Ctrl+A selects the data around the cursor — the whole sheet on a second press
- Dragging along the row or column headers selects a run of them, `3:7` or `C:E`
- Ctrl+Home and Ctrl+End jump to the two corners, Ctrl+Page Up/Down change sheet
- A new icon

## [0.2.0] - 2026-10-11

The interface now comes in English, Chinese and Japanese, and opens in your
system language. The button at the right of the toolbar switches it, and what
you pick is remembered.

- Simplified Chinese and Japanese throughout
- A language switcher in the toolbar
- The notes shown when an update is available are in that language too

## [0.1.1] - 2026-10-11

- The update dialog now shows how far the download has got, instead of a line
  that only said "downloading".

## [0.1.0] - 2026-10-11

First release: a portable Windows build of SheetLens that can update itself in
place. Unpack it anywhere and run `SheetLens.exe`; it keeps itself current from
the releases on this repository.

- Opens `.xlsx`, `.xlsm`, `.xls`, `.xlsb` and `.ods`
- Sheets, search, go-to, column filters and a formula bar
- Cell, row/column and rectangular-block selection, with Ctrl+C in Excel's own
  format
- Ctrl+wheel zoom on Excel's own steps
- Excel-faithful layout: column widths, merges, borders, fills, fonts, the print
  area, and the shapes and pictures floating above the cells
