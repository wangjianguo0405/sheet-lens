# Changelog

**English** | [简体中文](CHANGELOG.zh-CN.md) | [日本語](CHANGELOG.ja.md)

One section per release. The section for the version being published becomes
both the release notes on GitHub and the notes the app shows when it offers the
update — so write it for whoever is being asked to install it, not for whoever
wrote the code.

Within a section the entries are grouped: **Added** for what a sheet could not
do before, **Improved** for what it now does differently, and **Fixed** for what
it used to get wrong. A release with nothing to say under one of them leaves
that group out.

## [0.3.6] - 2026-10-13

The sheets can be picked from the bottom too, the way Excel keeps them.

### Added

- A strip of sheet tabs along the foot of the sheet area, beside the list on the
  left. Either can be shown or hidden on its own, and the choice is remembered
- Each tab carries the colour the workbook gives it, so the sheets are told
  apart the way they are in Excel
- The name box takes a reference: type `C16`, or `Sheet1!C16` to go to another
  sheet with it, and press Enter

### Improved

- While the tabs are showing, the sheet's horizontal scrollbar sits at the end
  of the strip, and the handle beside it makes it longer or shorter
- The formula bar's scrollbar is the same one as everywhere else in the app
- The toolbar has given up its Go To button — the name box and Ctrl+G were two
  ways to do the same thing already

### Fixed

- Scrolling no longer slows down as it passes the last of the data. The view was
  asking for the empty cells beyond it a block at a time and waiting on the
  answer

## [0.3.5] - 2026-10-13

The empty part of a sheet behaves itself.

### Improved

- Past the end of the data the scrollbar moves the way Excel's does: the range
  stands still until the view reaches its end, instead of growing outwards from
  the middle

### Fixed

- A table's rules and shading no longer run on past its data into the blank part
  of the sheet
- A sheet's drawings stay inside the sheet instead of being drawn again out to
  the right of the data
- The first frame after switching sheets is right first time, without having to
  scroll it to wake it up

## [0.3.4] - 2026-10-12

A sheet no longer ends where its data does.

### Added

- A sheet can be scrolled past its last cell, into the empty part of it, as
  Excel's can be. The blank area is made as the view reaches its end and given
  back as the view leaves it

### Fixed

- A copied merged cell is outlined around the whole merged block, not just
  around the corner the cursor sits on
- A sheet whose filter hides its last row no longer comes up grey from corner to
  corner
- The update prompt shows a release's notes as they are written, without the
  markers that group them

## [0.3.3] - 2026-10-12

Ctrl and the arrows now move the cursor the way Excel's do, and a copy says so.

### Added

- Ctrl+arrow jumps to the edge of the data — the far end of a run of filled
  cells, or the next filled cell past a blank one — and Ctrl+Shift+arrow extends
  the selection there
- Ctrl+C draws the dashed rule Excel draws around what was copied, and it crawls
  until Escape takes it away from wherever the focus happens to be

### Fixed

- Switching away from a sheet and back no longer moves the view: each sheet
  keeps the position it was left at
- A copy no longer leaves the keyboard pointing at nothing

## [0.3.2] - 2026-10-12

Two things that went wrong on a long sheet.

### Fixed

- Borders no longer stop partway down a heavily formatted sheet. The parser
  gave up on reading styles past a certain point, and what it dropped read as
  missing borders to the right of the table
- The row numbers stay out of the corner where the row and column headers meet
  while the sheet is scrolled

## [0.3.1] - 2026-10-12

Updating replaces more of the folder now. It used to swap only `SheetLens.exe`,
leaving the README and the changelog beside it describing whatever version
happened to be unpacked there first.

### Improved

- An update rewrites the README, the changelog and the licence in the folder too

### Fixed

- The README now describes the grid as it is — drawn on a canvas since 0.3.0,
  not rendered in CSS

## [0.3.0] - 2026-10-12

The grid is painted onto a canvas instead of being built out of an element per
cell, which is what used to make a long sheet slow to scroll and slow to click
in. The rest of this release is the keyboard catching up with it.

### Added

- Shift with the arrow keys, Page Up/Down or Home/End extends the selection, and
  Ctrl+A selects the data around the cursor — the whole sheet on a second press
- Dragging along the row or column headers selects a run of them, `3:7` or `C:E`
- Ctrl+Home and Ctrl+End jump to the two corners, Ctrl+Page Up/Down change sheet

### Improved

- Large sheets are responsive again: scrolling, selecting and clicking no longer
  stall the way they did
- A new icon

## [0.2.0] - 2026-10-11

The interface now comes in English, Chinese and Japanese, and opens in your
system language. The button at the right of the toolbar switches it, and what
you pick is remembered.

### Added

- Simplified Chinese and Japanese throughout
- A language switcher in the toolbar
- The notes shown when an update is available are in that language too

## [0.1.1] - 2026-10-11

### Improved

- The update dialog now shows how far the download has got, instead of a line
  that only said "downloading".

## [0.1.0] - 2026-10-11

First release: a portable Windows build of SheetLens that can update itself in
place. Unpack it anywhere and run `SheetLens.exe`; it keeps itself current from
the releases on this repository.

### Added

- Opens `.xlsx`, `.xlsm`, `.xls`, `.xlsb` and `.ods`
- Sheets, search, go-to, column filters and a formula bar
- Cell, row/column and rectangular-block selection, with Ctrl+C in Excel's own
  format
- Ctrl+wheel zoom on Excel's own steps
- Excel-faithful layout: column widths, merges, borders, fills, fonts, the print
  area, and the shapes and pictures floating above the cells
