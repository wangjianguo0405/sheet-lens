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

## [0.3.8] - 2026-10-13

A cell's text can be read and copied without leaving the grid, rows are fitted
to what is in them, and a filter lands where the workbook says it does.

### Added

- A cell's text opens over the cell when it is double-clicked — selectable and
  copyable with the mouse, its coloured runs and all, and not changeable
- A row's height is fitted to what is in it by double-clicking the bottom border
  of its header
- The File menu's list of recent workbooks can be emptied of an entry, or
  cleared whole

### Improved

- A Shift-click on a row or column header extends the run, the way it does on a
  cell
- Dragging a row height with several rows selected sets them all, and a
  double-click on the border fits each of them to its own content
- A sheet keeps the filter it was left with when you leave it and come back,
  as Excel does

### Fixed

- The filter button sits on the row the workbook's own filter names. A sheet
  with saved custom views was read from the last of those instead: a WBS whose
  filter is `M1:FC66` was taken as some view's `A7:EY121`, seven rows lower
- A filter hides the rows inside its range and nothing else. What sits under a
  table used to vanish with it
- The filter banner's row count and its Clear button are translated like the
  rest of the interface

## [0.3.7] - 2026-10-13

Recent workbooks are one click away, and the sheet tabs are painted the way
Excel paints them.

### Added

- A File menu, where Open used to be: open a workbook, open one of the fifty
  opened lately — ten listed in the menu and the rest under More — or close the
  one showing
- A tab is closed by a middle click, the gesture browsers put on one
- A recent whose file has gone is greyed, with the reason in its tooltip.
  Clicking it tries again, and opening it puts the entry right

### Improved

- A sheet the workbook gives a tab colour is filled with it, name and all, in
  white type or black as the colour needs. The current sheet is not filled: it
  keeps the raised tab, washed in that colour with the rule along its foot
- The interface's green is the one Excel draws today, `#107C41`, where it had
  Excel 2013's `#217346`
- The toolbar gave up its Search button. Ctrl+F opens the search panel, and the
  button was the second way to do one thing
- A coloured sheet tab lights up when the pointer is over it, as the plain ones
  already did

### Fixed

- The More submenu stays inside the window: once it is taller than the space it
  has, it scrolls, rather than drawing its last entries out of reach

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
