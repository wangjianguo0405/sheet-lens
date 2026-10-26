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

## [1.0.10] - 2026-10-26

### Added

- The fixed sheets in the sidebar's list can be dragged into the order you want them in; the rest of the list keeps the workbook's own order

### Improved

- The start page's two columns are one height whatever they hold, and scroll as a whole — the fixed run stays at the top of its column while the rest goes by
- The start page's history is listed in full: the "More" step before the list would scroll is gone

## [1.0.9] - 2026-10-26

### Added

- The sheet list on the left keeps the sheets you fix at its top, under a heading of their own: a workbook of many sheets no longer means hunting for the same two every time. A fixed sheet is remembered for its workbook, and nothing in the file changes

### Improved

- A fixed sheet is the same gesture as a fixed workbook: the heading over the fixed ones, the most recently fixed at the top, and the pin showing under the pointer and staying once it is fixed
- The start page's two columns scroll at one height, so their two lists no longer end at different points

## [1.0.8] - 2026-10-23

### Added

- A ribbon under the file tabs: view (sidebar, sheet tabs, freeze, split), font (family, size, bold, italic, underline), alignment, the review fill, and data (filter, outline). Its edits last for the session and are never written back to the file
- Excel's split view: every row band and column band gets its own scrollbar, drawn in the sheet tab row while the tabs are shown; a split across rows only, or columns only, gives the two panes it should, and a split saved in the file opens as one
- The keyboard a sheet is read with: Shift+Space and Ctrl+Space take a whole row or column, Ctrl+Shift+Home/End carry the selection to the sheet's ends, Enter and Tab step the cursor, Alt+PageUp/PageDown move a screen sideways, Ctrl+Tab walks the open workbooks, Ctrl+B/I/U set the selection's bold, italic and underline, Ctrl+Shift+L switches the filter, and F5 opens Go To
- The list the sheet tab arrows open now has a filter: type part of a name to narrow it, and Enter takes the first match
- The start page keeps folders of its own: pin the folder a workbook was opened from, and remove one together with the workbooks it holds
- The font list is every font installed on this machine, not a fixed handful
- A double press on the empty part of the file tab strip opens another workbook, as the "+" beside it does

### Improved

- Hover hints are the app's own: they show at once instead of after the system's pause, look like the rest of the app, and never stand between the pointer and the control under it
- The start page separates what is fixed from what was opened lately, and its shortcut list reads in three columns

### Fixed

- Ctrl+F with the search panel already open puts the cursor back in its field instead of closing the panel
- In page break preview a merged block no longer shows a pale outline around itself at 125% and 150% zoom

## [1.0.7] - 2026-10-22

### Added

- A workbook opens where it was left: the same sheet, view, zoom and cursor, instead of always the first sheet at A1
- Conditional formatting for date periods — today, yesterday, this week, last month and the rest — is drawn as Excel draws it
- Justified and distributed text spreads across the cell, and accounting underlines run the whole width of the cell
- A workbook that replaces Excel's indexed colours gets its own palette, and a cell marked right-to-left is laid out that way

### Improved

- Each of the thirteen border styles draws its own pattern; six dashed variants used to share one
- The row-number column widens as the visible row numbers gain digits

### Fixed

- A negative data bar is drawn black unless the file names a colour; it used to take the positive bar's colour

## [1.0.6] - 2026-10-22

### Improved

- A filter column's drop-down lists the values the other columns' filters leave on screen, as Excel does, instead of the whole column

### Fixed

- An in-cell list whose values the file escapes as character references now opens, and a list that names its range (a defined name such as `対象`) offers that range's values
- The in-cell drop-down arrow sits on the cell's right edge, along the bottom, as Excel draws it — and the press that opens the list follows it there
- A workbook with saved custom views opens at its own frozen rows, scroll position and column widths rather than at the last saved view's, and draws the filter buttons Excel shows
- A narrow filter header keeps its text instead of cutting it down to the few pixels left beside the button

## [1.0.5] - 2026-10-22

### Added

- The Unhide dialog filters the list and can bring several sheets back at once: Shift-click takes a run of names, everything picked returns together, and the last one picked becomes the current sheet

### Improved

- Hairline borders are dotted, the way Excel draws them, and a dashed or dotted edge keeps its rhythm from cell to cell instead of joining two dots into a wider one
- A filter button is drawn only on the columns the file gives one — a merged header carries a single button at its right edge — and it never grows past its row

### Fixed

- A sheet with a fixed print scale whose print area is wider than the paper now breaks pages where Excel breaks them
- The page number in Page Break Preview is the size Excel draws it, and shrinks to stay inside a narrow or short page
- A gridline no longer shows through the gaps of a dashed or dotted border

## [1.0.4] - 2026-10-21

### Added

- Outline groups as Excel draws them: the bracket beside the rows or columns a group holds, the +/− button on its summary line, and the level buttons in the corner of the gutter that step through the levels a sheet has — a change that lasts for the session only
- Page Break Preview now breaks pages where Excel breaks them: the paper, the margins, the print scale and the rows and columns repeated on every page are read from the file; a break the paper makes is drawn as Excel's dashed rule, one placed by hand as a solid rule, and every page carries its number across the middle
- The square in the top-left corner now selects the whole sheet, as it does in Excel

### Fixed

- A row or column scrolled under the header no longer draws its rule across the top-left corner and the grouping gutter — the header covers them, and a separator stays inside its own band
- A filter button no longer floats on top of a frozen row or column
- A sheet that only asks to fit one page is no longer split into four: a fit with no page count named means one page, which is what the file leaves out when it means one

## [1.0.3] - 2026-10-20

### Added

- Pattern fills — Excel's hatches — on cells and on shapes: a cell's pattern is pinned to the sheet and scales with the zoom, a shape's to its own corner and at a fixed size
- Radial gradient fills, drawn along the gradient's own path: through the corners of a circle, the edges of a rectangle, or an ellipse for any other shape
- Notes: the red corner mark on a cell that has one, the note's text and Excel's leader on hover, and a Comments list in the sidebar that shows an entry's note in the sheet
- A shape can be dragged aside to see the cells it covers; the move lasts until the file is closed and is never written back
- A sheet opens at the view the file was left at — its zoom and its first visible cell, frozen panes included

### Fixed

- Zooming out paints the whole window: the right and bottom of the sheet no longer keep the previous zoom's scale, and the row and column labels and the cells behind them follow the zoom across the full view

## [1.0.2] - 2026-10-20

### Added

- Diagonal borders, drawn corner to corner — across the whole region when the cell is merged
- Superscript and subscript text at Excel's own size and offset: two thirds of the font size, raised half an em or lowered a fifth
- A column the file itself closes with a filter such as "greater than 20" shows Excel's funnel and blue row numbers
- A file tab's right-click menu groups the copies under one submenu, opens the workbook with whichever app the system uses for it, and ends with the file's location
- Copying, or failing to open a file or its folder, now says so in a notice at the corner of the window

### Fixed

- A row's stored style now applies only when the file marks the row custom, as Excel does

## [1.0.1] - 2026-10-20

### Added

- A sheet tab's right-click menu copies the sheet's name, ready for a message or a bug report
- A file tab's right-click menu opens the workbook's folder with the file selected — the item Excel's own workbook tab carries

### Fixed

- A filtered column's button draws Excel's own mark again: the funnel high on the right with a small caret under its left, instead of a centred funnel with the caret directly beneath it

## [1.0.0] - 2026-10-19

SheetLens opens a workbook and draws what Excel draws — the grid, the text in
its cells, and the shapes and pictures above them — without opening Excel.

### Added

- `xlsx`, `xlsm`, `xls`, `xlsb` and `ods` workbooks; the older formats have no zip to read, so they fall back to plain defaults
- A grid drawn from the workbook's own XML, not from a spreadsheet library: column widths, row heights and merged regions at the sizes Excel gives them
- Cell styling as Excel draws it: borders, fills, fonts, alignment (centre-across included), indent, wrapped, shrunk and spilling text, and Excel's own number formats
- Rich text with a run's own bold, italic, underline, strikethrough, face or size inside a single value, each run overriding the cell's own decoration
- Text turned in its cell — 90°, 180°, and one character under the next — placed where Excel puts it
- Rows and columns the file hides take no room and their headings are not drawn, so the numbering jumps exactly where Excel's does
- Sheets as Excel shows them: hidden sheets, tab colours, groups hidden together, and frozen rows and columns that stay while the rest scrolls
- A tab strip that behaves like Excel's: file tabs drag to reorder and close on a middle click, sheet tabs drag to move, a “…” at either end reaches the ones that no longer fit, a tab's right-click copies its name or path, and a sheet tab's name opens on double-click to select and copy
- The start page: recently opened and pinned workbooks, a list of the app's shortcuts, and a way to forget an entry or clear the list
- Several workbooks open at once, each with a File menu, reload, a full list when the strip overflows, and the sheet count in the status bar
- Selection by cell, row, column and block, with Ctrl for several regions and Shift for a run; Ctrl+A; the name box lists the areas, the status bar counts them, and Ctrl+C copies them all in Excel's format
- The name box and Go To take a reference or a defined name, another sheet's included
- Excel's formula bar, which also expands to show a value too long for its cell
- A double-click or F2 opens a cell's value over the grid, selectable and copyable without leaving the sheet
- Column widths and row heights drag from a heading's border, and a double-click fits them to their content; both are kept while the workbook is open
- Search a sheet or the whole workbook; the arrow keys walk the results, and the field offers the words searched this session
- Column filters with Excel's own button on the header, a distinct-value list, the count of hidden rows in the status bar, and blue row numbers where the filter is working
- Conditional formatting for formula, comparison, text, duplicate, rank and average rules, plus colour scales, data bars and icon sets
- A validation list opens in the cell, and an input prompt shows Excel's note beside it
- Shapes from ECMA-376's 187-preset table and from geometry the file supplies: groups, crops, fills and gradients, line joins, connectors and arrows, and text with its own font, size, colour and paragraph direction
- Pictures above the grid, drawn through their crop
- Reader-side marking that is never written back: a paint-bucket fill from Excel's palette, and hidden rows, columns and sheets
- Excel's view control: Page Break Preview with the print area shaded and boxed, and Ctrl+wheel zoom on Excel's steps
- Large workbooks open and switch sheets without freezing the window
- English, Chinese and Japanese, chosen from the system
- A portable Windows build that updates itself from signed releases
