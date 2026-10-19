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
