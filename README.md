# SheetLens

**English** | [简体中文](README.zh-CN.md) | [日本語](README.ja.md)

A fast, local-first Excel viewer. Open a workbook and see what Excel shows —
without waiting for Excel.

Built with Tauri 2, Rust ([calamine](https://github.com/tafia/calamine) for cell
values) and React 19 + Tailwind. The grid is its own renderer rather than a
spreadsheet library: it reads the workbook's XML directly for layout — column
widths, row heights, merges, borders, fills, fonts, the print area, and the
shapes and pictures floating above the cells — and paints Excel's rendering onto
a canvas.

`xlsx`, `xlsm`, `xls`, `xlsb` and `ods` all open; the older formats have no zip to
read, so they fall back to plain defaults.

## What it does

- Several workbooks open at once, each a tab that drags to reorder and closes on
  a middle click or from its own menu; a File menu reopens one from the list of
  those opened lately, and a start page pins the workbooks and folders you keep
  coming back to
- A folder tree of your own, beside the grid or on the start page: open folders
  a level at a time, find a workbook by name across every root, open one with a
  click
- The workbook's sheet list: hidden and pinned sheets, tab colours, hide and
  unhide, and a sheet tab strip with a "…" at either end
- Workbook search, go-to, column filters and a formula bar
- Cell, row/column and rectangular-block selection, by dragging or with Shift
  and the arrows; Ctrl+A takes the region around the cursor, then the whole sheet
- Ctrl+C in Excel's own format: tabs between columns, newlines between rows
- A ribbon for formatting — font, size, bold, italic, underline, alignment,
  fill and font colour, freeze and split — for the session only, never written
  back to the file
- Freeze panes and Excel's split view, and each sheet opens where the file left it
- Conditional formatting, notes, validation lists and outline groups drawn as
  Excel draws them
- Ctrl+wheel zoom, on Excel's zoom steps, with its bottom-right control
- Shapes and pictures above the cells: callouts, connectors, braces, gradients —
  drawn from ECMA-376's preset geometry table rather than reverse-engineered
- The print area as page-break preview shows it: shading outside it, and the
  heavy blue rule on its boundary
- Merged regions, text spilling into empty neighbours, indentation, and borders
  resolved the way Excel resolves a shared edge
- English, Chinese and Japanese — the system's language at first, switchable in
  the app
- A portable Windows build that updates itself from signed releases

## Running it

```bash
pnpm install
pnpm tauri dev                   # development, with hot reload
pnpm tauri build --no-bundle     # portable exe in src-tauri/target/release
```

## Tests

```bash
cd src-tauri && cargo test
```

The suite runs against `samples/sample-design.xlsx`, a synthetic workbook built
to exercise the parser: four sheets, a merged title, explicit column widths, an
autofilter, borders, fills, alignment and indent, and a 12,000-row sheet for
paging.

Seven ignored tests read a workbook from your own machine instead
(`SHEETLENS_XLSX`), which is the only way to check parsing against documents
that are real — the profile switch, a long sheet's row styles, rich text,
drawings, conditional formatting, and two dumps: a browser fixture and the
resolved cell borders.

```bash
SHEETLENS_XLSX="/path/to/book.xlsx" SHEETLENS_SHEET="Sheet1" \
  cargo test --release profile_switch -- --ignored --nocapture
```

## How Excel fidelity is checked

Against Excel itself, on real documents:

- `Shape.CopyPicture` puts Excel's own rendering on the clipboard — a bitmap, or
  SVG with real path commands when shapes are selected. The fastest way to see
  what a callout or a brace *should* look like.
- `ExportAsFixedFormat` writes the print area to a vector PDF.
- The ignored dump tests print every sheet's shapes, anchors and colours, or the
  resolved width and position of a shape, so a difference can be read as numbers
  instead of guessed at.
- `Ctrl+Shift+D` in a development build shows the same resolved numbers on
  screen: the sheet's font and its measured digit width, the default column and
  row, the selected cell's resolved style and borders, and every shape's box.

Most of what makes this hard is that Excel's layout units are its own. A column
width is a multiple of the workbook font's maximum digit width — 8px for
游ゴシック, 7px for Calibri — and nothing in the file records which; row heights
are points; and the cell's padding is derived from that same digit.

## Not implemented

Bookmarks and comparing two workbooks. Threaded comments and legacy VML form
controls are not read — the older notes Excel keeps in VML are. A preset the
table does not name, or one whose path uses a command the renderer does not
handle, falls back to a rectangle. WMF and EMF pictures cannot be drawn at all.
