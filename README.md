# SheetLens

**English** | [简体中文](README.zh-CN.md) | [日本語](README.ja.md)

A fast, local-first Excel viewer. Open a workbook and see what Excel shows —
without waiting for Excel.

Built with Tauri 2, Rust ([calamine](https://github.com/tafia/calamine) for cell
values) and React 19 + Tailwind. The grid is its own renderer rather than a
spreadsheet library: it reads the workbook's XML directly for layout — column
widths, row heights, merges, borders, fills, fonts, the print area, and the
shapes and pictures floating above the cells — and reproduces Excel's rendering
in CSS.

It was built against Japanese design documents (画面定義書 / 機能定義書), so CJK
fonts, phonetic runs and the layout conventions those documents use are the case
it is tuned for. `xlsx`, `xlsm`, `xls`, `xlsb` and `ods` all open; the older
formats have no zip to read, so they fall back to plain defaults.

## What it does

- Sheets, search, go-to, column filters and a formula bar
- Cell, row/column and rectangular-block selection, with Ctrl+C in Excel's own
  format: tabs between columns, newlines between rows
- Ctrl+wheel zoom, on Excel's zoom steps, with its bottom-right control
- Shapes and pictures above the cells: callouts, connectors, braces, gradients —
  drawn from ECMA-376's preset geometry table rather than reverse-engineered
- The print area as page-break preview shows it: shading outside it, and the
  heavy blue rule on its boundary
- Merged regions, text spilling into empty neighbours, indentation, and borders
  resolved the way Excel resolves a shared edge

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

Two ignored tests read a workbook from your own machine instead, which is the
only way to check parsing against documents that are real:

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

Regex search, frozen panes, split view, bookmarks, recents, compare; conditional
formatting, and legacy VML (comments, form controls). Among shapes: presets other
than `rect`, `ellipse`, `pentagon`, `line`, the braces and `borderCallout2` fall
back to a rectangle, radial gradients are drawn as linear ones, `pattFill` as no
fill, and a callout's tail is a straight three-point polyline. WMF and EMF
pictures cannot be drawn at all.
