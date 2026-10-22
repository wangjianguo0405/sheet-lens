//! Column widths, row heights, merges, autofilter and per-cell styling.
//!
//! calamine parses cell values only — its `Metadata` carries sheet names and
//! types, and `Range::width()` counts cells rather than pixels. All of this
//! lives in the workbook's XML, so we unzip it and read it directly.

use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::{BufRead, BufReader, Read, Seek};

use quick_xml::events::{BytesRef, BytesStart, Event};
use quick_xml::Reader as XmlReader;
use quick_xml::XmlVersion;
use serde::Serialize;
use zip::read::ZipArchive;

use crate::drawing;

/// Excel's stock column width when a sheet declares nothing.
const STOCK_COL_WIDTH_CHARS: f64 = 8.43;
const STOCK_ROW_HEIGHT_PT: f64 = 15.0;
/// Excel's "automatic" colour, and what most plain borders really mean.
pub(crate) const BLACK: &str = "#000000";
const WHITE: &str = "#FFFFFF";

/// The legacy indexed palette (ECMA-376 §18.8.27) that `indexed="N"` refers to.
/// Most workbooks leave it alone; the 23 in the corpus that override it write
/// `<colors><indexedColors>` at the end of styles.xml.
const INDEXED: [&str; 64] = [
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "000000",
    "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "800000", "008000",
    "000080", "808000", "800080", "008080", "C0C0C0", "808080", "9999FF", "993366", "FFFFCC",
    "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF", "000080", "FF00FF", "FFFF00", "00FFFF",
    "800080", "800000", "008080", "0000FF", "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF",
    "FF99CC", "CC99FF", "FFCC99", "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600",
    "666699", "969696", "003366", "339966", "003300", "333300", "993300", "993366", "333399",
    "333333",
];

/// The two colour tables a `<color>` can name: the theme's twelve slots and the
/// workbook's indexed palette. Both travel together wherever a colour is read.
#[derive(Clone, Copy)]
pub(crate) struct Palette<'a> {
    pub theme: &'a [Option<String>],
    pub indexed: &'a [Option<String>],
}

impl Palette<'_> {
    /// For callers that resolve no colours, and for tests.
    #[cfg(test)]
    pub(crate) const EMPTY: Palette<'static> = Palette {
        theme: &[],
        indexed: &[],
    };
}

/// The workbook's indexed palette, merged over the stock table.
///
/// Excel writes `<indexedColors>` after the fonts, fills and borders that refer
/// to it, so the walk cannot discover it in time; it is read once before the
/// walk instead. A file with 64 entries redefines every index, one with 56
/// starts at 8 (the first eight are fixed).
fn indexed_palette(xml: &str) -> Vec<Option<String>> {
    let mut palette: Vec<Option<String>> = INDEXED
        .iter()
        .map(|hex| Some(format!("#{hex}")))
        .collect();
    let Some(start) = xml.find("<indexedColors") else {
        return palette;
    };
    let Some(open_end) = xml[start..].find('>').map(|at| start + at + 1) else {
        return palette;
    };
    let Some(end) = xml[open_end..]
        .find("</indexedColors>")
        .map(|at| open_end + at)
    else {
        return palette;
    };
    let mut colors = Vec::new();
    let mut rest = &xml[open_end..end];
    while let Some(at) = rest.find("rgb=\"") {
        rest = &rest[at + 5..];
        let Some(close) = rest.find('"') else { break };
        if let Some(color) = parse_color(&rest[..close]) {
            colors.push(color);
        }
        rest = &rest[close + 1..];
    }
    let offset = if colors.len() >= 64 { 0 } else { 8 };
    for (index, color) in colors.into_iter().enumerate() {
        if let Some(slot) = palette.get_mut(offset + index) {
            *slot = Some(color);
        }
    }
    palette
}

/// The theme's twelve colours, in the order `theme="N"` indexes them.
///
/// That order is *not* the order `theme1.xml` lists them in — the file puts
/// dk1 first while the index starts at lt1.
fn theme_slots(scheme: &HashMap<String, String>) -> Vec<Option<String>> {
    const ORDER: [&str; 12] = [
        "lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5",
        "accent6", "hlink", "folHlink",
    ];
    ORDER
        .iter()
        .map(|name| scheme.get(*name).cloned())
        .collect()
}
/// Excel's stock font. Cells inherit this size, so the grid renders text at the
/// same scale as the workbook rather than picking its own.
const STOCK_FONT_PT: f64 = 11.0;

/// Widest range we are willing to expand into per-column entries.
const MAX_EXPANDED_COLUMNS: u32 = 4096;
/// Guards against a pathological file turning into unbounded memory.
const MAX_MERGES: usize = 20_000;
/// Excel's own sheet limits, so a bogus dimension cannot blow up the grid.
const MAX_SHEET_ROWS: u32 = 1_048_576;
const MAX_SHEET_COLS: u32 = 16_384;

/// Column widths in pixels, for the one place that cannot wait for a proper
/// answer.
///
/// A column's stored width is in units of the *workbook font's* maximum digit
/// width, and that metric is in no part of the file — Excel itself reads it off
/// the installed font. The grid measures it in the renderer and does this
/// conversion there; `drawing.rs` needs pixels while still inside the parser,
/// so it borrows Calibri's 7. Everything anchored to two cells is described as
/// a cell plus an offset and never comes through here, so the approximation
/// only reaches the handful of shapes anchored to a single cell.
fn chars_to_px(chars: f64) -> f64 {
    (chars * 7.0 + 5.0).round()
}

/// Row heights are in points; CSS pixels are 96/72 of a point.
fn points_to_px(points: f64) -> f64 {
    (points * 4.0 / 3.0).round()
}

/// Font sizes keep their fraction: 11pt is 14.67px, and rounding to 15 would
/// visibly change how much text a column fits.
fn points_to_font_px(points: f64) -> f64 {
    (points * 4.0 / 3.0 * 100.0).round() / 100.0
}

/// Horizontal/vertical placement copied from the cell's style. `None` means the
/// style did not say, leaving the grid to pick a sensible default.
#[derive(Serialize, Clone, Copy, Default, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct Align {
    pub horizontal: Option<&'static str>,
    pub vertical: Option<&'static str>,
    /// Indent levels. One level is three spaces of the cell's font, and Excel
    /// spends it on the side the text is pinned to — the right for a
    /// right-aligned cell, the left otherwise. Leaving this out put every
    /// right-aligned, indented number a whole unit too close to the border.
    pub indent: u8,
    pub wrap: bool,
    /// `<alignment shrinkToFit>`: Excel scales the font down until the cell's
    /// one line fits. The standard makes it inapplicable to a cell holding
    /// several lines, so a wrapping cell ignores it.
    pub shrink: bool,
    /// `<alignment textRotation>`, as degrees turned counter-clockwise — a
    /// negative angle turns the other way. Excel stores 0–90 for one direction
    /// and 91–180 for the angles its own dialog calls −1 to −90, so 180 is the
    /// "-90" a reader arrives at there.
    pub rotation: i16,
    /// `textRotation="255"`: the characters one below the next, each upright,
    /// rather than the line turned as a whole. It is not an angle, which is why
    /// it has a field of its own.
    pub stacked: bool,
}

/// One side of a cell's border, already reduced to what CSS needs.
#[derive(Serialize, Clone, PartialEq, Eq, Hash, Debug)]
#[serde(rename_all = "camelCase")]
pub struct BorderEdge {
    /// A CSS `border-style`.
    pub line: &'static str,
    pub width: u8,
    /// Excel's border precedence. Two cells share an edge and only one line is
    /// drawn, so the heavier of the two wins.
    pub rank: u8,
    pub color: Option<String>,
}

#[derive(Serialize, Clone, Default, PartialEq, Eq, Hash, Debug)]
#[serde(rename_all = "camelCase")]
pub struct CellBorder {
    pub top: Option<BorderEdge>,
    pub right: Option<BorderEdge>,
    pub bottom: Option<BorderEdge>,
    pub left: Option<BorderEdge>,
    /// Excel's one diagonal rule. `<diagonal>` carries its style; the two flags
    /// on `<border>` say which corners it joins, and both may be set — that is
    /// the crossing pair. Measured off `CopyPicture`: on a merged region the
    /// line runs across the whole region, not just the anchor cell.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub diagonal: Option<BorderEdge>,
    #[serde(skip_serializing_if = "is_false")]
    pub diagonal_up: bool,
    #[serde(skip_serializing_if = "is_false")]
    pub diagonal_down: bool,
}

/// One of Excel's 18 hatch patterns, as `<patternFill patternType>` spells it.
///
/// The masks themselves are in the renderer, measured off Excel's own picture:
/// an 8×8 tile of the pattern colour over the background colour, anchored at
/// the sheet's own origin so the lines run unbroken across cell edges.
#[derive(Serialize, Clone, PartialEq, Eq, Hash, Debug)]
#[serde(rename_all = "camelCase")]
pub struct CellPattern {
    /// The `<patternFill patternType>` name, e.g. `lightUp` or `gray0625`.
    pub kind: &'static str,
    /// The pattern's own colour — `<fgColor>`, Excel's `PatternColor`.
    pub fg: String,
    /// The ground under the pattern — `<bgColor>`, Excel's `Interior.Color`.
    pub bg: String,
}

/// One `<fill>` entry in the styles table.
#[derive(Clone)]
enum FillSpec {
    Solid(String),
    Pattern(CellPattern),
}

#[derive(Serialize, Clone, Default, Debug)]
#[serde(rename_all = "camelCase")]
pub struct CellStyle {
    pub horizontal: Option<&'static str>,
    pub vertical: Option<&'static str>,
    /// Indent levels; the renderer turns each into three spaces of the font.
    pub indent: u8,
    pub wrap: bool,
    pub shrink: bool,
    /// Degrees turned counter-clockwise, from `<alignment textRotation>`; left
    /// out entirely for the cell that draws flat, which is nearly all of them.
    #[serde(skip_serializing_if = "is_zero")]
    pub rotation: i16,
    /// `textRotation="255"`: characters one below the next, each upright.
    #[serde(skip_serializing_if = "is_false")]
    pub stacked: bool,
    pub bold: bool,
    pub italic: bool,
    pub strike: bool,
    /// `single` or `double`; absent when the font has no underline.
    pub underline: Option<&'static str>,
    /// `superscript` or `subscript` for a font that draws raised or lowered at
    /// two thirds of its size. Excel keeps the stored size; the scale is a
    /// drawing rule, measured off `CopyPicture`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vert_align: Option<&'static str>,
    /// Only set when the cell's font differs from the workbook default, which
    /// is the difference between a 22pt title and body text.
    pub size_px: Option<f64>,
    /// Likewise only set when it differs from the workbook's own family. It
    /// matters twice over: Excel draws the cell in this font, and measures an
    /// `indent` in this font's spaces — ＭＳ ゴシック's space is a full half an
    /// em where 游ゴシック's is under a third, so the two indent differently.
    pub family: Option<String>,
    pub color: Option<String>,
    /// The cell's background, from a solid fill.
    pub fill: Option<String>,
    /// A hatch pattern in place of a flat background. A cell has one or the
    /// other, never both.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pattern: Option<CellPattern>,
    /// Its number format, when one applies. Dates are stored as serial numbers
    /// and only this says how to show them.
    pub format: Option<String>,
    pub border: Option<CellBorder>,
}

/// Serde leaves a field out when it still holds the value a flat cell has.
fn is_zero(value: &i16) -> bool {
    *value == 0
}

fn is_false(value: &bool) -> bool {
    !*value
}

// Hand-written so styles can be deduplicated in a map. `size_px` is a finite
// measurement, so claiming `Eq` is sound here even though `f64` is only
// `PartialEq`.
impl PartialEq for CellStyle {
    fn eq(&self, other: &Self) -> bool {
        self.horizontal == other.horizontal
            && self.vertical == other.vertical
            && self.indent == other.indent
            && self.wrap == other.wrap
            && self.shrink == other.shrink
            && self.rotation == other.rotation
            && self.stacked == other.stacked
            && self.bold == other.bold
            && self.italic == other.italic
            && self.strike == other.strike
            && self.underline == other.underline
            && self.vert_align == other.vert_align
            && self.size_px == other.size_px
            && self.family == other.family
            && self.color == other.color
            && self.fill == other.fill
            && self.pattern == other.pattern
            && self.border == other.border
    }
}

impl Eq for CellStyle {}

impl std::hash::Hash for CellStyle {
    fn hash<H: std::hash::Hasher>(&self, state: &mut H) {
        self.horizontal.hash(state);
        self.vertical.hash(state);
        self.indent.hash(state);
        self.wrap.hash(state);
        self.shrink.hash(state);
        self.rotation.hash(state);
        self.stacked.hash(state);
        self.bold.hash(state);
        self.italic.hash(state);
        self.strike.hash(state);
        self.underline.hash(state);
        self.vert_align.hash(state);
        self.size_px.map(f64::to_bits).hash(state);
        self.family.hash(state);
        self.color.hash(state);
        self.fill.hash(state);
        self.pattern.hash(state);
        self.border.hash(state);
    }
}

impl CellStyle {
    pub(crate) fn is_empty(&self) -> bool {
        self.horizontal.is_none()
            && self.vertical.is_none()
            && self.indent == 0
            && !self.wrap
            && !self.shrink
            && self.rotation == 0
            && !self.stacked
            && !self.bold
            && !self.italic
            && !self.strike
            && self.underline.is_none()
            && self.vert_align.is_none()
            && self.size_px.is_none()
            && self.family.is_none()
            && self.color.is_none()
            && self.fill.is_none()
            && self.pattern.is_none()
            && self.format.is_none()
            && self.border.is_none()
    }
}

/// A merged block, anchored at its top-left cell.
#[derive(Serialize, Clone, Debug)]
#[serde(rename_all = "camelCase")]
pub struct MergeRect {
    pub row: u32,
    pub col: u32,
    pub rows: u32,
    pub cols: u32,
}

/// One data-validation rule and the cells it governs.
#[derive(Serialize, Clone, Debug, Default)]
#[serde(rename_all = "camelCase")]
pub struct DataValidation {
    pub ranges: Vec<MergeRect>,
    /// `list`, `whole`, `decimal`, `date`, `time`, `textLength`, `custom` or
    /// `none`.
    pub kind: String,
    /// The rule's own formula, as written: a quoted list such as `"a,b"`, a
    /// range, or an expression.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formula1: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formula2: Option<String>,
    pub allow_blank: bool,
    /// The file's `showDropDown` — **inverted**, as Excel writes it: setting it
    /// hides the in-cell dropdown (`InCellDropdown = False` writes `"1"`).
    pub hide_dropdown: bool,
    pub show_input: bool,
    pub show_error: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub prompt_title: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub prompt: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_title: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error: Option<String>,
}

/// The sheet's autofilter: `row` holds the filter buttons, `rows` counts the
/// whole range including that header row.
#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct AutoFilter {
    pub row: u32,
    pub col: u32,
    pub rows: u32,
    pub cols: u32,
    /// Absolute column index -> values the file already filters on. `colId` in
    /// the XML is relative to the filter range, so it is rebased here.
    pub applied: HashMap<u32, Vec<String>>,
    /// Columns filtered by a condition (`<customFilters>`, e.g. "> 20") rather
    /// than by a value list. Only the button's funnel depends on this: the rows
    /// such a filter hides are already the file's `hidden` rows.
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub custom: Vec<u32>,
    /// Columns whose drop-down the file hides (`showButton="0"`), absolute.
    /// Excel writes one for every column a merged header covers, so that the
    /// region shows a single button at its right edge rather than one per
    /// column.
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub hidden_buttons: Vec<u32>,
}

impl SheetLayout {
    /// What Excel prints when the sheet defines no `_xlnm.Print_Area`, and so
    /// also what it shades around.
    ///
    /// Excel's used range — what `Ctrl+End` jumps to — covers every cell the
    /// sheet *formats*, not only the ones holding a value. Measured in Excel on
    /// a cover sheet whose values stop at row 36: `Ctrl+End` lands on `AW44`,
    /// the shading and the page's edge both run to row 44, and the earlier
    /// reading that only value-bearing cells counted put that edge four rows
    /// too high. Merged ranges already extend the extent, which is what holds a
    /// merged banner open to the sheet's last column.
    fn print_region(&self) -> Option<MergeRect> {
        (self.rows > 0 && self.cols > 0).then_some(MergeRect {
            row: 0,
            col: 0,
            rows: self.rows,
            cols: self.cols,
        })
    }

    /// The style index for a cell, following Excel's cell → row → column rule.
    pub fn style_index(&self, row: u32, col: u32) -> Option<u32> {
        if let Some(found) = self
            .cell_runs
            .get(&row)
            .and_then(|runs| runs.iter().find(|run| col >= run[0] && col <= run[1]))
        {
            return Some(found[2]);
        }
        if let Some(found) = self.row_styles.get(&row) {
            return Some(*found);
        }
        self.col_styles
            .iter()
            .find(|run| col >= run.from && col <= run.to)
            .map(|run| run.style)
    }

    /// The cell's number format, when one applies.
    pub fn format_at(&self, row: u32, col: u32) -> Option<&str> {
        let index = self.style_index(row, col)?;
        self.styles.get(index as usize)?.format.as_deref()
    }
}

/// A run of columns sharing a default style.
#[derive(Serialize, Clone, Debug)]
#[serde(rename_all = "camelCase")]
pub struct ColumnStyle {
    pub from: u32,
    pub to: u32,
    pub style: u32,
}

impl Styles {
    /// The viewer's style for one `cellXfs` entry, or `None` when the entry
    /// says nothing this grid would draw differently.
    fn cell_style(&self, index: usize, default_font_px: f64) -> Option<CellStyle> {
        let entry = self.resolved(self.xfs.get(index)?);
        let font = self.fonts.get(entry.font_id);
        let size_px = font
            .and_then(|font| font.size_pt)
            .map(points_to_font_px)
            .filter(|px| (px - default_font_px).abs() > 0.01);
        let default_family = self.fonts.first().and_then(|font| font.name.as_deref());
        let family = font
            .and_then(|font| font.name.clone())
            .filter(|name| Some(name.as_str()) != default_family);
        let (fill, pattern) = match self.fills.get(entry.fill_id).cloned().flatten() {
            Some(FillSpec::Solid(color)) => (Some(color), None),
            Some(FillSpec::Pattern(pattern)) => (None, Some(pattern)),
            None => (None, None),
        };
        let style = CellStyle {
            horizontal: entry.align.horizontal,
            vertical: entry.align.vertical,
            indent: entry.align.indent,
            wrap: entry.align.wrap,
            shrink: entry.align.shrink,
            rotation: entry.align.rotation,
            stacked: entry.align.stacked,
            bold: font.is_some_and(|font| font.bold),
            italic: font.is_some_and(|font| font.italic),
            strike: font.is_some_and(|font| font.strike),
            underline: font.and_then(|font| font.underline),
            vert_align: font.and_then(|font| font.vert_align),
            size_px,
            family,
            color: font.and_then(|font| font.color.clone()),
            fill,
            pattern,
            format: self
                .num_formats
                .get(&entry.num_fmt_id)
                .cloned()
                .or_else(|| builtin_date_format(entry.num_fmt_id).map(str::to_owned)),
            border: self.borders.get(entry.border_id).and_then(|b| b.clone()),
        };
        (!style.is_empty()).then_some(style)
    }

    /// A cell format with the properties it does not apply taken from the named
    /// style it points at.
    ///
    /// Excel's inheritance is one level deeper than cell → row → column: every
    /// `cellXfs` entry carries an `xfId` into `cellStyleXfs`, plus an `apply…`
    /// flag per property saying whether its own value is the one to use. With
    /// the flag off, the ids written beside it are leftovers from whatever the
    /// format was built on, and reading those draws the cell in the wrong font.
    ///
    /// Only an explicit `apply…="0"` inherits — see `Xf` for what a missing
    /// flag means.
    fn resolved(&self, entry: &Xf) -> Xf {
        let Some(named) = self.style_xfs.get(entry.style_id) else {
            return entry.clone();
        };
        let pick = |apply: Option<bool>, own: usize, named: usize| match apply {
            Some(false) => named,
            _ => own,
        };
        Xf {
            align: match entry.apply_align {
                Some(false) => named.align,
                _ => entry.align,
            },
            border_id: pick(entry.apply_border, entry.border_id, named.border_id),
            font_id: pick(entry.apply_font, entry.font_id, named.font_id),
            fill_id: pick(entry.apply_fill, entry.fill_id, named.fill_id),
            num_fmt_id: match entry.apply_number {
                Some(false) => named.num_fmt_id,
                _ => entry.num_fmt_id,
            },
            ..entry.clone()
        }
    }
}

/// Stores `style` once and returns its index, so the payload carries each
/// distinct style a single time.
pub(crate) fn intern(
    styles: &mut Vec<CellStyle>,
    ids: &mut HashMap<CellStyle, u32>,
    style: CellStyle,
) -> u32 {
    if let Some(found) = ids.get(&style) {
        return *found;
    }
    styles.push(style.clone());
    let index = styles.len() as u32 - 1;
    ids.insert(style, index);
    index
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct SheetLayout {
    /// The sheet's real extent. calamine's range only spans cells that hold a
    /// value, so a column kept alive purely by styling or a merge — which is
    /// how tables close off their right edge — vanishes from it.
    pub rows: u32,
    pub cols: u32,
    /// Columns arrive in the file's own unit — a multiple of the workbook
    /// font's maximum digit width — and stay in it: only the renderer, which
    /// can measure that font, knows how many pixels one of them is.
    pub default_col_width_chars: f64,
    pub default_row_height_px: f64,
    /// The workbook's default font. Rendered at the file's own size so the grid
    /// matches Excel's scale instead of imposing one — and measured for the
    /// digit width that turns column widths into pixels.
    pub font_family: Option<String>,
    pub font_size_px: f64,
    /// Excel's View > Gridlines setting. Most design documents turn it off, and
    /// then only the borders the author drew are visible.
    pub show_grid_lines: bool,
    /// The zoom the sheet is left at, in percent (`<sheetView zoomScale>`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub zoom_scale: Option<u16>,
    /// The first cell of the scrolling view, as `[row, col]` — Excel opens the
    /// sheet scrolled here rather than at A1.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub view_top_left: Option<[u32; 2]>,
    /// How many rows and columns the sheet view freezes at the top and left.
    /// Excel's `state="frozen"` pane; a *split* pane leaves nothing pinned in
    /// this viewer, since it repeats one region rather than holding rows still.
    pub frozen_rows: u32,
    pub frozen_cols: u32,
    /// Column index -> width in characters, for columns with an explicit width.
    pub col_widths_chars: HashMap<u32, f64>,
    /// Row index -> height in pixels, for rows with an explicit height.
    pub row_heights_px: HashMap<u32, f64>,
    pub merges: Vec<MergeRect>,
    /// Data-validation rules: what a cell accepts, and the prompt or dropdown
    /// Excel offers for it.
    pub validations: Vec<DataValidation>,
    /// The workbook's defined names as this sheet sees them, name -> formula.
    /// A list validation may name one instead of writing out a range.
    #[serde(skip_serializing_if = "HashMap::is_empty")]
    pub names: HashMap<String, String>,
    /// Rows the file hides. Excel draws none of them and skips their numbers —
    /// there is no marker at the boundary, so the gap is the only sign. A
    /// hidden row keeps whatever height it had, which is what unhiding puts
    /// back; one that never had a height comes back at the sheet's default.
    pub hidden_rows: Vec<u32>,
    /// Columns the file hides, which it writes as `width="0"` as well.
    pub hidden_cols: Vec<u32>,
    pub auto_filter: Option<AutoFilter>,
    /// `_xlnm.Print_Area`, or the used range where the sheet defines none —
    /// Excel prints the used range in that case and shades around it just the
    /// same. The shade replaces each cell's own background rather than sitting
    /// under it.
    pub print_area: Option<MergeRect>,
    /// The distinct cell styles in use. Cells index into this rather than each
    /// carrying a copy — repeating one style per cell made the payload 4.6 MB
    /// on a sheet with 18k styled cells.
    pub styles: Vec<CellStyle>,
    /// Per row, runs of consecutive columns sharing one style, as
    /// `[fromCol, toCol, styleIndex]`. A flat per-cell map cost one entry per
    /// cell — 110k of them on one sheet — while runs collapse a filled row
    /// into a single entry.
    pub cell_runs: HashMap<u32, Vec<[u32; 3]>>,
    /// Rows that declare a default style, applied to the cells in them that
    /// declare none of their own. Excel inherits cell -> row -> column, and
    /// reading only the cell level left holes in filled rows.
    pub row_styles: HashMap<u32, u32>,
    /// Column runs that declare a default style, same rule.
    pub col_styles: Vec<ColumnStyle>,
    /// The shapes and pictures the sheet floats above its cells, in the order
    /// Excel draws them — later ones sit on top.
    pub drawings: Vec<crate::drawing::Shape>,
    /// The `r:id` of the sheet's drawing part, resolved while parsing and then
    /// used to read it in. Not part of what the grid sees.
    #[serde(skip)]
    pub drawing_rel: Option<String>,
    /// The `r:id` of the sheet's legacy VML part, the same way — its notes are
    /// folded into `notes` before this is cleared.
    #[serde(skip)]
    pub legacy_rel: Option<String>,
    /// Cells carrying a note. Excel marks each with a small red triangle in the
    /// cell's top-right corner, which is all a note draws until it is hovered.
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub notes: Vec<NoteCell>,
    /// Cells whose text is written as several runs, each with its own colour or
    /// font. Per cell, not per string: a file may spell the same text twice with
    /// different formatting — a breadcrumb whose current page is red — and the
    /// text alone cannot tell the two apart. Sparse by nature.
    pub rich_cells: Vec<RichCell>,
    /// Cells a data-bar rule draws a bar in, with the share of the cell's width
    /// the bar takes. Sparse by nature.
    pub data_bars: Vec<DataBarCell>,
    /// Cells an icon-set rule marks, with the family and the level drawn.
    pub icons: Vec<IconCell>,
    /// Excel's outline (Data > Group). Absent when the sheet groups nothing —
    /// which is every workbook in the corpus — so the payload of those files is
    /// untouched.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub outline: Option<Outline>,
    /// How the sheet is laid out for printing. Page Break Preview draws its
    /// breaks and page numbers from this, and the page layout view lays the
    /// sheet out on these pages.
    pub page_setup: PageSetup,
}

/// The paper a sheet prints on, and where Excel breaks it into pages.
///
/// Everything here is the file's own answer. Where the file says nothing the
/// ECMA-376 defaults apply — Letter paper, portrait, 100%, the standard
/// margins — because that is what a reader that never met this printer has to
/// assume. Excel writes `pageMargins` into every sheet and `pageSetup` into
/// almost all of them, so the defaults rarely decide anything.
#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct PageSetup {
    /// `ST_PaperSize`; 9 is A4, 1 is Letter.
    pub paper_size: u16,
    pub landscape: bool,
    /// Percent, 10–400. Ignored when `fit_to_page` is set.
    pub scale: u16,
    /// `<pageSetUpPr fitToPage="1"/>`: the sheet is scaled to fit the given
    /// number of pages rather than printed at a fixed percentage.
    pub fit_to_page: bool,
    /// How many pages wide and tall the fit may use. ECMA-376 defaults both to
    /// one, and Excel leaves the attributes out when they are one — so a sheet
    /// that only says `fitToPage="1"` means "one page", not "as many as it
    /// takes". An explicit `0` is what asks for automatic.
    pub fit_to_width: u16,
    pub fit_to_height: u16,
    /// `overThenDown`: pages number across before they number down.
    pub page_order_over_then_down: bool,
    pub margins: PageMargins,
    pub print_options: PrintOptions,
    /// Rows and columns a manual break falls after, zero-based. Excel stores
    /// only the breaks somebody placed; automatic ones are worked out from the
    /// paper and are not in the file at all.
    pub row_breaks: Vec<u32>,
    pub col_breaks: Vec<u32>,
    /// `_xlnm.Print_Titles`: the rows and columns repeated on every page.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub print_titles: Option<PrintTitles>,
    /// Dots per inch the sheet is laid out for, from the printer settings the
    /// file was saved with (`printerSettings*.bin`). `None` when the workbook
    /// carries none, or when the stored value is a resolution code
    /// (`DMRES_HIGH` and friends) rather than a DPI.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub print_dpi: Option<u16>,
}

impl Default for PageSetup {
    fn default() -> Self {
        Self {
            paper_size: 1,
            landscape: false,
            scale: 100,
            fit_to_page: false,
            fit_to_width: 1,
            fit_to_height: 1,
            page_order_over_then_down: false,
            margins: PageMargins::default(),
            print_options: PrintOptions::default(),
            row_breaks: Vec::new(),
            col_breaks: Vec::new(),
            print_titles: None,
            print_dpi: None,
        }
    }
}

/// Page margins in inches, as `<pageMargins>` writes them.
#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct PageMargins {
    pub left: f64,
    pub right: f64,
    pub top: f64,
    pub bottom: f64,
    pub header: f64,
    pub footer: f64,
}

impl Default for PageMargins {
    fn default() -> Self {
        Self {
            left: 0.75,
            right: 0.75,
            top: 1.0,
            bottom: 1.0,
            header: 0.5,
            footer: 0.5,
        }
    }
}

/// What a page carries besides the cells themselves.
#[derive(Serialize, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct PrintOptions {
    pub horizontal_centered: bool,
    pub vertical_centered: bool,
    pub headings: bool,
    pub grid_lines: bool,
}

/// Rows and columns repeated on every page, as zero-based inclusive bounds.
#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct PrintTitles {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rows: Option<[u32; 2]>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cols: Option<[u32; 2]>,
}

/// The sheet's outline groups, as the file records them.
///
/// A group is not written as a range: `<row outlineLevel="1">` marks each
/// detail row, and the row beside the group carries `collapsed="1"` when the
/// group is closed. Detail rows are `hidden="1"` while collapsed, but the
/// levels are what say where a group begins and ends, so they are kept as they
/// are and the renderer builds the bands.
#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct Outline {
    pub rows: Vec<OutlineRow>,
    pub cols: Vec<OutlineCol>,
    /// Excel's default is a summary below and to the right of the detail; the
    /// other side is what `<outlinePr summaryBelow="0">` asks for.
    pub summary_below: bool,
    pub summary_right: bool,
}

/// One row's outline facts. Rows at level 0 with nothing else to say are not
/// listed at all.
#[derive(Serialize, Clone)]
pub struct OutlineRow {
    pub row: u32,
    /// 1-based; 0 is outside any group.
    pub level: u8,
    /// Set on the summary row of a closed group.
    pub collapsed: bool,
}

/// One `<col>` run's outline facts, in the file's own `min`/`max` terms.
#[derive(Serialize, Clone)]
pub struct OutlineCol {
    pub from: u32,
    pub to: u32,
    pub level: u8,
    pub collapsed: bool,
}

/// One cell's note: where Excel marks it, and where hovering shows it.
///
/// The wording comes from the sheet's `comments1.xml`; the VML only has the
/// shape. The box is anchored in cells, as `[col, dx, row, dy, col, dx, row, dy]`
/// — the offsets are pixels, which is what reproduces Excel's own `Left/Top/
/// Width/Height` from its anchor (15 px in, 96 pt wide for a 14-character
/// column, and so on).
#[derive(Serialize, Clone)]
pub struct NoteCell {
    pub row: u32,
    pub col: u32,
    #[serde(skip_serializing_if = "String::is_empty")]
    pub text: String,
    /// The first run's font, which is what the note is written in.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font: Option<NoteFont>,
    pub anchor: [f64; 8],
}

/// A note's font, from the `<rPr>` on the first run of its text.
#[derive(Serialize, Clone, PartialEq, Debug)]
#[serde(rename_all = "camelCase")]
pub struct NoteFont {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub family: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub size_px: Option<f64>,
    pub bold: bool,
    pub italic: bool,
}

/// One cell's icon: the family's name and which of its marks the value takes.
#[derive(Serialize, Clone)]
pub struct IconCell {
    pub row: u32,
    pub col: u32,
    pub family: String,
    pub level: u8,
}

/// One cell's data bar: where it runs across the cell, and what it is drawn in.
#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct DataBarCell {
    pub row: u32,
    pub col: u32,
    pub color: String,
    /// Where the bar starts and ends, as shares of the cell's width.
    pub from: f64,
    pub to: f64,
    /// Whether the bar lies left of the axis, which flips its gradient so the
    /// dark end stays on the axis.
    pub negative: bool,
    /// The axis line's share of the cell, when Excel draws one.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub axis: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub axis_color: Option<String>,
}

/// One cell's runs. The position is the cell's own, because two cells spelling
/// the same text may carry different runs.
#[derive(Serialize, Clone)]
pub struct RichCell {
    pub row: u32,
    pub col: u32,
    pub runs: Vec<RichRun>,
}

/// One run of a cell's rich text.
///
/// A cell shows the concatenation of its runs, so the text is nothing new; what
/// a run adds is formatting for part of it, which is how a document makes one
/// word inside a sentence red, or bold, or a size larger.
#[derive(Serialize, Clone)]
pub struct RichRun {
    pub text: String,
    /// `#rrggbb`, or `None` when the run takes the cell's own colour.
    pub color: Option<String>,
    /// What the run's own `<rPr>` says. `None` when the run has none, which
    /// means the whole of the cell's font — see `RichFont`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font: Option<RichFont>,
}

/// A run's `<rPr>`: the font the run draws in, where it differs from the cell's.
///
/// The two halves of this inherit differently, and Excel settles both from its
/// own files:
///
/// - `bold`/`italic`/`strike`/`underline` are what the run *is*, and an absent
///   one means off rather than the cell's. Writing a rich cell whose first three
///   characters are bold leaves the rest of the cell's font alone and gives the
///   *other* run an `<rPr>` naming the font without `<b/>`, and that run draws
///   plain even though the cell's font is bold. `rt.xlsx` row 1 in the fixture,
///   and its row 11 (a bold cell with the run turned back off) show the same
///   thing from both sides. `Some(false)` is an rPr that spells the off out —
///   which is what a run has to say to be plain inside a bold cell it shares
///   with a bold one — and is kept apart from `None` so that an rPr saying
///   nothing at all can still be dropped.
/// - `family`/`size_px` are overrides, and stay `None` when the rPr omits them.
///   Excel writes both on every rPr it produces, so no Excel file can tell the
///   difference; inheriting is what a hand-written `<rPr><i/></rPr>` means, and
///   the alternative — the workbook default — would draw it at the wrong size.
#[derive(Serialize, Clone, Default, PartialEq, Debug)]
#[serde(rename_all = "camelCase")]
pub struct RichFont {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub family: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub size_px: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub strike: Option<bool>,
    /// `single`, `double`, or `none` for an rPr that spells the absence out;
    /// `None` means the rPr did not mention it at all.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub underline: Option<&'static str>,
    /// `superscript` or `subscript` for this run alone.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vert_align: Option<&'static str>,
}

impl RichFont {
    /// True when the rPr says nothing the cell would not already draw, which is
    /// what lets a restated font be dropped rather than kept.
    fn is_empty(&self) -> bool {
        self.family.is_none()
            && self.size_px.is_none()
            && self.bold.is_none()
            && self.italic.is_none()
            && self.strike.is_none()
            && self.underline.is_none()
            && self.vert_align.is_none()
    }
}

impl SheetLayout {
    pub(crate) fn empty() -> Self {
        Self {
            rows: 0,
            cols: 0,
            default_col_width_chars: STOCK_COL_WIDTH_CHARS,
            default_row_height_px: points_to_px(STOCK_ROW_HEIGHT_PT),
            font_family: None,
            font_size_px: points_to_font_px(STOCK_FONT_PT),
            show_grid_lines: true,
            zoom_scale: None,
            view_top_left: None,
            frozen_rows: 0,
            frozen_cols: 0,
            col_widths_chars: HashMap::new(),
            row_heights_px: HashMap::new(),
            merges: Vec::new(),
            validations: Vec::new(),
            names: HashMap::new(),
            hidden_rows: Vec::new(),
            hidden_cols: Vec::new(),
            auto_filter: None,
            print_area: None,
            styles: Vec::new(),
            cell_runs: HashMap::new(),
            row_styles: HashMap::new(),
            col_styles: Vec::new(),
            drawings: Vec::new(),
            drawing_rel: None,
            legacy_rel: None,
            notes: Vec::new(),
            rich_cells: Vec::new(),
            data_bars: Vec::new(),
            icons: Vec::new(),
            outline: None,
            page_setup: PageSetup::default(),
        }
    }
}

/// One `<font>` entry from the styles table.
#[derive(Clone, Default)]
struct FontInfo {
    name: Option<String>,
    size_pt: Option<f64>,
    bold: bool,
    italic: bool,
    strike: bool,
    /// `single` or `double`. The two accounting forms draw the same line, only
    /// spanning the cell rather than the text, and land on one of these.
    underline: Option<&'static str>,
    /// `superscript` or `subscript`; the cell's whole text draws raised or
    /// lowered at two thirds of its size.
    vert_align: Option<&'static str>,
    color: Option<String>,
}

/// The parts of a `cellXfs` entry this viewer cares about.
#[derive(Clone, Default)]
struct Xf {
    align: Align,
    border_id: usize,
    font_id: usize,
    fill_id: usize,
    num_fmt_id: u32,
    /// `xfId`: the named style in `cellStyleXfs` this format inherits from.
    style_id: usize,
    /// One `apply…` flag per property, kept as `None` when the file writes no
    /// such attribute — which is not the same as writing `apply…="0"`. Only the
    /// zero gives the property up to the named style; an entry that says
    /// nothing is using the ids beside it, which is how every producer that is
    /// not Excel writes its formats.
    apply_font: Option<bool>,
    apply_fill: Option<bool>,
    apply_border: Option<bool>,
    apply_number: Option<bool>,
    apply_align: Option<bool>,
}

/// Everything this viewer needs out of `xl/styles.xml`.
#[derive(Default)]
struct Styles {
    xfs: Vec<Xf>,
    /// `cellStyleXfs`: the workbook's named styles, which a cell format borrows
    /// the properties it does not apply from.
    style_xfs: Vec<Xf>,
    borders: Vec<Option<CellBorder>>,
    fonts: Vec<FontInfo>,
    /// What each `<fill>` entry paints: a flat colour, or one of Excel's 18
    /// hatch patterns (6 of the 118 corpus files use a pattern on a cell).
    fills: Vec<Option<FillSpec>>,
    /// Custom `numFmtId` -> format code. Built-in ids are looked up separately.
    num_formats: HashMap<u32, String>,
    /// The workbook's `<dxfs>`: what each conditional-format rule paints.
    dxfs: Vec<crate::cond::Patch>,
    /// The workbook's own indexed palette, when it writes `<indexedColors>`.
    /// Empty means the stock table in `INDEXED` applies.
    indexed: Vec<Option<String>>,
}

/// Workbook-wide parse results. The styles and theme tables are shared by every
/// sheet, and re-reading them per switch cost tens — once over 100 — milliseconds.
#[derive(Default)]
pub struct Cache {
    loaded: bool,
    theme: Vec<Option<String>>,
    /// The theme's own line widths, which a shape's `lnRef idx` selects from.
    ln_widths: Vec<f64>,
    /// The faces the theme's font scheme names, which a shape's
    /// `a:latin typeface="+mn-ea"` and its like point at.
    theme_fonts: HashMap<String, String>,
    styles: Styles,
    /// Shared strings split into coloured runs, one slot per `<si>` in the
    /// order the sheet's cells index them. Nearly every string in every sheet is
    /// shared, so this is read once per workbook rather than once per sheet.
    rich_strings: Vec<Option<Vec<RichRun>>>,
    /// The same table in the order the sheet's cells index it. Only a formula
    /// being evaluated needs to look a string up by number — everything drawn
    /// carries its text already.
    shared_strings: Vec<String>,
}

impl Cache {
    /// What an `dxfId` in a conditional-format rule points at.
    pub(crate) fn patches(&self) -> &[crate::cond::Patch] {
        &self.styles.dxfs
    }

    pub(crate) fn shared_strings(&self) -> &[String] {
        &self.shared_strings
    }

    /// The two colour tables together, for anything resolving a `<color>`.
    pub(crate) fn palette(&self) -> Palette<'_> {
        Palette {
            theme: &self.theme,
            indexed: &self.styles.indexed,
        }
    }

    pub fn load(&mut self, path: &str) -> Result<(), String> {
        if self.loaded {
            return Ok(());
        }
        let file = File::open(path).map_err(|e| e.to_string())?;
        let mut zip = ZipArchive::new(BufReader::new(file)).map_err(|e| e.to_string())?;
        self.theme = theme_slots(&read_theme(&mut zip));
        self.ln_widths = read_line_widths(&mut zip);
        self.theme_fonts = read_theme_fonts(&mut zip);
        self.styles = read_styles(&mut zip, &self.theme);
        // Absent in a workbook that writes its strings inline, which is why
        // this is a lookup on failure rather than an error.
        if let Ok(xml) = read_entry(&mut zip, "xl/sharedStrings.xml") {
            let palette = Palette {
                theme: &self.theme,
                indexed: &self.styles.indexed,
            };
            self.rich_strings = rich_strings(&xml, palette);
            self.shared_strings = shared_strings(&xml);
        }
        self.loaded = true;
        Ok(())
    }
}

/// How much of a sheet's part is read when looking for its tab colour.
/// `<sheetPr>` is the first thing after the root element, and a sheet's XML
/// runs to megabytes.
const SHEET_HEAD_BYTES: u64 = 8192;

/// The sheets the workbook keeps out of sight: `<sheet state="hidden">`.
///
/// Excel hides the tab, and its own Unhide list is the only way back, so a
/// viewer that drew these would show a workbook that does not look like the one
/// Excel opens. `state="veryHidden"` is hidden the same way here — Excel keeps
/// those out of its Unhide list because a macro is usually what put them there,
/// and nothing in the corpus uses the value.
pub fn hidden_sheets(path: &str) -> HashSet<String> {
    let Ok(file) = File::open(path) else {
        return HashSet::new();
    };
    let Ok(mut zip) = ZipArchive::new(BufReader::new(file)) else {
        return HashSet::new();
    };
    let Ok(workbook) = read_entry(&mut zip, "xl/workbook.xml") else {
        return HashSet::new();
    };
    hidden_in_workbook(&workbook)
}

/// The names `workbook.xml` marks as hidden. Split out from the read so the rule
/// can be tested on the XML alone.
fn hidden_in_workbook(xml: &str) -> HashSet<String> {
    let mut hidden = HashSet::new();
    for element in scan(xml, b"sheet") {
        let (Some(name), Some(state)) = (element.get("name"), element.get("state")) else {
            continue;
        };
        if state != "visible" {
            hidden.insert(name.clone());
        }
    }
    hidden
}

/// Every sheet's tab colour, keyed by sheet name.
///
/// Read once at open, before any layout: the tab strip draws all of them at
/// once, and the colour lives in each sheet's own part rather than in
/// `workbook.xml`. A workbook whose colour cannot be read simply has none.
pub fn tab_colors(path: &str) -> HashMap<String, String> {
    let mut colors = HashMap::new();
    let Ok(file) = File::open(path) else {
        return colors;
    };
    let Ok(mut zip) = ZipArchive::new(BufReader::new(file)) else {
        return colors;
    };
    let Ok(workbook) = read_entry(&mut zip, "xl/workbook.xml") else {
        return colors;
    };
    let Ok(rels) = read_entry(&mut zip, "xl/_rels/workbook.xml.rels") else {
        return colors;
    };
    let theme = theme_slots(&read_theme(&mut zip));
    let indexed = read_entry(&mut zip, "xl/styles.xml")
        .map(|xml| indexed_palette(&xml))
        .unwrap_or_default();
    let palette = Palette {
        theme: &theme,
        indexed: &indexed,
    };

    for element in scan(&workbook, b"sheet") {
        let Some(name) = element.get("name") else {
            continue;
        };
        let Ok(part) = sheet_part(&workbook, &rels, name) else {
            continue;
        };
        let Ok(entry) = zip.by_name(&part) else {
            continue;
        };
        let mut head = Vec::new();
        if entry.take(SHEET_HEAD_BYTES).read_to_end(&mut head).is_err() {
            continue;
        }
        // The byte run is a document cut off mid-element, which `scan` is
        // already happy with: it keeps what it read and stops.
        if let Some(color) = tab_color(&String::from_utf8_lossy(&head), palette) {
            colors.insert(name.clone(), color);
        }
    }
    colors
}

/// `<sheetPr><tabColor rgb="FF00B050"/></sheetPr>`, resolved the way a cell's
/// colour is: `rgb`, `theme` and `indexed` are all three in use.
fn tab_color(head: &str, palette: Palette) -> Option<String> {
    let mut element = BytesStart::new("color");
    for (key, value) in scan(head, b"tabColor").first()? {
        element.push_attribute((key.as_str(), value.as_str()));
    }
    color_of(&element, palette)
}

/// Reads one sheet's layout, reusing the workbook-wide tables in `cache`.
pub fn read_cached(cache: &mut Cache, path: &str, sheet: &str) -> Result<SheetLayout, String> {
    cache.load(path)?;
    let file = File::open(path).map_err(|e| e.to_string())?;
    let mut zip = ZipArchive::new(BufReader::new(file)).map_err(|e| e.to_string())?;
    let workbook = read_entry(&mut zip, "xl/workbook.xml")?;
    let rels = read_entry(&mut zip, "xl/_rels/workbook.xml.rels")?;
    let part = sheet_part(&workbook, &rels, sheet)?;
    // The sheet is read whole rather than streamed straight out of the archive,
    // because a conditional format asks a question that spans the sheet — what
    // does the cell three columns over hold — and only the walk itself can stay
    // local to one row.
    let sheet_xml = read_entry(&mut zip, &part)?;
    let mut layout = parse_sheet(sheet_xml.as_bytes(), &cache.styles, &cache.rich_strings)?;
    // A sheet with no `_xlnm.Print_Area` still gets shaded by Excel, which falls
    // back to the range it would print. Leaving this `None` painted those sheets
    // white to the edge while every other sheet in the same workbook greyed out
    // past its last cell.
    layout.print_area = print_area(&workbook, sheet).or_else(|| layout.print_region());
    layout.page_setup.print_titles = print_titles(&workbook, sheet);
    layout.names = visible_names(&workbook, sheet);
    layout.page_setup.print_dpi = read_print_dpi(&mut zip, &part);
    read_drawings(&mut zip, cache, &part, &mut layout);
    read_notes(&mut zip, &part, &mut layout);
    // Last, so that the extent a rule is clamped to is the finished one.
    crate::cond::apply(&sheet_xml, cache, &mut layout);
    Ok(layout)
}

/// Reads the sheet's legacy VML part: the notes on it, and the cells Excel
/// marks for them.
///
/// Notes are the only VML these documents use — the shape types a legacy form
/// control writes are not read. A note draws nothing much: a small red
/// triangle in its cell's top-right corner, which is also all that goes into
/// Excel's own picture of the range.
fn read_notes<R: Read + Seek>(zip: &mut ZipArchive<R>, part: &str, layout: &mut SheetLayout) {
    let Some(rel) = layout.legacy_rel.take() else {
        return;
    };
    let Some(vml_part) = part_rels(zip, part).remove(&rel) else {
        return;
    };
    let Ok(xml) = read_entry(zip, &vml_part) else {
        return;
    };
    let mut notes = parse_notes(&xml);
    // The wording is a part of its own; a note without one still gets its mark.
    if let Some(comments_part) = part_by_type(zip, part, "/comments") {
        if let Ok(comments) = read_entry(zip, &comments_part) {
            let by_cell = parse_comments(&comments);
            for note in notes.iter_mut() {
                if let Some((text, font)) = by_cell.get(&(note.row, note.col)) {
                    note.text = text.clone();
                    note.font = font.clone();
                }
            }
        }
    }
    layout.notes = notes;
}

/// The part a worksheet relates to by relationship type — its comments, which
/// no element in the sheet names directly.
fn part_by_type<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
    part: &str,
    suffix: &str,
) -> Option<String> {
    let (dir, base) = part.rsplit_once('/')?;
    let rels = read_entry(zip, &format!("{dir}/_rels/{base}.rels")).ok()?;
    scan(&rels, b"Relationship").into_iter().find_map(|element| {
        let kind = element.get("type")?;
        if !kind.ends_with(suffix) {
            return None;
        }
        Some(resolve_target(dir, element.get("target")?))
    })
}

/// The DPI the sheet was last laid out for: the DEVMODE in the workbook's
/// printer settings part. Excel writes one of these per sheet when a sheet is
/// printed, and the 600 dpi it carries is what the page breaks are snapped to.
fn read_print_dpi<R: Read + Seek>(zip: &mut ZipArchive<R>, part: &str) -> Option<u16> {
    let bin = part_by_type(zip, part, "/printerSettings")?;
    let mut entry = zip.by_name(&bin).ok()?;
    let mut bytes = Vec::new();
    entry.read_to_end(&mut bytes).ok()?;
    devmode_dpi(&bytes)
}

/// `dmPrintQuality` (offset 90) is a positive DPI, or a `DMRES_*` code when
/// negative — in which case `dmYResolution` (offset 96) is the real number.
fn devmode_dpi(bytes: &[u8]) -> Option<u16> {
    let word = |at: usize| bytes.get(at..at + 2).map(|b| i16::from_le_bytes([b[0], b[1]]));
    let quality = word(90)?;
    let dpi = if quality > 0 { quality } else { word(96)? };
    (72..=4800).contains(&dpi).then_some(dpi as u16)
}

/// Walks a VML drawing for its `<x:ClientData ObjectType="Note">` shapes. The
/// row and column are 0-based.
fn parse_notes(xml: &str) -> Vec<NoteCell> {
    let mut reader = XmlReader::from_str(xml);
    let mut buf = Vec::new();
    let mut notes = Vec::new();
    let mut in_shape = false;
    let mut is_note = false;
    let mut row: Option<u32> = None;
    let mut col: Option<u32> = None;
    let mut number_slot: Option<bool> = None;
    let mut in_anchor = false;
    let mut anchor_text = String::new();

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element)) => match element.local_name().as_ref() {
                b"shape" => {
                    in_shape = true;
                    is_note = false;
                    row = None;
                    col = None;
                }
                b"ClientData" if in_shape => {
                    is_note = attr(&element, b"ObjectType").as_deref() == Some("Note");
                }
                b"Row" if in_shape => number_slot = Some(true),
                b"Column" if in_shape => number_slot = Some(false),
                b"Anchor" if in_shape => {
                    in_anchor = true;
                    anchor_text.clear();
                }
                _ => {}
            },
            Ok(Event::End(element)) => match element.local_name().as_ref() {
                b"Row" | b"Column" => number_slot = None,
                b"Anchor" => in_anchor = false,
                b"shape" => {
                    if is_note {
                        if let (Some(row), Some(col)) = (row, col) {
                            notes.push(NoteCell {
                                row,
                                col,
                                text: String::new(),
                                font: None,
                                anchor: parse_anchor(&anchor_text, row, col),
                            });
                        }
                    }
                    in_shape = false;
                }
                _ => {}
            },
            Ok(Event::Text(content)) => {
                if let Some(is_row) = number_slot {
                    if let Ok(value) = content.xml_content(XmlVersion::default()) {
                        let parsed = value.trim().parse::<u32>().ok();
                        if is_row {
                            row = parsed;
                        } else {
                            col = parsed;
                        }
                    }
                } else if in_anchor {
                    if let Ok(value) = content.xml_content(XmlVersion::default()) {
                        anchor_text.push_str(&value);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }
    notes
}

/// The eight anchor numbers, or a one-cell box on the note's own cell when the
/// file has none (or writes something unreadable).
fn parse_anchor(text: &str, row: u32, col: u32) -> [f64; 8] {
    let values: Vec<f64> = text
        .split(',')
        .filter_map(|value| value.trim().parse::<f64>().ok())
        .collect();
    if values.len() == 8 {
        let mut anchor = [0.0; 8];
        anchor.copy_from_slice(&values);
        return anchor;
    }
    [col as f64, 0.0, row as f64, 0.0, (col + 1) as f64, 0.0, (row + 1) as f64, 0.0]
}

/// `comments1.xml`: the wording of each note, by the cell it hangs on, with the
/// font its first run names.
fn parse_comments(xml: &str) -> HashMap<(u32, u32), (String, Option<NoteFont>)> {
    let mut reader = XmlReader::from_str(xml);
    let mut buf = Vec::new();
    let mut out = HashMap::new();
    let mut current: Option<(u32, u32)> = None;
    let mut in_text = false;
    let mut text = String::new();
    let mut font: Option<RichFont> = None;
    let mut first_font: Option<RichFont> = None;

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element)) => match element.local_name().as_ref() {
                b"comment" => {
                    current = attr(&element, b"ref").and_then(|reference| parse_cell(&reference));
                    text.clear();
                    first_font = None;
                }
                b"text" if current.is_some() => in_text = true,
                b"rPr" if in_text => font = Some(RichFont::default()),
                b"rFont" | b"sz" | b"b" | b"i" | b"u" | b"strike" | b"vertAlign" if in_text => {
                    read_run_font(&element, font.as_mut());
                }
                _ => {}
            },
            Ok(Event::Empty(element)) => match element.local_name().as_ref() {
                b"rFont" | b"sz" | b"b" | b"i" | b"u" | b"strike" | b"vertAlign" if in_text => {
                    read_run_font(&element, font.as_mut());
                }
                _ => {}
            },
            Ok(Event::End(element)) => match element.local_name().as_ref() {
                b"rPr" => {
                    if first_font.is_none() {
                        first_font = font.take();
                    } else {
                        font = None;
                    }
                }
                b"text" => in_text = false,
                b"comment" => {
                    if let Some(cell) = current.take() {
                        out.insert(
                            cell,
                            (
                                text.trim().to_owned(),
                                first_font.take().map(|font| NoteFont {
                                    family: font.family,
                                    size_px: font.size_px,
                                    bold: font.bold.unwrap_or(false),
                                    italic: font.italic.unwrap_or(false),
                                }),
                            ),
                        );
                    }
                    text.clear();
                }
                _ => {}
            },
            Ok(Event::Text(content)) if in_text => {
                if let Ok(value) = content.xml_content(XmlVersion::default()) {
                    text.push_str(&value);
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }
    out
}

/// Reads the sheet's drawing part, if it has one, and hangs its shapes on the
/// layout.
///
/// Every failure here is silent on purpose: a sheet whose shapes cannot be read
/// still shows its cells, and the cells are the part that matters.
fn read_drawings<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
    cache: &Cache,
    part: &str,
    layout: &mut SheetLayout,
) {
    let Some(rel) = layout.drawing_rel.take() else {
        return;
    };
    let Some(drawing_part) = part_rels(zip, part).remove(&rel) else {
        return;
    };
    let Ok(xml) = read_entry(zip, &drawing_part) else {
        return;
    };

    // The shape pass works in pixels and cannot wait for the renderer's
    // measurement, so this is the one place left that converts — see
    // `chars_to_px` for what that costs and what it does not.
    let col_widths_px: HashMap<u32, f64> = layout
        .col_widths_chars
        .iter()
        .map(|(col, chars)| (*col, chars_to_px(*chars)))
        .collect();
    let grid = drawing::Grid::new(
        layout.cols,
        &col_widths_px,
        chars_to_px(layout.default_col_width_chars),
        layout.rows,
        &layout.row_heights_px,
        layout.default_row_height_px,
    );
    let mut shapes = drawing::parse(&xml, &grid, &cache.theme, &cache.ln_widths, &cache.theme_fonts);

    // A picture names its image by relationship id; resolve them all in one read.
    let media = part_rels(zip, &drawing_part);
    for shape in &mut shapes {
        if let Some(rel) = shape.media.take() {
            shape.media = media.get(&rel).cloned();
        }
    }

    // A shape can be anchored past the used range — a brace drawn below the last
    // row, say — and Excel scrolls to it, so the grid has to reach that far too.
    for shape in &shapes {
        layout.cols = layout.cols.max(shape.to_col + 1);
        layout.rows = layout.rows.max(shape.to_row + 1);
    }
    layout.cols = layout.cols.min(MAX_SHEET_COLS);
    layout.rows = layout.rows.min(MAX_SHEET_ROWS);
    layout.drawings = shapes;
}

/// Every relationship a part declares, as `r:id` -> the target's part name.
fn part_rels<R: Read + Seek>(zip: &mut ZipArchive<R>, part: &str) -> HashMap<String, String> {
    let Some((dir, base)) = part.rsplit_once('/') else {
        return HashMap::new();
    };
    let Ok(rels) = read_entry(zip, &format!("{dir}/_rels/{base}.rels")) else {
        return HashMap::new();
    };
    scan(&rels, b"Relationship")
        .into_iter()
        .filter_map(|element| {
            let id = element.get("id")?.clone();
            let target = resolve_target(dir, element.get("target")?);
            Some((id, target))
        })
        .collect()
}

/// Resolves a relationship target against the folder its part lives in, so
/// `../drawings/drawing1.xml` from `xl/worksheets` is `xl/drawings/drawing1.xml`.
fn resolve_target(dir: &str, target: &str) -> String {
    let mut parts: Vec<&str> = if target.starts_with('/') {
        Vec::new()
    } else {
        dir.split('/').collect()
    };
    for segment in target.trim_start_matches('/').split('/') {
        match segment {
            ".." => {
                parts.pop();
            }
            "" | "." => {}
            other => parts.push(other),
        }
    }
    parts.join("/")
}

/// The character an entity reference stands for.
///
/// quick-xml hands `&lt;` back as its own event rather than as part of the text
/// around it, so a parser that reads only `Event::Text` drops it silently — which
/// is how `<>` inside a label came out as nothing at all, and with it every
/// coloured run keyed by that label.
pub(crate) fn entity_text(reference: &BytesRef) -> Option<char> {
    if let Ok(Some(ch)) = reference.resolve_char_ref() {
        return Some(ch);
    }
    match reference.decode().ok()?.as_ref() {
        "lt" => Some('<'),
        "gt" => Some('>'),
        "amp" => Some('&'),
        "apos" => Some('\''),
        "quot" => Some('"'),
        _ => None,
    }
}

/// The shared strings in the order the sheets index them into.
///
/// A cell that carries a shared string holds its *number*, so anything that has
/// to compare a cell's text — a conditional format's formula — needs the table
/// itself and not just the runs above.
fn shared_strings(xml: &str) -> Vec<String> {
    let mut reader = XmlReader::from_str(xml);
    let mut buf = Vec::new();
    let mut out: Vec<String> = Vec::new();
    let mut in_text = false;
    let mut in_phonetic = false;

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element)) | Ok(Event::Empty(element)) => {
                match element.local_name().as_ref() {
                    b"si" => out.push(String::new()),
                    // A phonetic reading spells the same text out again in
                    // `<t>` elements of its own and is no part of the string.
                    b"rPh" => in_phonetic = true,
                    b"t" => in_text = true,
                    _ => {}
                }
            }
            Ok(Event::Text(text)) if in_text && !in_phonetic => {
                if let (Some(slot), Ok(text)) = (out.last_mut(), text.xml_content(XmlVersion::default()))
                {
                    slot.push_str(&text);
                }
            }
            Ok(Event::GeneralRef(reference)) if in_text && !in_phonetic => {
                if let (Some(slot), Some(ch)) = (out.last_mut(), entity_text(&reference)) {
                    slot.push(ch);
                }
            }
            Ok(Event::End(element)) => match element.local_name().as_ref() {
                b"t" => in_text = false,
                b"rPh" => in_phonetic = false,
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    out
}

/// Records a run's colour, which is written as `<color …/>` beside the font
/// properties of its `<rPr>`.
fn set_run_color(element: &BytesStart, run: Option<&mut RichRun>, palette: Palette) {
    if let (Some(run), Some(color)) = (run, color_of(element, palette)) {
        run.color = Some(color);
    }
}

/// Records one property of a run's `<rPr>` — the run's own font, as distinct from
/// the cell's. Unknown names are ignored, and so is a property that arrives with
/// no `<rPr>` open.
fn read_run_font(element: &BytesStart, font: Option<&mut RichFont>) {
    let Some(font) = font else { return };
    match element.local_name().as_ref() {
        b"rFont" => font.family = attr(element, b"val"),
        b"sz" => {
            font.size_px = attr(element, b"val")
                .and_then(|v| v.parse::<f64>().ok())
                .map(points_to_font_px);
        }
        b"b" | b"i" | b"strike" => {
            let on = !matches!(attr(element, b"val").as_deref(), Some("0" | "false"));
            match element.local_name().as_ref() {
                b"b" => font.bold = Some(on),
                b"i" => font.italic = Some(on),
                _ => font.strike = Some(on),
            }
        }
        b"u" => {
            // The two accounting forms draw the same line, spanning the cell
            // rather than the text; they land on single and double. An explicit
            // `none` is a value of its own — a run inside an underlined cell has
            // to be able to say that it draws no line.
            font.underline = Some(match attr(element, b"val").as_deref() {
                Some("none") => "none",
                Some("double" | "doubleAccounting") => "double",
                _ => "single",
            });
        }
        b"vertAlign" => {
            font.vert_align = match attr(element, b"val").as_deref() {
                Some("superscript") => Some("superscript"),
                Some("subscript") => Some("subscript"),
                _ => None,
            }
        }
        _ => {}
    }
}

/// The shared strings that are split into runs, one slot per `<si>` in the
/// order the sheet's cells index them.
///
/// A cell shows the concatenation of its runs either way, so the text is nothing
/// new; what a run adds is formatting for part of it — how one word inside a
/// sentence comes out red, or bold, or underlined. The slots line up with
/// `shared_strings`, which is what lets a cell's `<v>` index find its own runs:
/// two cells can spell the same text with different runs, so the text alone
/// cannot identify them.
///
/// Only strings whose runs really say something are kept. A run that merely
/// restates the cell's own colour draws exactly like the plain text, and keeping
/// those would put most of the workbook in the table.
fn rich_strings(xml: &str, palette: Palette) -> Vec<Option<Vec<RichRun>>> {
    let mut out: Vec<Option<Vec<RichRun>>> = Vec::new();
    let mut reader = XmlReader::from_str(xml);
    let mut buf = Vec::new();
    let mut runs: Option<Vec<RichRun>> = None;
    let mut run: Option<RichRun> = None;
    // The `<rPr>` being read, which belongs to the run that is open.
    let mut font: Option<RichFont> = None;
    let mut in_text = false;

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element)) => match element.local_name().as_ref() {
                b"si" => {
                    runs = Some(Vec::new());
                    out.push(None);
                }
                b"r" => {
                    run = Some(RichRun {
                        text: String::new(),
                        color: None,
                        font: None,
                    });
                }
                b"rPr" => font = Some(RichFont::default()),
                b"t" => in_text = true,
                b"color" => set_run_color(&element, run.as_mut(), palette),
                _ => read_run_font(&element, font.as_mut()),
            },
            // A font property and a run's colour are written as empty elements,
            // which never get a start/end pair — but `<b></b>` is the same
            // property spelled the long way, hence both arms.
            Ok(Event::Empty(element)) => match element.local_name().as_ref() {
                b"si" => out.push(None),
                b"color" => set_run_color(&element, run.as_mut(), palette),
                _ => read_run_font(&element, font.as_mut()),
            },
            // An entity arrives as its own event, so the text around it has to be
            // stitched back together here or the run loses those characters and
            // stops matching the cell it belongs to.
            Ok(Event::GeneralRef(reference)) => {
                if let (Some(run), Some(ch)) = (run.as_mut(), entity_text(&reference)) {
                    run.text.push(ch);
                }
            }
            // Only text inside a `<r>` counts; a phonetic `<rPh>` sets the same
            // flag but has no run open, so its reading is left out.
            Ok(Event::Text(text)) if in_text => {
                if let (Some(run), Ok(text)) =
                    (run.as_mut(), text.xml_content(XmlVersion::default()))
                {
                    run.text.push_str(&text);
                }
            }
            Ok(Event::End(element)) => match element.local_name().as_ref() {
                b"t" => in_text = false,
                b"rPr" => {
                    if let (Some(run), Some(font)) = (run.as_mut(), font.take()) {
                        run.font = (!font.is_empty()).then_some(font);
                    }
                }
                b"r" => {
                    if let (Some(list), Some(run)) = (runs.as_mut(), run.take()) {
                        list.push(run);
                    }
                }
                b"si" => {
                    if let Some(list) = runs.take() {
                        if list.iter().any(|run| run.color.is_some() || run.font.is_some()) {
                            if let Some(slot) = out.last_mut() {
                                *slot = Some(list);
                            }
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    out
}

/// Every `_xlnm.Print_Area` in `workbook.xml`, keyed by the sheet it belongs to.
///
/// The sheet a name belongs to is its `localSheetId` — an index into the
/// workbook's own sheet order — and that is the sheet Excel applies it to. The
/// formula usually names that sheet as well, and in all 291 workbooks of the
/// corpus that set a print area the two agree; where they do not, the scope is
/// the one to believe. A defined name rewritten without its formula, or a sheet
/// named with a `!` in it — which the formula splits on and cannot escape —
/// are the two ways they come apart.
fn print_area_names(workbook: &str) -> HashMap<String, String> {
    defined_names(workbook, "_xlnm.Print_Area")
}

/// Every `<definedName>` in `workbook.xml`: the sheet it is scoped to (`None`
/// for a workbook-wide name), its name, and its formula.
fn defined_name_entries(workbook: &str) -> Vec<(Option<usize>, String, String)> {
    let mut out = Vec::new();
    let mut reader = XmlReader::from_str(workbook);
    let mut buf = Vec::new();
    let mut open = false;
    let mut scope: Option<usize> = None;
    let mut name = String::new();
    let mut formula = String::new();

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element)) if element.local_name().as_ref() == b"definedName" => {
                open = true;
                scope = attr(&element, b"localSheetId").and_then(|v| v.parse::<usize>().ok());
                name = attr(&element, b"name").unwrap_or_default();
                formula.clear();
            }
            Ok(Event::Text(text)) if open => {
                if let Ok(content) = text.xml_content(XmlVersion::default()) {
                    formula.push_str(&content);
                }
            }
            // An escaped character arrives as its own event, as it does in a
            // shared string, so the formula has to be stitched back together.
            Ok(Event::GeneralRef(reference)) if open => {
                if let Some(ch) = entity_text(&reference) {
                    formula.push(ch);
                }
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"definedName" => {
                open = false;
                out.push((scope.take(), name.trim().to_owned(), formula.trim().to_owned()));
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    out
}

/// Every defined name called `which` in `workbook.xml`, keyed by the sheet it
/// belongs to. The scope is `localSheetId`, an index into the workbook's own
/// sheet order; a formula that names the sheet as well usually agrees with it.
fn defined_names(workbook: &str, which: &str) -> HashMap<String, String> {
    let sheets: Vec<String> = scan(workbook, b"sheet")
        .iter()
        .filter_map(|element| element.get("name").cloned())
        .collect();

    let mut out = HashMap::new();
    for (scope, name, formula) in defined_name_entries(workbook) {
        if name != which {
            continue;
        }
        // `<sheet>!<range>`, the sheet name quoted only when it must be — and a
        // local name may leave the sheet out altogether, since the scope
        // already says which it is.
        let (named, range) = match formula.split_once('!') {
            Some((owner, range)) => (Some(owner.trim_matches('\'')), range.trim()),
            None => (None, formula.as_str()),
        };
        let owner = scope
            .and_then(|index| sheets.get(index).cloned())
            .or_else(|| named.map(str::to_owned));
        if let Some(owner) = owner {
            out.insert(owner, range.to_owned());
        }
    }
    out
}

/// The names one sheet can see: the workbook-wide ones, plus those scoped to
/// it, which win — the order Excel resolves them in.
///
/// A data-validation list may name one of these instead of writing out a range
/// (`<formula1>対象</formula1>`, with `対象` defined as `LIST!$H$2:$H$23`),
/// which is what a sheet that keeps its choices on a hidden list does.
fn visible_names(workbook: &str, sheet: &str) -> HashMap<String, String> {
    let sheets: Vec<String> = scan(workbook, b"sheet")
        .iter()
        .filter_map(|element| element.get("name").cloned())
        .collect();
    let sheet_index = sheets.iter().position(|candidate| candidate == sheet);

    let mut out = HashMap::new();
    for (scope, name, formula) in defined_name_entries(workbook) {
        match scope {
            Some(owner) if Some(owner) == sheet_index => {
                out.insert(name, formula);
            }
            None => {
                out.entry(name).or_insert(formula);
            }
            _ => {}
        }
    }
    out
}

/// The sheet's `_xlnm.Print_Area`, if it has one.
fn print_area(workbook: &str, sheet_name: &str) -> Option<MergeRect> {
    print_area_names(workbook)
        .get(sheet_name)
        .and_then(|range| parse_range(range))
}

/// The sheet's `_xlnm.Print_Titles`: the rows and columns repeated at the top
/// and left of every printed page.
///
/// The formula lists the ranges separated by commas — `Sheet!$A:$C` for the
/// repeating columns, `Sheet!$1:$2` for the repeating rows, in either order.
fn print_titles(workbook: &str, sheet_name: &str) -> Option<PrintTitles> {
    let formula = defined_names(workbook, "_xlnm.Print_Titles").remove(sheet_name)?;
    let mut titles = PrintTitles {
        rows: None,
        cols: None,
    };
    for part in formula.split(',') {
        // A local name may leave the sheet out; the scope already says which.
        let range = part.rsplit_once('!').map_or(part, |(_, tail)| tail).trim();
        match title_axis(range) {
            Some((true, first, last)) => titles.rows = Some([first, last]),
            Some((false, first, last)) => titles.cols = Some([first, last]),
            None => {}
        }
    }
    (titles.rows.is_some() || titles.cols.is_some()).then_some(titles)
}

/// A `$A:$C` or `$1:$2` range, as zero-based inclusive bounds and which axis
/// it lies on — `true` for rows.
fn title_axis(range: &str) -> Option<(bool, u32, u32)> {
    let (start, end) = range.split_once(':')?;
    if let (Some(first), Some(last)) = (column_index(start), column_index(end)) {
        return Some((false, first, last));
    }
    if let (Some(first), Some(last)) = (row_index(start), row_index(end)) {
        return Some((true, first, last));
    }
    None
}

fn column_index(text: &str) -> Option<u32> {
    let mut value = 0u32;
    for ch in text.chars() {
        match ch {
            '$' => {}
            'A'..='Z' | 'a'..='z' => {
                value = value * 26 + (ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1)
            }
            _ => return None,
        }
    }
    (value > 0).then(|| value - 1)
}

fn row_index(text: &str) -> Option<u32> {
    let digits: String = text.chars().filter(|ch| *ch != '$').collect();
    if digits.is_empty() || !digits.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    digits.parse::<u32>().ok().filter(|row| *row > 0).map(|row| row - 1)
}

/// The sheets that define a `_xlnm.Print_Area` of their own.
///
/// Read at open, alongside the tab colours, because the answer decides which
/// view the workbook opens in: Excel's Page Break Preview shades around the
/// print area and boxes it, and a workbook where nobody ever set one belongs in
/// the ordinary view rather than in a page the app guessed at.
///
/// `print_area` above cannot answer this — it falls back to the used range, so
/// by the time a layout exists *every* sheet has an area. This walks
/// `workbook.xml` itself, where the difference is still visible.
pub fn print_area_sheets(path: &str) -> HashSet<String> {
    let Ok(file) = File::open(path) else {
        return HashSet::new();
    };
    let Ok(mut zip) = ZipArchive::new(BufReader::new(file)) else {
        return HashSet::new();
    };
    let Ok(workbook) = read_entry(&mut zip, "xl/workbook.xml") else {
        return HashSet::new();
    };
    print_area_names(&workbook).into_keys().collect()
}

/// Resolves a sheet name to its part inside the zip via workbook relationships.
fn sheet_part(workbook: &str, rels: &str, sheet_name: &str) -> Result<String, String> {
    let mut targets = HashMap::new();
    for element in scan(rels, b"Relationship") {
        if let (Some(id), Some(target)) = (element.get("id"), element.get("target")) {
            targets.insert(id.clone(), target.clone());
        }
    }

    let rid = scan(workbook, b"sheet")
        .into_iter()
        .find(|element| element.get("name").is_some_and(|n| n == sheet_name))
        .and_then(|element| element.get("id").cloned())
        .ok_or_else(|| format!("sheet not found in workbook.xml: {sheet_name}"))?;

    let target = targets
        .get(&rid)
        .ok_or_else(|| format!("no relationship {rid} for sheet {sheet_name}"))?;

    let trimmed = target.trim_start_matches('/');
    Ok(if trimmed.starts_with("xl/") {
        trimmed.to_owned()
    } else {
        format!("xl/{trimmed}")
    })
}

/// Reads `xl/styles.xml`: the `cellXfs` table, the `borders` it points into, and
/// the `fonts` those point into.
fn read_styles<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
    theme: &[Option<String>],
) -> Styles {
    let Ok(xml) = read_entry(zip, "xl/styles.xml") else {
        return Styles::default();
    };
    // The palette is written at the end of the part, so it is read before the
    // walk rather than discovered by it.
    let indexed = indexed_palette(&xml);
    let mut styles = parse_styles(
        &xml,
        Palette {
            theme,
            indexed: &indexed,
        },
    );
    styles.indexed = indexed;
    styles
}

/// Walks `xl/styles.xml` — the split that lets the walk be tested on a string.
fn parse_styles(xml: &str, palette: Palette) -> Styles {
    let mut reader = XmlReader::from_str(xml);
    let mut buf = Vec::new();
    let mut styles = Styles::default();

    let mut in_borders = false;
    let mut in_fonts = false;
    let mut in_fills = false;
    let mut in_num_fmts = false;
    let mut in_cell_xfs = false;
    let mut in_named_xfs = false;
    let mut border: Option<CellBorder> = None;
    let mut edge: Option<Side> = None;
    let mut font: Option<FontInfo> = None;
    let mut fill: Option<String> = None;
    let mut fill_pattern: Option<&'static str> = None;
    let mut fill_fg: Option<String> = None;
    let mut fill_bg: Option<String> = None;
    let mut xf: Option<Xf> = None;
    let mut named_xf: Option<Xf> = None;

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element)) => match element.local_name().as_ref() {
                b"borders" => in_borders = true,
                b"border" if in_borders => {
                    border = Some(CellBorder {
                        diagonal_up: flag_attr(&element, b"diagonalUp").unwrap_or(false),
                        diagonal_down: flag_attr(&element, b"diagonalDown").unwrap_or(false),
                        ..CellBorder::default()
                    })
                }
                b"left" | b"right" | b"top" | b"bottom" | b"diagonal" if in_borders => {
                    edge = set_edge(&element, border.as_mut(), &element.local_name().as_ref());
                }
                b"color" if in_borders => color_edge(&element, border.as_mut(), edge, palette),
                // A solid `<patternFill>` paints its `fgColor`; gradients and
                // the stock none/gray125 entries have no flat colour to show.
                b"fills" => in_fills = true,
                b"fill" if in_fills => {
                    fill = None;
                    fill_pattern = None;
                    fill_fg = None;
                    fill_bg = None;
                }
                b"patternFill" if in_fills => {
                    fill_pattern =
                        attr(&element, b"patternType").and_then(|kind| pattern_kind(&kind));
                }
                b"fgColor" if in_fills => {
                    // A solid fill paints its `fgColor`; a pattern draws it as
                    // the hatch colour over `bgColor`.
                    fill_fg = color_of(&element, palette);
                    if fill_pattern == Some("solid") {
                        fill = fill_fg.clone();
                    }
                }
                b"bgColor" if in_fills => fill_bg = color_of(&element, palette),
                b"fonts" => in_fonts = true,
                b"numFmts" => in_num_fmts = true,
                b"numFmt" if in_num_fmts => record_num_format(&element, &mut styles),
                b"font" if in_fonts => font = Some(FontInfo::default()),
                b"b" | b"i" | b"u" | b"strike" if in_fonts => {
                    read_font_prop(&element, font.as_mut())
                }
                b"vertAlign" if in_fonts => {
                    if let Some(slot) = font.as_mut() {
                        slot.vert_align = match attr(&element, b"val").as_deref() {
                            Some("superscript") => Some("superscript"),
                            Some("subscript") => Some("subscript"),
                            _ => None,
                        };
                    }
                }
                b"sz" if in_fonts => {
                    if let Some(slot) = font.as_mut() {
                        slot.size_pt = num_attr(&element, b"val");
                    }
                }
                b"name" if in_fonts => {
                    if let Some(slot) = font.as_mut() {
                        slot.name = attr(&element, b"val");
                    }
                }
                b"color" if in_fonts => {
                    if let Some(slot) = font.as_mut() {
                        slot.color = color_of(&element, palette);
                    }
                }
                b"cellStyleXfs" => in_named_xfs = true,
                b"cellXfs" => in_cell_xfs = true,
                b"xf" if in_cell_xfs => xf = Some(new_xf(&element)),
                // A named style's own entry, which a cell format inherits from.
                // Its properties are read the same way; only where they land
                // differs.
                b"xf" if in_named_xfs => named_xf = Some(new_xf(&element)),
                b"alignment" if in_cell_xfs => {
                    if let Some(slot) = xf.as_mut() {
                        read_alignment(&element, &mut slot.align);
                    }
                }
                b"alignment" if in_named_xfs => {
                    if let Some(slot) = named_xf.as_mut() {
                        read_alignment(&element, &mut slot.align);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(element)) => match element.local_name().as_ref() {
                // A self-closing element has no colour child to wait for.
                b"border" if in_borders => styles.borders.push(border.take().filter(|b| {
                    b.top.is_some()
                        || b.right.is_some()
                        || b.bottom.is_some()
                        || b.left.is_some()
                        || b.diagonal.is_some()
                })),
                b"left" | b"right" | b"top" | b"bottom" | b"diagonal" if in_borders => {
                    set_edge(&element, border.as_mut(), &element.local_name().as_ref());
                }
                // Excel writes the colour self-closing, e.g. `<color rgb="FF000000"/>`.
                b"color" if in_borders => color_edge(&element, border.as_mut(), edge, palette),
                b"fill" if in_fills => styles.fills.push(None),
                b"patternFill" if in_fills => {
                    fill_pattern =
                        attr(&element, b"patternType").and_then(|kind| pattern_kind(&kind));
                }
                b"fgColor" if in_fills => {
                    fill_fg = color_of(&element, palette);
                    if fill_pattern == Some("solid") {
                        fill = fill_fg.clone();
                    }
                }
                b"bgColor" if in_fills => fill_bg = color_of(&element, palette),
                b"font" if in_fonts => styles.fonts.push(FontInfo::default()),
                b"numFmt" if in_num_fmts => record_num_format(&element, &mut styles),
                b"b" | b"i" | b"u" | b"strike" if in_fonts => {
                    read_font_prop(&element, font.as_mut())
                }
                b"vertAlign" if in_fonts => {
                    if let Some(slot) = font.as_mut() {
                        slot.vert_align = match attr(&element, b"val").as_deref() {
                            Some("superscript") => Some("superscript"),
                            Some("subscript") => Some("subscript"),
                            _ => None,
                        };
                    }
                }
                b"sz" if in_fonts => {
                    if let Some(slot) = font.as_mut() {
                        slot.size_pt = num_attr(&element, b"val");
                    }
                }
                b"name" if in_fonts => {
                    if let Some(slot) = font.as_mut() {
                        slot.name = attr(&element, b"val");
                    }
                }
                // Excel writes font colours self-closing too.
                b"color" if in_fonts => {
                    if let Some(slot) = font.as_mut() {
                        slot.color = color_of(&element, palette);
                    }
                }
                b"xf" if in_cell_xfs => styles.xfs.push(new_xf(&element)),
                b"xf" if in_named_xfs => styles.style_xfs.push(new_xf(&element)),
                b"alignment" if in_cell_xfs => {
                    if let Some(slot) = xf.as_mut() {
                        read_alignment(&element, &mut slot.align);
                    }
                }
                b"alignment" if in_named_xfs => {
                    if let Some(slot) = named_xf.as_mut() {
                        read_alignment(&element, &mut slot.align);
                    }
                }
                _ => {}
            },
            Ok(Event::End(element)) => match element.local_name().as_ref() {
                b"borders" => in_borders = false,
                // The stock entry has a `<left/>`-style placeholder for every
                // side; normalise it to "no border" rather than a border with
                // four empty edges.
                b"border" if in_borders => styles.borders.push(border.take().filter(|b| {
                    b.top.is_some()
                        || b.right.is_some()
                        || b.bottom.is_some()
                        || b.left.is_some()
                        || b.diagonal.is_some()
                })),
                b"left" | b"right" | b"top" | b"bottom" | b"diagonal" if in_borders => {
                    edge = None
                }
                b"fonts" => in_fonts = false,
                b"numFmts" => in_num_fmts = false,
                b"fills" => in_fills = false,
                b"fill" if in_fills => {
                    let spec = match fill_pattern.take() {
                        Some("solid") => fill.take().map(FillSpec::Solid),
                        Some(kind) => Some(FillSpec::Pattern(CellPattern {
                            kind,
                            // A file that leaves the colours out means Excel's
                            // automatic pair: black hatch on white.
                            fg: fill_fg.take().unwrap_or_else(|| "#000000".to_owned()),
                            bg: fill_bg.take().unwrap_or_else(|| "#FFFFFF".to_owned()),
                        })),
                        None => None,
                    };
                    fill = None;
                    fill_fg = None;
                    fill_bg = None;
                    styles.fills.push(spec);
                }
                b"font" if in_fonts => styles.fonts.push(font.take().unwrap_or_default()),
                b"cellStyleXfs" => in_named_xfs = false,
                b"cellXfs" => in_cell_xfs = false,
                b"xf" if in_cell_xfs => {
                    if let Some(slot) = xf.take() {
                        styles.xfs.push(slot);
                    }
                }
                b"xf" if in_named_xfs => {
                    if let Some(slot) = named_xf.take() {
                        styles.style_xfs.push(slot);
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    // The `dxfs` are what the sheet's conditional formats paint, and they are
    // read in a walk of their own: a differential format is its own small schema,
    // and folding it into this one would mean carrying its state through every
    // branch above.
    styles.dxfs = crate::cond::dxfs(xml, palette);
    styles
}

#[derive(Clone, Copy, PartialEq)]
pub(crate) enum Side {
    Top,
    Right,
    Bottom,
    Left,
    Diagonal,
}

/// `<numFmt numFmtId="164" formatCode="..."/>` — a format the workbook defines
/// itself, such as a Japanese date.
fn record_num_format(element: &BytesStart, styles: &mut Styles) {
    if let (Some(id), Some(code)) = (
        attr(element, b"numFmtId").and_then(|v| v.parse::<u32>().ok()),
        attr(element, b"formatCode"),
    ) {
        styles.num_formats.insert(id, code);
    }
}

/// An `apply…` flag, kept as `None` when the file writes no such attribute —
/// which is not the same as writing one that is off.
fn flag_attr(element: &BytesStart, name: &[u8]) -> Option<bool> {
    attr(element, name).map(|value| value == "1" || value == "true")
}

/// `<b/>`, `<i/>`, `<strike/>` and `<u val="…"/>` — the font properties whose
/// mere presence means "on" and whose `val` can say otherwise.
///
/// 44 of 60 workbooks in the corpus use `<i/>` and 42 use `<u/>`, and neither
/// was read: every italic and every underlined cell drew plain. `val` is
/// spelled out only to turn one *off* or to pick the double line, so the
/// default is on.
fn read_font_prop(element: &BytesStart, font: Option<&mut FontInfo>) {
    let Some(font) = font else { return };
    match element.local_name().as_ref() {
        b"b" | b"i" | b"strike" => {
            let on = !matches!(attr(element, b"val").as_deref(), Some("0" | "false"));
            match element.local_name().as_ref() {
                b"b" => font.bold = on,
                b"i" => font.italic = on,
                _ => font.strike = on,
            }
        }
        b"u" => {
            // The two accounting forms draw the same line, spanning the cell
            // rather than the text; they land on single and double.
            font.underline = match attr(element, b"val").as_deref() {
                Some("none") => None,
                Some("double" | "doubleAccounting") => Some("double"),
                _ => Some("single"),
            };
        }
        _ => {}
    }
}

fn new_xf(element: &BytesStart) -> Xf {
    Xf {
        align: Align::default(),
        border_id: attr(element, b"borderId")
            .and_then(|v| v.parse::<usize>().ok())
            .unwrap_or(0),
        font_id: attr(element, b"fontId")
            .and_then(|v| v.parse::<usize>().ok())
            .unwrap_or(0),
        fill_id: attr(element, b"fillId")
            .and_then(|v| v.parse::<usize>().ok())
            .unwrap_or(0),
        num_fmt_id: attr(element, b"numFmtId")
            .and_then(|v| v.parse::<u32>().ok())
            .unwrap_or(0),
        style_id: attr(element, b"xfId")
            .and_then(|v| v.parse::<usize>().ok())
            .unwrap_or(0),
        apply_font: flag_attr(element, b"applyFont"),
        apply_fill: flag_attr(element, b"applyFill"),
        apply_border: flag_attr(element, b"applyBorder"),
        apply_number: flag_attr(element, b"applyNumberFormat"),
        apply_align: flag_attr(element, b"applyAlignment"),
    }
}

/// The built-in formats that carry a date or time. The rest are numeric or
/// plain, and are left to the default rendering.
fn builtin_date_format(id: u32) -> Option<&'static str> {
    Some(match id {
        14 => "m/d/yyyy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yyyy h:mm",
        45 => "mm:ss",
        _ => return None,
    })
}

/// Records one side's style and returns which side is open, so a nested
/// `<color>` knows where to land.
pub(crate) fn set_edge(
    element: &BytesStart,
    border: Option<&mut CellBorder>,
    name: &[u8],
) -> Option<Side> {
    let side = match name {
        b"top" => Side::Top,
        b"right" => Side::Right,
        b"bottom" => Side::Bottom,
        b"left" => Side::Left,
        b"diagonal" => Side::Diagonal,
        _ => return None,
    };
    let border = border?;
    let edge = attr(element, b"style").and_then(|name| edge_style(&name));
    let slot = match side {
        Side::Top => &mut border.top,
        Side::Right => &mut border.right,
        Side::Bottom => &mut border.bottom,
        Side::Left => &mut border.left,
        Side::Diagonal => &mut border.diagonal,
    };
    *slot = edge;
    Some(side)
}

pub(crate) fn color_edge(
    element: &BytesStart,
    border: Option<&mut CellBorder>,
    side: Option<Side>,
    palette: Palette,
) {
    let (Some(border), Some(side)) = (border, side) else {
        return;
    };
    let Some(color) = color_of(element, palette) else {
        return;
    };
    let slot = match side {
        Side::Top => &mut border.top,
        Side::Right => &mut border.right,
        Side::Bottom => &mut border.bottom,
        Side::Left => &mut border.left,
        Side::Diagonal => &mut border.diagonal,
    };
    if let Some(edge) = slot.as_mut() {
        edge.color = Some(color);
    }
}

/// `<patternFill patternType>` reduced to the names the renderer has tiles for.
/// `none` and anything unknown paint nothing.
fn pattern_kind(name: &str) -> Option<&'static str> {
    Some(match name {
        "solid" => "solid",
        "lightGray" => "lightGray",
        "mediumGray" => "mediumGray",
        "darkGray" => "darkGray",
        "gray125" => "gray125",
        "gray0625" => "gray0625",
        "darkHorizontal" => "darkHorizontal",
        "darkVertical" => "darkVertical",
        "darkDown" => "darkDown",
        "darkUp" => "darkUp",
        "darkGrid" => "darkGrid",
        "darkTrellis" => "darkTrellis",
        "lightHorizontal" => "lightHorizontal",
        "lightVertical" => "lightVertical",
        "lightDown" => "lightDown",
        "lightUp" => "lightUp",
        "lightGrid" => "lightGrid",
        "lightTrellis" => "lightTrellis",
        _ => return None,
    })
}

/// Excel's border names reduced to a CSS style, a width and a precedence rank.
/// The rank follows Excel's own ordering so the heavier of two shared edges wins.
fn edge_style(name: &str) -> Option<BorderEdge> {
    let (line, width, rank) = match name {
        // Excel's hairline is a dotted rule, not a thin solid one: a 1:1
        // CopyPicture of a hair-bordered row shows 1 px dots with 1 px gaps.
        "hair" => ("dotted", 1, 1),
        "dotted" => ("dotted", 1, 2),
        "dashed" => ("dashed", 1, 3),
        "thin" => ("solid", 1, 4),
        "mediumDashed" => ("dashed", 2, 5),
        "dashDot" => ("dashed", 1, 6),
        "mediumDashDot" => ("dashed", 2, 7),
        "dashDotDot" => ("dashed", 1, 8),
        "medium" => ("solid", 2, 9),
        "mediumDashDotDot" => ("dashed", 2, 10),
        "slantDashDot" => ("dashed", 1, 11),
        "thick" => ("solid", 3, 12),
        "double" => ("double", 3, 13),
        _ => return None,
    };
    Some(BorderEdge {
        line,
        width,
        rank,
        color: None,
    })
}

/// `FF000000` (ARGB) or `000000` (RGB) to `#000000`.
///
/// The hex digits are validated: `sysClr` carries a *name* in `val`, and
/// "window" is six characters long, so a length check alone would happily
/// accept it and produce a colour that is not one.
fn parse_color(value: &str) -> Option<String> {
    let hex = value.trim().trim_start_matches('#');
    if !hex.chars().all(|c| c.is_ascii_hexdigit()) {
        return None;
    }
    match hex.len() {
        8 => Some(format!("#{}", &hex[2..])),
        6 => Some(format!("#{hex}")),
        _ => None,
    }
}

/// The twelve `<a:clrScheme>` entries, keyed by their role name.
fn read_theme<R: Read + Seek>(zip: &mut ZipArchive<R>) -> HashMap<String, String> {
    let Ok(xml) = read_entry(zip, "xl/theme/theme1.xml") else {
        return HashMap::new();
    };
    parse_theme(&xml)
}

/// The faces a theme's font scheme names, keyed the way a drawing refers to
/// them: `mn-lt`, `mn-ea`, `mj-lt`, `mj-ea`.
///
/// An east asian slot left empty means "whatever the latin one is" — which is
/// how a stock Office theme reads, since it says nothing about Japanese at all
/// — so that fallback is filled in here rather than left for every reader.
fn read_theme_fonts<R: Read + Seek>(zip: &mut ZipArchive<R>) -> HashMap<String, String> {
    let mut fonts = HashMap::new();
    let Ok(xml) = read_entry(zip, "xl/theme/theme1.xml") else {
        return fonts;
    };
    let mut reader = XmlReader::from_str(&xml);
    let mut buf = Vec::new();
    // Which of the two schemes is open — `mj` or `mn` — and the latin face it
    // named, which an empty east asian slot falls back to.
    let mut scheme: Option<&str> = None;
    let mut latin: Option<String> = None;

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element)) | Ok(Event::Empty(element)) => {
                match element.local_name().as_ref() {
                    b"majorFont" => {
                        scheme = Some("mj");
                        latin = None;
                    }
                    b"minorFont" => {
                        scheme = Some("mn");
                        latin = None;
                    }
                    // Only the scheme's own face counts: the per-script ones
                    // that follow carry their face as an attribute.
                    b"latin" if scheme.is_some() && latin.is_none() => {
                        if let Some(face) = attr(&element, b"typeface").filter(|f| !f.is_empty()) {
                            fonts.insert(format!("{}-lt", scheme.unwrap_or_default()), face.clone());
                            latin = Some(face);
                        }
                    }
                    b"ea" if scheme.is_some() => {
                        if let Some(face) = attr(&element, b"typeface").filter(|f| !f.is_empty()) {
                            fonts.insert(format!("{}-ea", scheme.unwrap_or_default()), face);
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(element))
                if matches!(element.local_name().as_ref(), b"majorFont" | b"minorFont") =>
            {
                if let (Some(scheme), Some(latin)) = (scheme, latin.take()) {
                    fonts.entry(format!("{scheme}-ea")).or_insert(latin);
                }
                scheme = None;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    fonts
}

/// The theme's `lnStyleLst` widths, in pixels, which a shape's `lnRef idx="n"`
/// indexes. Excel's own theme ships 0.5pt, 1pt and 1.5pt, but a document may
/// carry its own theme.
fn read_line_widths<R: Read + Seek>(zip: &mut ZipArchive<R>) -> Vec<f64> {
    let Ok(xml) = read_entry(zip, "xl/theme/theme1.xml") else {
        return Vec::new();
    };
    let mut reader = XmlReader::from_str(&xml);
    let mut buf = Vec::new();
    let mut inside = false;
    let mut widths = Vec::new();

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element)) | Ok(Event::Empty(element)) => {
                match element.local_name().as_ref() {
                    b"lnStyleLst" => inside = true,
                    b"ln" if inside => {
                        // Line weights are EMU, like every other length here.
                        // `points_to_px` would round 1pt to a whole pixel, which
                        // is too coarse for a line that is drawn to be looked at.
                        if let Some(width) = num_attr(&element, b"w") {
                            widths.push(width / 9525.0);
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"lnStyleLst" => {
                inside = false;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    widths
}

fn parse_theme(xml: &str) -> HashMap<String, String> {
    let mut reader = XmlReader::from_str(xml);
    let mut buf = Vec::new();
    let mut scheme = HashMap::new();
    let mut slot: Option<String> = None;
    let mut in_scheme = false;

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            // `<a:dk1>` opens a slot; the colour element inside it fills it.
            Ok(Event::Start(element)) | Ok(Event::Empty(element)) => {
                let name = String::from_utf8_lossy(element.local_name().as_ref()).into_owned();
                if name == "clrScheme" {
                    in_scheme = true;
                    continue;
                }
                if !in_scheme {
                    continue;
                }
                // `val` is a hex colour on srgbClr but a system *name* on
                // sysClr, where the literal lives in `lastClr` instead.
                let value = match name.as_str() {
                    "srgbClr" => attr(&element, b"val"),
                    "sysClr" => attr(&element, b"lastClr"),
                    _ => None,
                };
                match value.as_deref().and_then(parse_color) {
                    Some(hex) => {
                        if let Some(role) = slot.take() {
                            scheme.insert(role, hex);
                        }
                    }
                    None if name != "srgbClr" && name != "sysClr" && slot.is_none() => {
                        slot = Some(name);
                    }
                    None => {}
                }
            }
            Ok(Event::End(element)) => {
                let name = String::from_utf8_lossy(element.local_name().as_ref()).into_owned();
                if name == "clrScheme" {
                    in_scheme = false;
                } else if slot.as_deref() == Some(name.as_str()) {
                    slot = None;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    scheme
}

/// Resolves a `<color …/>` however it names its colour: `rgb`, `theme`,
/// `indexed`, or `auto`. The theme and indexed forms are what real documents
/// lean on, and reading only `rgb` left most of their colours unresolved.
pub(crate) fn color_of(element: &BytesStart, palette: Palette) -> Option<String> {
    if let Some(rgb) = attr(element, b"rgb").as_deref().and_then(parse_color) {
        return Some(tint(rgb, attr(element, b"tint")));
    }
    if attr(element, b"auto").is_some_and(|value| value == "1" || value == "true") {
        return Some(BLACK.to_owned());
    }
    if let Some(index) = attr(element, b"theme").and_then(|v| v.parse::<usize>().ok()) {
        return match palette.theme.get(index).and_then(|slot| slot.clone()) {
            Some(base) => Some(tint(base, attr(element, b"tint"))),
            // Without a parsed theme, slot 1 is the text colour and 0 the page —
            // and the tint applies to those two just the same.
            None => match index {
                1 => Some(tint(BLACK.to_owned(), attr(element, b"tint"))),
                0 => Some(tint(WHITE.to_owned(), attr(element, b"tint"))),
                _ => None,
            },
        };
    }
    if let Some(index) = attr(element, b"indexed").and_then(|v| v.parse::<usize>().ok()) {
        if let Some(color) = palette.indexed.get(index).and_then(|slot| slot.clone()) {
            return Some(color);
        }
        return match index {
            i if i < INDEXED.len() => Some(format!("#{}", INDEXED[i])),
            64 => Some(BLACK.to_owned()),
            65 => Some(WHITE.to_owned()),
            _ => None,
        };
    }
    None
}

/// Excel lightens or darkens a colour by walking its HSL luminance.
///
/// The drawing parser needs the same arithmetic: `<a:tint>` means exactly this.
pub(crate) fn tint(hex: String, tint: Option<String>) -> String {
    let Some(tint) = tint.and_then(|value| value.parse::<f64>().ok()) else {
        return hex;
    };
    if tint == 0.0 || hex.len() != 7 {
        return hex;
    }

    let channel = |range: std::ops::Range<usize>| {
        u8::from_str_radix(&hex[range], 16).unwrap_or(0) as f64 / 255.0
    };
    let (r, g, b) = (channel(1..3), channel(3..5), channel(5..7));

    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let mut l = (max + min) / 2.0;
    let delta = max - min;
    let s = if delta == 0.0 {
        0.0
    } else {
        delta / (1.0 - (2.0 * l - 1.0).abs())
    };
    let mut h = if delta == 0.0 {
        0.0
    } else if max == r {
        60.0 * (((g - b) / delta) % 6.0)
    } else if max == g {
        60.0 * ((b - r) / delta + 2.0)
    } else {
        60.0 * ((r - g) / delta + 4.0)
    };
    h = ((h % 360.0) + 360.0) % 360.0;

    l = if tint < 0.0 {
        l * (1.0 + tint)
    } else {
        l * (1.0 - tint) + tint
    };

    let c = (1.0 - (2.0 * l - 1.0).abs()) * s;
    let x = c * (1.0 - (((h / 60.0) % 2.0) - 1.0).abs());
    let m = l - c / 2.0;
    let (r1, g1, b1) = match h as u32 / 60 {
        0 => (c, x, 0.0),
        1 => (x, c, 0.0),
        2 => (0.0, c, x),
        3 => (0.0, x, c),
        4 => (x, 0.0, c),
        _ => (c, 0.0, x),
    };
    let byte = |value: f64| ((value + m) * 255.0).round().clamp(0.0, 255.0) as u8;
    format!("#{:02X}{:02X}{:02X}", byte(r1), byte(g1), byte(b1))
}

fn read_alignment(element: &BytesStart, slot: &mut Align) {
    if let Some(value) = attr(element, b"horizontal") {
        slot.horizontal = match value.as_str() {
            "left" | "justify" | "distributed" => Some("left"),
            "center" => Some("center"),
            // Center Across Selection is not "centred inside the cell": Excel
            // centres the text across the run of neighbours carrying the same
            // alignment and drops the borders between them, so a button drawn
            // as four cells reads as one box. The grid needs it kept apart.
            "centerContinuous" => Some("centerContinuous"),
            "right" => Some("right"),
            // "general" and "fill" fall through to the grid's own default.
            _ => None,
        };
    }
    if let Some(value) = attr(element, b"vertical") {
        slot.vertical = match value.as_str() {
            "top" => Some("top"),
            "center" | "justify" | "distributed" => Some("middle"),
            "bottom" => Some("bottom"),
            _ => None,
        };
    }
    if let Some(value) = attr(element, b"wrapText") {
        slot.wrap = value == "1" || value == "true";
    }
    if let Some(value) = attr(element, b"shrinkToFit") {
        slot.shrink = value == "1" || value == "true";
    }
    if let Some(value) = attr(element, b"textRotation").and_then(|v| v.parse::<i16>().ok()) {
        match value {
            // 255 is the stacked form, not an angle.
            255 => slot.stacked = true,
            // 91–180 is the clockwise half, written as the dial's -1…-90.
            91..=180 => slot.rotation = 90 - value,
            0..=90 => slot.rotation = value,
            _ => {}
        }
    }
    if let Some(value) = attr(element, b"indent") {
        slot.indent = value.trim().parse::<u8>().unwrap_or(0);
    }
}

/// Turns one row's styled cells into runs and files them.
///
/// Sorted and merged exactly as the whole-sheet pass this replaced did: adjacent
/// columns sharing a style become one run, and the sort is what keeps the runs
/// ordered when a file writes its cells out of order. Only that row is held, so
/// nothing here scales with the sheet.
fn collect_runs(layout: &mut SheetLayout, row: u32, cells: &mut Vec<(u32, u32)>) {
    if cells.is_empty() {
        return;
    }
    cells.sort_unstable();
    let mut runs: Vec<[u32; 3]> = Vec::with_capacity(cells.len());
    for &(col, style) in cells.iter() {
        match runs.last_mut() {
            Some(last) if last[2] == style && last[1] + 1 == col => last[1] = col,
            _ => runs.push([col, col, style]),
        }
    }
    layout.cell_runs.insert(row, runs);
    cells.clear();
}

fn parse_sheet<R: BufRead>(
    reader: R,
    styles: &Styles,
    rich: &[Option<Vec<RichRun>>],
) -> Result<SheetLayout, String> {
    let mut xml = XmlReader::from_reader(reader);
    let mut layout = SheetLayout::empty();
    let mut buf = Vec::new();
    let mut filter_col: Option<u32> = None;
    // `<autoFilter>` is not only the sheet's own: a custom view carries the
    // filter it was saved with, and an extension block can too. The schema puts
    // the sheet's own before either of those, but they were each overwriting it
    // as they went past, so the last one won — a WBS whose filter is `M1:FC66`
    // got the row a saved view had been looking at instead, and its filter
    // button landed seven rows down. Counting the nesting is what tells them
    // apart.
    let mut borrowed_filters: u32 = 0;
    let mut declared_default_col: Option<f64> = None;
    let mut base_col_width: Option<f64> = None;
    let mut declared_default_row: Option<f64> = None;
    // Outline groups: the level each row sits at, and where a closed group's
    // summary row is. Collected as the sheet streams; the renderer turns them
    // into bands and buttons.
    let mut outline_rows: Vec<OutlineRow> = Vec::new();
    let mut outline_cols: Vec<OutlineCol> = Vec::new();
    let mut outline_summary_below = true;
    let mut outline_summary_right = true;
    let mut extent_rows: u32 = 0;
    let mut extent_cols: u32 = 0;
    // Distinct styles seen so far, so each is stored once.
    let mut style_ids: HashMap<CellStyle, u32> = HashMap::new();
    // Runs are built one row at a time as the sheet streams past, so nothing is
    // held for the sheet as a whole and the buffer never holds more than the row
    // being read. The whole-sheet map this replaces could promise neither, which
    // is why it needed a cap — and the cap silently dropped the style of every
    // cell past it, which on a document that fills its sheet means the borders
    // from that row down.
    let mut pending_row: u32 = u32::MAX;
    let mut pending: Vec<(u32, u32)> = Vec::new();
    // Cells only carry a size when their font differs from this.
    let default_font_px = styles
        .fonts
        .first()
        .and_then(|font| font.size_pt)
        .map(points_to_font_px)
        .unwrap_or_else(|| points_to_font_px(STOCK_FONT_PT));

    // The cell whose shared string is streaming past, and whether its `<v>` is
    // the thing the text now belongs to. A rich-text cell is identified by the
    // shared-string index in that value, not by the text it spells out.
    let mut rich_cell: Option<(u32, u32)> = None;
    let mut in_cell_value = false;

    // The rule being read, and which of its `<formula1>`/`<formula2>` the text
    // now belongs to — the formulas are element text, not attributes.
    let mut validation: Option<DataValidation> = None;
    let mut formula_slot: Option<u8> = None;

    // Which `<brk>` list is being read. Both carry bare `<brk>` children, so
    // the container is the only thing that says whether an id names a row or a
    // column.
    let mut in_row_breaks = false;
    let mut in_col_breaks = false;

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(element)) | Ok(Event::Empty(element)) => {
                match element.local_name().as_ref() {
                    b"sheetFormatPr" => {
                        declared_default_col = attr(&element, b"defaultColWidth")
                            .and_then(|v| v.parse::<f64>().ok());
                        base_col_width = attr(&element, b"baseColWidth")
                            .and_then(|v| v.parse::<f64>().ok());
                        declared_default_row = attr(&element, b"defaultRowHeight")
                            .and_then(|v| v.parse::<f64>().ok());
                    }
                    // Which side of a group its summary row sits on. The
                    // defaults are below and right; Excel writes the attribute
                    // only when the author asked for the other side.
                    b"outlinePr" => {
                        outline_summary_below =
                            flag_attr(&element, b"summaryBelow").unwrap_or(true);
                        outline_summary_right =
                            flag_attr(&element, b"summaryRight").unwrap_or(true);
                    }
                    b"sheetView" => {
                        // Absent means the default, which is on.
                        if let Some(value) = attr(&element, b"showGridLines") {
                            layout.show_grid_lines = value != "0" && value != "false";
                        }
                        // Where the sheet is left: Excel opens it at this zoom
                        // and scrolled to this cell, not at 100% and A1.
                        if let Some(value) = attr(&element, b"zoomScale")
                            .and_then(|v| v.parse::<u16>().ok())
                        {
                            layout.zoom_scale = Some(value);
                        }
                        if let Some((row, col)) =
                            attr(&element, b"topLeftCell").and_then(|cell| parse_cell(&cell))
                        {
                            layout.view_top_left = Some([row, col]);
                        }
                    }
                    // A frozen view holds the rows above the split and the
                    // columns to its left still while the rest scrolls. The
                    // split counts are authoritative; `topLeftCell` names the
                    // scrolling region's first cell, which is the same count
                    // when a writer leaves them out.
                    //
                    // Each custom sheet view keeps a pane of its own, and those
                    // are somebody's saved views rather than the one the file
                    // opens at — a workbook of thirteen of them left the grid
                    // frozen and scrolled to the last one's corner.
                    b"pane" => {
                        if borrowed_filters == 0
                            && attr(&element, b"state").as_deref() == Some("frozen")
                        {
                            let split = |name: &[u8]| {
                                attr(&element, name)
                                    .and_then(|value| value.parse::<u32>().ok())
                                    .unwrap_or(0)
                            };
                            layout.frozen_rows = split(b"ySplit");
                            layout.frozen_cols = split(b"xSplit");
                            if layout.frozen_rows == 0 && layout.frozen_cols == 0 {
                                if let Some((row, col)) =
                                    attr(&element, b"topLeftCell").and_then(|cell| parse_cell(&cell))
                                {
                                    layout.frozen_rows = row;
                                    layout.frozen_cols = col;
                                }
                            }
                            // A frozen pane's `topLeftCell` is the scrolling
                            // region's first cell — where the sheet is left.
                            if let Some((row, col)) =
                                attr(&element, b"topLeftCell").and_then(|cell| parse_cell(&cell))
                            {
                                layout.view_top_left = Some([row, col]);
                            }
                        }
                    }
                    // How the sheet prints. Page Break Preview draws from
                    // this: the paper decides how many columns fit across,
                    // the margins how much of it may be used, and the manual
                    // breaks where a page ends regardless.
                    b"pageMargins" => {
                        let m = &mut layout.page_setup.margins;
                        if let Some(v) = num_attr(&element, b"left") {
                            m.left = v;
                        }
                        if let Some(v) = num_attr(&element, b"right") {
                            m.right = v;
                        }
                        if let Some(v) = num_attr(&element, b"top") {
                            m.top = v;
                        }
                        if let Some(v) = num_attr(&element, b"bottom") {
                            m.bottom = v;
                        }
                        if let Some(v) = num_attr(&element, b"header") {
                            m.header = v;
                        }
                        if let Some(v) = num_attr(&element, b"footer") {
                            m.footer = v;
                        }
                    }
                    b"pageSetup" => {
                        let setup = &mut layout.page_setup;
                        if let Some(v) = attr(&element, b"paperSize")
                            .and_then(|v| v.parse::<u16>().ok())
                        {
                            setup.paper_size = v;
                        }
                        if let Some(v) = attr(&element, b"orientation") {
                            setup.landscape = v == "landscape";
                        }
                        if let Some(v) = attr(&element, b"scale").and_then(|v| v.parse::<u16>().ok())
                        {
                            setup.scale = v;
                        }
                        if let Some(v) = attr(&element, b"fitToWidth")
                            .and_then(|v| v.parse::<u16>().ok())
                        {
                            setup.fit_to_width = v;
                        }
                        if let Some(v) = attr(&element, b"fitToHeight")
                            .and_then(|v| v.parse::<u16>().ok())
                        {
                            setup.fit_to_height = v;
                        }
                        setup.page_order_over_then_down =
                            attr(&element, b"pageOrder").as_deref() == Some("overThenDown");
                    }
                    // What turns `scale` off: fit-to-page lives on
                    // `<pageSetUpPr>` under `<sheetPr>`, not on `<pageSetup>`.
                    b"pageSetUpPr" => {
                        layout.page_setup.fit_to_page =
                            flag_attr(&element, b"fitToPage").unwrap_or(false);
                    }
                    b"printOptions" => {
                        let options = &mut layout.page_setup.print_options;
                        options.horizontal_centered =
                            flag_attr(&element, b"horizontalCentered").unwrap_or(false);
                        options.vertical_centered =
                            flag_attr(&element, b"verticalCentered").unwrap_or(false);
                        options.headings = flag_attr(&element, b"headings").unwrap_or(false);
                        options.grid_lines = flag_attr(&element, b"gridLines").unwrap_or(false);
                    }
                    b"rowBreaks" => in_row_breaks = true,
                    b"colBreaks" => in_col_breaks = true,
                    // Manual breaks only. Automatic ones are not in the file —
                    // Excel works them out from the paper — and a stale
                    // `man="0"` entry would add a second line where the paper
                    // already breaks.
                    b"brk" if in_row_breaks || in_col_breaks => {
                        if flag_attr(&element, b"man").unwrap_or(false) {
                            if let Some(id) = attr(&element, b"id")
                                .and_then(|v| v.parse::<u32>().ok())
                            {
                                let list = if in_row_breaks {
                                    &mut layout.page_setup.row_breaks
                                } else {
                                    &mut layout.page_setup.col_breaks
                                };
                                list.push(id.saturating_sub(1));
                            }
                        }
                    }
                    b"col" => {
                        let min = attr(&element, b"min")                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(1);
                        let max = attr(&element, b"max")
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(min);
                        if min == 0 || max < min {
                            continue;
                        }
                        let last = max.min(min.saturating_add(MAX_EXPANDED_COLUMNS - 1));

                        // Outline level and the closed-group marker, which
                        // may sit on a column that carries no width at all.
                        let level = attr(&element, b"outlineLevel")
                            .and_then(|v| v.parse::<u8>().ok())
                            .unwrap_or(0)
                            .min(7);
                        let collapsed = flag_attr(&element, b"collapsed").unwrap_or(false);
                        if level > 0 || collapsed {
                            outline_cols.push(OutlineCol {
                                from: min - 1,
                                to: last - 1,
                                level,
                                collapsed,
                            });
                        }

                        // A column run carries its default style in `style`,
                        // which is not the `s` a cell uses.
                        if let Some(index) = attr(&element, b"style")
                            .and_then(|v| v.parse::<usize>().ok())
                        {
                            if let Some(style) = styles.cell_style(index, default_font_px) {
                                let style = intern(&mut layout.styles, &mut style_ids, style);
                                layout.col_styles.push(ColumnStyle {
                                    from: min - 1,
                                    to: last - 1,
                                    style,
                                });
                            }
                        }

                        // Excel writes `width="0"` on a hidden column as well,
                        // so it already takes no room; the flag is what lets the
                        // renderer skip its letter and put a width back if the
                        // reader asks to see it.
                        if attr(&element, b"hidden").is_some_and(|v| v == "1" || v == "true") {
                            for column in min..=last {
                                layout.hidden_cols.push(column - 1);
                            }
                        }

                        let Some(width) = attr(&element, b"width")
                            .and_then(|v| v.parse::<f64>().ok())
                        else {
                            continue;
                        };
                        // A bare `<col min="1" max="16384" width="8.43"/>` just
                        // restates the default; expanding it would be waste.
                        let custom = attr(&element, b"customWidth")
                            .is_some_and(|v| v == "1" || v == "true");
                        if max - min + 1 > MAX_EXPANDED_COLUMNS && !custom {
                            continue;
                        }
                        for column in min..=last {
                            layout.col_widths_chars.insert(column - 1, width);
                        }
                    }
                    b"row" => {
                        // Whatever is pending belongs to the row that just ended.
                        collect_runs(&mut layout, pending_row, &mut pending);
                        let Some(row) =
                            attr(&element, b"r").and_then(|v| v.parse::<u32>().ok())
                        else {
                            continue;
                        };
                        if row == 0 {
                            continue;
                        }
                        // A hidden row keeps no height of its own, so without
                        // this the reader sees a row Excel does not draw at all.
                        if attr(&element, b"hidden").is_some_and(|v| v == "1" || v == "true") {
                            layout.hidden_rows.push(row - 1);
                        }
                        // The row's place in the outline. A detail row carries
                        // the level; the summary row beside a closed group
                        // carries `collapsed` and no level at all.
                        let level = attr(&element, b"outlineLevel")
                            .and_then(|v| v.parse::<u8>().ok())
                            .unwrap_or(0)
                            .min(7);
                        let collapsed = flag_attr(&element, b"collapsed").unwrap_or(false);
                        if level > 0 || collapsed {
                            outline_rows.push(OutlineRow {
                                row: row - 1,
                                level,
                                collapsed,
                            });
                        }
                        // Rows can carry a style for the cells in them that have
                        // none of their own — how a filled row reads as solid.
                        // The id only counts when `customFormat` says so; Excel
                        // writes the flag on every row it styles (10,655 rows
                        // across 40 workbooks, none without it), so this only
                        // tells other writers' output apart.
                        if flag_attr(&element, b"customFormat").unwrap_or(false) {
                            if let Some(index) = attr(&element, b"s")
                                .and_then(|v| v.parse::<usize>().ok())
                            {
                                if let Some(style) = styles.cell_style(index, default_font_px) {
                                    let style = intern(&mut layout.styles, &mut style_ids, style);
                                    layout.row_styles.insert(row - 1, style);
                                }
                            }
                        }
                        if let Some(height) =
                            attr(&element, b"ht").and_then(|v| v.parse::<f64>().ok())
                        {
                            layout.row_heights_px.insert(row - 1, points_to_px(height));
                        }
                        extent_rows = extent_rows.max(row);
                    }
                    // A cell's value: for a shared string it is the index into
                    // the table, which is what identifies the runs.
                    b"v" => in_cell_value = rich_cell.is_some(),
                    b"c" => {
                        // Cell display values come from calamine; this pass only
                        // recovers the style, and only when it says something.
                        let Some((row, col)) =
                            attr(&element, b"r").and_then(|r| parse_cell(&r))
                        else {
                            continue;
                        };
                        // A shared-string cell names its string by number; the
                        // runs are looked up by that number when `<v>` arrives,
                        // because two cells may spell the same text with
                        // different runs.
                        rich_cell = (attr(&element, b"t").as_deref() == Some("s"))
                            .then_some((row, col));
                        // The sheet's extent is where its cells are, styled or
                        // not: a conditional format has to reach an unstyled
                        // cell, and it is clamped to this extent.
                        extent_rows = extent_rows.max(row + 1);
                        extent_cols = extent_cols.max(col + 1);
                        let Some(index) = attr(&element, b"s")
                            .and_then(|v| v.parse::<usize>().ok())
                        else {
                            continue;
                        };
                        let Some(style) = styles.cell_style(index, default_font_px) else {
                            continue;
                        };
                        // Rows arrive in order, but a file is free to write its
                        // cells out of order or spread a row over several
                        // elements, so the buffer follows the cell's own row
                        // rather than where the cell appeared.
                        if row != pending_row {
                            collect_runs(&mut layout, pending_row, &mut pending);
                            pending_row = row;
                        }
                        pending.push((col, intern(&mut layout.styles, &mut style_ids, style)));
                    }
                    b"mergeCell" => {
                        if layout.merges.len() >= MAX_MERGES {
                            continue;
                        }
                        if let Some(reference) = attr(&element, b"ref") {
                            if let Some(rect) = parse_range(&reference) {
                                extent_rows = extent_rows.max(rect.row + rect.rows);
                                extent_cols = extent_cols.max(rect.col + rect.cols);
                                layout.merges.push(rect);
                            }
                        }
                    }
                    b"customSheetViews" | b"extLst" => borrowed_filters += 1,
                    b"autoFilter" => {
                        if borrowed_filters == 0 {
                            if let Some(rect) =
                                attr(&element, b"ref").and_then(|r| parse_range(&r))
                            {
                                layout.auto_filter = Some(AutoFilter {
                                    row: rect.row,
                                    col: rect.col,
                                    rows: rect.rows,
                                    cols: rect.cols,
                                    applied: HashMap::new(),
                                    custom: Vec::new(),
                                    hidden_buttons: Vec::new(),
                                });
                            }
                        }
                    }
                    // The sheet only points at its shapes; the part itself is read
                    // once the sheet's own extent is known.
                    b"drawing" => layout.drawing_rel = attr(&element, b"id"),
                    // Notes live in a legacy VML part, not in the drawing.
                    b"legacyDrawing" => layout.legacy_rel = attr(&element, b"id"),
                    b"filterColumn" => {
                        // A custom sheet view's filter is relative to a range
                        // of its own; reading its columns as this filter's
                        // would hide buttons on columns Excel still shows.
                        let offset = (borrowed_filters == 0)
                            .then(|| attr(&element, b"colId").and_then(|v| v.parse::<u32>().ok()))
                            .flatten();
                        // ECMA-376 defaults `showButton` to true; "0" is the
                        // file saying Excel draws no drop-down on that column
                        // — a merged header's covered columns are written that
                        // way, so only the region's right edge carries one.
                        if let (Some(offset), Some(filter)) =
                            (offset, layout.auto_filter.as_mut())
                        {
                            let shown = match attr(&element, b"showButton") {
                                Some(value) => value != "0" && value != "false",
                                None => true,
                            };
                            let column = filter.col + offset;
                            if !shown && !filter.hidden_buttons.contains(&column) {
                                filter.hidden_buttons.push(column);
                            }
                        }
                        filter_col = offset;
                    }
                    b"customFilter" => {
                        // A condition — "greater than 20" — rather than a value
                        // list. What it keeps is the file's own cached `hidden`
                        // answer; the column is recorded only so the filter
                        // button draws Excel's funnel instead of its caret.
                        if let (Some(offset), Some(filter)) =
                            (filter_col, layout.auto_filter.as_mut())
                        {
                            let column = filter.col + offset;
                            if !filter.custom.contains(&column) {
                                filter.custom.push(column);
                            }
                        }
                    }
                    b"filter" => {
                        // `<filter val="…"/>` only carries a value when the file
                        // was saved with that column filtered.
                        if let (Some(offset), Some(value)) = (filter_col, attr(&element, b"val")) {
                            if let Some(filter) = layout.auto_filter.as_mut() {
                                filter
                                    .applied
                                    .entry(filter.col + offset)
                                    .or_default()
                                    .push(value);
                            }
                        }
                    }
                    b"dataValidation" => {
                        let flag = |name: &[u8]| {
                            attr(&element, name).is_some_and(|v| v == "1" || v == "true")
                        };
                        validation = Some(DataValidation {
                            ranges: attr(&element, b"sqref")
                                .map(|refs| {
                                    refs.split_whitespace()
                                        .filter_map(parse_range)
                                        .collect()
                                })
                                .unwrap_or_default(),
                            kind: attr(&element, b"type")
                                .unwrap_or_else(|| "none".to_owned()),
                            formula1: None,
                            formula2: None,
                            allow_blank: flag(b"allowBlank"),
                            // Excel's own quirk: `InCellDropdown = False`
                            // writes `showDropDown="1"`.
                            hide_dropdown: flag(b"showDropDown"),
                            show_input: flag(b"showInputMessage"),
                            show_error: flag(b"showErrorMessage"),
                            prompt_title: attr(&element, b"promptTitle"),
                            prompt: attr(&element, b"prompt"),
                            error_title: attr(&element, b"errorTitle"),
                            error: attr(&element, b"error"),
                        });
                    }
                    b"formula1" if validation.is_some() => formula_slot = Some(1),
                    b"formula2" if validation.is_some() => formula_slot = Some(2),
                    _ => {}
                }
            }
            Ok(Event::End(element)) => {
                match element.local_name().as_ref() {
                    b"filterColumn" => filter_col = None,
                    b"rowBreaks" => in_row_breaks = false,
                    b"colBreaks" => in_col_breaks = false,
                    b"v" => in_cell_value = false,
                    b"c" => rich_cell = None,
                    b"formula1" | b"formula2" => {
                        // The text arrives in pieces when the file escapes its
                        // characters as entity references, so it is trimmed
                        // once, here, rather than per fragment.
                        if let Some(rule) = validation.as_mut() {
                            let slot = if formula_slot == Some(1) {
                                &mut rule.formula1
                            } else {
                                &mut rule.formula2
                            };
                            let trimmed = slot.as_deref().map(str::trim).unwrap_or("");
                            *slot = (!trimmed.is_empty()).then(|| trimmed.to_owned());
                        }
                        formula_slot = None;
                    }
                    b"dataValidation" => {
                        if let Some(rule) = validation.take() {
                            layout.validations.push(rule);
                        }
                    }
                    b"customSheetViews" | b"extLst" => {
                        borrowed_filters = borrowed_filters.saturating_sub(1)
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(text)) if in_cell_value => {
                if let (Some((row, col)), Ok(value)) =
                    (rich_cell, text.xml_content(XmlVersion::default()))
                {
                    if let Some(Some(runs)) = value
                        .trim()
                        .parse::<usize>()
                        .ok()
                        .and_then(|index| rich.get(index))
                    {
                        layout.rich_cells.push(RichCell {
                            row,
                            col,
                            runs: runs.clone(),
                        });
                    }
                }
                in_cell_value = false;
            }
            Ok(Event::Text(text)) if formula_slot.is_some() => {
                if let (Some(rule), Ok(content)) =
                    (validation.as_mut(), text.xml_content(XmlVersion::default()))
                {
                    let slot = if formula_slot == Some(1) {
                        &mut rule.formula1
                    } else {
                        &mut rule.formula2
                    };
                    slot.get_or_insert_with(String::new).push_str(&content);
                }
            }
            // A character the file writes as an entity (`&#23550;`) arrives as
            // its own event: without this the list of values breaks at the
            // first such character and the rule reads as next to empty.
            Ok(Event::GeneralRef(reference)) if formula_slot.is_some() => {
                if let (Some(rule), Some(ch)) = (validation.as_mut(), entity_text(&reference)) {
                    let slot = if formula_slot == Some(1) {
                        &mut rule.formula1
                    } else {
                        &mut rule.formula2
                    };
                    slot.get_or_insert_with(String::new).push(ch);
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(e.to_string()),
            _ => {}
        }
    }

    // `baseColWidth` is what Excel writes for an untouched sheet, and it is the
    // width every undeclared column really has — falling back to something
    // wider would silently disagree with Excel.
    layout.default_col_width_chars = declared_default_col
        .or(base_col_width)
        .unwrap_or(STOCK_COL_WIDTH_CHARS)
        .max(1.0);
    layout.default_row_height_px =
        points_to_px(declared_default_row.unwrap_or(STOCK_ROW_HEIGHT_PT).max(1.0));

    layout.rows = extent_rows.min(MAX_SHEET_ROWS);
    layout.cols = extent_cols.min(MAX_SHEET_COLS);
    // A file writes its rows in order, but nothing makes it, and the payload is
    // compared byte for byte when a change is checked.
    layout.hidden_rows.sort_unstable();
    layout.hidden_cols.sort_unstable();

    if !outline_rows.is_empty() || !outline_cols.is_empty() {
        outline_rows.sort_by_key(|entry| entry.row);
        outline_cols.sort_by_key(|entry| entry.from);
        layout.outline = Some(Outline {
            rows: outline_rows,
            cols: outline_cols,
            summary_below: outline_summary_below,
            summary_right: outline_summary_right,
        });
    }

    // The last row is ended by the end of the sheet rather than by a `<row>`.
    collect_runs(&mut layout, pending_row, &mut pending);

    // The workbook's first font entry is the "Normal" style, so it is what
    // every unstyled cell inherits.
    if let Some(normal) = styles.fonts.first() {
        layout.font_family = normal.name.clone();
        if let Some(size) = normal.size_pt {
            layout.font_size_px = points_to_font_px(size.max(1.0));
        }
    }

    Ok(layout)
}

/// Parses an A1 range such as `A1:C1` or a single cell such as `B7`.
pub(crate) fn parse_range(reference: &str) -> Option<MergeRect> {
    let (start, end) = reference.split_once(':').unwrap_or((reference, reference));
    let (row1, col1) = parse_cell(start)?;
    let (row2, col2) = parse_cell(end)?;
    if row2 < row1 || col2 < col1 {
        return None;
    }
    Some(MergeRect {
        row: row1,
        col: col1,
        rows: row2 - row1 + 1,
        cols: col2 - col1 + 1,
    })
}

/// Parses `B7` into zero-based `(row, col)`, tolerating `$` anchors.
pub(crate) fn parse_cell(reference: &str) -> Option<(u32, u32)> {
    let mut col = 0u32;
    let mut digits = String::new();
    for ch in reference.chars() {
        match ch {
            '$' => {}
            'A'..='Z' | 'a'..='z' => {
                if !digits.is_empty() {
                    return None;
                }
                col = col * 26 + (ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
            }
            '0'..='9' => digits.push(ch),
            _ => return None,
        }
    }
    let row: u32 = digits.parse().ok()?;
    if row == 0 || col == 0 {
        return None;
    }
    Some((row - 1, col - 1))
}

fn read_entry<R: Read + Seek>(zip: &mut ZipArchive<R>, name: &str) -> Result<String, String> {
    let mut entry = zip.by_name(name).map_err(|e| format!("{name}: {e}"))?;
    let mut out = String::new();
    entry.read_to_string(&mut out).map_err(|e| e.to_string())?;
    Ok(out)
}

/// First attribute matching `name`, ignoring any namespace prefix.
pub(crate) fn attr(element: &BytesStart, name: &[u8]) -> Option<String> {
    element.attributes().flatten().find_map(|attribute| {
        (attribute.key.local_name().as_ref() == name)
            .then(|| String::from_utf8_lossy(&attribute.value).into_owned())
    })
}

pub(crate) fn num_attr(element: &BytesStart, name: &[u8]) -> Option<f64> {
    attr(element, name).and_then(|value| value.trim().parse::<f64>().ok())
}

/// Every `<element>` in `xml`, as lower-cased attribute name/value pairs.
fn scan(xml: &str, element: &[u8]) -> Vec<HashMap<String, String>> {
    let mut reader = XmlReader::from_str(xml);
    let mut out = Vec::new();
    let mut buf = Vec::new();

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == element {
                    out.push(
                        e.attributes()
                            .flatten()
                            .map(|attribute| {
                                (
                                    String::from_utf8_lossy(attribute.key.local_name().as_ref())
                                        .to_lowercase(),
                                    String::from_utf8_lossy(&attribute.value).into_owned(),
                                )
                            })
                            .collect(),
                    );
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::Path;

    fn sample_path() -> &'static str {
        concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../samples/sample-design.xlsx"
        )
    }

    fn sample(sheet: &str) -> Option<SheetLayout> {
        if !Path::new(sample_path()).exists() {
            eprintln!("skipping: {} not found", sample_path());
            return None;
        }
        Some(read_cached(&mut Cache::default(), sample_path(), sheet).expect("layout should parse"))
    }

    /// The print DPI comes from the DEVMODE Excel stores beside the sheet: a
    /// positive `dmPrintQuality`, or `dmYResolution` when that field holds a
    /// `DMRES_*` code instead.
    #[test]
    fn printer_dpi_comes_from_the_stored_devmode() {
        let mut devmode = vec![0u8; 220];
        devmode[90..92].copy_from_slice(&600i16.to_le_bytes());
        assert_eq!(devmode_dpi(&devmode), Some(600));

        devmode[90..92].copy_from_slice(&(-1i16).to_le_bytes());
        devmode[96..98].copy_from_slice(&1200i16.to_le_bytes());
        assert_eq!(devmode_dpi(&devmode), Some(1200));

        devmode[96..98].copy_from_slice(&0i16.to_le_bytes());
        assert_eq!(devmode_dpi(&devmode), None);
    }

    /// The style stored for one cell, or `None` when it carries none.
    fn style_at(layout: &SheetLayout, row: u32, col: u32) -> Option<&CellStyle> {
        layout.styles.get(style_index(layout, row, col)? as usize)
    }

    fn style_index(layout: &SheetLayout, row: u32, col: u32) -> Option<u32> {
        layout
            .cell_runs
            .get(&row)?
            .iter()
            .find(|run| col >= run[0] && col <= run[1])
            .map(|run| run[2])
    }

    #[test]
    fn converts_excel_units_to_pixels() {
        // Only the shape pass converts here now; the grid does it from the
        // font it measured. 8.43 is the familiar 64px stock column on Calibri.
        assert_eq!(chars_to_px(8.43), 64.0);
        assert_eq!(chars_to_px(8.0), 61.0);
        assert_eq!(chars_to_px(44.0), 313.0);
        assert_eq!(points_to_px(15.0), 20.0);
        assert_eq!(points_to_px(30.0), 40.0);
    }

    #[test]
    fn reads_explicit_column_widths_and_row_heights() {
        let Some(layout) = sample("Logic") else { return };
        // Column widths stay in the file's own unit — eight characters wide,
        // not the 61px that unit happens to be on Calibri.
        assert_eq!(layout.col_widths_chars.get(&0), Some(&8.0));
        assert_eq!(layout.col_widths_chars.get(&1), Some(&14.0));
        assert_eq!(layout.col_widths_chars.get(&2), Some(&44.0));
        // Header row was given a 30pt height.
        assert_eq!(layout.row_heights_px.get(&0), Some(&40.0));
        // A row nobody touched gets no entry at all.
        assert_eq!(layout.row_heights_px.get(&5), None);
    }

    #[test]
    fn undeclared_widths_fall_back_to_the_sheets_own_base_width() {
        let Some(layout) = sample("Logic") else { return };
        // The sheet writes baseColWidth="8" and never defaultColWidth, so
        // undeclared columns must be eight characters — matching Excel, not
        // something wider.
        assert_eq!(layout.default_col_width_chars, 8.0);
        assert_eq!(layout.default_row_height_px, 20.0);
    }

    #[test]
    fn stock_width_ranges_are_not_expanded_column_by_column() {
        let Some(layout) = sample("Summary") else { return };
        // Two explicit columns, not sixteen thousand.
        assert_eq!(layout.col_widths_chars.len(), 2);
        assert_eq!(layout.col_widths_chars.get(&0), Some(&16.0));
        assert_eq!(layout.col_widths_chars.get(&1), Some(&46.0));
    }

    #[test]
    fn reads_merged_regions() {
        let Some(layout) = sample("Summary") else { return };
        assert_eq!(layout.merges.len(), 1);
        let merge = &layout.merges[0];
        assert_eq!((merge.row, merge.col, merge.rows, merge.cols), (0, 0, 1, 2));
    }

    #[test]
    fn reads_the_autofilter_range_and_applied_values() {
        let Some(layout) = sample("Query") else { return };
        let filter = layout.auto_filter.expect("autofilter should be present");
        assert_eq!(
            (filter.row, filter.col, filter.rows, filter.cols),
            (0, 0, 401, 5)
        );
        // colId is relative to the filter range and must be rebased.
        assert_eq!(
            filter.applied.get(&4).map(Vec::as_slice),
            Some(&["要確認".to_owned()][..])
        );
    }

    #[test]
    fn reads_cell_alignment_and_border_together() {
        let Some(layout) = sample("Summary") else { return };
        let title = style_at(&layout, 0, 0).expect("title should be styled");
        assert_eq!(title.horizontal, Some("center"));
        assert_eq!(title.vertical, Some("middle"));

        let border = title.border.as_ref().expect("title should be bordered");
        for edge in [&border.top, &border.right, &border.bottom, &border.left] {
            let edge = edge.as_ref().expect("every side should be set");
            assert_eq!(edge.line, "solid");
            assert_eq!(edge.width, 1);
            assert_eq!(edge.rank, 4);
            assert_eq!(edge.color.as_deref(), Some("#000000"));
        }

        // A cell with no style at all is simply absent.
        assert!(style_at(&layout, 1, 0).is_none());
    }

    /// `indent` is what holds a right-aligned number the right distance off the
    /// border. The real screen definitions use it on dozens of styles, and
    /// skipping it left every one of them a whole unit too far right.
    #[test]
    fn reads_the_indent_level_of_an_alignment() {
        let mut reader = XmlReader::from_str(
            r#"<alignment horizontal="right" vertical="center" indent="1" shrinkToFit="1"/>"#,
        );
        let mut buf = Vec::new();
        let mut align = Align::default();
        loop {
            buf.clear();
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(element)) | Ok(Event::Empty(element)) => {
                    read_alignment(&element, &mut align);
                    break;
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
        assert_eq!(align.horizontal, Some("right"));
        assert_eq!(align.vertical, Some("middle"));
        assert_eq!(align.indent, 1);
        assert!(align.shrink);
    }

    /// `textRotation` carries two directions in one attribute: 0–90 turn one
    /// way, and 91–180 are the angles the dialog calls −1 to −90, written as
    /// 90 + degrees. 255 is not an angle at all — it stacks the characters one
    /// under the next. Seventeen of the corpus's files use one of the three.
    #[test]
    fn reads_the_angles_text_rotation_stores() {
        let read = |xml: &str| -> Align {
            let mut reader = XmlReader::from_str(xml);
            let mut buf = Vec::new();
            let mut align = Align::default();
            loop {
                buf.clear();
                match reader.read_event_into(&mut buf) {
                    Ok(Event::Start(element)) | Ok(Event::Empty(element)) => {
                        read_alignment(&element, &mut align);
                        break;
                    }
                    Ok(Event::Eof) | Err(_) => break,
                    _ => {}
                }
            }
            align
        };

        assert_eq!(read(r#"<alignment textRotation="90"/>"#).rotation, 90);
        assert_eq!(read(r#"<alignment textRotation="45"/>"#).rotation, 45);
        // The dialog's -90, which is what a column heading read downwards is.
        assert_eq!(read(r#"<alignment textRotation="180"/>"#).rotation, -90);
        assert_eq!(read(r#"<alignment textRotation="135"/>"#).rotation, -45);

        // 255 is the stacked form rather than an angle, and the cell that says
        // nothing keeps both off.
        let stacked = read(r#"<alignment textRotation="255"/>"#);
        assert!(stacked.stacked);
        assert_eq!(stacked.rotation, 0);
        assert_eq!(read(r#"<alignment wrapText="1"/>"#).rotation, 0);
    }

    /// Center Across Selection has to reach the grid as itself. Collapsing it
    /// into `center` centred each button label inside its own cell — a 14pt
    /// sliver — instead of across the box the run of formatted cells draws.
    #[test]
    fn keeps_center_across_selection_apart_from_center() {
        for (attribute, expected) in [
            (r#"horizontal="centerContinuous""#, "centerContinuous"),
            (r#"horizontal="center""#, "center"),
        ] {
            let xml = format!(r#"<alignment {attribute}/>"#);
            let mut reader = XmlReader::from_str(&xml);
            let mut buf = Vec::new();
            let mut align = Align::default();
            loop {
                buf.clear();
                match reader.read_event_into(&mut buf) {
                    Ok(Event::Start(element)) | Ok(Event::Empty(element)) => {
                        read_alignment(&element, &mut align);
                        break;
                    }
                    Ok(Event::Eof) | Err(_) => break,
                    _ => {}
                }
            }
            assert_eq!(align.horizontal, Some(expected));
        }
    }

    #[test]
    fn reads_the_workbooks_default_font() {
        let Some(layout) = sample("Summary") else { return };
        // openpyxl writes the stock Calibri 11, which is 14.67px — not 15.
        assert_eq!(layout.font_family.as_deref(), Some("Calibri"));
        assert_eq!(layout.font_size_px, 14.67);
    }

    #[test]
    fn points_to_font_px_keeps_the_fraction() {
        assert_eq!(points_to_font_px(11.0), 14.67);
        assert_eq!(points_to_font_px(9.0), 12.0);
        // Row heights still snap to whole pixels.
        assert_eq!(points_to_px(11.0), 15.0);
    }

    #[test]
    fn the_extent_covers_cells_that_only_carry_a_style() {
        let Some(layout) = sample("History") else { return };
        // Column D holds a border and no values anywhere. calamine's range stops
        // at column C, so relying on it dropped the column — and with it the
        // line closing off the table's right edge.
        assert_eq!(layout.cols, 4);
    }

    #[test]
    fn the_layout_extent_covers_the_sheets_rows_and_named_columns() {
        let Some(layout) = sample("Summary") else { return };
        // Six rows are written out, and the A1:B1 merge names two columns.
        assert_eq!((layout.rows, layout.cols), (6, 2));
    }

    #[test]
    fn reads_a_note_and_where_it_sits() {
        let xml = r#"<xml xmlns:v="v" xmlns:x="x">
            <v:shape id="_x0000_s1025"><v:textbox><div>hello</div></v:textbox>
              <x:ClientData ObjectType="Note"><x:Anchor>1, 2, 3, 4, 5, 6, 7, 8</x:Anchor>
                <x:Row>4</x:Row><x:Column>7</x:Column></x:ClientData></v:shape>
            <v:shape id="_x0000_s1026"><x:ClientData ObjectType="Button">
                <x:Row>1</x:Row><x:Column>1</x:Column></x:ClientData></v:shape>
        </xml>"#;
        let notes = parse_notes(xml);
        assert_eq!(notes.len(), 1, "a form control is not a note");
        assert_eq!((notes[0].row, notes[0].col), (4, 7));
        // The anchor's numbers pass through as the file wrote them.
        assert_eq!(notes[0].anchor, [1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0]);
    }

    #[test]
    fn reads_note_text_by_cell() {
        let xml = r#"<comments><authors><author>me</author></authors><commentList>
            <comment ref="B3" authorId="0"><text>
              <r><rPr><b/><sz val="9"/></rPr><t>hello</t></r><r><t xml:space="preserve"> world</t></r>
            </text></comment>
        </commentList></comments>"#;
        let map = parse_comments(xml);
        assert_eq!(map.len(), 1);
        let (text, font) = map.get(&(2, 1)).expect("B3 has a note");
        assert_eq!(text, "hello world");
        let font = font.as_ref().expect("the first run names a font");
        assert_eq!(font.size_px, Some(12.0), "9pt is 12px");
        assert!(font.bold);
    }

    #[test]
    fn reads_a_hatch_fill() {
        let xml = r#"<styleSheet>
            <fills count="5">
                <fill><patternFill patternType="none"/></fill>
                <fill><patternFill patternType="gray125"/></fill>
                <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor rgb="FF000000"/></patternFill></fill>
                <fill><patternFill patternType="lightUp"><fgColor rgb="FFFF0000"/><bgColor rgb="FFFFFF00"/></patternFill></fill>
                <fill><patternFill patternType="lightDown"><fgColor rgb="FF0000FF"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>
            </fills>
            <cellXfs count="5"><xf fillId="0"/><xf fillId="1"/><xf fillId="2"/><xf fillId="3"/><xf fillId="4"/></cellXfs>
        </styleSheet>"#;
        let styles = parse_styles(xml, Palette::EMPTY);
        assert!(styles.cell_style(0, 14.67).is_none());

        // A pattern with no colours is Excel's automatic pair: black on white.
        let gray = styles.cell_style(1, 14.67).unwrap().pattern.unwrap();
        assert_eq!(gray.kind, "gray125");
        assert_eq!(gray.fg, "#000000");
        assert_eq!(gray.bg, "#FFFFFF");

        // A solid fill is still a colour and never a pattern.
        let solid = styles.cell_style(2, 14.67).unwrap();
        assert_eq!(solid.fill.as_deref(), Some("#FFFF00"));
        assert!(solid.pattern.is_none());

        let up = styles.cell_style(3, 14.67).unwrap().pattern.unwrap();
        assert_eq!((up.kind, up.fg.as_str(), up.bg.as_str()), ("lightUp", "#FF0000", "#FFFF00"));
        let down = styles.cell_style(4, 14.67).unwrap().pattern.unwrap();
        assert_eq!(down.kind, "lightDown");
    }

    #[test]
    fn reads_a_diagonal_border() {
        let xml = r#"<styleSheet><borders count="3">
            <border><left/><right/><top/><bottom/><diagonal/></border>
            <border diagonalDown="1"><left/><right/><top/><bottom/><diagonal style="thick"><color rgb="FF0000FF"/></diagonal></border>
            <border diagonalUp="1" diagonalDown="1"><left/><right/><top/><bottom/><diagonal style="dashed"><color rgb="FFFF0000"/></diagonal></border>
        </borders><cellXfs count="3"><xf borderId="0"/><xf borderId="1"/><xf borderId="2"/></cellXfs></styleSheet>"#;
        let styles = parse_styles(xml, Palette::EMPTY);

        // The stock placeholder every border carries is not a rule.
        assert!(styles.cell_style(0, 14.67).is_none());

        let down = styles.cell_style(1, 14.67).unwrap().border.unwrap();
        assert!(down.diagonal_down);
        assert!(!down.diagonal_up);
        let edge = down.diagonal.expect("the diagonal style is kept");
        assert_eq!(edge.line, "solid");
        assert_eq!(edge.width, 3);
        assert_eq!(edge.color.as_deref(), Some("#0000FF"));

        let both = styles.cell_style(2, 14.67).unwrap().border.unwrap();
        assert!(both.diagonal_up && both.diagonal_down);
        assert_eq!(both.diagonal.unwrap().line, "dashed");
    }

    #[test]
    fn reads_superscript_and_subscript_fonts() {
        let xml = r#"<styleSheet>
            <fonts count="3">
                <font><sz val="11"/><name val="MS Gothic"/></font>
                <font><vertAlign val="superscript"/><sz val="11"/><name val="MS Gothic"/></font>
                <font><vertAlign val="subscript"/><sz val="11"/><name val="MS Gothic"/></font>
            </fonts>
            <cellXfs count="3"><xf fontId="0"/><xf fontId="1"/><xf fontId="2"/></cellXfs>
        </styleSheet>"#;
        let styles = parse_styles(xml, Palette::EMPTY);
        assert!(styles.cell_style(0, 14.67).is_none());
        assert_eq!(
            styles.cell_style(1, 14.67).unwrap().vert_align,
            Some("superscript")
        );
        assert_eq!(
            styles.cell_style(2, 14.67).unwrap().vert_align,
            Some("subscript")
        );

        // A run says the same thing in its own `<rPr>`.
        let xml = r#"<sst><si><r><t>X</t></r><r><rPr><vertAlign val="superscript"/><sz val="11"/></rPr><t>2</t></r></si></sst>"#;
        let runs = rich_strings(xml, Palette::EMPTY);
        let run = &runs[0].as_ref().unwrap()[1];
        assert_eq!(run.font.as_ref().unwrap().vert_align, Some("superscript"));
    }

    #[test]
    fn a_rows_style_counts_only_with_custom_format() {
        let xml = r#"<styleSheet>
            <fonts count="1"><font><sz val="11"/></font></fonts>
            <fills count="2">
                <fill><patternFill patternType="none"/></fill>
                <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
            </fills>
            <cellXfs count="2"><xf fontId="0" fillId="0"/><xf fontId="0" fillId="1"/></cellXfs>
        </styleSheet>"#;
        let styles = parse_styles(xml, Palette::EMPTY);
        let sheet = br#"<worksheet><sheetData>
            <row r="1" s="1"/>
            <row r="2" s="1" customFormat="1"/>
        </sheetData></worksheet>"#;
        let layout = parse_sheet(&sheet[..], &styles, &[]).expect("layout");
        // Row 1 names a style but `customFormat` is off, so Excel applies none.
        assert_eq!(layout.row_styles.get(&0), None);
        let style = layout.row_styles.get(&1).expect("row 2 is styled");
        assert_eq!(
            layout.styles[*style as usize].fill.as_deref(),
            Some("#FFFF00")
        );
    }

    #[test]
    fn reads_the_columns_a_custom_filter_hides() {
        let cache = Cache::default();
        let xml = br#"<worksheet><autoFilter ref="A1:C5"><filterColumn colId="2"><customFilters and="1"><customFilter operator="greaterThan" val="20"/></customFilters></filterColumn></autoFilter><sheetData/></worksheet>"#;
        let layout = parse_sheet(&xml[..], &cache.styles, &[]).expect("layout");
        let filter = layout.auto_filter.expect("autofilter should be present");
        // `colId` is relative to the range, so it is rebased like the values.
        assert_eq!(filter.custom, vec![2]);
        // A condition is not a value list: nothing goes in `applied`.
        assert!(filter.applied.is_empty());
    }

    #[test]
    fn reads_a_data_validation_rule() {
        let cache = Cache::default();
        // Excel's `showDropDown` is inverted: writing it hides the in-cell
        // dropdown, which is what `InCellDropdown = False` produces.
        let xml = br#"<worksheet><dataValidations><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" showDropDown="1" promptTitle="Pick" prompt="Choose" sqref="B1 C3:D4"><formula1>"a,b"</formula1></dataValidation></dataValidations><sheetData/></worksheet>"#;
        let layout = parse_sheet(&xml[..], &cache.styles, &[]).expect("layout");
        assert_eq!(layout.validations.len(), 1);
        let rule = &layout.validations[0];
        assert_eq!(rule.kind, "list");
        assert_eq!(rule.formula1.as_deref(), Some("\"a,b\""));
        assert!(rule.allow_blank);
        assert!(rule.hide_dropdown);
        assert!(rule.show_input);
        assert_eq!(rule.prompt_title.as_deref(), Some("Pick"));
        assert_eq!(rule.prompt.as_deref(), Some("Choose"));
        assert_eq!(rule.ranges.len(), 2);
        let block = &rule.ranges[1];
        assert_eq!((block.row, block.col, block.rows, block.cols), (2, 2, 2, 2));
    }

    #[test]
    fn reads_a_validation_list_escaped_as_entities() {
        let cache = Cache::default();
        // A file that escapes its characters delivers the list in pieces, with
        // an entity event between them. Keeping only the last text fragment
        // left a lone quote, and the in-cell dropdown then offered nothing.
        let xml = br#"<worksheet><dataValidations><dataValidation type="list" sqref="B1"><formula1>"&#23550;&#35937;,&#23550;&#35937;&#22806;"</formula1></dataValidation></dataValidations><sheetData/></worksheet>"#;
        let layout = parse_sheet(&xml[..], &cache.styles, &[]).expect("layout");
        assert_eq!(
            layout.validations[0].formula1.as_deref(),
            Some("\"対象,対象外\"")
        );
    }

    #[test]
    fn reads_the_rows_and_columns_a_sheet_view_freezes() {
        let cache = Cache::default();
        let xml = br#"<worksheet><sheetViews><sheetView><pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/></sheetView></sheetViews><sheetData/></worksheet>"#;
        let layout = parse_sheet(&xml[..], &cache.styles, &[]).expect("layout");
        assert_eq!((layout.frozen_rows, layout.frozen_cols), (2, 1));

        // A writer that leaves the splits out still names the scrolling
        // region's first cell; the count is what that cell's position implies.
        let xml = br#"<worksheet><sheetViews><sheetView><pane topLeftCell="B3" state="frozen"/></sheetView></sheetViews><sheetData/></worksheet>"#;
        let layout = parse_sheet(&xml[..], &cache.styles, &[]).expect("layout");
        assert_eq!((layout.frozen_rows, layout.frozen_cols), (2, 1));

        // A split pane is not a frozen one: nothing is held still.
        let xml = br#"<worksheet><sheetViews><sheetView><pane xSplit="1" ySplit="2" state="split"/></sheetView></sheetViews><sheetData/></worksheet>"#;
        let layout = parse_sheet(&xml[..], &cache.styles, &[]).expect("layout");
        assert_eq!((layout.frozen_rows, layout.frozen_cols), (0, 0));
    }

    /// Custom sheet views each keep a pane of their own, and they are not where
    /// the file opens: one workbook carried thirteen, and reading the last one
    /// froze seven rows and scrolled the sheet to its corner.
    #[test]
    fn a_custom_sheet_views_pane_is_not_the_view_the_sheet_opens_at() {
        let cache = Cache::default();
        let xml = br#"<worksheet><sheetViews><sheetView><pane xSplit="14" ySplit="4" topLeftCell="O5" state="frozen"/>
                <autoFilter ref="O4:ET69"><filterColumn colId="1" showButton="0"/></autoFilter>
            </sheetView></sheetViews>
            <customSheetViews><customSheetView guid="{1}" showAutoFilter="1">
                <pane xSplit="14" ySplit="7" topLeftCell="BW45" state="frozen"/>
                <autoFilter ref="A7:EV119"><filterColumn colId="0" showButton="0"/></autoFilter>
            </customSheetView></customSheetViews><sheetData/></worksheet>"#;
        let layout = parse_sheet(&xml[..], &cache.styles, &[]).expect("layout");
        assert_eq!((layout.frozen_rows, layout.frozen_cols), (4, 14));
        assert_eq!(layout.view_top_left, Some([4, 14]));
        // Its filter columns are relative to its own range, not this one's.
        let filter = layout.auto_filter.expect("the sheet view's filter");
        assert_eq!(filter.hidden_buttons, vec![15]);
    }

    #[test]
    fn reads_the_outline_levels_and_the_summary_side() {
        let cache = Cache::default();
        let xml = br#"<worksheet><sheetPr><outlinePr summaryBelow="0" summaryRight="0"/></sheetPr><cols>
            <col min="2" max="4" hidden="1" outlineLevel="1" width="0" customWidth="1"/>
            <col min="5" max="5" collapsed="1" width="9"/>
        </cols><sheetData>
            <row r="3" hidden="1" outlineLevel="2"></row>
            <row r="4" collapsed="1"></row>
        </sheetData></worksheet>"#;
        let layout = parse_sheet(&xml[..], &cache.styles, &[]).expect("layout");
        let outline = layout.outline.expect("outline should be present");
        assert!(!outline.summary_below);
        assert!(!outline.summary_right);
        assert_eq!(outline.rows.len(), 2);
        assert_eq!((outline.rows[0].row, outline.rows[0].level), (2, 2));
        assert!(!outline.rows[0].collapsed);
        assert_eq!(outline.rows[1].row, 3);
        assert!(outline.rows[1].collapsed);
        assert_eq!(outline.cols.len(), 2);
        assert_eq!((outline.cols[0].from, outline.cols[0].to), (1, 3));
        assert_eq!(outline.cols[0].level, 1);
        assert!(outline.cols[1].collapsed);

        // A sheet that groups nothing writes no outline at all.
        let plain = br#"<worksheet><sheetData><row r="1"></row></sheetData></worksheet>"#;
        let layout = parse_sheet(&plain[..], &cache.styles, &[]).expect("layout");
        assert!(layout.outline.is_none());
    }

    #[test]
    fn reads_the_gridline_visibility_setting() {
        let Some(shown) = sample("Summary") else { return };
        // No showGridLines attribute means Excel's default, which is on.
        assert!(shown.show_grid_lines);
        assert!(!sample("History").unwrap().show_grid_lines);
    }

    #[test]
    fn reads_a_cells_own_font_size() {
        let Some(layout) = sample("Summary") else { return };
        let title = style_at(&layout, 0, 0).expect("title should be styled");
        assert!(title.bold);
        // 22pt is 29.33px, against a 14.67px default — rendering it at the
        // default size is what made headings look nothing like Excel.
        assert_eq!(title.size_px, Some(29.33));
        assert_eq!(title.color.as_deref(), Some("#1F3864"));
        // Body cells keep the default and store no size at all.
        assert!(style_at(&layout, 2, 0).is_none());
    }

    #[test]
    fn reads_a_cells_background_fill() {
        let Some(layout) = sample("Summary") else { return };
        let title = style_at(&layout, 0, 0).expect("title should be styled");
        assert_eq!(title.fill.as_deref(), Some("#DEEBF7"));
        // A cell that declares no fill carries none.
        assert!(style_at(&layout, 2, 0).is_none());
    }

    /// Excel's styling runs one level deeper than cell → row → column: every
    /// `cellXfs` entry carries an `xfId` into `cellStyleXfs` and an `apply…`
    /// flag per property. A format that says `applyFont="0"` is asking for the
    /// named style's font and leaving its own `fontId` behind as a leftover —
    /// which is how a heading built on a named style came out as body text.
    ///
    /// Each property is decided on its own, and the two directions are told
    /// apart by their values: the named style is bold 22pt on red, and every
    /// entry below that declares its own carries the plain font and green.
    #[test]
    fn a_cell_format_borrows_the_properties_of_the_named_style_it_points_at() {
        let xml = r#"<styleSheet>
            <fonts count="2">
                <font><sz val="11"/><name val="Calibri"/></font>
                <font><b/><sz val="22"/><name val="Calibri"/></font>
            </fonts>
            <fills count="3">
                <fill><patternFill patternType="none"/></fill>
                <fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill>
                <fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill>
            </fills>
            <borders count="1"><border/></borders>
            <cellStyleXfs count="2">
                <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
                <xf numFmtId="0" fontId="1" fillId="1" borderId="0"/>
            </cellStyleXfs>
            <cellXfs count="5">
                <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
                <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="1" applyFont="0" applyFill="0"/>
                <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="1"/>
                <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="1" applyFont="0" applyFill="1"/>
                <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="1" applyFont="1" applyFill="0"/>
            </cellXfs>
        </styleSheet>"#;
        let styles = parse_styles(xml, Palette::EMPTY);
        let default_font_px = points_to_font_px(11.0);
        let style = |index: usize| {
            styles
                .cell_style(index, default_font_px)
                .expect("a styled entry")
        };

        // Both declined, so both come from the named style.
        let borrowed = style(1);
        assert!(borrowed.bold);
        assert_eq!(borrowed.size_px, Some(29.33));
        assert_eq!(borrowed.fill.as_deref(), Some("#FF0000"));

        // No flag at all: the entry's own ids stand. This is how a producer
        // that is not Excel writes its formats, and inheriting here would empty
        // most of a workbook's styling.
        let own = style(2);
        assert!(own.bold, "its own font is the 22pt one");
        assert_eq!(own.fill.as_deref(), Some("#00FF00"));

        // One of each on the same entry — the font is declined and the fill is
        // not — so the two must resolve in opposite directions.
        let mixed = style(3);
        assert!(mixed.bold, "the named style's bold font");
        assert_eq!(mixed.fill.as_deref(), Some("#00FF00"), "its own fill");

        let other_way = style(4);
        assert!(other_way.bold, "its own bold font");
        assert_eq!(other_way.fill.as_deref(), Some("#FF0000"), "the named fill");
    }

    /// `<i/>`, `<u/>` and `<strike/>` are the three font properties nothing
    /// read — 44, 42 and 23 of the 60 corpus workbooks use them, and every one
    /// of those cells drew plain. Presence is "on"; `val` is written only to
    /// turn one *off* or to pick the double underline.
    #[test]
    fn reads_the_font_properties_that_are_on_by_presence() {
        let xml = r#"<styleSheet>
            <fonts count="5">
                <font><sz val="11"/><name val="Calibri"/></font>
                <font><i/></font>
                <font><u/></font>
                <font><u val="double"/><strike/></font>
                <font><b val="0"/><i val="0"/><u val="none"/></font>
            </fonts>
            <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
            <borders count="1"><border/></borders>
            <cellXfs count="5">
                <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
                <xf numFmtId="0" fontId="1" fillId="0" borderId="0"/>
                <xf numFmtId="0" fontId="2" fillId="0" borderId="0"/>
                <xf numFmtId="0" fontId="3" fillId="0" borderId="0"/>
                <xf numFmtId="0" fontId="4" fillId="0" borderId="0"/>
            </cellXfs>
        </styleSheet>"#;
        let styles = parse_styles(xml, Palette::EMPTY);
        let default_font_px = points_to_font_px(11.0);
        let style = |index: usize| {
            styles
                .cell_style(index, default_font_px)
                .expect("a styled entry")
        };

        let italic = style(1);
        assert!(italic.italic);
        assert!(!italic.strike);
        assert_eq!(italic.underline, None);

        let underlined = style(2);
        assert_eq!(underlined.underline, Some("single"));
        assert!(!underlined.italic);

        // `val` picks the double line, and the two elements are independent.
        let doubled = style(3);
        assert_eq!(doubled.underline, Some("double"));
        assert!(doubled.strike);

        // Turning all three off leaves nothing to draw, which is what proves
        // the zeroes were read — an unread `<b val="0"/>` would leave it bold.
        assert!(styles.cell_style(4, default_font_px).is_none());
    }

    #[test]
    fn resolves_every_way_a_colour_can_be_written() {
        fn element(xml: &str) -> BytesStart<'static> {
            let mut reader = XmlReader::from_str(xml);
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf) {
                    Ok(Event::Start(e)) | Ok(Event::Empty(e)) => return e.into_owned(),
                    Ok(Event::Eof) => panic!("no element in {xml}"),
                    Ok(_) => {}
                    Err(e) => panic!("{e}"),
                }
            }
        }
        let theme = theme_slots(&HashMap::from([
            ("lt1".to_owned(), "#FFFFFF".to_owned()),
            ("dk1".to_owned(), "#000000".to_owned()),
            ("accent5".to_owned(), "#5B9BD5".to_owned()),
        ]));
        let palette = Palette {
            theme: &theme,
            indexed: &[],
        };

        // Literal.
        assert_eq!(
            color_of(&element(r#"<color rgb="FF808080"/>"#), palette).as_deref(),
            Some("#808080")
        );
        // Plain borders are written as "automatic", which Excel draws black.
        assert_eq!(
            color_of(&element(r#"<color auto="1"/>"#), palette).as_deref(),
            Some("#000000")
        );
        // The legacy palette; 12 is blue, 8 is black, 64 is automatic.
        assert_eq!(
            color_of(&element(r#"<color indexed="12"/>"#), palette).as_deref(),
            Some("#0000FF")
        );
        assert_eq!(
            color_of(&element(r#"<color indexed="8"/>"#), palette).as_deref(),
            Some("#000000")
        );
        // 31 is the pale blue these design documents fill their headings with,
        // and it is the one slot the palette is easy to get wrong: read from a
        // neighbouring entry it comes out grey and the fill looks plausible.
        assert_eq!(
            color_of(&element(r#"<color indexed="31"/>"#), palette).as_deref(),
            Some("#CCCCFF")
        );
        assert_eq!(
            color_of(&element(r#"<color indexed="64"/>"#), palette).as_deref(),
            Some("#000000")
        );
        // Theme slots are indexed from lt1, not from the order theme1.xml lists.
        assert_eq!(
            color_of(&element(r#"<color theme="1"/>"#), palette).as_deref(),
            Some("#000000")
        );
        assert_eq!(
            color_of(&element(r#"<color theme="0"/>"#), palette).as_deref(),
            Some("#FFFFFF")
        );
        // Nothing resolvable still falls back to the grid's own colour.
        assert_eq!(color_of(&element(r#"<color theme="9"/>"#), palette), None);
    }

    #[test]
    fn theme_system_colours_read_lastclr_not_the_system_name() {
        let scheme = parse_theme(
            r#"<a:theme><a:themeElements><a:clrScheme name="Office">
               <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
               <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
               <a:dk2><a:srgbClr val="44546A"/></a:dk2>
               </a:clrScheme></a:themeElements></a:theme>"#,
        );
        assert_eq!(scheme.get("dk1").map(String::as_str), Some("#000000"));
        // "window" is six characters long, which a length-only check on hex
        // would accept — and it then tinted to black instead of grey.
        assert_eq!(scheme.get("lt1").map(String::as_str), Some("#FFFFFF"));
        assert_eq!(scheme.get("dk2").map(String::as_str), Some("#44546A"));
    }

    #[test]
    fn colours_that_are_not_hex_are_rejected() {
        assert_eq!(parse_color("window"), None);
        assert_eq!(parse_color("windowText"), None);
        assert_eq!(parse_color(""), None);
        assert_eq!(parse_color("FF808080").as_deref(), Some("#808080"));
    }

    #[test]
    fn tint_lightens_and_darkens() {
        // Whites and blacks are the easy cases to reason about.
        assert_eq!(tint("#000000".to_owned(), Some("1".to_owned())), "#FFFFFF");
        assert_eq!(tint("#FFFFFF".to_owned(), Some("-1".to_owned())), "#000000");
        // A tint of zero changes nothing, and neither does a missing one.
        assert_eq!(tint("#5B9BD5".to_owned(), Some("0".to_owned())), "#5B9BD5");
        assert_eq!(tint("#5B9BD5".to_owned(), None), "#5B9BD5");
        // Lightening a blue keeps its hue and raises its luminance.
        let lighter = tint("#5B9BD5".to_owned(), Some("0.8".to_owned()));
        assert_eq!(lighter, "#DEEBF7");
    }

    #[test]
    fn ranks_borders_so_the_heavier_shared_edge_wins() {
        assert!(edge_style("thick").unwrap().rank > edge_style("thin").unwrap().rank);
        assert!(edge_style("medium").unwrap().rank > edge_style("dashed").unwrap().rank);
        assert!(edge_style("double").unwrap().rank > edge_style("thick").unwrap().rank);
        assert!(edge_style("none").is_none());
    }

    #[test]
    fn parses_argb_and_rgb_colours() {
        assert_eq!(parse_color("FF000000").as_deref(), Some("#000000"));
        assert_eq!(parse_color("00FF00").as_deref(), Some("#00FF00"));
        assert_eq!(parse_color("#ABCDEF").as_deref(), Some("#ABCDEF"));
        assert_eq!(parse_color("nope"), None);
    }

    #[test]
    fn plain_sheets_report_no_merges_or_filter() {
        let Some(layout) = sample("Logic") else { return };
        assert!(layout.merges.is_empty());
        assert!(layout.auto_filter.is_none());
    }

    /// The fallback is the sheet's used range — every cell it formats, with
    /// merged ranges folded in during the parse. On a cover sheet whose values
    /// stop at row 36, `Ctrl+End` lands on `AW44`, which is what put the page's
    /// edge back down on row 44.
    #[test]
    fn the_print_fallback_is_the_sheets_used_range() {
        let mut layout = SheetLayout::empty();
        // That cover sheet, to scale: formatted out to 44 rows and 49 columns.
        layout.rows = 44;
        layout.cols = 49;

        let area = layout.print_region().expect("a fallback is always there");
        assert_eq!((area.row, area.col, area.rows, area.cols), (0, 0, 44, 49));
    }

    /// The sheet a print area belongs to is its `localSheetId`, not the name
    /// its formula spells out. Every workbook of the corpus has the two
    /// agreeing, so this pins the reading that has to hold where they do not:
    /// a name rewritten without its formula, and a sheet whose name carries the
    /// `!` the formula would have to split on.
    #[test]
    fn reads_a_print_area_by_the_sheet_it_is_scoped_to() {
        let xml = r#"<workbook>
            <sheets>
                <sheet name="Cover" sheetId="1"/>
                <sheet name="Interface" sheetId="2"/>
                <sheet name="A!B" sheetId="3"/>
            </sheets>
            <definedNames>
                <definedName name="_xlnm.Print_Area" localSheetId="1">Cover!$A$1:$K$20</definedName>
                <definedName name="_xlnm.Print_Area" localSheetId="2">$B$2:$C$5</definedName>
                <definedName name="_xlnm.Print_Area">Cover!$D$4:$E$9</definedName>
                <definedName name="_xlnm.Print_Titles">Cover!$1:$1</definedName>
            </definedNames>
        </workbook>"#;
        let areas = print_area_names(xml);

        // The scope says Interface and the formula still says Cover: the scope wins.
        assert_eq!(areas.get("Interface").map(String::as_str), Some("$A$1:$K$20"));
        // A local name may leave the sheet out of its formula altogether.
        assert_eq!(areas.get("A!B").map(String::as_str), Some("$B$2:$C$5"));
        // A name with no scope says which sheet it is for in the formula.
        assert_eq!(areas.get("Cover").map(String::as_str), Some("$D$4:$E$9"));
        // Only print areas are worth keeping.
        assert_eq!(areas.len(), 3);
    }

    /// A list validation may name its range instead of spelling it out, and a
    /// name scoped to a sheet has to win over the workbook-wide one — the
    /// order Excel resolves a cell's formula in.
    #[test]
    fn a_sheet_sees_the_workbook_names_and_its_own() {
        let xml = r#"<workbook>
            <sheets>
                <sheet name="One" sheetId="1"/>
                <sheet name="Two" sheetId="2"/>
            </sheets>
            <definedNames>
                <definedName name="対象">LIST!$H$2:$H$23</definedName>
                <definedName name="対象" localSheetId="1">LIST!$I$2:$I$3</definedName>
                <definedName name="OnlyTwo" localSheetId="1">One!$A$1</definedName>
            </definedNames>
        </workbook>"#;

        let one = visible_names(xml, "One");
        // The workbook-wide name, until this sheet defines its own.
        assert_eq!(one.get("対象").map(String::as_str), Some("LIST!$H$2:$H$23"));
        assert!(!one.contains_key("OnlyTwo"));

        let two = visible_names(xml, "Two");
        assert_eq!(two.get("対象").map(String::as_str), Some("LIST!$I$2:$I$3"));
        assert_eq!(two.get("OnlyTwo").map(String::as_str), Some("One!$A$1"));
    }

    /// Excel prints the used range when a sheet defines no `_xlnm.Print_Area`,
    /// and shades around it just the same — so the sheets that never had one
    /// still grey out past their last cell instead of running white to the edge.
    #[test]
    fn without_a_print_area_the_used_range_stands_in() {
        let Some(layout) = sample("Summary") else { return };
        let area = layout.print_area.expect("the fallback should stand in");
        assert!(area.rows >= layout.rows && area.cols >= layout.cols);
        // Guards against a vacuous pass on a degenerate sheet.
        assert!(area.rows > 1 && area.cols > 1);
    }

    /// The paper, the margins and the breaks a sheet prints on all come from
    /// the file; Page Break Preview lays the pages out from them.
    #[test]
    fn reads_the_paper_and_the_breaks_a_sheet_prints_on() {
        let cache = Cache::default();
        let xml = br#"<worksheet>
            <sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
            <printOptions horizontalCentered="1" gridLines="1"/>
            <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            <pageSetup paperSize="9" orientation="landscape" scale="80" fitToWidth="1" fitToHeight="2" pageOrder="overThenDown"/>
            <rowBreaks count="1" manualBreakCount="1"><brk id="5" max="16383" man="1"/></rowBreaks>
            <colBreaks count="1" manualBreakCount="1"><brk id="3" max="1048575" man="1"/></colBreaks>
            <sheetData/>
        </worksheet>"#;
        let layout = parse_sheet(&xml[..], &cache.styles, &[]).expect("layout");
        let setup = &layout.page_setup;

        assert_eq!(setup.paper_size, 9);
        assert!(setup.landscape);
        assert_eq!(setup.scale, 80);
        assert!(setup.fit_to_page, "fitToPage lives on pageSetUpPr");
        assert_eq!((setup.fit_to_width, setup.fit_to_height), (1, 2));
        assert!(setup.page_order_over_then_down);
        assert_eq!(setup.margins.left, 0.7);
        assert_eq!(setup.margins.top, 0.75);
        assert!(setup.print_options.horizontal_centered);
        assert!(setup.print_options.grid_lines);
        assert!(!setup.print_options.headings);
        // `<brk id>` is the row or column the break falls after, counted from 1.
        assert_eq!(setup.row_breaks, vec![4]);
        assert_eq!(setup.col_breaks, vec![2]);
    }

    /// A stored break is only a page break if it was placed by hand; an
    /// automatic one is worked out from the paper and must not be replayed
    /// from a stale entry.
    #[test]
    fn a_break_nobody_placed_is_not_a_page_break() {
        let cache = Cache::default();
        let xml = br#"<worksheet><rowBreaks><brk id="12"/></rowBreaks><sheetData/></worksheet>"#;
        let layout = parse_sheet(&xml[..], &cache.styles, &[]).expect("layout");
        assert!(layout.page_setup.row_breaks.is_empty());
    }

    /// A sheet that turns fit-to-page on and says nothing else fits one page:
    /// ECMA-376 defaults `fitToWidth` and `fitToHeight` to one, and Excel
    /// leaves both attributes out when they are one.
    #[test]
    fn fit_to_page_alone_means_one_page() {
        let cache = Cache::default();
        let xml = br#"<worksheet><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><pageSetup paperSize="9" orientation="landscape"/><sheetData/></worksheet>"#;
        let layout = parse_sheet(&xml[..], &cache.styles, &[]).expect("layout");
        assert!(layout.page_setup.fit_to_page);
        assert_eq!(layout.page_setup.fit_to_width, 1);
        assert_eq!(layout.page_setup.fit_to_height, 1);
    }

    /// `_xlnm.Print_Titles` lists the repeating columns and rows in one
    /// formula — each range spelled either `$A:$C` or `$1:$2`, sheet-qualified.
    #[test]
    fn reads_the_rows_and_columns_every_page_repeats() {
        let xml = r#"<workbook>
            <sheets>
                <sheet name="Cover" sheetId="1"/>
                <sheet name="Body" sheetId="2"/>
            </sheets>
            <definedNames>
                <definedName name="_xlnm.Print_Titles" localSheetId="1">Body!$A:$C,Body!$1:$2</definedName>
            </definedNames>
        </workbook>"#;
        let titles = print_titles(xml, "Body").expect("titles");
        assert_eq!(titles.cols, Some([0, 2]));
        assert_eq!(titles.rows, Some([0, 1]));
        assert!(print_titles(xml, "Cover").is_none());
    }

    /// Only the strings that colour part of themselves keep their runs, the
    /// slots line up with the order the cells index them, and an escaped
    /// character inside one is put back so the runs spell the text the cell
    /// reports.
    #[test]
    fn keeps_the_runs_of_a_string_that_colours_a_word() {
        let xml = r#"<sst><si><t>plain</t></si><si><r><rPr><sz val="9"/></rPr><t>a</t></r><r><rPr><color rgb="FFFF0000"/></rPr><t>red</t></r><r><t>w&lt;&gt;ord</t></r></si></sst>"#;
        let runs = rich_strings(xml, Palette::EMPTY);
        assert_eq!(runs.len(), 2, "one slot per <si>, whether or not it is kept");
        assert!(runs[0].is_none(), "a string with no colour is not worth a table entry");
        let list = runs[1].as_ref().expect("the coloured string keeps its runs");
        assert_eq!(list.len(), 3);
        assert_eq!(list[0].color, None);
        assert_eq!(list[1].color.as_deref(), Some("#FF0000"));
        assert_eq!(list[2].color, None);
        let text: String = list.iter().map(|run| run.text.as_str()).collect();
        assert_eq!(text, "aredw<>ord");
    }

    /// A run's own `<rPr>` is read, and a string distinguished only by its font
    /// keeps its runs — the corpus writes far more of those than of coloured
    /// ones: `顧客管理 ▼` makes the marker bold and the breadcrumbs underline
    /// their links.
    #[test]
    fn reads_the_font_a_run_names() {
        let xml = r#"<sst>
            <si><t>plain</t></si>
            <si><r><t>a</t></r><r><rPr><b/><i/><u/><strike/><rFont val="MS Gothic"/><sz val="18"/></rPr><t>b</t></r></si>
            <si><r><rPr><u val="double"/></rPr><t>c</t></r></si>
            <si><r><rPr><b val="0"/><u val="none"/></rPr><t>d</t></r></si>
        </sst>"#;
        let runs = rich_strings(xml, Palette::EMPTY);
        assert_eq!(
            runs.iter().flatten().count(),
            3,
            "a string with no formatting is not worth a table entry"
        );

        let loaded = runs[1].as_ref().expect("the formatted string keeps its runs");
        assert_eq!(loaded[0].font, None, "a run with no rPr takes the cell's font");
        let font = loaded[1].font.as_ref().expect("the second run names a font");
        assert_eq!(font.family.as_deref(), Some("MS Gothic"));
        assert_eq!(font.size_px, Some(24.0), "18pt is 24px");
        assert_eq!(font.bold, Some(true));
        assert_eq!(font.italic, Some(true));
        assert_eq!(font.strike, Some(true));
        assert_eq!(font.underline, Some("single"));

        // An rPr that turns both off has to survive as a decision of its own:
        // dropped, the run would inherit a bold, underlined cell's formatting.
        let off = runs[3]
            .as_ref()
            .and_then(|list| list[0].font.clone())
            .expect("kept");
        assert_eq!(off.bold, Some(false));
        assert_eq!(off.underline, Some("none"));

        assert_eq!(
            runs[2]
                .as_ref()
                .and_then(|list| list[0].font.as_ref())
                .and_then(|f| f.underline),
            Some("double")
        );
    }

    /// The tabs a file keeps out of sight: the flag is `state`, and a sheet
    /// that says nothing is shown — which is what makes the attribute worth
    /// reading rather than assuming.
    #[test]
    fn reads_the_sheets_the_workbook_hides() {
        let xml = r#"<workbook><sheets>
            <sheet name="Cover" sheetId="1" r:id="rId1"/>
            <sheet name="DFD" sheetId="2" state="hidden" r:id="rId2"/>
            <sheet name="入力規則" sheetId="3" state="veryHidden" r:id="rId3"/>
            <sheet name="Interface" sheetId="4" state="visible" r:id="rId4"/>
        </sheets></workbook>"#;
        let hidden = hidden_in_workbook(xml);
        assert!(hidden.contains("DFD"));
        assert!(hidden.contains("入力規則"), "veryHidden is hidden too");
        assert!(!hidden.contains("Cover"), "no state at all means shown");
        assert!(!hidden.contains("Interface"), "and so does saying visible");
    }

    /// The style pass has to reach the end of the sheet.
    ///
    /// It used to stop at a ceiling on the number of styled cells it would keep,
    /// and every cell past that fell back to its row's default style. On a
    /// document that fills its sheet that is invisible through most of the
    /// table — the row default carries the same borders — and shows only where
    /// the row default has none, which is past the table's right edge. The rows
    /// recorded have to reach the sheet's own extent, not some fraction of it.
    #[test]
    #[ignore = "reads a workbook from the developer's machine"]
    fn every_row_of_a_long_sheet_keeps_its_own_style() {
        let Ok(path) = std::env::var("SHEETLENS_XLSX") else {
            return;
        };
        let sheet =
            std::env::var("SHEETLENS_SHEET").unwrap_or_else(|_| "Spec".to_string());
        let layout = read_cached(&mut Cache::default(), &path, &sheet).expect("layout");

        let deepest = layout.cell_runs.keys().max().map(|row| row + 1).unwrap_or(0);
        assert_eq!(
            deepest, layout.rows,
            "{sheet}: styles stop at row {deepest} of {}",
            layout.rows
        );
    }

    /// Dumps the coloured-run table a workbook yields, and — the part that
    /// matters — whether the cells really report the text those runs are keyed
    /// by. The renderer looks a cell's runs up by its own text, so a key no cell
    /// reports would leave the feature silently doing nothing.
    ///
    /// Ignored by default: it reads the workbook named by `SHEETLENS_XLSX`.
    #[test]
    #[ignore = "reads a workbook from the developer's machine"]
    fn dump_rich_text() {
        let Ok(path) = std::env::var("SHEETLENS_XLSX") else {
            return;
        };
        let book = crate::excel::load(&path).expect("open");
        let mut cache = Cache::default();
        for sheet in &book.info.sheets {
            let Ok(layout) = read_cached(&mut cache, &path, &sheet.name) else {
                continue;
            };
            if layout.rich_cells.is_empty() {
                continue;
            }
            println!("--- {} : {} rich cells", sheet.name, layout.rich_cells.len());

            let cells = book
                .cells(&sheet.name, 0, 0, sheet.rows.min(400), sheet.cols.min(64))
                .unwrap_or_default();
            let mut landed = 0;
            for cell in layout.rich_cells.iter().take(5) {
                let text = cells
                    .get(cell.row as usize)
                    .and_then(|row| row.get(cell.col as usize))
                    .map(|value| value.text.as_str())
                    .unwrap_or("");
                if !text.is_empty() {
                    landed += 1;
                }
                let joined: String = cell.runs.iter().map(|run| run.text.as_str()).collect();
                println!(
                    "    ({}, {}) {:?}… — {} runs, cell text {:?}…",
                    cell.row + 1,
                    cell.col + 1,
                    joined.chars().take(28).collect::<String>(),
                    cell.runs.len(),
                    text.chars().take(28).collect::<String>()
                );
                if joined != text {
                    println!(
                        "      runs and text differ: {} vs {} chars",
                        joined.chars().count(),
                        text.chars().count()
                    );
                }
                for run in &cell.runs {
                    let what: Vec<&str> = [
                        run.font.as_ref().and_then(|f| f.family.as_deref()),
                        run.font.as_ref().and_then(|f| f.underline),
                        run.font.as_ref().is_some_and(|f| f.bold == Some(true)).then_some("bold"),
                        run.font.as_ref().is_some_and(|f| f.italic == Some(true)).then_some("italic"),
                        run.font.as_ref().is_some_and(|f| f.strike == Some(true)).then_some("strike"),
                    ]
                    .into_iter()
                    .flatten()
                    .collect();
                    println!(
                        "        {:?} — colour {:?}, {:?}",
                        run.text.chars().take(20).collect::<String>(),
                        run.color.as_deref(),
                        what
                    );
                }
            }
            let plain = cells.iter().flatten().filter(|cell| cell.text != "").count();
            println!(
                "    {landed} of the first 5 rich cells sit on a non-empty cell; {plain} non-empty cells in the window"
            );
        }
    }

    #[test]
    fn parses_a1_references_with_anchors_and_single_cells() {
        let rect = parse_range("A1:C1").unwrap();
        assert_eq!((rect.row, rect.col, rect.rows, rect.cols), (0, 0, 1, 3));

        let single = parse_range("$B$7").unwrap();
        assert_eq!(
            (single.row, single.col, single.rows, single.cols),
            (6, 1, 1, 1)
        );

        let block = parse_range("B5:B8").unwrap();
        assert_eq!((block.row, block.col, block.rows, block.cols), (4, 1, 4, 1));

        assert!(parse_range("7B").is_none());
        assert!(parse_range("A0").is_none());
    }

    #[test]
    fn unknown_sheet_is_an_error_not_a_panic() {
        if !Path::new(sample_path()).exists() {
            return;
        }
        assert!(read_cached(&mut Cache::default(), sample_path(), "no such sheet").is_err());
    }
}
