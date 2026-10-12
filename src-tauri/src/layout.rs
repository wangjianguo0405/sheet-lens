//! Column widths, row heights, merges, autofilter and per-cell styling.
//!
//! calamine parses cell values only — its `Metadata` carries sheet names and
//! types, and `Range::width()` counts cells rather than pixels. All of this
//! lives in the workbook's XML, so we unzip it and read it directly.

use std::collections::HashMap;
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
/// None of the workbooks seen so far override it, so the stock table applies.
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
}

#[derive(Serialize, Clone, Default, Debug)]
#[serde(rename_all = "camelCase")]
pub struct CellStyle {
    pub horizontal: Option<&'static str>,
    pub vertical: Option<&'static str>,
    /// Indent levels; the renderer turns each into three spaces of the font.
    pub indent: u8,
    pub wrap: bool,
    pub bold: bool,
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
    /// Its number format, when one applies. Dates are stored as serial numbers
    /// and only this says how to show them.
    pub format: Option<String>,
    pub border: Option<CellBorder>,
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
            && self.bold == other.bold
            && self.size_px == other.size_px
            && self.family == other.family
            && self.color == other.color
            && self.fill == other.fill
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
        self.bold.hash(state);
        self.size_px.map(f64::to_bits).hash(state);
        self.family.hash(state);
        self.color.hash(state);
        self.fill.hash(state);
        self.border.hash(state);
    }
}

impl CellStyle {
    fn is_empty(&self) -> bool {
        self.horizontal.is_none()
            && self.vertical.is_none()
            && self.indent == 0
            && !self.wrap
            && !self.bold
            && self.size_px.is_none()
            && self.family.is_none()
            && self.color.is_none()
            && self.fill.is_none()
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
}

impl SheetLayout {
    /// What Excel prints when the sheet defines no `_xlnm.Print_Area`, and so
    /// also what it shades around.
    ///
    /// Excel's used range — what `Ctrl+End` jumps to — covers every cell the
    /// sheet *formats*, not only the ones holding a value. Checked on `表紙` of
    /// ESCUS0102, where the values stop at row 36 and a `Ctrl+End` lands on
    /// `AW44`: the shading and the page's edge both run to row 44, and the
    /// earlier reading that only value-bearing cells counted put that edge four
    /// rows too high. Merged ranges already extend the extent, which is what
    /// holds a merged banner open to the sheet's last column.
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
        let entry = self.xfs.get(index)?;
        let font = self.fonts.get(entry.font_id);
        let size_px = font
            .and_then(|font| font.size_pt)
            .map(points_to_font_px)
            .filter(|px| (px - default_font_px).abs() > 0.01);
        let default_family = self.fonts.first().and_then(|font| font.name.as_deref());
        let family = font
            .and_then(|font| font.name.clone())
            .filter(|name| Some(name.as_str()) != default_family);
        let style = CellStyle {
            horizontal: entry.align.horizontal,
            vertical: entry.align.vertical,
            indent: entry.align.indent,
            wrap: entry.align.wrap,
            bold: font.is_some_and(|font| font.bold),
            size_px,
            family,
            color: font.and_then(|font| font.color.clone()),
            fill: self.fills.get(entry.fill_id).and_then(|f| f.clone()),
            format: self
                .num_formats
                .get(&entry.num_fmt_id)
                .cloned()
                .or_else(|| builtin_date_format(entry.num_fmt_id).map(str::to_owned)),
            border: self.borders.get(entry.border_id).and_then(|b| b.clone()),
        };
        (!style.is_empty()).then_some(style)
    }
}

/// Stores `style` once and returns its index, so the payload carries each
/// distinct style a single time.
fn intern(
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
    /// The box around the cells that hold *something* — a `<v>`, or an `<is>`
    /// for inline text — and the merged regions, which count to Excel even with
    /// no value in them.
    ///
    /// Narrower than the extent above, and it is what the grid draws the file's
    /// formatting inside: a row's own style is written for all 16,384 columns of
    /// that row and a column's for all 1,048,576 rows, so a table's rules and
    /// fills otherwise run on over the formatted-but-empty cells a document
    /// keeps beside it — a ladder of borders out in what reads as blank sheet.
    /// Excel paints the same box and no more.
    pub content_rows: u32,
    pub content_cols: u32,
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
    /// Column index -> width in characters, for columns with an explicit width.
    pub col_widths_chars: HashMap<u32, f64>,
    /// Row index -> height in pixels, for rows with an explicit height.
    pub row_heights_px: HashMap<u32, f64>,
    pub merges: Vec<MergeRect>,
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
    /// Cells whose text is written as several runs with colours of their own,
    /// keyed by the string they spell out — which is what a cell holds, so the
    /// renderer can find a cell's runs without the sheet recording anything per
    /// cell. Sparse by nature: only strings that really carry a colour appear.
    pub rich_text: HashMap<String, Vec<RichRun>>,
}

/// One coloured run of a cell's rich text.
///
/// A cell shows the concatenation of its runs, so the text is nothing new; what
/// a run adds is a colour for part of it, which is how a document makes one word
/// inside a sentence red.
#[derive(Serialize, Clone)]
pub struct RichRun {
    pub text: String,
    /// `#rrggbb`, or `None` when the run takes the cell's own colour.
    pub color: Option<String>,
}

impl SheetLayout {
    fn empty() -> Self {
        Self {
            rows: 0,
            cols: 0,
            content_rows: 0,
            content_cols: 0,
            default_col_width_chars: STOCK_COL_WIDTH_CHARS,
            default_row_height_px: points_to_px(STOCK_ROW_HEIGHT_PT),
            font_family: None,
            font_size_px: points_to_font_px(STOCK_FONT_PT),
            show_grid_lines: true,
            col_widths_chars: HashMap::new(),
            row_heights_px: HashMap::new(),
            merges: Vec::new(),
            auto_filter: None,
            print_area: None,
            styles: Vec::new(),
            cell_runs: HashMap::new(),
            row_styles: HashMap::new(),
            col_styles: Vec::new(),
            drawings: Vec::new(),
            drawing_rel: None,
            rich_text: HashMap::new(),
        }
    }
}

/// One `<font>` entry from the styles table.
#[derive(Clone, Default)]
struct FontInfo {
    name: Option<String>,
    size_pt: Option<f64>,
    bold: bool,
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
}

/// Everything this viewer needs out of `xl/styles.xml`.
#[derive(Default)]
struct Styles {
    xfs: Vec<Xf>,
    borders: Vec<Option<CellBorder>>,
    fonts: Vec<FontInfo>,
    /// Background colour per `<fill>` entry; `None` for patterns and gradients.
    fills: Vec<Option<String>>,
    /// Custom `numFmtId` -> format code. Built-in ids are looked up separately.
    num_formats: HashMap<u32, String>,
}

/// Workbook-wide parse results. The styles and theme tables are shared by every
/// sheet, and re-reading them per switch cost tens — once over 100 — milliseconds.
#[derive(Default)]
pub struct Cache {
    loaded: bool,
    theme: Vec<Option<String>>,
    /// The theme's own line widths, which a shape's `lnRef idx` selects from.
    ln_widths: Vec<f64>,
    styles: Styles,
    /// Shared strings split into coloured runs, keyed by the text they spell
    /// out. Nearly every string in every sheet is shared, so this is read once
    /// per workbook rather than once per sheet.
    rich_strings: HashMap<String, Vec<RichRun>>,
}

impl Cache {
    pub fn load(&mut self, path: &str) -> Result<(), String> {
        if self.loaded {
            return Ok(());
        }
        let file = File::open(path).map_err(|e| e.to_string())?;
        let mut zip = ZipArchive::new(BufReader::new(file)).map_err(|e| e.to_string())?;
        self.theme = theme_slots(&read_theme(&mut zip));
        self.ln_widths = read_line_widths(&mut zip);
        self.styles = read_styles(&mut zip, &self.theme);
        // Absent in a workbook that writes its strings inline, which is why
        // this is a lookup on failure rather than an error.
        if let Ok(xml) = read_entry(&mut zip, "xl/sharedStrings.xml") {
            self.rich_strings = rich_strings(&xml, &self.theme);
        }
        self.loaded = true;
        Ok(())
    }
}

/// Reads one sheet's layout, reusing the workbook-wide tables in `cache`.
pub fn read_cached(cache: &mut Cache, path: &str, sheet: &str) -> Result<SheetLayout, String> {
    cache.load(path)?;
    let file = File::open(path).map_err(|e| e.to_string())?;
    let mut zip = ZipArchive::new(BufReader::new(file)).map_err(|e| e.to_string())?;
    let workbook = read_entry(&mut zip, "xl/workbook.xml")?;
    let rels = read_entry(&mut zip, "xl/_rels/workbook.xml.rels")?;
    let part = sheet_part(&workbook, &rels, sheet)?;
    let entry = zip.by_name(&part).map_err(|e| e.to_string())?;
    let mut layout = parse_sheet(BufReader::new(entry), &cache.styles)?;
    // A sheet with no `_xlnm.Print_Area` still gets shaded by Excel, which falls
    // back to the range it would print. Leaving this `None` painted those sheets
    // white to the edge while every other sheet in the same workbook greyed out
    // past its last cell.
    layout.print_area = print_area(&workbook, sheet).or_else(|| layout.print_region());
    // Workbook-wide, so every sheet carries the same table and the renderer
    // simply looks up whatever text it is about to draw.
    layout.rich_text = cache.rich_strings.clone();
    read_drawings(&mut zip, cache, &part, &mut layout);
    Ok(layout)
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
    let mut shapes = drawing::parse(&xml, &grid, &cache.theme, &cache.ln_widths);

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

/// The shared strings that are split into coloured runs, keyed by the text they
/// spell out.
///
/// A cell shows the concatenation of its runs either way, so the text is nothing
/// new; what a run adds is a colour for part of it — how one word inside a
/// sentence comes out red. Keying by the whole string is what lets the renderer
/// find a cell's runs without the sheet having to record anything per cell, and
/// it costs nothing: the runs already spell the key out.
///
/// Only strings whose runs really carry a colour are kept. A run that merely
/// restates the cell's own colour draws exactly like the plain text, and keeping
/// those would put most of the workbook in the table.
fn rich_strings(xml: &str, theme: &[Option<String>]) -> HashMap<String, Vec<RichRun>> {
    let mut out = HashMap::new();
    let mut reader = XmlReader::from_str(xml);
    let mut buf = Vec::new();
    let mut runs: Option<Vec<RichRun>> = None;
    let mut run: Option<RichRun> = None;
    let mut in_text = false;

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element)) => match element.local_name().as_ref() {
                b"si" => runs = Some(Vec::new()),
                b"r" => {
                    run = Some(RichRun {
                        text: String::new(),
                        color: None,
                    });
                }
                b"t" => in_text = true,
                b"color" => {
                    if let (Some(run), Some(color)) = (run.as_mut(), color_of(&element, theme)) {
                        run.color = Some(color);
                    }
                }
                _ => {}
            },
            // `<color …/>` is how a run names its colour, and an empty element
            // never gets a start/end pair.
            Ok(Event::Empty(element)) if element.local_name().as_ref() == b"color" => {
                if let (Some(run), Some(color)) = (run.as_mut(), color_of(&element, theme)) {
                    run.color = Some(color);
                }
            }
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
                b"r" => {
                    if let (Some(list), Some(run)) = (runs.as_mut(), run.take()) {
                        list.push(run);
                    }
                }
                b"si" => {
                    if let Some(list) = runs.take() {
                        if list.iter().any(|run| run.color.is_some()) {
                            let text: String = list.iter().map(|run| run.text.as_str()).collect();
                            out.insert(text, list);
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

/// The sheet's `_xlnm.Print_Area`, if it has one.
///
/// The range is the element's *text*, not an attribute, so this walks events
/// instead of reusing `scan`.
fn print_area(workbook: &str, sheet_name: &str) -> Option<MergeRect> {
    let mut reader = XmlReader::from_str(workbook);
    let mut buf = Vec::new();
    let mut inside = false;

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element)) if element.local_name().as_ref() == b"definedName" => {
                inside = attr(&element, b"name").as_deref() == Some("_xlnm.Print_Area");
            }
            Ok(Event::Text(text)) if inside => {
                inside = false;
                let content = text.xml_content(XmlVersion::default()).ok()?;
                // `<sheet>!<range>`, the sheet name quoted only when it must be.
                let (owner, range) = content.split_once('!')?;
                if owner.trim_matches('\'') == sheet_name {
                    return parse_range(range);
                }
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"definedName" => {
                inside = false;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    None
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

    let mut reader = XmlReader::from_str(&xml);
    let mut buf = Vec::new();
    let mut styles = Styles::default();

    let mut in_borders = false;
    let mut in_fonts = false;
    let mut in_fills = false;
    let mut in_num_fmts = false;
    let mut in_cell_xfs = false;
    let mut border: Option<CellBorder> = None;
    let mut edge: Option<Side> = None;
    let mut font: Option<FontInfo> = None;
    let mut fill: Option<String> = None;
    let mut fill_is_solid = false;
    let mut xf: Option<Xf> = None;

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element)) => match element.local_name().as_ref() {
                b"borders" => in_borders = true,
                b"border" if in_borders => border = Some(CellBorder::default()),
                b"left" | b"right" | b"top" | b"bottom" if in_borders => {
                    edge = set_edge(&element, border.as_mut(), &element.local_name().as_ref());
                }
                b"color" if in_borders => color_edge(&element, border.as_mut(), edge, theme),
                // A solid `<patternFill>` paints its `fgColor`; gradients and
                // the stock none/gray125 entries have no flat colour to show.
                b"fills" => in_fills = true,
                b"fill" if in_fills => {
                    fill = None;
                    fill_is_solid = false;
                }
                b"patternFill" if in_fills => {
                    fill_is_solid = attr(&element, b"patternType").as_deref() == Some("solid");
                }
                b"fgColor" if in_fills => {
                    if fill_is_solid {
                        fill = color_of(&element, theme);
                    }
                }
                b"fonts" => in_fonts = true,
                b"numFmts" => in_num_fmts = true,
                b"numFmt" if in_num_fmts => record_num_format(&element, &mut styles),
                b"font" if in_fonts => font = Some(FontInfo::default()),
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
                        slot.color = color_of(&element, theme);
                    }
                }
                b"cellXfs" => in_cell_xfs = true,
                b"xf" if in_cell_xfs => xf = Some(new_xf(&element)),
                b"alignment" if in_cell_xfs => {
                    if let Some(slot) = xf.as_mut() {
                        read_alignment(&element, &mut slot.align);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(element)) => match element.local_name().as_ref() {
                // A self-closing element has no colour child to wait for.
                b"border" if in_borders => styles.borders.push(None),
                b"left" | b"right" | b"top" | b"bottom" if in_borders => {
                    set_edge(&element, border.as_mut(), &element.local_name().as_ref());
                }
                // Excel writes the colour self-closing, e.g. `<color rgb="FF000000"/>`.
                b"color" if in_borders => color_edge(&element, border.as_mut(), edge, theme),
                b"fill" if in_fills => styles.fills.push(None),
                b"patternFill" if in_fills => {
                    fill_is_solid = attr(&element, b"patternType").as_deref() == Some("solid");
                }
                b"fgColor" if in_fills => {
                    if fill_is_solid {
                        fill = color_of(&element, theme);
                    }
                }
                b"font" if in_fonts => styles.fonts.push(FontInfo::default()),
                b"numFmt" if in_num_fmts => record_num_format(&element, &mut styles),
                b"b" if in_fonts => {
                    if let Some(slot) = font.as_mut() {
                        slot.bold = !matches!(attr(&element, b"val").as_deref(), Some("0" | "false"));
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
                        slot.color = color_of(&element, theme);
                    }
                }
                b"xf" if in_cell_xfs => styles.xfs.push(new_xf(&element)),
                b"alignment" if in_cell_xfs => {
                    if let Some(slot) = xf.as_mut() {
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
                })),
                b"left" | b"right" | b"top" | b"bottom" if in_borders => edge = None,
                b"fonts" => in_fonts = false,
                b"numFmts" => in_num_fmts = false,
                b"fills" => in_fills = false,
                b"fill" if in_fills => styles.fills.push(fill.take()),
                b"font" if in_fonts => styles.fonts.push(font.take().unwrap_or_default()),
                b"cellXfs" => in_cell_xfs = false,
                b"xf" if in_cell_xfs => {
                    if let Some(slot) = xf.take() {
                        styles.xfs.push(slot);
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    styles
}

#[derive(Clone, Copy, PartialEq)]
enum Side {
    Top,
    Right,
    Bottom,
    Left,
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
fn set_edge(element: &BytesStart, border: Option<&mut CellBorder>, name: &[u8]) -> Option<Side> {
    let side = match name {
        b"top" => Side::Top,
        b"right" => Side::Right,
        b"bottom" => Side::Bottom,
        b"left" => Side::Left,
        _ => return None,
    };
    let border = border?;
    let edge = attr(element, b"style").and_then(|name| edge_style(&name));
    let slot = match side {
        Side::Top => &mut border.top,
        Side::Right => &mut border.right,
        Side::Bottom => &mut border.bottom,
        Side::Left => &mut border.left,
    };
    *slot = edge;
    Some(side)
}

fn color_edge(
    element: &BytesStart,
    border: Option<&mut CellBorder>,
    side: Option<Side>,
    theme: &[Option<String>],
) {
    let (Some(border), Some(side)) = (border, side) else {
        return;
    };
    let Some(color) = color_of(element, theme) else {
        return;
    };
    let slot = match side {
        Side::Top => &mut border.top,
        Side::Right => &mut border.right,
        Side::Bottom => &mut border.bottom,
        Side::Left => &mut border.left,
    };
    if let Some(edge) = slot.as_mut() {
        edge.color = Some(color);
    }
}

/// Excel's border names reduced to a CSS style, a width and a precedence rank.
/// The rank follows Excel's own ordering so the heavier of two shared edges wins.
fn edge_style(name: &str) -> Option<BorderEdge> {
    let (line, width, rank) = match name {
        "hair" => ("solid", 1, 1),
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
fn color_of(element: &BytesStart, theme: &[Option<String>]) -> Option<String> {
    if let Some(rgb) = attr(element, b"rgb").as_deref().and_then(parse_color) {
        return Some(tint(rgb, attr(element, b"tint")));
    }
    if attr(element, b"auto").is_some_and(|value| value == "1" || value == "true") {
        return Some(BLACK.to_owned());
    }
    if let Some(index) = attr(element, b"theme").and_then(|v| v.parse::<usize>().ok()) {
        return match theme.get(index).and_then(|slot| slot.clone()) {
            Some(base) => Some(tint(base, attr(element, b"tint"))),
            // Without a parsed theme, slot 1 is the text colour and 0 the page.
            None => match index {
                1 => Some(BLACK.to_owned()),
                0 => Some(WHITE.to_owned()),
                _ => None,
            },
        };
    }
    if let Some(index) = attr(element, b"indexed").and_then(|v| v.parse::<usize>().ok()) {
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

fn parse_sheet<R: BufRead>(reader: R, styles: &Styles) -> Result<SheetLayout, String> {
    let mut xml = XmlReader::from_reader(reader);
    let mut layout = SheetLayout::empty();
    let mut buf = Vec::new();
    let mut filter_col: Option<u32> = None;
    let mut declared_default_col: Option<f64> = None;
    let mut base_col_width: Option<f64> = None;
    let mut declared_default_row: Option<f64> = None;
    let mut extent_rows: u32 = 0;
    let mut extent_cols: u32 = 0;
    // Where the last cell that held something ended, and the cell the walk is
    // inside of: its children are what say whether it holds anything.
    let mut content_rows: u32 = 0;
    let mut content_cols: u32 = 0;
    let mut cell_at: Option<(u32, u32)> = None;
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
                    b"sheetView" => {
                        // Absent means the default, which is on.
                        if let Some(value) = attr(&element, b"showGridLines") {
                            layout.show_grid_lines = value != "0" && value != "false";
                        }
                    }
                    b"col" => {
                        let min = attr(&element, b"min")
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(1);
                        let max = attr(&element, b"max")
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(min);
                        if min == 0 || max < min {
                            continue;
                        }
                        let last = max.min(min.saturating_add(MAX_EXPANDED_COLUMNS - 1));

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
                        // Rows can carry a style for the cells in them that have
                        // none of their own — how a filled row reads as solid.
                        if let Some(index) = attr(&element, b"s")
                            .and_then(|v| v.parse::<usize>().ok())
                        {
                            if let Some(style) = styles.cell_style(index, default_font_px) {
                                let style = intern(&mut layout.styles, &mut style_ids, style);
                                layout.row_styles.insert(row - 1, style);
                            }
                        }
                        if let Some(height) =
                            attr(&element, b"ht").and_then(|v| v.parse::<f64>().ok())
                        {
                            layout.row_heights_px.insert(row - 1, points_to_px(height));
                        }
                        extent_rows = extent_rows.max(row);
                    }
                    b"c" => {
                        // Cell display values come from calamine; this pass only
                        // recovers the style, and only when it says something.
                        let Some((row, col)) =
                            attr(&element, b"r").and_then(|r| parse_cell(&r))
                        else {
                            continue;
                        };
                        // Where a cell is, whether or not it turns out to have a
                        // style: its children decide whether it holds anything,
                        // and that box is what the grid draws formatting inside.
                        cell_at = Some((row, col));
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
                        extent_rows = extent_rows.max(row + 1);
                        extent_cols = extent_cols.max(col + 1);
                    }
                    // A cell's children are its value: a `<v>` for a stored one,
                    // an `<is>` for text written inline. Either is what makes the
                    // cell content, and so part of the box the formatting is
                    // drawn inside.
                    b"v" | b"is" => {
                        if let Some((row, col)) = cell_at {
                            content_rows = content_rows.max(row + 1);
                            content_cols = content_cols.max(col + 1);
                        }
                    }
                    b"mergeCell" => {
                        if layout.merges.len() >= MAX_MERGES {
                            continue;
                        }
                        if let Some(reference) = attr(&element, b"ref") {
                            if let Some(rect) = parse_range(&reference) {
                                extent_rows = extent_rows.max(rect.row + rect.rows);
                                extent_cols = extent_cols.max(rect.col + rect.cols);
                                // A merge is content to Excel even with nothing
                                // in it — it is what holds a banner open.
                                content_rows = content_rows.max(rect.row + rect.rows);
                                content_cols = content_cols.max(rect.col + rect.cols);
                                layout.merges.push(rect);
                            }
                        }
                    }
                    b"autoFilter" => {
                        if let Some(rect) =
                            attr(&element, b"ref").and_then(|r| parse_range(&r))
                        {
                            layout.auto_filter = Some(AutoFilter {
                                row: rect.row,
                                col: rect.col,
                                rows: rect.rows,
                                cols: rect.cols,
                                applied: HashMap::new(),
                            });
                        }
                    }
                    // The sheet only points at its shapes; the part itself is read
                    // once the sheet's own extent is known.
                    b"drawing" => layout.drawing_rel = attr(&element, b"id"),
                    b"filterColumn" => {
                        filter_col = attr(&element, b"colId").and_then(|v| v.parse::<u32>().ok());
                    }                    b"filter" => {
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
                    _ => {}
                }
            }
            Ok(Event::End(element)) => {
                match element.local_name().as_ref() {
                    b"filterColumn" => filter_col = None,
                    b"c" => cell_at = None,
                    _ => {}
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
    layout.content_rows = content_rows.min(MAX_SHEET_ROWS);
    layout.content_cols = content_cols.min(MAX_SHEET_COLS);

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
fn parse_range(reference: &str) -> Option<MergeRect> {
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
fn parse_cell(reference: &str) -> Option<(u32, u32)> {
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
        let Some(layout) = sample("ロジック") else { return };
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
        let Some(layout) = sample("ロジック") else { return };
        // The sheet writes baseColWidth="8" and never defaultColWidth, so
        // undeclared columns must be eight characters — matching Excel, not
        // something wider.
        assert_eq!(layout.default_col_width_chars, 8.0);
        assert_eq!(layout.default_row_height_px, 20.0);
    }

    #[test]
    fn stock_width_ranges_are_not_expanded_column_by_column() {
        let Some(layout) = sample("基本情報") else { return };
        // Two explicit columns, not sixteen thousand.
        assert_eq!(layout.col_widths_chars.len(), 2);
        assert_eq!(layout.col_widths_chars.get(&0), Some(&16.0));
        assert_eq!(layout.col_widths_chars.get(&1), Some(&46.0));
    }

    #[test]
    fn reads_merged_regions() {
        let Some(layout) = sample("基本情報") else { return };
        assert_eq!(layout.merges.len(), 1);
        let merge = &layout.merges[0];
        assert_eq!((merge.row, merge.col, merge.rows, merge.cols), (0, 0, 1, 2));
    }

    #[test]
    fn reads_the_autofilter_range_and_applied_values() {
        let Some(layout) = sample("DB抽出仕様") else { return };
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
        let Some(layout) = sample("基本情報") else { return };
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
            r#"<alignment horizontal="right" vertical="center" indent="1"/>"#,
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
        let Some(layout) = sample("基本情報") else { return };
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
        let Some(layout) = sample("変更履歴") else { return };
        // Column D holds a border and no values anywhere. calamine's range stops
        // at column C, so relying on it dropped the column — and with it the
        // line closing off the table's right edge.
        assert_eq!(layout.cols, 4);
    }

    #[test]
    fn the_layout_extent_covers_the_sheets_rows_and_named_columns() {
        let Some(layout) = sample("基本情報") else { return };
        // Six rows are written out, and the A1:B1 merge names two columns.
        assert_eq!((layout.rows, layout.cols), (6, 2));
    }

    #[test]
    fn reads_the_gridline_visibility_setting() {
        let Some(shown) = sample("基本情報") else { return };
        // No showGridLines attribute means Excel's default, which is on.
        assert!(shown.show_grid_lines);
        assert!(!sample("変更履歴").unwrap().show_grid_lines);
    }

    #[test]
    fn reads_a_cells_own_font_size() {
        let Some(layout) = sample("基本情報") else { return };
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
        let Some(layout) = sample("基本情報") else { return };
        let title = style_at(&layout, 0, 0).expect("title should be styled");
        assert_eq!(title.fill.as_deref(), Some("#DEEBF7"));
        // A cell that declares no fill carries none.
        assert!(style_at(&layout, 2, 0).is_none());
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

        // Literal.
        assert_eq!(
            color_of(&element(r#"<color rgb="FF808080"/>"#), &theme).as_deref(),
            Some("#808080")
        );
        // Plain borders are written as "automatic", which Excel draws black.
        assert_eq!(
            color_of(&element(r#"<color auto="1"/>"#), &theme).as_deref(),
            Some("#000000")
        );
        // The legacy palette; 12 is blue, 8 is black, 64 is automatic.
        assert_eq!(
            color_of(&element(r#"<color indexed="12"/>"#), &theme).as_deref(),
            Some("#0000FF")
        );
        assert_eq!(
            color_of(&element(r#"<color indexed="8"/>"#), &theme).as_deref(),
            Some("#000000")
        );
        // 31 is the pale blue these design documents fill their headings with,
        // and it is the one slot the palette is easy to get wrong: read from a
        // neighbouring entry it comes out grey and the fill looks plausible.
        assert_eq!(
            color_of(&element(r#"<color indexed="31"/>"#), &theme).as_deref(),
            Some("#CCCCFF")
        );
        assert_eq!(
            color_of(&element(r#"<color indexed="64"/>"#), &theme).as_deref(),
            Some("#000000")
        );
        // Theme slots are indexed from lt1, not from the order theme1.xml lists.
        assert_eq!(
            color_of(&element(r#"<color theme="1"/>"#), &theme).as_deref(),
            Some("#000000")
        );
        assert_eq!(
            color_of(&element(r#"<color theme="0"/>"#), &theme).as_deref(),
            Some("#FFFFFF")
        );
        // Nothing resolvable still falls back to the grid's own colour.
        assert_eq!(color_of(&element(r#"<color theme="9"/>"#), &theme), None);
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
        let Some(layout) = sample("ロジック") else { return };
        assert!(layout.merges.is_empty());
        assert!(layout.auto_filter.is_none());
    }

    /// The fallback is the sheet's used range — every cell it formats, with
    /// merged ranges folded in during the parse. `Ctrl+End` on `表紙` of
    /// ESCUS0102 lands on `AW44` while its values stop at row 36, which is what
    /// put the page's edge back down on row 44.
    #[test]
    fn the_print_fallback_is_the_sheets_used_range() {
        let mut layout = SheetLayout::empty();
        // `表紙` of ESCUS0102, to scale: formatted out to 44 rows and 49 columns.
        layout.rows = 44;
        layout.cols = 49;

        let area = layout.print_region().expect("a fallback is always there");
        assert_eq!((area.row, area.col, area.rows, area.cols), (0, 0, 44, 49));
    }

    /// Excel prints the used range when a sheet defines no `_xlnm.Print_Area`,
    /// and shades around it just the same — so the sheets that never had one
    /// still grey out past their last cell instead of running white to the edge.
    #[test]
    fn without_a_print_area_the_used_range_stands_in() {
        let Some(layout) = sample("基本情報") else { return };
        let area = layout.print_area.expect("the fallback should stand in");
        assert!(area.rows >= layout.rows && area.cols >= layout.cols);
        // Guards against a vacuous pass on a degenerate sheet.
        assert!(area.rows > 1 && area.cols > 1);
    }

    /// Only the strings that colour part of themselves are kept, they are keyed
    /// by the text they spell out — which is what the cell holds, so the renderer
    /// can find a cell's runs without the sheet recording anything — and an
    /// escaped character inside one is put back, since the key has to match the
    /// text the cell reports.
    #[test]
    fn keeps_the_runs_of_a_string_that_colours_a_word() {
        let xml = r#"<sst><si><t>plain</t></si><si><r><rPr><sz val="9"/></rPr><t>a</t></r><r><rPr><color rgb="FFFF0000"/></rPr><t>red</t></r><r><t>w&lt;&gt;ord</t></r></si></sst>"#;
        let runs = rich_strings(xml, &[]);
        assert_eq!(runs.len(), 1, "a string with no colour is not worth a table entry");
        let list = runs.get("aredw<>ord").expect("keyed by the whole text");
        assert_eq!(list.len(), 3);
        assert_eq!(list[0].color, None);
        assert_eq!(list[1].color.as_deref(), Some("#FF0000"));
        assert_eq!(list[2].color, None);
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
            std::env::var("SHEETLENS_SHEET").unwrap_or_else(|_| "機能定義書".to_string());
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
            if layout.rich_text.is_empty() {
                continue;
            }
            println!("--- {} : {} rich strings", sheet.name, layout.rich_text.len());
            for (text, runs) in layout.rich_text.iter().take(5) {
                let head: String = text.chars().take(28).collect();
                let colours: Vec<&str> = runs.iter().filter_map(|run| run.color.as_deref()).collect();
                println!(
                    "    {:?}… — {} runs, {} chars, colours {:?}",
                    head,
                    runs.len(),
                    text.chars().count(),
                    colours
                );
            }

            let cells = book
                .cells(&sheet.name, 0, 0, sheet.rows.min(400), sheet.cols.min(64))
                .unwrap_or_default();
            let matched = cells
                .iter()
                .flatten()
                .filter(|cell| layout.rich_text.contains_key(&cell.text))
                .count();
            let plain = cells.iter().flatten().filter(|cell| cell.text != "").count();
            println!("    {matched} of {plain} non-empty cells found their runs");
            // What the nearest cell text looks like, when nothing matched.
            let longest = cells
                .iter()
                .flatten()
                .filter(|cell| !cell.text.is_empty())
                .max_by_key(|cell| cell.text.chars().count());
            if let Some(cell) = longest {
                let head: String = cell.text.chars().take(28).collect();
                let tail: String = cell.text.chars().skip(cell.text.chars().count().saturating_sub(24)).collect();
                println!(
                    "    longest cell text: {:?}… ({} chars) tail {:?}",
                    head,
                    cell.text.chars().count(),
                    tail
                );
                if let Some((key, _)) = layout
                    .rich_text
                    .iter()
                    .max_by_key(|(text, _)| text.chars().count())
                {
                    let tail: String = key.chars().skip(key.chars().count().saturating_sub(24)).collect();
                    println!("    longest key tail: {:?}", tail);
                    let at = cell
                        .text
                        .chars()
                        .zip(key.chars())
                        .position(|(a, b)| a != b);
                    let around = |text: &str| -> String {
                        text.chars().skip(at.unwrap_or(0).saturating_sub(8)).take(40).collect()
                    };
                    println!(
                        "    first difference at char {:?}; cell {} chars, key {} chars",
                        at,
                        cell.text.chars().count(),
                        key.chars().count()
                    );
                    println!("      cell: {:?}", around(&cell.text));
                    println!("      key : {:?}", around(key));
                }
            }
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
