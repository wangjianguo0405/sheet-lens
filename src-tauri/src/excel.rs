use std::collections::{HashMap, HashSet};
use std::fmt::Write as _;
use std::path::Path;

use calamine::{open_workbook_auto, Data, ExcelDateTime, Range, Reader};
use serde::Serialize;

use crate::layout::{self, SheetLayout};
use crate::numfmt;

const MAX_HITS: usize = 5000;
const HIT_PREVIEW_CHARS: usize = 200;

#[derive(Serialize, Clone)]
pub struct Cell {
    /// "empty" | "text" | "number" | "bool" | "date" | "error"
    pub kind: &'static str,
    pub text: String,
}

#[derive(Serialize, Clone)]
pub struct SheetInfo {
    pub name: String,
    pub rows: u32,
    pub cols: u32,
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct WorkbookInfo {
    pub path: String,
    pub file_name: String,
    pub sheets: Vec<SheetInfo>,
}

#[derive(Serialize)]
pub struct Hit {
    pub sheet: String,
    pub row: u32,
    pub col: u32,
    pub text: String,
}

#[derive(Serialize)]
pub struct ValueCount {
    pub value: String,
    pub count: u32,
}

struct Sheet {
    range: Range<Data>,
    /// Top-left cell of the stored range, in absolute sheet coordinates.
    origin: (u32, u32),
}

pub struct Book {
    pub info: WorkbookInfo,
    sheets: HashMap<String, Sheet>,
    /// Parsed lazily — layout needs a second pass over the zip, and most opens
    /// only ever look at one sheet.
    layouts: HashMap<String, SheetLayout>,
    /// Theme and style tables, shared by every sheet in the workbook.
    layout_cache: layout::Cache,
}

pub fn load(path: &str) -> Result<Book, String> {
    let mut workbook = open_workbook_auto(path).map_err(|e| e.to_string())?;
    let names = workbook.sheet_names();

    let mut sheets = HashMap::with_capacity(names.len());
    let mut infos = Vec::with_capacity(names.len());

    for name in names {
        let range = workbook
            .worksheet_range(&name)
            .map_err(|e| format!("{name}: {e}"))?;
        let (rows, cols) = range.get_size();
        let origin = range.start().unwrap_or((0, 0));
        infos.push(SheetInfo {
            name: name.clone(),
            rows: rows as u32,
            cols: cols as u32,
        });
        sheets.insert(name, Sheet { range, origin });
    }

    let file_name = Path::new(path)
        .file_name()
        .map(|n| n.to_string_lossy().into_owned())
        .unwrap_or_else(|| path.to_owned());

    Ok(Book {
        info: WorkbookInfo {
            path: path.to_owned(),
            file_name,
            sheets: infos,
        },
        sheets,
        layouts: HashMap::new(),
        layout_cache: layout::Cache::default(),
    })
}

impl Book {
    pub fn layout(&mut self, sheet: &str) -> Result<SheetLayout, String> {
        if let Some(cached) = self.layouts.get(sheet) {
            return Ok(cached.clone());
        }
        let path = self.info.path.clone();
        let parsed = layout::read_cached(&mut self.layout_cache, &path, sheet)?;
        self.layouts.insert(sheet.to_owned(), parsed.clone());
        Ok(parsed)
    }

    pub fn has_layout(&self, sheet: &str) -> bool {
        self.layouts.contains_key(sheet)
    }

    /// Used by the background warmer, which parses with its own cache and then
    /// hands the finished result over.
    pub fn put_layout(&mut self, sheet: String, layout: SheetLayout) {
        self.layouts.insert(sheet, layout);
    }
}

impl Book {
    /// A window of cells, addressed by zero-based offset from the sheet's top-left.
    pub fn cells(
        &self,
        sheet: &str,
        row: u32,
        col: u32,
        rows: u32,
        cols: u32,
    ) -> Result<Vec<Vec<Cell>>, String> {
        let height_width = self
            .sheets
            .get(sheet)
            .ok_or_else(|| format!("sheet not found: {sheet}"))?;
        let (height, width) = height_width.range.get_size();
        let (origin_row, origin_col) = height_width.origin;

        // The grid asks for a block by its absolute position on the sheet and
        // reads the reply back at that same position, so the window has to be
        // the requested one laid over the used range rather than an offset
        // into it. Adding the origin here shifted every value by the distance
        // from A1 to the first used cell — invisible on a sheet that starts at
        // A1, but a document whose content begins at B2 came out a row up and
        // a column left of its own styling.
        let row_end = row
            .saturating_add(rows)
            .min(origin_row.saturating_add(height as u32));
        let col_end = col
            .saturating_add(cols)
            .min(origin_col.saturating_add(width as u32));

        // The number format lives in the layout, which the grid loads before
        // it asks for any cells.
        let layout = self.layouts.get(sheet);

        let mut out = Vec::with_capacity(row_end.saturating_sub(row) as usize);
        for r in row..row_end {
            let mut line = Vec::with_capacity(col_end.saturating_sub(col) as usize);
            for c in col..col_end {
                let value = height_width.range.get_value((r, c));
                let formatted = value
                    .zip(layout.and_then(|found| found.format_at(r, c)))
                    .and_then(|(value, code)| numfmt::apply(value, code));
                line.push(match formatted {
                    Some(text) => Cell { kind: "date", text },
                    None => cell_at(value),
                });
            }
            out.push(line);
        }
        Ok(out)
    }

    /// Distinct display values in one column, most common first.
    pub fn column_values(
        &self,
        sheet: &str,
        col: u32,
        first_row: u32,
        last_row: u32,
        query: Option<&str>,
        limit: usize,
    ) -> Result<Vec<ValueCount>, String> {
        let data = self
            .sheets
            .get(sheet)
            .ok_or_else(|| format!("sheet not found: {sheet}"))?;
        let (height, width) = data.range.get_size();
        let (origin_row, origin_col) = data.origin;
        // Absolute coordinates, as the grid names them — see `cells`.
        if col < origin_col || col >= origin_col.saturating_add(width as u32) {
            return Ok(Vec::new());
        }
        let start = first_row.max(origin_row);
        let end = last_row
            .saturating_add(1)
            .min(origin_row.saturating_add(height as u32));
        let needle = query.map(str::to_lowercase);

        let mut counts: HashMap<String, u32> = HashMap::new();
        let mut text = String::new();
        for r in start..end {
            write_text(data.range.get_value((r, col)), &mut text);
            if let Some(needle) = &needle {
                if !text.to_lowercase().contains(needle) {
                    continue;
                }
            }
            *counts.entry(text.clone()).or_insert(0) += 1;
        }

        let mut out: Vec<ValueCount> = counts
            .into_iter()
            .map(|(value, count)| ValueCount { value, count })
            .collect();
        // Frequency first so the useful values survive the truncation, then
        // alphabetical so the list does not reshuffle between calls.
        out.sort_by(|a, b| b.count.cmp(&a.count).then_with(|| a.value.cmp(&b.value)));
        out.truncate(limit);
        Ok(out)
    }

    /// Rows where every criterion matches on its column's display text.
    /// An empty criterion list matches everything.
    pub fn filter_rows(
        &self,
        sheet: &str,
        first_row: u32,
        last_row: u32,
        criteria: &[(u32, Vec<String>)],
        limit: usize,
    ) -> Result<Vec<u32>, String> {
        let data = self
            .sheets
            .get(sheet)
            .ok_or_else(|| format!("sheet not found: {sheet}"))?;
        let (height, _) = data.range.get_size();
        let (origin_row, _) = data.origin;
        // Absolute coordinates, as the grid names them — see `cells`.
        let start = first_row.max(origin_row);
        let end = last_row
            .saturating_add(1)
            .min(origin_row.saturating_add(height as u32));

        // Excel's filter does not distinguish case, so neither do we.
        let wanted: Vec<(u32, HashSet<String>)> = criteria
            .iter()
            .map(|(col, values)| {
                (
                    *col,
                    values.iter().map(|value| value.to_lowercase()).collect(),
                )
            })
            .collect();

        let mut out = Vec::new();
        let mut text = String::new();
        'rows: for r in start..end {
            for (col, accepted) in &wanted {
                write_text(data.range.get_value((r, *col)), &mut text);
                if !accepted.contains(&text.to_lowercase()) {
                    continue 'rows;
                }
            }
            out.push(r);
            if out.len() >= limit {
                break;
            }
        }
        Ok(out)
    }

    /// Display text along one axis, for copying a whole row or column. Set
    /// exactly one of `row`/`col`; `from`/`to` are the other axis' bounds.
    pub fn axis_values(
        &self,
        sheet: &str,
        row: Option<u32>,
        col: Option<u32>,
        from: u32,
        to: u32,
        limit: usize,
    ) -> Result<Vec<String>, String> {
        let data = self
            .sheets
            .get(sheet)
            .ok_or_else(|| format!("sheet not found: {sheet}"))?;
        let last = to.min(from.saturating_add(limit.saturating_sub(1) as u32));

        let mut out = Vec::new();
        let mut text = String::new();
        for index in from..=last {
            let (r, c) = match (row, col) {
                (Some(r), _) => (r, index),
                (_, Some(c)) => (index, c),
                _ => return Err("give either a row or a column".to_owned()),
            };
            // Absolute coordinates, as the grid names them — see `cells`.
            // calamine answers `None` past the used range, which reads back as
            // the blank Excel copies alongside the cells that are there.
            write_text(data.range.get_value((r, c)), &mut text);
            out.push(text.clone());
        }

        // Excel does not copy the trailing blanks out to the sheet edge.
        while out.last().is_some_and(|value| value.is_empty()) {
            out.pop();
        }
        Ok(out)
    }

    pub fn search(&self, sheet: Option<&str>, matcher: &Matcher) -> Vec<Hit> {        let mut hits = Vec::new();
        let mut text = String::new();

        for info in &self.info.sheets {
            if sheet.is_some_and(|only| only != info.name) {
                continue;
            }
            let Some(data) = self.sheets.get(&info.name) else {
                continue;
            };
            let (height, width) = data.range.get_size();
            let (origin_row, origin_col) = data.origin;

            // Walked in absolute coordinates so the positions in the hits are
            // the ones the grid jumps to — see `cells`.
            for r in origin_row..origin_row + height as u32 {
                for c in origin_col..origin_col + width as u32 {
                    write_text(data.range.get_value((r, c)), &mut text);
                    if text.is_empty() || !matcher.is_match(&text) {
                        continue;
                    }
                    hits.push(Hit {
                        sheet: info.name.clone(),
                        row: r,
                        col: c,
                        text: preview(&text),
                    });
                    if hits.len() >= MAX_HITS {
                        return hits;
                    }
                }
            }
        }
        hits
    }
}

pub struct Matcher {
    needle: String,
    case_sensitive: bool,
    whole_cell: bool,
}

impl Matcher {
    pub fn new(query: &str, case_sensitive: bool, whole_cell: bool) -> Self {
        Self {
            needle: if case_sensitive {
                query.to_owned()
            } else {
                query.to_lowercase()
            },
            case_sensitive,
            whole_cell,
        }
    }

    fn is_match(&self, text: &str) -> bool {
        if self.needle.is_empty() {
            return false;
        }
        if self.case_sensitive {
            if self.whole_cell {
                text == self.needle
            } else {
                text.contains(&self.needle)
            }
        } else {
            let text = text.to_lowercase();
            if self.whole_cell {
                text == self.needle
            } else {
                text.contains(&self.needle)
            }
        }
    }
}

fn cell_at(value: Option<&Data>) -> Cell {
    let kind = match value {
        None | Some(Data::Empty) => "empty",
        Some(Data::String(_)) => "text",
        Some(Data::Int(_) | Data::Float(_)) => "number",
        Some(Data::Bool(_)) => "bool",
        Some(Data::DateTime(_) | Data::DateTimeIso(_) | Data::DurationIso(_)) => "date",
        Some(Data::Error(_)) => "error",
    };
    let mut text = String::new();
    write_text(value, &mut text);
    Cell { kind, text }
}

fn write_text(value: Option<&Data>, out: &mut String) {
    out.clear();
    match value {
        None | Some(Data::Empty) => {}
        Some(Data::String(s)) => out.push_str(s),
        Some(Data::Int(i)) => {
            let _ = write!(out, "{i}");
        }
        Some(Data::Float(f)) => write_float(*f, out),
        Some(Data::Bool(b)) => out.push_str(if *b { "TRUE" } else { "FALSE" }),
        Some(Data::DateTime(dt)) => write_datetime(dt, out),
        Some(Data::DateTimeIso(s) | Data::DurationIso(s)) => out.push_str(s),
        Some(Data::Error(e)) => {
            let _ = write!(out, "{e:?}");
        }
    }
}

fn write_float(f: f64, out: &mut String) {
    if f.fract() == 0.0 && f.abs() < 1e15 {
        let _ = write!(out, "{}", f as i64);
    } else {
        let _ = write!(out, "{f}");
    }
}

fn write_datetime(dt: &ExcelDateTime, out: &mut String) {
    let Some(datetime) = dt.as_datetime() else {
        let _ = write!(out, "{dt}");
        return;
    };
    let date = datetime.date();
    let time = datetime.time();
    if format!("{time}") == "00:00:00" {
        let _ = write!(out, "{date}");
    } else {
        let _ = write!(out, "{date} {time}");
    }
}

fn preview(text: &str) -> String {
    if text.chars().count() <= HIT_PREVIEW_CHARS {
        return text.to_owned();
    }
    let mut out: String = text.chars().take(HIT_PREVIEW_CHARS).collect();
    out.push('…');
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample() -> Option<Book> {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../samples/sample-design.xlsx"
        );
        if !Path::new(path).exists() {
            eprintln!("skipping: {path} not found");
            return None;
        }
        Some(load(path).expect("sample workbook should load"))
    }

    #[test]
    #[ignore = "reads a workbook from the developer's machine"]
    fn dump_real_workbook() {
        let path = r"D:\svn\fbs\Marusan\branches\202610_現物継承サービス案件\01doc\01.設計書\01.基本設計\04.営業支援システム\03.画面定義\ESCON0101.契約詳細_画面定義書.xlsx";
        let started = std::time::Instant::now();
        let mut book = load(path).expect("workbook should load");
        println!("open_workbook total: {:?}", started.elapsed());

        {
            use std::io::Read;
            let file = std::fs::File::open(path).expect("open");
            let mut zip =
                zip::read::ZipArchive::new(std::io::BufReader::new(file)).expect("zip");
            let mut entry = zip.by_name("xl/worksheets/sheet6.xml").expect("entry");
            println!(
                "raw entry: declared {} bytes, compression {:?}",
                entry.size(),
                entry.compression()
            );
            let mut buf = Vec::new();
            let read = entry.read_to_end(&mut buf).expect("read");
            println!("raw read_to_end: {read} bytes");
        }
        let names: Vec<String> = book
            .info
            .sheets
            .iter()
            .map(|sheet| sheet.name.clone())
            .collect();
        println!("sheets ({}): {:?}", names.len(), names);

        // What calamine thinks the used range is, per sheet.
        for sheet in book.info.sheets.iter().take(3) {
            println!("size: {} = {} rows x {} cols", sheet.name, sheet.rows, sheet.cols);
        }

        let first = names.first().cloned().expect("at least one sheet");
        let layout = book.layout(&first).expect("layout should parse");

        // How long a sheet switch has to wait before it can draw anything.
        for name in names.iter().take(6) {
            let started = std::time::Instant::now();
            let _ = book.layout(name).expect("layout");
            println!("layout {name}: {:?}", started.elapsed());
        }
        let started = std::time::Instant::now();
        let _ = book.cells(&first, 0, 0, 200, 20).expect("cells");
        println!("first cells ({first}): {:?}", started.elapsed());
        println!("--- {first} ---");
        println!("extent      : {} rows x {} cols", layout.rows, layout.cols);
        println!("default col : {} chars", layout.default_col_width_chars);
        println!("default row : {}px", layout.default_row_height_px);
        println!("font        : {:?} @ {}px", layout.font_family, layout.font_size_px);
        println!("cols        : {:?}", layout.col_widths_chars);
        println!(
            "rows 0..6   : {:?}",
            (0..6)
                .map(|row| layout.row_heights_px.get(&row).copied())
                .collect::<Vec<_>>()
        );
        println!("merges      : {}", layout.merges.len());
        println!(
            "styles      : {} distinct over {} runs",
            layout.styles.len(),
            layout.cell_runs.values().map(Vec::len).sum::<usize>()
        );
        let mut fills: Vec<String> = layout
            .styles
            .iter()
            .filter_map(|style| style.fill.clone())
            .collect();
        fills.sort();
        fills.dedup();
        println!("fills       : {fills:?}");
        let mut border_colors: Vec<String> = layout
            .styles
            .iter()
            .filter_map(|style| style.border.as_ref())
            .flat_map(|border| [&border.top, &border.right, &border.bottom, &border.left])
            .filter_map(|edge| edge.as_ref())
            .filter_map(|edge| edge.color.clone())
            .collect();
        border_colors.sort();
        border_colors.dedup();
        println!("border inks : {border_colors:?}");
        for name in names.iter().take(14) {
            let Ok(parsed) = book.layout(name) else { continue };
            let mut fills: Vec<String> = parsed
                .styles
                .iter()
                .filter_map(|style| style.fill.clone())
                .collect();
            fills.sort();
            fills.dedup();
            println!("fills {name}: {fills:?}");
            println!(
                "   payload {name}: {} KB ({} runs, {} row styles, {} col runs)",
                serde_json::to_string(&parsed).map(|j| j.len() / 1024).unwrap_or(0),
                parsed.cell_runs.values().map(Vec::len).sum::<usize>(),
                parsed.row_styles.len(),
                parsed.col_styles.len(),
            );
        }
        for style in layout.styles.iter().take(4) {
            println!("  style -> {style:?}");
        }
    }

    /// Where the time goes between clicking a sheet and seeing it.
    #[test]
    #[ignore]
    fn profile_switch() {
        let Ok(path) = std::env::var("SHEETLENS_XLSX") else {
            return;
        };
        let sheet = std::env::var("SHEETLENS_SHEET").unwrap_or_else(|_| "Sheet1".into());

        let started = std::time::Instant::now();
        let mut book = load(&path).expect("workbook should load");
        println!("open            : {:?}", started.elapsed());

        let started = std::time::Instant::now();
        let layout = book.layout(&sheet).expect("layout");
        println!(
            "layout, cold    : {:?}   payload {} KB",
            started.elapsed(),
            serde_json::to_string(&layout).map(|j| j.len() / 1024).unwrap_or(0)
        );

        let started = std::time::Instant::now();
        let _ = book.layout(&sheet).expect("layout");
        println!("layout, cached  : {:?}", started.elapsed());

        for (row, col) in [(0u32, 0u32), (128, 0), (256, 0)] {
            let started = std::time::Instant::now();
            let block = book.cells(&sheet, row, col, 128, 32).expect("cells");
            let built = started.elapsed();
            let started = std::time::Instant::now();
            let json = serde_json::to_string(&block).expect("json");
            let serialise = started.elapsed();
            let filled = block
                .iter()
                .flatten()
                .filter(|cell| !cell.text.is_empty())
                .count();
            println!(
                "cells r{row:<3} c{col} : {built:?} build, {serialise:?} json, {} lines, {filled} filled, payload {} KB",
                block.len(),
                json.len() / 1024
            );
        }

        // How much of a viewport-sized window draws nothing at all — no text,
        // no fill and no border — which is the share the grid could leave out
        // instead of committing a div for every one.
        let window = book.cells(&sheet, 0, 0, 60, 70).expect("window");
        let mut blank = 0usize;
        let mut total = 0usize;
        for (row, line) in window.iter().enumerate() {
            for (col, cell) in line.iter().enumerate() {
                total += 1;
                let style = layout
                    .style_index(row as u32, col as u32)
                    .and_then(|index| layout.styles.get(index as usize));
                let draws = !cell.text.is_empty()
                    || style.is_some_and(|s| s.fill.is_some() || s.border.is_some());
                if !draws {
                    blank += 1;
                }
            }
        }
        println!(
            "blank cells in a 60x70 window: {blank} of {total} ({:.0}% skippable)",
            100.0 * blank as f64 / (total.max(1)) as f64
        );
    }

    #[test]
    fn lists_every_sheet_with_its_size() {
        let Some(book) = sample() else { return };
        let names: Vec<_> = book.info.sheets.iter().map(|s| s.name.as_str()).collect();
        assert_eq!(names, ["基本情報", "DB抽出仕様", "ロジック", "変更履歴"]);

        let logic = book.info.sheets.iter().find(|s| s.name == "ロジック").unwrap();
        assert_eq!((logic.rows, logic.cols), (12001, 4));
    }

    #[test]
    fn reads_cells_by_offset_and_keeps_numbers_typed() {
        let Some(book) = sample() else { return };

        // Row 0 is the merged title, so the header row is row 1.
        let title = book.cells("基本情報", 0, 0, 1, 1).unwrap();
        assert_eq!(title[0][0].text, "基本情報");

        let head = book.cells("基本情報", 1, 0, 2, 2).unwrap();
        assert_eq!(head[0][0].text, "項目");
        assert_eq!(head[0][1].text, "内容");
        assert_eq!(head[1][1].text, "ゴールベースチャート作成");

        // Row 0 holds the header, so row 3 is the record whose "行" is 3.
        let logic = book.cells("ロジック", 3, 0, 1, 4).unwrap();
        assert_eq!(logic[0][0].kind, "number");
        assert_eq!(logic[0][0].text, "3");
        assert_eq!(logic[0][1].text, "Step 3");
        assert_eq!(logic[0][2].text, "処理説明 3");
        assert_eq!(logic[0][3].kind, "number");
        assert_eq!(logic[0][3].text, "4.5");
    }

    /// A one-sheet book whose content starts at B2 rather than A1.
    ///
    /// calamine positions its range at the first used cell while the grid
    /// addresses cells by their place on the sheet, so this is the shape where
    /// those two conventions come apart. Every other fixture here starts at
    /// A1, which is why none of them caught it.
    fn offset_sample() -> Book {
        let range: Range<Data> = Range::from_sparse(vec![
            calamine::Cell::new((1, 1), Data::String("先頭".to_owned())),
            calamine::Cell::new((2, 1), Data::String("次".to_owned())),
        ]);
        let origin = range.start().unwrap_or((0, 0));
        let (rows, cols) = range.get_size();
        Book {
            info: WorkbookInfo {
                path: "fixture".to_owned(),
                file_name: "fixture".to_owned(),
                sheets: vec![SheetInfo {
                    name: "Sheet1".to_owned(),
                    rows: rows as u32,
                    cols: cols as u32,
                }],
            },
            sheets: HashMap::from([("Sheet1".to_owned(), Sheet { range, origin })]),
            layouts: HashMap::new(),
            layout_cache: layout::Cache::default(),
        }
    }

    #[test]
    fn addresses_cells_where_the_grid_draws_them_not_where_the_range_starts() {
        let book = offset_sample();

        // B2 is row 1, column 1, so the window comes back with it in slot
        // [1][1] — adding the range's origin put it in [0][0], one row up and
        // one column left of the styling that belonged to it.
        let block = book.cells("Sheet1", 0, 0, 4, 2).unwrap();
        assert_eq!(block[0][0].text, "");
        assert_eq!(block[1][1].text, "先頭");
        assert_eq!(block[2][1].text, "次");

        // Asking for one cell by its own address finds it.
        let cell = book.cells("Sheet1", 1, 1, 1, 1).unwrap();
        assert_eq!(cell[0][0].text, "先頭");

        // And a hit points at the cell it was found in.
        let hits = book.search(None, &Matcher::new("次", true, true));
        assert_eq!((hits[0].row, hits[0].col), (2, 1));
    }

    #[test]
    fn paging_clamps_at_the_last_row() {
        let Some(book) = sample() else { return };
        let tail = book.cells("変更履歴", 2, 0, 128, 32).unwrap();
        assert_eq!(tail.len(), 1);
        assert_eq!(tail[0][2].text, "定時定額払戻開始日 の条件を追加");
    }

    #[test]
    fn search_covers_japanese_text_across_sheets() {
        let Some(book) = sample() else { return };
        let hits = book.search(None, &Matcher::new("定時定額払戻開始日", false, false));
        assert!(hits.iter().any(|h| h.sheet == "DB抽出仕様"));
        assert!(hits.iter().any(|h| h.sheet == "変更履歴"));

        let scoped = book.search(Some("変更履歴"), &Matcher::new("定時定額払戻開始日", false, false));
        assert!(scoped.iter().all(|h| h.sheet == "変更履歴"));

        assert!(book.search(None, &Matcher::new("", false, false)).is_empty());
    }

    #[test]
    fn reads_a_row_or_column_for_copying() {
        let Some(book) = sample() else { return };

        // The header row, with the trailing blanks past column E trimmed off.
        let row = book.axis_values("DB抽出仕様", Some(0), None, 0, 31, 64).unwrap();
        assert_eq!(row, ["No", "テーブル", "カラム", "条件", "備考"]);

        // Column A of the same sheet is a run of numbers.
        let column = book
            .axis_values("DB抽出仕様", None, Some(0), 0, 4, 64)
            .unwrap();
        assert_eq!(column, ["No", "1", "2", "3", "4"]);

        assert!(book.axis_values("DB抽出仕様", None, None, 0, 4, 64).is_err());
    }

    #[test]
    fn lists_distinct_column_values_by_frequency() {
        let Some(book) = sample() else { return };

        // The autofilter covers A1:E401, so data rows are 1..=400.
        let values = book
            .column_values("DB抽出仕様", 4, 1, 400, None, 50)
            .unwrap();
        let flagged = values
            .iter()
            .find(|entry| entry.value == "要確認")
            .expect("要確認 should appear");
        // Every 11th row, 11..=396.
        assert_eq!(flagged.count, 36);

        let filtered = book
            .column_values("DB抽出仕様", 4, 1, 400, Some("要確"), 50)
            .unwrap();
        assert_eq!(filtered.len(), 1);
        assert_eq!(filtered[0].value, "要確認");
    }

    #[test]
    fn filters_rows_by_column_value() {
        let Some(book) = sample() else { return };

        let matching = book
            .filter_rows("DB抽出仕様", 1, 400, &[(4, vec!["要確認".to_owned()])], 1000)
            .unwrap();
        assert_eq!(matching.len(), 36);
        assert_eq!(matching[0], 11);
        assert_eq!(matching[35], 396);

        // An empty criterion list matches every row in range.
        let all = book.filter_rows("DB抽出仕様", 1, 400, &[], 1000).unwrap();
        assert_eq!(all.len(), 400);

        // Two criteria must both hold, so nothing survives here.
        let none = book
            .filter_rows(
                "DB抽出仕様",
                1,
                400,
                &[(4, vec!["要確認".to_owned()]), (0, vec!["999".to_owned()])],
                1000,
            )
            .unwrap();
        assert!(none.is_empty());
    }
}
