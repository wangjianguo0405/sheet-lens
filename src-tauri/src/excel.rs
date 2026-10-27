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
#[serde(rename_all = "camelCase")]
pub struct SheetInfo {
    pub name: String,
    pub rows: u32,
    pub cols: u32,
    /// The sheet tab's own colour (`<tabColor>`), as `#RRGGBB`. Excel shows it
    /// as a strip under the tab, and a tab with none has none.
    pub tab_color: Option<String>,
    /// Whether the file hides the sheet (`<sheet state="hidden">`). Excel draws
    /// no tab for one of these and reaches it only through Unhide.
    pub hidden: bool,
    /// Whether the sheet defines a `_xlnm.Print_Area` of its own. A workbook
    /// with none anywhere opens in the ordinary view rather than under the page
    /// shading, so this is read before any layout is.
    pub print_area: bool,
    /// How the file left this sheet: the view it was in, whether its tab was
    /// the selected one, and where its cursor was. Absent for a sheet whose
    /// part could not be read.
    #[serde(flatten)]
    pub state: layout::SheetState,
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

/// A rectangle on the sheet, in zero-based cells.
#[derive(Serialize, Clone, Debug, PartialEq)]
pub struct Block {
    pub row: u32,
    pub col: u32,
    pub rows: u32,
    pub cols: u32,
}

/// One cell on the sheet, in zero-based coordinates.
#[derive(Serialize, Clone, Debug, PartialEq)]
pub struct CellPos {
    pub row: u32,
    pub col: u32,
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
    // One pass over the zip for every sheet's tab colour, up front: the tab
    // strip draws them all at once and a sheet's own layout is read lazily.
    let tab_colors = layout::tab_colors(path);
    // Same pass, same reason: which view the workbook opens in is decided from
    // this before anything is drawn.
    let print_areas = layout::print_area_sheets(path);
    // And which sheets the file keeps off the tab strip, which the strip needs
    // before it draws a single tab.
    let hidden = layout::hidden_sheets(path);
    // And how each sheet was left — the view, the selected tab and the cursor —
    // which decides what the workbook opens on.
    let states = layout::sheet_states(path);

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
            tab_color: tab_colors.get(&name).cloned(),
            hidden: hidden.contains(&name),
            print_area: print_areas.contains(&name),
            state: states.get(&name).cloned().unwrap_or_default(),
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

    /// Excel's current region: the block reachable from a cell without crossing
    /// a wholly blank row or column.
    ///
    /// The two scans it starts from — up and down the anchor's column, then left
    /// and right along the anchor's row — are only the beginning, because a
    /// column reached along that row can carry the block further up again. What
    /// Excel returns is the closure of the two, so the edges are walked to a
    /// fixed point rather than once each.
    pub fn current_region(&self, sheet: &str, row: u32, col: u32) -> Result<Block, String> {
        let stored = self
            .sheets
            .get(sheet)
            .ok_or_else(|| format!("sheet not found: {sheet}"))?;
        let (height, width) = stored.range.get_size();
        let (origin_row, origin_col) = stored.origin;
        let end_row = origin_row + height as u32;
        let end_col = origin_col + width as u32;

        let filled = |r: u32, c: u32| cell_at(stored.range.get_value((r, c))).kind != "empty";

        // A blank cell is a region of its own, which is what Excel gives for
        // Ctrl+A out in the empty part of a sheet — and an anchor off the stored
        // range cannot be walked outwards from at all.
        if row >= end_row || col >= end_col || !filled(row, col) {
            return Ok(Block {
                row,
                col,
                rows: 1,
                cols: 1,
            });
        }

        let (mut top, mut bottom) = (row, row);
        let (mut left, mut right) = (col, col);
        loop {
            let before = (top, bottom, left, right);
            while top > origin_row && (left..=right).any(|c| filled(top - 1, c)) {
                top -= 1;
            }
            while bottom + 1 < end_row && (left..=right).any(|c| filled(bottom + 1, c)) {
                bottom += 1;
            }
            while left > origin_col && (top..=bottom).any(|r| filled(r, left - 1)) {
                left -= 1;
            }
            while right + 1 < end_col && (top..=bottom).any(|r| filled(r, right + 1)) {
                right += 1;
            }
            if (top, bottom, left, right) == before {
                break;
            }
        }

        Ok(Block {
            row: top,
            col: left,
            rows: bottom - top + 1,
            cols: right - left + 1,
        })
    }

    /// Where Ctrl+arrow lands, in Excel's own rule.
    ///
    /// From a filled cell it is the far end of the contiguous run of filled
    /// cells in that direction; from a blank one it is the nearest filled cell,
    /// which is the near edge of the run beyond the blank. That is what makes
    /// the presses walk: the first leaves a block at its bottom, the next takes
    /// the cursor off the block to the top of the one below it.
    ///
    /// `None` means there was nothing ahead at all, which Excel answers with the
    /// sheet's last row or column — a number only the caller knows, since the
    /// stored range stops at the last cell the file mentions.
    pub fn edge_cell(
        &self,
        sheet: &str,
        row: u32,
        col: u32,
        dir: &str,
    ) -> Result<Option<CellPos>, String> {
        let stored = self
            .sheets
            .get(sheet)
            .ok_or_else(|| format!("sheet not found: {sheet}"))?;
        let (height, width) = stored.range.get_size();
        let (origin_row, origin_col) = stored.origin;
        let end_row = origin_row + height as u32;
        let end_col = origin_col + width as u32;
        if height == 0 || width == 0 {
            return Ok(None);
        }

        let filled = |r: u32, c: u32| cell_at(stored.range.get_value((r, c))).kind != "empty";

        let (drow, dcol): (i64, i64) = match dir {
            "up" => (-1, 0),
            "down" => (1, 0),
            "left" => (0, -1),
            "right" => (0, 1),
            other => return Err(format!("unknown direction: {other}")),
        };

        // The walk runs along one axis with the other coordinate held, so a line
        // whose held coordinate lies off the stored range holds nothing at all —
        // a row past the sheet's last one has no cells in any column.
        let horizontal = drow == 0;
        let held_ok = if horizontal {
            row >= origin_row && row < end_row
        } else {
            col >= origin_col && col < end_col
        };
        if !held_ok {
            return Ok(None);
        }
        let step: i64 = if horizontal { dcol } else { drow };
        let axis = if horizontal { col } else { row } as i64;
        let (near, far) = if horizontal {
            (origin_col as i64, end_col as i64 - 1)
        } else {
            (origin_row as i64, end_row as i64 - 1)
        };

        // The cursor may stand outside the stored range entirely — in the blank
        // area past the last column of a table, say — and the walk still has to
        // run towards it to find the data's edge. So only passing the far end of
        // the range in the direction of travel ends the walk; merely being
        // beyond the range to begin with does not.
        let at = |n: i64| -> Option<(u32, u32)> {
            let along = axis + step * n;
            if (step > 0 && along > far) || (step < 0 && along < near) {
                return None;
            }
            Some(if horizontal {
                (row, along as u32)
            } else {
                (along as u32, col)
            })
        };

        // A filled cell runs to the end of its own run. A blank one — and a
        // filled one whose neighbour is blank, which is how the next press
        // leaves a block — takes the nearest filled cell instead, so a press
        // from a blank row drops onto the block's near edge rather than flying
        // over it to the far one.
        if filled(row, col) {
            let mut last = None;
            let mut n = 1;
            while let Some((r, c)) = at(n) {
                if !filled(r, c) {
                    break;
                }
                last = Some((r, c));
                n += 1;
            }
            if let Some((r, c)) = last {
                return Ok(Some(CellPos { row: r, col: c }));
            }
        }

        let mut n = 1;
        while let Some((r, c)) = at(n) {
            if filled(r, c) {
                return Ok(Some(CellPos { row: r, col: c }));
            }
            n += 1;
        }
        Ok(None)
    }

    /// Distinct display values in one column, most common first, counted over
    /// the rows the other columns' filters leave visible.
    ///
    /// That is the list Excel's drop-down offers: a column's own criterion is
    /// not applied to it — the list is how the criterion gets changed, and a
    /// list narrowed by its own filter could never offer back what was turned
    /// off — while every other column's is.
    pub fn column_values(
        &self,
        sheet: &str,
        col: u32,
        first_row: u32,
        last_row: u32,
        criteria: &[(u32, Vec<String>)],
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
        let wanted: Vec<(u32, HashSet<String>)> = criteria
            .iter()
            .filter(|(criterion_col, _)| *criterion_col != col)
            .map(|(criterion_col, values)| {
                (
                    *criterion_col,
                    values.iter().map(|value| value.to_lowercase()).collect(),
                )
            })
            .collect();

        let mut counts: HashMap<String, u32> = HashMap::new();
        let mut text = String::new();
        'rows: for r in start..end {
            for (criterion_col, accepted) in &wanted {
                write_text(data.range.get_value((r, *criterion_col)), &mut text);
                if !accepted.contains(&text.to_lowercase()) {
                    continue 'rows;
                }
            }
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
            // A hidden sheet is not searched: Excel cannot take you to one, and
            // a hit is a place the grid is about to put the cursor.
            if info.hidden {
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

    /// Writes a workbook's real parser output to JSON, for the browser harness.
    ///
    /// The grid can be driven outside the Tauri shell — `pnpm dev` serves the
    /// same UI — but `invoke` has to come from somewhere, and hand-made payloads
    /// hide exactly the things worth testing: widths, merges, print areas and
    /// shapes all come off the parser. So this dumps what the parser actually
    /// returned and the harness answers `open_workbook`/`sheet_layout`/`get_cells`
    /// out of it.
    ///
    /// `SHEETLENS_XLSX` names the workbook and `SHEETLENS_OUT` the JSON file to
    /// write; `SHEETLENS_SHEETS` restricts it to some of the sheets, since a
    /// design document can be several hundred thousand cells.
    #[test]
    #[ignore]
    fn dump_web_fixture() {
        let Ok(path) = std::env::var("SHEETLENS_XLSX") else {
            return;
        };
        let Ok(out) = std::env::var("SHEETLENS_OUT") else {
            eprintln!("set SHEETLENS_OUT to the file to write");
            return;
        };
        let only: Option<Vec<String>> = std::env::var("SHEETLENS_SHEETS")
            .ok()
            .map(|list| list.split(',').map(|s| s.trim().to_owned()).collect());
        // The whole point of a sweep over hundreds of workbooks is the page
        // layout; serialising every cell as well turns it into gigabytes.
        let layout_only = std::env::var("SHEETLENS_LAYOUT_ONLY").is_ok();
        // The reader's short date, for the cells whose format is the built-in
        // one — a fixture dumped without it gets the American pattern.
        if let Ok(pattern) = std::env::var("SHEETLENS_SHORT_DATE") {
            crate::layout::set_short_date(pattern);
        }

        let mut book = load(&path).expect("workbook should load");
        let mut sheets = serde_json::Map::new();
        for info in book.info.sheets.clone() {
            if let Some(names) = &only {
                if !names.iter().any(|n| n == &info.name) {
                    continue;
                }
            }
            let layout = book.layout(&info.name).expect("layout");
            if layout_only {
                sheets.insert(info.name.clone(), serde_json::json!({ "layout": layout }));
                continue;
            }
            let rows = layout.rows.max(info.rows);
            let cols = layout.cols.max(info.cols);
            let cells = book.cells(&info.name, 0, 0, rows, cols).expect("cells");
            // Padded out to the layout's own extent: `cells` stops where
            // calamine's range does, and everything past it is blank anyway.
            let mut grid: Vec<Vec<Cell>> = cells;
            grid.resize(rows as usize, Vec::new());
            for line in grid.iter_mut() {
                line.resize(cols as usize, Cell { kind: "empty", text: String::new() });
            }
            eprintln!(
                "{}: {rows}×{cols}, {} shapes",
                info.name,
                layout.drawings.len()
            );
            sheets.insert(
                info.name.clone(),
                serde_json::json!({ "layout": layout, "cells": grid }),
            );
        }

        let payload = serde_json::json!({ "info": book.info, "sheets": sheets });
        std::fs::write(&out, serde_json::to_vec(&payload).expect("json")).expect("write fixture");
        eprintln!("wrote {out}");
    }

    #[test]
    fn lists_every_sheet_with_its_size() {
        let Some(book) = sample() else { return };
        let names: Vec<_> = book.info.sheets.iter().map(|s| s.name.as_str()).collect();
        assert_eq!(names, ["Summary", "Query", "Logic", "History"]);

        let logic = book.info.sheets.iter().find(|s| s.name == "Logic").unwrap();
        assert_eq!((logic.rows, logic.cols), (12001, 4));
    }

    /// Ctrl+A's first stage. Checked against the invariants rather than one
    /// document's shape, since the sample is the only workbook a test may open.
    #[test]
    fn a_current_region_is_bounded_by_blank_rows_and_columns() {
        let Some(book) = sample() else { return };
        for sheet in ["Summary", "Logic", "History"] {
            for (row, col) in [(0, 0), (2, 1), (17, 3), (500, 0), (11999, 3)] {
                let region = book.current_region(sheet, row, col).expect("sheet exists");
                assert!(region.rows >= 1 && region.cols >= 1, "{sheet}");
                // The cell asked about is inside what came back.
                assert!(
                    row >= region.row && row < region.row + region.rows,
                    "{sheet}: row {row} outside {region:?}"
                );
                assert!(
                    col >= region.col && col < region.col + region.cols,
                    "{sheet}: col {col} outside {region:?}"
                );

                // An empty reply means the position is off the stored range,
                // which is the scan's own boundary rather than a blank cell.
                let blank = |r: u32, c: u32| -> Option<bool> {
                    let line = book.cells(sheet, r, c, 1, 1).ok()?.pop()?;
                    Some(line.into_iter().next()?.kind == "empty")
                };
                // A region ends where a wholly blank row or column begins, which
                // is the whole point of it — an arbitrary rectangle would not.
                if region.row > 0 {
                    for c in region.col..region.col + region.cols {
                        if let Some(empty) = blank(region.row - 1, c) {
                            assert!(empty, "{sheet}: ({},{c}) above the region is not blank", region.row - 1);
                        }
                    }
                }
                if region.col > 0 {
                    for r in region.row..region.row + region.rows {
                        if let Some(empty) = blank(r, region.col - 1) {
                            assert!(empty, "{sheet}: ({r},{}) before the region is not blank", region.col - 1);
                        }
                    }
                }
                for c in region.col..region.col + region.cols {
                    if let Some(empty) = blank(region.row + region.rows, c) {
                        assert!(empty, "{sheet}: ({},{c}) below the region is not blank", region.row + region.rows);
                    }
                }
            }
        }
    }

    #[test]
    fn reads_cells_by_offset_and_keeps_numbers_typed() {
        let Some(book) = sample() else { return };

        // Row 0 is the merged title, so the header row is row 1.
        let title = book.cells("Summary", 0, 0, 1, 1).unwrap();
        assert_eq!(title[0][0].text, "Summary");

        let head = book.cells("Summary", 1, 0, 2, 2).unwrap();
        assert_eq!(head[0][0].text, "項目");
        assert_eq!(head[0][1].text, "内容");
        assert_eq!(head[1][1].text, "サンプル機能A");

        // Row 0 holds the header, so row 3 is the record whose "行" is 3.
        let logic = book.cells("Logic", 3, 0, 1, 4).unwrap();
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
                    tab_color: None,
                    hidden: false,
                    print_area: false,
                    state: layout::SheetState::default(),
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

    /// A small sheet whose content starts at A3 and reaches column D, with a
    /// blank row splitting column A in two — the shapes Ctrl+arrow has to get
    /// right, including the blank area past the data it can be pressed from.
    fn gapped_sheet() -> Book {
        let holds = [
            (2u32, 0u32),
            (3, 0),
            (4, 0),
            (6, 0),
            (7, 0),
            (8, 0),
            (8, 3),
        ];
        let range: Range<Data> = Range::from_sparse(
            holds
                .into_iter()
                .map(|(row, col)| calamine::Cell::new((row, col), Data::String("x".to_owned())))
                .collect(),
        );
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
                    tab_color: None,
                    hidden: false,
                    print_area: false,
                    state: layout::SheetState::default(),
                }],
            },
            sheets: HashMap::from([("Sheet1".to_owned(), Sheet { range, origin })]),
            layouts: HashMap::new(),
            layout_cache: layout::Cache::default(),
        }
    }

    #[test]
    #[ignore]
    fn dump_edges() {
        let Ok(path) = std::env::var("SHEETLENS_XLSX") else {
            return;
        };
        let book = load(&path).expect("workbook should load");
        let sheet = std::env::var("SHEETLENS_SHEET").unwrap_or_else(|_| "Sheet1".into());
        for info in &book.info.sheets {
            println!("sheet {} {}x{}", info.name, info.rows, info.cols);
        }
        let stored = &book.sheets[&sheet];
        let (height, width) = stored.range.get_size();
        let (origin_row, origin_col) = stored.origin;
        let end_row = origin_row as i64 + height as i64;
        let end_col = origin_col as i64 + width as i64;

        // Spread over the used range and a little past both of its edges: the
        // blank area beyond the data is where a jump inwards has to work from.
        let spread = |from: i64, to: i64| -> Vec<u32> {
            let mut out = vec![from, from + 1, to - 1, to, to + 1, to + 3];
            for i in 0..6 {
                out.push(from + (to - from) * i / 5);
            }
            out.into_iter()
                .filter(|v| *v >= 0)
                .map(|v| v as u32)
                .collect()
        };
        let rows = spread(origin_row as i64, end_row);
        let cols = spread(origin_col as i64, end_col);
        for r in &rows {
            for c in &cols {
                println!("P {r} {c}");
            }
        }
        for r in &rows {
            for c in &cols {
                for dir in ["up", "down", "left", "right"] {
                    match book.edge_cell(&sheet, *r, *c, dir).expect("sheet exists") {
                        Some(pos) => println!("R {r} {c} {dir} {} {}", pos.row, pos.col),
                        None => println!("R {r} {c} {dir} - -"),
                    }
                }
            }
        }
    }

    #[test]
    fn ctrl_arrow_is_the_edge_of_the_run_then_the_next_cell_past_a_gap() {
        let book = gapped_sheet();
        let sheet = "Sheet1";
        let at = |row, col| CellPos { row, col };

        // From inside a run it is the last filled cell before the blank one: the
        // end of the *run*, which the probe against Excel's own `Range.End`
        // confirms — from A1 of `A1:A5` it is A5, not the A10 that lies beyond
        // the following gap.
        assert_eq!(book.edge_cell(sheet, 2, 0, "down"), Ok(Some(at(4, 0))));
        // Standing on the last cell of a run, the next press leaves it for the
        // top of the one past the gap rather than staying put.
        assert_eq!(book.edge_cell(sheet, 4, 0, "down"), Ok(Some(at(6, 0))));
        assert_eq!(book.edge_cell(sheet, 5, 0, "down"), Ok(Some(at(6, 0))));
        // From the gap itself, the nearest cell either way.
        assert_eq!(book.edge_cell(sheet, 5, 0, "up"), Ok(Some(at(4, 0))));
        assert_eq!(book.edge_cell(sheet, 6, 0, "up"), Ok(Some(at(4, 0))));
        // The far end of the lower run, asked upwards, is its own top.
        assert_eq!(book.edge_cell(sheet, 7, 0, "up"), Ok(Some(at(6, 0))));
        // Above the data, and still on the sheet: the nearest filled cell below
        // is the data's own top. A walk bounded by the range's near edge cannot
        // answer this — its first step lands outside the range altogether.
        assert_eq!(book.edge_cell(sheet, 0, 0, "down"), Ok(Some(at(2, 0))));
        assert_eq!(book.edge_cell(sheet, 1, 0, "up"), Ok(None));
        assert_eq!(book.edge_cell(sheet, 8, 0, "down"), Ok(None));

        // Left from the blank columns past the data crosses C and B to reach D,
        // which is what a walk that stopped at the range's own edge got wrong:
        // it gave up at once and left the grid to fall back to column A. Excel
        // answers the same shape as B1 — `D1.End(xlToLeft)` on the probe sheet.
        assert_eq!(book.edge_cell(sheet, 8, 4, "left"), Ok(Some(at(8, 3))));
        assert_eq!(book.edge_cell(sheet, 8, 7, "left"), Ok(Some(at(8, 3))));
        // A filled cell whose neighbour is blank leaves its own run for the
        // nearest cell beyond the gap, rather than flying to the sheet's edge.
        assert_eq!(book.edge_cell(sheet, 8, 3, "left"), Ok(Some(at(8, 0))));
        // Across a row whose only content is at column A.
        assert_eq!(book.edge_cell(sheet, 2, 4, "left"), Ok(Some(at(2, 0))));
        // Nothing to the right of the data at all, so the walk ends there and
        // the caller answers with the sheet's own last column.
        assert_eq!(book.edge_cell(sheet, 8, 3, "right"), Ok(None));
        assert_eq!(book.edge_cell(sheet, 2, 0, "right"), Ok(None));
    }

    #[test]
    fn paging_clamps_at_the_last_row() {
        let Some(book) = sample() else { return };
        let tail = book.cells("History", 2, 0, 128, 32).unwrap();
        assert_eq!(tail.len(), 1);
        assert_eq!(tail[0][2].text, "サンプル日付項目 の条件を追加");
    }

    #[test]
    fn search_covers_japanese_text_across_sheets() {
        let Some(book) = sample() else { return };
        let hits = book.search(None, &Matcher::new("サンプル日付項目", false, false));
        assert!(hits.iter().any(|h| h.sheet == "Query"));
        assert!(hits.iter().any(|h| h.sheet == "History"));

        let scoped = book.search(Some("History"), &Matcher::new("サンプル日付項目", false, false));
        assert!(scoped.iter().all(|h| h.sheet == "History"));

        assert!(book.search(None, &Matcher::new("", false, false)).is_empty());
    }

    /// A hidden sheet is not searched. Excel cannot put the cursor on one — the
    /// sheet cannot even be activated — and a hit is exactly a place the grid is
    /// about to go.
    #[test]
    fn search_leaves_the_hidden_sheets_alone() {
        let range = |text: &str| -> Range<Data> {
            Range::from_sparse(vec![calamine::Cell::new((0, 0), Data::String(text.to_owned()))])
        };
        let info = |name: &str, hidden: bool| SheetInfo {
            name: name.to_owned(),
            rows: 1,
            cols: 1,
            tab_color: None,
            hidden,
            print_area: false,
            state: layout::SheetState::default(),
        };
        let book = Book {
            info: WorkbookInfo {
                path: "fixture".to_owned(),
                file_name: "fixture".to_owned(),
                sheets: vec![info("表示", false), info("非表示", true)],
            },
            sheets: HashMap::from([
                (
                    "表示".to_owned(),
                    Sheet { range: range("さがす"), origin: (0, 0) },
                ),
                (
                    "非表示".to_owned(),
                    Sheet { range: range("さがす"), origin: (0, 0) },
                ),
            ]),
            layouts: HashMap::new(),
            layout_cache: layout::Cache::default(),
        };

        let hits = book.search(None, &Matcher::new("さがす", false, false));
        assert_eq!(hits.len(), 1);
        assert_eq!(hits[0].sheet, "表示");
        // Named outright it is still not read: the search panel can be pointed
        // at a sheet, and the answer is not a cursor on a hidden one.
        assert!(book
            .search(Some("非表示"), &Matcher::new("さがす", false, false))
            .is_empty());
    }

    #[test]
    fn reads_a_row_or_column_for_copying() {
        let Some(book) = sample() else { return };

        // The header row, with the trailing blanks past column E trimmed off.
        let row = book.axis_values("Query", Some(0), None, 0, 31, 64).unwrap();
        assert_eq!(row, ["No", "テーブル", "カラム", "条件", "備考"]);

        // Column A of the same sheet is a run of numbers.
        let column = book
            .axis_values("Query", None, Some(0), 0, 4, 64)
            .unwrap();
        assert_eq!(column, ["No", "1", "2", "3", "4"]);

        assert!(book.axis_values("Query", None, None, 0, 4, 64).is_err());
    }

    #[test]
    fn lists_distinct_column_values_by_frequency() {
        let Some(book) = sample() else { return };

        // The autofilter covers A1:E401, so data rows are 1..=400.
        let values = book
            .column_values("Query", 4, 1, 400, &[], None, 50)
            .unwrap();
        let flagged = values
            .iter()
            .find(|entry| entry.value == "要確認")
            .expect("要確認 should appear");
        // Every 11th row, 11..=396.
        assert_eq!(flagged.count, 36);

        let filtered = book
            .column_values("Query", 4, 1, 400, &[], Some("要確"), 50)
            .unwrap();
        assert_eq!(filtered.len(), 1);
        assert_eq!(filtered[0].value, "要確認");
    }

    #[test]
    fn column_values_follow_the_other_columns_filters() {
        let Some(book) = sample() else { return };
        let criteria = [(4, vec!["要確認".to_owned()])];

        // Another column lists only the rows the filter leaves — 36 of them.
        let ids = book.column_values("Query", 0, 1, 400, &criteria, None, 500).unwrap();
        assert_eq!(ids.iter().map(|entry| entry.count).sum::<u32>(), 36);

        // The filtered column keeps its own full list: its criterion is how the
        // list gets changed, so it must not narrow it.
        let own = book.column_values("Query", 4, 1, 400, &criteria, None, 50).unwrap();
        let flagged = own
            .iter()
            .find(|entry| entry.value == "要確認")
            .expect("要確認 should appear");
        assert_eq!(flagged.count, 36);
        assert!(own.len() > 1, "the column's other values stay reachable");
    }

    #[test]
    fn filters_rows_by_column_value() {
        let Some(book) = sample() else { return };

        let matching = book
            .filter_rows("Query", 1, 400, &[(4, vec!["要確認".to_owned()])], 1000)
            .unwrap();
        assert_eq!(matching.len(), 36);
        assert_eq!(matching[0], 11);
        assert_eq!(matching[35], 396);

        // An empty criterion list matches every row in range.
        let all = book.filter_rows("Query", 1, 400, &[], 1000).unwrap();
        assert_eq!(all.len(), 400);

        // Two criteria must both hold, so nothing survives here.
        let none = book
            .filter_rows(
                "Query",
                1,
                400,
                &[(4, vec!["要確認".to_owned()]), (0, vec!["999".to_owned()])],
                1000,
            )
            .unwrap();
        assert!(none.is_empty());
    }
}
