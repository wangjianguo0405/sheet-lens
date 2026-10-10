mod drawing;
mod excel;
mod layout;
mod numfmt;

use std::collections::HashMap;
use std::io::{BufReader, Read};
use std::sync::{Arc, Mutex};

use excel::{Book, Cell, Hit, Matcher, ValueCount, WorkbookInfo};
use layout::SheetLayout;
use serde::Deserialize;
use tauri::State;

/// Distinct values returned for a filter dropdown.
const MAX_VALUES: usize = 1000;
/// Upper bound on rows a filter may return, to keep the IPC payload sane.
const MAX_FILTER_ROWS: usize = 200_000;
/// Upper bound on cells a single copy can pull back.
const MAX_AXIS_CELLS: usize = 4096;

#[derive(Deserialize)]
struct FilterCriterion {
    col: u32,
    values: Vec<String>,
}

/// Open workbooks, keyed by absolute path. Parsed sheets stay in memory so the
/// grid can page through them and search without re-reading the file.
#[derive(Default)]
struct Store(Mutex<HashMap<String, Book>>);

/// Shared, so the layout warmer can keep the store alive past the command that
/// started it. The lock is only ever held for short map operations.
type Shared = Arc<Store>;

/// Parses every sheet's layout ahead of time.
///
/// One design document's `機能定義書` sheet is 2.5 MB of XML and takes ~100 ms
/// to walk, which is a visible stall every time the user switches to it. Doing
/// the work up front costs nothing the user is waiting on.
fn warm_layouts(store: Shared, path: String, names: Vec<String>) {
    let mut cache = layout::Cache::default();
    for name in names {
        {
            let books = store.0.lock().unwrap();
            match books.get(&path) {
                // Already parsed on demand, or the workbook was closed.
                Some(book) if book.has_layout(&name) => continue,
                Some(_) => {}
                None => return,
            }
        }
        let Ok(parsed) = layout::read_cached(&mut cache, &path, &name) else {
            continue;
        };
        if let Some(book) = store.0.lock().unwrap().get_mut(&path) {
            book.put_layout(name, parsed);
        }
    }
}

#[tauri::command]
fn open_workbook(path: String, store: State<'_, Shared>) -> Result<WorkbookInfo, String> {
    let book = excel::load(&path)?;
    let info = book.info.clone();
    store.0.lock().unwrap().insert(path.clone(), book);

    let names: Vec<String> = info
        .sheets
        .iter()
        .map(|sheet| sheet.name.clone())
        .collect();
    let shared = Arc::clone(&store);
    let warm = path.clone();
    std::thread::spawn(move || warm_layouts(shared, warm, names));

    Ok(info)
}

#[tauri::command]
fn get_cells(
    path: String,
    sheet: String,
    row: u32,
    col: u32,
    rows: u32,
    cols: u32,
    store: State<'_, Shared>,
) -> Result<Vec<Vec<Cell>>, String> {
    let books = store.0.lock().unwrap();
    let book = books
        .get(&path)
        .ok_or_else(|| format!("workbook not open: {path}"))?;
    book.cells(&sheet, row, col, rows, cols)
}

#[tauri::command]
fn search_workbook(
    path: String,
    query: String,
    sheet: Option<String>,
    case_sensitive: bool,
    whole_cell: bool,
    store: State<'_, Shared>,
) -> Result<Vec<Hit>, String> {
    let books = store.0.lock().unwrap();
    let book = books
        .get(&path)
        .ok_or_else(|| format!("workbook not open: {path}"))?;
    let matcher = Matcher::new(&query, case_sensitive, whole_cell);
    Ok(book.search(sheet.as_deref(), &matcher))
}

#[tauri::command]
fn sheet_layout(
    path: String,
    sheet: String,
    store: State<'_, Shared>,
) -> Result<SheetLayout, String> {
    let mut books = store.0.lock().unwrap();
    let book = books
        .get_mut(&path)
        .ok_or_else(|| format!("workbook not open: {path}"))?;
    book.layout(&sheet)
}

#[tauri::command]
fn column_values(
    path: String,
    sheet: String,
    col: u32,
    first_row: u32,
    last_row: u32,
    query: Option<String>,
    store: State<'_, Shared>,
) -> Result<Vec<ValueCount>, String> {
    let books = store.0.lock().unwrap();
    let book = books
        .get(&path)
        .ok_or_else(|| format!("workbook not open: {path}"))?;
    book.column_values(&sheet, col, first_row, last_row, query.as_deref(), MAX_VALUES)
}

#[tauri::command]
fn filter_rows(
    path: String,
    sheet: String,
    first_row: u32,
    last_row: u32,
    criteria: Vec<FilterCriterion>,
    store: State<'_, Shared>,
) -> Result<Vec<u32>, String> {
    let books = store.0.lock().unwrap();
    let book = books
        .get(&path)
        .ok_or_else(|| format!("workbook not open: {path}"))?;
    let criteria: Vec<(u32, Vec<String>)> = criteria
        .into_iter()
        .map(|criterion| (criterion.col, criterion.values))
        .collect();
    book.filter_rows(&sheet, first_row, last_row, &criteria, MAX_FILTER_ROWS)
}

#[tauri::command]
fn axis_values(
    path: String,
    sheet: String,
    row: Option<u32>,
    col: Option<u32>,
    from: u32,
    to: u32,
    store: State<'_, Shared>,
) -> Result<Vec<String>, String> {
    let books = store.0.lock().unwrap();
    let book = books
        .get(&path)
        .ok_or_else(|| format!("workbook not open: {path}"))?;
    book.axis_values(&sheet, row, col, from, to, MAX_AXIS_CELLS)
}

#[tauri::command]
fn close_workbook(path: String, store: State<'_, Shared>) {
    store.0.lock().unwrap().remove(&path);
}

/// Reads one embedded image out of a workbook, as base64.
///
/// Pictures are asked for one at a time and only once they are on screen: a
/// design document carries around twenty screenshots, and folding them into the
/// layout payload would cost megabytes on every sheet switch.
#[tauri::command]
fn read_media(path: String, part: String) -> Result<String, String> {
    let file = std::fs::File::open(&path).map_err(|e| e.to_string())?;
    let mut zip =
        zip::ZipArchive::new(BufReader::new(file)).map_err(|e| e.to_string())?;
    // A lookup in the archive's own directory, so a crafted name cannot escape it.
    let mut entry = zip.by_name(&part).map_err(|e| e.to_string())?;
    let mut bytes = Vec::new();
    entry.read_to_end(&mut bytes).map_err(|e| e.to_string())?;
    Ok(base64(&bytes))
}

/// Standard base64, so the renderer can put the bytes straight into a `data:` URL.
fn base64(bytes: &[u8]) -> String {    const ALPHABET: &[u8; 64] =
        b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut out = String::with_capacity((bytes.len() + 2) / 3 * 4);
    for chunk in bytes.chunks(3) {
        let group = (u32::from(chunk[0]) << 16)
            | (u32::from(chunk.get(1).copied().unwrap_or(0)) << 8)
            | u32::from(chunk.get(2).copied().unwrap_or(0));
        let digit = |shift: u32| ALPHABET[(group >> shift) as usize & 63] as char;
        out.push(digit(18));
        out.push(digit(12));
        out.push(if chunk.len() > 1 { digit(6) } else { '=' });
        out.push(if chunk.len() > 2 { digit(0) } else { '=' });
    }
    out
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_dialog::init())
        .manage(Arc::new(Store::default()))
        .invoke_handler(tauri::generate_handler![
            open_workbook,
            get_cells,
            search_workbook,
            sheet_layout,
            column_values,
            filter_rows,
            axis_values,
            read_media,
            close_workbook
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
