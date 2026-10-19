mod cond;
mod drawing;
mod excel;
mod layout;
mod numfmt;
mod update;

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
/// A 2.5 MB sheet is roughly 100 ms of XML to walk, a visible stall every time
/// the user switches to it. Doing the work up front costs nothing the user is
/// waiting on.
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

/// Every command here is declared `async` so that it runs on the thread pool
/// instead of on the thread that received it. Tauri dispatches a plain `fn`
/// inline (`run_invoke_handler` in `webview/mod.rs`), and the webview's messages
/// arrive on the main one, so a synchronous command holds the whole IPC up for
/// as long as it runs — a sheet's layout, a workbook-wide search or a screenful
/// of pictures is long enough to be felt.
#[tauri::command(async)]
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

#[tauri::command(async)]
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

#[tauri::command(async)]
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

#[tauri::command(async)]
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

#[tauri::command(async)]
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

#[tauri::command(async)]
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

#[tauri::command(async)]
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

#[tauri::command(async)]
fn current_region(
    path: String,
    sheet: String,
    row: u32,
    col: u32,
    store: State<'_, Shared>,
) -> Result<excel::Block, String> {
    let books = store.0.lock().unwrap();
    let book = books
        .get(&path)
        .ok_or_else(|| format!("workbook not open: {path}"))?;
    book.current_region(&sheet, row, col)
}

/// Where Ctrl+arrow lands: the edge of the run of filled cells in that
/// direction, or the nearest filled cell when the one it starts from is blank.
/// A null reply is nothing ahead at all, which the grid reads as the sheet's own
/// edge — the stored range stops at the last cell the file mentions.
#[tauri::command(async)]
fn edge_cell(
    path: String,
    sheet: String,
    row: u32,
    col: u32,
    dir: String,
    store: State<'_, Shared>,
) -> Result<Option<excel::CellPos>, String> {
    let books = store.0.lock().unwrap();
    let book = books
        .get(&path)
        .ok_or_else(|| format!("workbook not open: {path}"))?;
    book.edge_cell(&sheet, row, col, &dir)
}

#[tauri::command(async)]
fn close_workbook(path: String, store: State<'_, Shared>) {
    store.0.lock().unwrap().remove(&path);
}

/// Reads embedded images out of a workbook, as base64, keyed by part.
///
/// Asked for in one call per sheet rather than one per picture: a design
/// document's sheet can float fifty screenshots, and every call would re-open
/// the archive and walk its directory to reach a single entry. The parts the
/// archive does not hold are left out rather than failing the call — a picture
/// that cannot be read is simply not drawn, and the cells still are.
#[tauri::command(async)]
fn read_media(path: String, parts: Vec<String>) -> Result<HashMap<String, String>, String> {
    let file = std::fs::File::open(&path).map_err(|e| e.to_string())?;
    let mut zip =
        zip::ZipArchive::new(BufReader::new(file)).map_err(|e| e.to_string())?;
    let mut out = HashMap::with_capacity(parts.len());
    for part in parts {
        // A lookup in the archive's own directory, so a crafted name cannot
        // escape it.
        let Ok(mut entry) = zip.by_name(&part) else {
            continue;
        };
        let mut bytes = Vec::new();
        if entry.read_to_end(&mut bytes).is_ok() {
            out.insert(part, base64(&bytes));
        }
    }
    Ok(out)
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
    // An update that was interrupted leaves files named for exactly this; a
    // healthy install has none, and nothing else in the folder is touched.
    update::sweep_leftovers();

    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_updater::Builder::new().build())
        .manage(Arc::new(Store::default()))
        .invoke_handler(tauri::generate_handler![
            open_workbook,
            get_cells,
            search_workbook,
            sheet_layout,
            column_values,
            filter_rows,
            axis_values,
            current_region,
            edge_cell,
            read_media,
            close_workbook,
            update::update_status,
            update::update_check,
            update::update_install,
            update::update_restart
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
