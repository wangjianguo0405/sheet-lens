mod cond;
mod drawing;
mod excel;
mod fonts;
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

    let names: Vec<String> = info.sheets.iter().map(|sheet| sheet.name.clone()).collect();
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
    let book = books.get(&path).ok_or_else(|| format!("workbook not open: {path}"))?;
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
    let book = books.get(&path).ok_or_else(|| format!("workbook not open: {path}"))?;
    let matcher = Matcher::new(&query, case_sensitive, whole_cell);
    Ok(book.search(sheet.as_deref(), &matcher))
}

#[tauri::command(async)]
fn sheet_layout(path: String, sheet: String, store: State<'_, Shared>) -> Result<SheetLayout, String> {
    let mut books = store.0.lock().unwrap();
    let book = books
        .get_mut(&path)
        .ok_or_else(|| format!("workbook not open: {path}"))?;
    book.layout(&sheet)
}

/// How this reader writes a short date — `yyyy/m/d` on a Japanese Windows,
/// `m/d/yyyy` on an American one. Excel's built-in short-date formats are drawn
/// in it rather than in any code the file carries, and only the webview knows
/// the reader's locale, so it sends the pattern once at startup.
///
/// Synchronous, unlike the rest: it stores a string, and it has to land before
/// the workbook that the next message opens is parsed — an `async` command
/// would go to the thread pool and could be overtaken.
#[tauri::command]
fn set_short_date(pattern: String) {
    layout::set_short_date(pattern);
}

#[tauri::command(async)]
fn column_values(
    path: String,
    sheet: String,
    col: u32,
    first_row: u32,
    last_row: u32,
    criteria: Vec<FilterCriterion>,
    query: Option<String>,
    store: State<'_, Shared>,
) -> Result<Vec<ValueCount>, String> {
    let books = store.0.lock().unwrap();
    let book = books.get(&path).ok_or_else(|| format!("workbook not open: {path}"))?;
    let criteria: Vec<(u32, Vec<String>)> = criteria
        .into_iter()
        .map(|criterion| (criterion.col, criterion.values))
        .collect();
    book.column_values(
        &sheet,
        col,
        first_row,
        last_row,
        &criteria,
        query.as_deref(),
        MAX_VALUES,
    )
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
    let book = books.get(&path).ok_or_else(|| format!("workbook not open: {path}"))?;
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
    let book = books.get(&path).ok_or_else(|| format!("workbook not open: {path}"))?;
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
    let book = books.get(&path).ok_or_else(|| format!("workbook not open: {path}"))?;
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
    let book = books.get(&path).ok_or_else(|| format!("workbook not open: {path}"))?;
    book.edge_cell(&sheet, row, col, &dir)
}

#[tauri::command(async)]
fn close_workbook(path: String, store: State<'_, Shared>) {
    store.0.lock().unwrap().remove(&path);
}

/// One level of a folder, for the file tree.
///
/// A level at a time rather than the whole tree: a root can be a checkout
/// hundreds of folders deep, and the tree only asks for a folder when the
/// reader opens it.
#[derive(serde::Serialize)]
#[serde(rename_all = "camelCase")]
struct DirEntry {
    name: String,
    path: String,
    dir: bool,
}

/// The extensions the reader can open. Mirrors `EXCEL_EXTENSIONS` on the
/// frontend: the tree lists folders beside these and nothing else, so a folder
/// full of `.svn` bookkeeping or lock files reads as what it is to this app.
const WORKBOOK_EXTENSIONS: [&str; 5] = ["xlsx", "xlsm", "xls", "xlsb", "ods"];

#[tauri::command(async)]
fn list_directory(path: String) -> Result<Vec<DirEntry>, String> {
    let entries = std::fs::read_dir(&path).map_err(|error| format!("{path} — {error}"))?;
    let mut out = Vec::new();
    for entry in entries.flatten() {
        let name = entry.file_name().to_string_lossy().into_owned();
        // `metadata` rather than `file_type`: a link to a folder should open
        // as one, and a link is not a folder to `file_type`.
        let Ok(meta) = entry.metadata() else {
            continue;
        };
        let dir = meta.is_dir();
        if !dir {
            let workbook = std::path::Path::new(&name)
                .extension()
                .and_then(|ext| ext.to_str())
                .is_some_and(|ext| WORKBOOK_EXTENSIONS.contains(&ext.to_ascii_lowercase().as_str()));
            if !workbook {
                continue;
            }
        }
        out.push(DirEntry {
            name,
            path: entry.path().to_string_lossy().into_owned(),
            dir,
        });
    }
    // Explorer's own order: folders first, then the workbooks, each by name.
    // A case-insensitive compare, so `Alpha` and `alpha` sit together.
    out.sort_by(|a, b| {
        b.dir
            .cmp(&a.dir)
            .then_with(|| a.name.to_lowercase().cmp(&b.name.to_lowercase()))
    });
    Ok(out)
}

/// How much of a root one file-name search walks before it stops: enough to
/// cover a checkout, near enough nothing that a root pointed at by mistake
/// cannot hold the reader waiting.
const MAX_SEARCH_DIRS: usize = 5_000;
const MAX_SEARCH_RESULTS: usize = 200;
const MAX_SEARCH_DEPTH: usize = 12;

/// One workbook a folder search found.
#[derive(serde::Serialize)]
#[serde(rename_all = "camelCase")]
struct FoundFile {
    name: String,
    path: String,
    /// The root the search reached it under.
    root: String,
    /// The folder it sits in, relative to that root.
    folder: String,
}

/// Finds workbooks by name under the given roots.
///
/// The tree lists a folder only once it is opened, so it cannot help find a
/// workbook three levels down. This walks instead, past the dot-folders a
/// working copy keeps its bookkeeping in, and gives up at a fixed number of
/// folders and results rather than walking however large the root turns out
/// to be.
#[tauri::command(async)]
fn search_tree(roots: Vec<String>, query: String) -> Vec<FoundFile> {
    let needle = query.trim().to_lowercase();
    let mut out = Vec::new();
    if needle.is_empty() {
        return out;
    }
    let mut budget = MAX_SEARCH_DIRS;
    for root in &roots {
        walk_workbooks(
            std::path::Path::new(root),
            std::path::Path::new(root),
            &needle,
            0,
            &mut budget,
            &mut out,
        );
        if budget == 0 || out.len() >= MAX_SEARCH_RESULTS {
            break;
        }
    }
    // Grouped by the folder they sit in, which is what a result row shows
    // after the name, then by name — the order the tree itself would list them.
    out.sort_by(|a, b| {
        a.folder
            .to_lowercase()
            .cmp(&b.folder.to_lowercase())
            .then_with(|| a.name.to_lowercase().cmp(&b.name.to_lowercase()))
    });
    out
}

fn walk_workbooks(
    root: &std::path::Path,
    dir: &std::path::Path,
    needle: &str,
    depth: usize,
    budget: &mut usize,
    out: &mut Vec<FoundFile>,
) {
    if *budget == 0 || depth > MAX_SEARCH_DEPTH || out.len() >= MAX_SEARCH_RESULTS {
        return;
    }
    *budget -= 1;
    let Ok(entries) = std::fs::read_dir(dir) else {
        // A folder this account cannot read is passed over, not reported: the
        // search is a sweep of the roots, and one refusal is not its answer.
        return;
    };
    for entry in entries.flatten() {
        if *budget == 0 || out.len() >= MAX_SEARCH_RESULTS {
            return;
        }
        let name = entry.file_name().to_string_lossy().into_owned();
        if name.starts_with('.') {
            continue;
        }
        let Ok(meta) = entry.metadata() else {
            continue;
        };
        let path = entry.path();
        if meta.is_dir() {
            walk_workbooks(root, &path, needle, depth + 1, budget, out);
            continue;
        }
        let workbook = std::path::Path::new(&name)
            .extension()
            .and_then(|ext| ext.to_str())
            .is_some_and(|ext| WORKBOOK_EXTENSIONS.contains(&ext.to_ascii_lowercase().as_str()));
        if !workbook || !name.to_lowercase().contains(needle) {
            continue;
        }
        let folder = path
            .strip_prefix(root)
            .ok()
            .and_then(|rel| rel.parent())
            .map(|rel| rel.to_string_lossy().into_owned())
            .unwrap_or_default();
        out.push(FoundFile {
            name,
            path: path.to_string_lossy().into_owned(),
            root: root.to_string_lossy().into_owned(),
            folder,
        });
    }
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
    let mut zip = zip::ZipArchive::new(BufReader::new(file)).map_err(|e| e.to_string())?;
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
fn base64(bytes: &[u8]) -> String {
    const ALPHABET: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
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
            set_short_date,
            column_values,
            filter_rows,
            axis_values,
            current_region,
            edge_cell,
            read_media,
            close_workbook,
            list_directory,
            search_tree,
            fonts::list_fonts,
            update::update_status,
            update::update_check,
            update::update_install,
            update::update_restart
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A scratch folder for one test, named so parallel runs cannot collide.
    fn scratch(name: &str) -> std::path::PathBuf {
        let dir = std::env::temp_dir().join(format!("sheetlens-{name}-{}", std::process::id()));
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(&dir).unwrap();
        dir
    }

    /// Folders first, then only the workbooks, each run by name — the order
    /// Explorer itself lists a folder in.
    #[test]
    fn lists_folders_before_workbooks_and_only_workbooks() {
        let dir = scratch("list-directory");
        std::fs::create_dir(dir.join("zeta")).unwrap();
        std::fs::create_dir(dir.join("Alpha")).unwrap();
        std::fs::write(dir.join("b.XLSX"), b"").unwrap();
        std::fs::write(dir.join("a.ods"), b"").unwrap();
        std::fs::write(dir.join("notes.txt"), b"").unwrap();
        std::fs::write(dir.join("no-extension"), b"").unwrap();

        let entries = list_directory(dir.to_string_lossy().into_owned()).unwrap();
        let listed: Vec<(String, bool)> = entries.iter().map(|e| (e.name.clone(), e.dir)).collect();
        assert_eq!(
            listed,
            vec![
                ("Alpha".to_owned(), true),
                ("zeta".to_owned(), true),
                ("a.ods".to_owned(), false),
                ("b.XLSX".to_owned(), false),
            ]
        );
        // The full path comes back, so the tree can ask for the next level.
        assert!(entries[0].path.ends_with("Alpha"));
        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn a_folder_that_is_not_there_is_an_error() {
        let dir = std::env::temp_dir().join("sheetlens-no-such-folder-for-sure");
        let _ = std::fs::remove_dir_all(&dir);
        assert!(list_directory(dir.to_string_lossy().into_owned()).is_err());
    }

    /// A search reaches past the levels the tree has not opened, matches names
    /// without case, and leaves the bookkeeping dot-folders alone.
    #[test]
    fn search_tree_walks_the_roots_for_workbooks() {
        let dir = scratch("search-tree");
        std::fs::create_dir_all(dir.join("project").join("design")).unwrap();
        std::fs::create_dir(dir.join(".svn")).unwrap();
        std::fs::write(dir.join("project").join("design").join("Cost Report.xlsx"), b"").unwrap();
        std::fs::write(dir.join("project").join("cost summary.ods"), b"").unwrap();
        std::fs::write(dir.join("project").join("notes.txt"), b"").unwrap();
        std::fs::write(dir.join(".svn").join("cost base.xlsx"), b"").unwrap();

        let roots = vec![dir.to_string_lossy().into_owned()];
        let found = search_tree(roots.clone(), "COST".to_owned());
        let listed: Vec<(&str, &str)> = found.iter().map(|f| (f.name.as_str(), f.folder.as_str())).collect();
        assert_eq!(listed.len(), 2);
        assert_eq!(listed[0], ("cost summary.ods", "project"));
        assert_eq!(listed[1].0, "Cost Report.xlsx");
        assert!(listed[1].1.ends_with("design"), "found in {:?}", listed[1].1);
        // Each result names its root, so the tree can hang it where it lives.
        assert!(found.iter().all(|f| f.root == roots[0]));
        std::fs::remove_dir_all(&dir).unwrap();
    }

    /// Nothing typed is no walk at all.
    #[test]
    fn search_tree_with_no_query_finds_nothing() {
        let dir = scratch("search-empty");
        std::fs::write(dir.join("a.xlsx"), b"").unwrap();
        assert!(search_tree(vec![dir.to_string_lossy().into_owned()], "  ".to_owned()).is_empty());
        std::fs::remove_dir_all(&dir).unwrap();
    }
}
