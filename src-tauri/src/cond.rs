//! Conditional formatting: the fill, font colour and border Excel paints over a
//! cell when a rule matches it.
//!
//! Three parts of the file have to be put together, which is why this is its own
//! module: which cells a rule covers (`<conditionalFormatting sqref>` in the
//! sheet), what it paints (`dxfId` into the workbook's `<dxfs>`), and *whether*
//! it matches — a formula, evaluated per cell with every relative reference in
//! it shifted from the top-left corner of the range it sits on. What comes out
//! is folded straight into the cell's style, so the renderer knows none of it.
//!
//! Only the rule types a formula expresses are read. Colour scales, data bars
//! and icon sets are drawings rather than formats, and a rule this cannot read
//! paints nothing rather than something wrong.

use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};

use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;
use quick_xml::XmlVersion;

use crate::layout::{
    attr, color_edge, color_of, entity_text, intern, parse_cell, parse_range, set_edge, Cache,
    CellBorder, CellStyle, MergeRect, SheetLayout, Side,
};

/// The most cells any one sheet's rules may be evaluated over. A rule written
/// across whole columns asks its question of a million rows, and only the
/// sheet's own extent is ever drawn.
const MAX_CONDITIONAL_CELLS: usize = 200_000;

/// What a `dxf` changes about a cell. Every field is optional and a `None` one
/// leaves whatever the cell already had — that is what "differential" means, and
/// why these are merged onto a style rather than replacing it.
#[derive(Clone, Default, PartialEq, Debug)]
pub struct Patch {
    pub fill: Option<String>,
    pub color: Option<String>,
    pub bold: Option<bool>,
    pub border: Option<CellBorder>,
}

/// One conditional-format rule, reduced to what this viewer can act on.
pub struct Rule {
    /// Lower runs first, and wins a property that two rules both set.
    priority: u32,
    /// A rule that fires with this set is the last one looked at for that cell.
    stop_if_true: bool,
    /// Index into the workbook's `<dxfs>`.
    patch: usize,
    /// The cells the rule covers. Each range is its own anchor: a rule's formula
    /// is written relative to the top-left of the range it sits on.
    ranges: Vec<MergeRect>,
    formula: Formula,
}

/// Reads a sheet's `<conditionalFormatting>` blocks.
///
/// A rule whose formula cannot be read is dropped here rather than at every
/// cell: there is nothing to evaluate it with, and a rule that never fires
/// paints nothing, which is what not being there looks like too.
pub fn rules(xml: &str) -> Vec<Rule> {
    let mut reader = XmlReader::from_str(xml);
    let mut buf = Vec::new();
    let mut out = Vec::new();

    let mut sqref: Vec<MergeRect> = Vec::new();
    let mut current: Option<(u32, bool, usize, String)> = None;
    let mut in_formula = false;
    let mut formula = String::new();

    loop {
        buf.clear();
        let event = reader.read_event_into(&mut buf);
        let empty = matches!(event, Ok(Event::Empty(_)));
        match event {
            Ok(Event::Start(element)) | Ok(Event::Empty(element)) => match element.local_name().as_ref() {
                b"conditionalFormatting" => {
                    sqref = sqref_of(attr(&element, b"sqref").as_deref());
                }
                // Only an expression rule carries a formula to evaluate, and
                // only a rule with a format has anything to paint.
                b"cfRule" if attr(&element, b"type").as_deref() == Some("expression") => {
                    current = attr(&element, b"dxfId")
                        .and_then(|v| v.parse::<usize>().ok())
                        .map(|patch| {
                            let priority = attr(&element, b"priority")
                                .and_then(|v| v.parse::<u32>().ok())
                                .unwrap_or(u32::MAX);
                            let stop = attr(&element, b"stopIfTrue")
                                .is_some_and(|v| v == "1" || v == "true");
                            (priority, stop, patch, String::new())
                        });
                }
                b"formula" if current.is_some() && !empty => {
                    in_formula = true;
                    formula.clear();
                }
                _ => {}
            },
            // A formula's `<` and `>` are written `&lt;` and `&gt;`, and each
            // arrives as an entity of its own.
            Ok(Event::Text(text)) if in_formula => {
                if let Ok(text) = text.xml_content(XmlVersion::default()) {
                    formula.push_str(&text);
                }
            }
            Ok(Event::GeneralRef(reference)) if in_formula => {
                if let Some(ch) = entity_text(&reference) {
                    formula.push(ch);
                }
            }
            Ok(Event::End(element)) => match element.local_name().as_ref() {
                b"formula" if in_formula => {
                    in_formula = false;
                    if let Some(current) = current.as_mut() {
                        current.3 = std::mem::take(&mut formula);
                    }
                }
                b"cfRule" => {
                    if let Some((priority, stop_if_true, patch, text)) = current.take() {
                        if let Some(formula) = Formula::parse(&text) {
                            out.push(Rule {
                                priority,
                                stop_if_true,
                                patch,
                                ranges: sqref.clone(),
                                formula,
                            });
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

/// `sqref` holds one or more A1 ranges, space separated.
fn sqref_of(sqref: Option<&str>) -> Vec<MergeRect> {
    sqref
        .unwrap_or_default()
        .split_whitespace()
        .filter_map(parse_range)
        .collect()
}

impl Rule {
    /// Every column the rule's formula can read, so the value pass knows which
    /// ones to keep. A relative reference reaches a column for every cell of its
    /// range, which is the range's own width either side of where it is written.
    fn columns(&self, out: &mut HashSet<u32>) {
        for rect in &self.ranges {
            self.formula.columns(rect.cols.max(1), out);
        }
    }
}

/// Reads the workbook's `<dxfs>`.
pub fn dxfs(xml: &str, theme: &[Option<String>]) -> Vec<Patch> {
    let mut reader = XmlReader::from_str(xml);
    let mut buf = Vec::new();
    let mut out = Vec::new();

    let mut in_dxfs = false;
    let mut patch: Option<Patch> = None;
    let mut part = Part::None;
    let mut edge: Option<Side> = None;
    let mut pattern_type: Option<String> = None;
    let mut fg: Option<String> = None;
    let mut bg: Option<String> = None;

    let finish = |patch: Option<Patch>,
                  pattern_type: Option<String>,
                  fg: Option<String>,
                  bg: Option<String>|
     -> Option<Patch> {
        patch.map(|mut patch| {
            patch.fill = dxf_fill(pattern_type.as_deref(), fg, bg);
            patch
        })
    };

    loop {
        buf.clear();
        let event = reader.read_event_into(&mut buf);
        let empty = matches!(event, Ok(Event::Empty(_)));
        match event {
            Ok(Event::Start(element)) | Ok(Event::Empty(element)) => {
                match element.local_name().as_ref() {
                    b"dxfs" if !empty => in_dxfs = true,
                    b"dxf" if in_dxfs => {
                        if empty {
                            out.push(Patch::default());
                        } else {
                            patch = Some(Patch::default());
                            part = Part::None;
                            edge = None;
                            pattern_type = None;
                            fg = None;
                            bg = None;
                        }
                    }
                    b"font" if patch.is_some() && !empty => part = Part::Font,
                    b"fill" if patch.is_some() && !empty => part = Part::Fill,
                    b"border" if patch.is_some() && !empty => {
                        part = Part::Border;
                        edge = None;
                        if let Some(patch) = patch.as_mut() {
                            patch.border = Some(CellBorder::default());
                        }
                    }
                    b"left" | b"right" | b"top" | b"bottom" if part == Part::Border => {
                        let border = patch.as_mut().and_then(|patch| patch.border.as_mut());
                        edge = set_edge(&element, border, element.local_name().as_ref());
                    }
                    // `<color>` means the font's colour in one part of a `dxf`
                    // and one side's in another, so which part this is decides
                    // where it lands.
                    b"color" if part == Part::Border => {
                        let border = patch.as_mut().and_then(|patch| patch.border.as_mut());
                        color_edge(&element, border, edge, theme);
                    }
                    b"color" if part == Part::Font => {
                        if let Some(patch) = patch.as_mut() {
                            patch.color = color_of(&element, theme);
                        }
                    }
                    b"b" if part == Part::Font => {
                        if let Some(patch) = patch.as_mut() {
                            let off = matches!(attr(&element, b"val").as_deref(), Some("0" | "false"));
                            patch.bold = Some(!off);
                        }
                    }
                    b"patternFill" if part == Part::Fill => {
                        pattern_type = attr(&element, b"patternType");
                    }
                    b"fgColor" if part == Part::Fill => fg = color_of(&element, theme),
                    b"bgColor" if part == Part::Fill => bg = color_of(&element, theme),
                    _ => {}
                }
            }
            Ok(Event::End(element)) => match element.local_name().as_ref() {
                b"dxfs" => in_dxfs = false,
                b"dxf" if in_dxfs => {
                    if let Some(done) = finish(
                        patch.take(),
                        pattern_type.take(),
                        fg.take(),
                        bg.take(),
                    ) {
                        out.push(done);
                    }
                }
                b"font" | b"fill" | b"border" => part = Part::None,
                b"left" | b"right" | b"top" | b"bottom" if part == Part::Border => edge = None,
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    for patch in out.iter_mut() {
        if patch.border.as_ref().is_some_and(no_edges) {
            patch.border = None;
        }
    }
    out
}

/// Which part of a `<dxf>` the parser is inside.
#[derive(Clone, Copy, PartialEq)]
enum Part {
    None,
    Font,
    Fill,
    Border,
}

fn no_edges(border: &CellBorder) -> bool {
    border.top.is_none() && border.right.is_none() && border.bottom.is_none() && border.left.is_none()
}

/// A differential fill names its colour in `bgColor` and usually leaves the
/// pattern type out altogether — Excel paints that solid, which is the shape
/// behind every "light red fill" conditional format. An explicit `none` paints
/// nothing, and any of the hatched patterns has no flat colour to show.
fn dxf_fill(pattern_type: Option<&str>, fg: Option<String>, bg: Option<String>) -> Option<String> {
    match pattern_type.filter(|kind| !kind.is_empty()) {
        Some("none") => None,
        Some("solid") => fg.or(bg),
        Some(_) => None,
        None => bg.or(fg),
    }
}

/// The value a cell holds, as far as a comparison is concerned. A cell the sheet
/// never mentions is `Blank` rather than missing: that is what an empty cell is
/// to a formula.
#[derive(Clone, Debug, PartialEq)]
enum Value {
    Number(f64),
    Text(String),
    Bool(bool),
    Blank,
}

/// Pulls the values the rules' formulas refer to out of the sheet.
///
/// Only the columns a formula can reach are kept, so this costs the sheet's rows
/// times a handful of columns rather than a map of every cell in it.
fn values(xml: &str, columns: &HashSet<u32>, shared: &[String]) -> HashMap<(u32, u32), Value> {
    let mut out = HashMap::new();
    let mut reader = XmlReader::from_str(xml);
    let mut buf = Vec::new();

    let mut cell: Option<(u32, u32, Option<String>)> = None;
    let mut in_value = false;
    let mut in_inline = false;
    let mut in_phonetic = false;
    let mut cached = String::new();
    let mut inline = String::new();

    loop {
        buf.clear();
        let event = reader.read_event_into(&mut buf);
        let empty = matches!(event, Ok(Event::Empty(_)));
        match event {
            Ok(Event::Start(element)) | Ok(Event::Empty(element)) => {
                match element.local_name().as_ref() {
                    // A cell with no children holds nothing, and leaving it open
                    // would hand its neighbours' values to it.
                    b"c" if cell.is_none() && !empty => {
                        cell = attr(&element, b"r")
                            .and_then(|reference| parse_cell(&reference))
                            .map(|(row, col)| (row, col, attr(&element, b"t")));
                        cached.clear();
                        inline.clear();
                    }
                    b"is" if !empty => in_inline = true,
                    b"v" if cell.is_some() && !empty => in_value = true,
                    // A phonetic reading is written with the same `<t>` as the
                    // text itself and is no part of the cell's value.
                    b"rPh" if !empty => in_phonetic = true,
                    _ => {}
                }
            }
            Ok(Event::Text(text)) if cell.is_some() && (in_value || in_inline) && !in_phonetic => {
                if let Ok(text) = text.xml_content(XmlVersion::default()) {
                    if in_value {
                        cached.push_str(&text);
                    } else {
                        inline.push_str(&text);
                    }
                }
            }
            Ok(Event::GeneralRef(reference))
                if cell.is_some() && (in_value || in_inline) && !in_phonetic =>
            {
                if let Some(ch) = entity_text(&reference) {
                    if in_value {
                        cached.push(ch);
                    } else {
                        inline.push(ch);
                    }
                }
            }
            Ok(Event::End(element)) => match element.local_name().as_ref() {
                b"is" => in_inline = false,
                b"rPh" => in_phonetic = false,
                b"v" => in_value = false,
                b"c" => {
                    if let Some((row, col, kind)) = cell.take() {
                        if columns.contains(&col) {
                            let text = match kind.as_deref() {
                                Some("inlineStr") => &inline,
                                _ => &cached,
                            };
                            if let Some(value) = cell_value(kind.as_deref(), text, shared) {
                                out.insert((row, col), value);
                            }
                        }
                    }
                    in_value = false;
                    in_inline = false;
                }
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    out
}

fn cell_value(kind: Option<&str>, text: &str, shared: &[String]) -> Option<Value> {
    match kind {
        // A shared string is an index into the workbook's own table.
        Some("s") => {
            let index = text.trim().parse::<usize>().ok()?;
            Some(Value::Text(shared.get(index)?.clone()))
        }
        // The cached result of a formula returning text, and an inline string,
        // are both just the text.
        Some("str" | "inlineStr") => Some(Value::Text(text.to_owned())),
        Some("b") => Some(Value::Bool(text.trim() != "0" && !text.trim().is_empty())),
        // An error has no value to compare against.
        Some("e") => None,
        _ => {
            let text = text.trim();
            if text.is_empty() {
                return None;
            }
            Some(
                text.parse::<f64>()
                    .map(Value::Number)
                    .unwrap_or_else(|_| Value::Text(text.to_owned())),
            )
        }
    }
}

/// Folds a sheet's conditional formats into the styles of the cells they match.
///
/// Everything the rules could read is gathered in one pass of its own, over the
/// columns their formulas actually name — so a sheet with no rules, which is
/// most of them, costs a substring search and nothing else.
pub fn apply(xml: &str, cache: &Cache, layout: &mut SheetLayout) {
    if !xml.contains("conditionalFormatting") {
        return;
    }
    let patches = cache.patches();
    if patches.is_empty() {
        return;
    }
    let rules = rules(xml);
    if rules.is_empty() {
        return;
    }
    let mut columns = HashSet::new();
    for rule in &rules {
        rule.columns(&mut columns);
    }
    // A rule that names no cell — `MOD(ROW(),43)=4`, which is how a document
    // keeps a stripe every so many rows — needs nothing looked up, and the walk
    // that would look it up is the expensive half of this.
    let values = if columns.is_empty() {
        HashMap::new()
    } else {
        values(xml, &columns, cache.shared_strings())
    };
    fold(layout, patches, &rules, &values);
}

/// Folds every rule's patch onto the cells it matches.
fn fold(
    layout: &mut SheetLayout,
    patches: &[Patch],
    rules: &[Rule],
    values: &HashMap<(u32, u32), Value>,
) {
    let mut ordered: Vec<&Rule> = rules.iter().filter(|rule| rule.patch < patches.len()).collect();
    ordered.sort_by_key(|rule| rule.priority);
    if ordered.is_empty() {
        return;
    }

    // Which rules reach which row, in the order they are to be looked at. Built
    // up first so that a rule over a range the sheet never reaches costs nothing.
    let mut spans: BTreeMap<u32, Vec<Span<'_>>> = BTreeMap::new();
    let mut budget = MAX_CONDITIONAL_CELLS;
    'rules: for rule in &ordered {
        for rect in &rule.ranges {
            let first_row = rect.row;
            let last_row = rect.row.saturating_add(rect.rows).min(layout.rows);
            let first_col = rect.col;
            let last_col = rect.col.saturating_add(rect.cols).min(layout.cols);
            if first_col >= last_col || first_row >= last_row {
                continue;
            }
            let width = (last_col - first_col) as usize;
            if budget < width {
                break 'rules;
            }
            for row in first_row..last_row {
                budget -= width;
                spans.entry(row).or_default().push(Span {
                    rule,
                    anchor_row: rect.row,
                    anchor_col: rect.col,
                    first_col,
                    last_col,
                });
            }
        }
    }

    let mut ids: HashMap<CellStyle, u32> = layout
        .styles
        .iter()
        .enumerate()
        .map(|(index, style)| (style.clone(), index as u32))
        .collect();
    let mut rebuilt: Vec<(u32, Vec<[u32; 3]>)> = Vec::new();

    for (&row, spans) in &spans {
        // Every cell a rule fires on. Properties accumulate in priority order,
        // keeping the first rule to set each one: Excel lets a lower-priority
        // rule supply the property a higher-priority one leaves alone. A rule
        // that stops when it fires ends the matter for that cell.
        let mut fired: BTreeMap<u32, Patch> = BTreeMap::new();
        let mut stopped: BTreeSet<u32> = BTreeSet::new();
        for span in spans {
            for col in span.first_col..span.last_col {
                if stopped.contains(&col) {
                    continue;
                }
                let context = Context {
                    row,
                    col,
                    anchor_row: span.anchor_row,
                    anchor_col: span.anchor_col,
                    values,
                };
                if span.rule.formula.eval(&context) != Some(Value::Bool(true)) {
                    continue;
                }
                merge(fired.entry(col).or_default(), &patches[span.rule.patch]);
                if span.rule.stop_if_true {
                    stopped.insert(col);
                }
            }
        }
        if fired.is_empty() {
            continue;
        }

        // Each firing cell's final style, resolved before the row's own runs are
        // touched — `style_index` reads them.
        let mut styles: Vec<(u32, u32)> = Vec::with_capacity(fired.len());
        for (&col, patch) in &fired {
            let base = layout
                .style_index(row, col)
                .and_then(|index| layout.styles.get(index as usize))
                .cloned()
                .unwrap_or_default();
            let patched = patched(&base, patch);
            if patched.is_empty() {
                continue;
            }
            styles.push((col, intern(&mut layout.styles, &mut ids, patched)));
        }
        if styles.is_empty() {
            continue;
        }

        let existing = layout.cell_runs.remove(&row).unwrap_or_default();
        rebuilt.push((row, rebuild_row(&existing, &styles)));
    }

    for (row, runs) in rebuilt {
        layout.cell_runs.insert(row, runs);
    }
}

/// One rule's range, as it applies to a single row.
struct Span<'a> {
    rule: &'a Rule,
    anchor_row: u32,
    anchor_col: u32,
    first_col: u32,
    last_col: u32,
}

/// Adds whatever the next rule sets and the running patch does not.
fn merge(into: &mut Patch, next: &Patch) {
    if into.fill.is_none() {
        into.fill = next.fill.clone();
    }
    if into.color.is_none() {
        into.color = next.color.clone();
    }
    if into.bold.is_none() {
        into.bold = next.bold;
    }
    if into.border.is_none() {
        into.border = next.border.clone();
    } else if let (Some(into), Some(next)) = (into.border.as_mut(), next.border.as_ref()) {
        if into.top.is_none() {
            into.top = next.top.clone();
        }
        if into.right.is_none() {
            into.right = next.right.clone();
        }
        if into.bottom.is_none() {
            into.bottom = next.bottom.clone();
        }
        if into.left.is_none() {
            into.left = next.left.clone();
        }
    }
}

/// The cell's own style with the patch laid over it.
fn patched(base: &CellStyle, patch: &Patch) -> CellStyle {
    let mut style = base.clone();
    if let Some(fill) = patch.fill.clone() {
        style.fill = Some(fill);
    }
    if let Some(color) = patch.color.clone() {
        style.color = Some(color);
    }
    if let Some(bold) = patch.bold {
        style.bold = bold;
    }
    if let Some(patch) = patch.border.as_ref() {
        let border = style.border.get_or_insert_with(CellBorder::default);
        if patch.top.is_some() {
            border.top = patch.top.clone();
        }
        if patch.right.is_some() {
            border.right = patch.right.clone();
        }
        if patch.bottom.is_some() {
            border.bottom = patch.bottom.clone();
        }
        if patch.left.is_some() {
            border.left = patch.left.clone();
        }
    }
    style
}

/// Rewrites one row's runs so the cells a rule fired on carry its style.
///
/// Runs are split around the changed cells rather than expanded one entry per
/// column: a row can be sixteen thousand columns wide and a style covering all
/// of it is still a single run.
fn rebuild_row(existing: &[[u32; 3]], fired: &[(u32, u32)]) -> Vec<[u32; 3]> {
    let changed: BTreeMap<u32, u32> = fired.iter().copied().collect();
    let mut pieces: Vec<[u32; 3]> = Vec::with_capacity(existing.len() + changed.len());

    for run in existing {
        let (from, to, style) = (run[0], run[1], run[2]);
        let mut cursor = from;
        for (&col, &patched) in changed.range(from..=to) {
            if col > cursor {
                pieces.push([cursor, col - 1, style]);
            }
            pieces.push([col, col, patched]);
            cursor = col + 1;
        }
        if cursor <= to {
            pieces.push([cursor, to, style]);
        }
    }
    for &(col, patched) in fired {
        if !existing.iter().any(|run| col >= run[0] && col <= run[1]) {
            pieces.push([col, col, patched]);
        }
    }

    pieces.sort_unstable();
    let mut runs: Vec<[u32; 3]> = Vec::with_capacity(pieces.len());
    for piece in pieces {
        match runs.last_mut() {
            Some(last) if last[2] == piece[2] && last[1] + 1 >= piece[0] => {
                last[1] = last[1].max(piece[1]);
            }
            _ => runs.push(piece),
        }
    }
    runs
}

/// Where a formula is being evaluated: the cell itself, the corner its relative
/// references are measured from, and the sheet's values.
struct Context<'a> {
    row: u32,
    col: u32,
    anchor_row: u32,
    anchor_col: u32,
    values: &'a HashMap<(u32, u32), Value>,
}

// ---------------------------------------------------------------------------
// The formula
// ---------------------------------------------------------------------------

#[derive(Clone, Debug)]
enum Formula {
    Number(f64),
    Text(String),
    Bool(bool),
    /// A cell reference. An absolute row or column addresses the sheet directly;
    /// a relative one is the distance from the rule's anchor, laid on the cell
    /// being asked about.
    Cell {
        row: i64,
        col: i64,
        row_absolute: bool,
        col_absolute: bool,
    },
    Call(String, Vec<Formula>),
    Negate(Box<Formula>),
    Binary(Op, Box<Formula>, Box<Formula>),
}

#[derive(Clone, Copy, PartialEq, Debug)]
enum Op {
    Add,
    Subtract,
    Multiply,
    Divide,
    Power,
    Concat,
    Equal,
    NotEqual,
    Less,
    LessOrEqual,
    Greater,
    GreaterOrEqual,
}

impl Formula {
    fn parse(text: &str) -> Option<Formula> {
        let tokens = tokenize(text)?;
        let mut parser = Parser { tokens, at: 0 };
        let formula = parser.expression()?;
        (parser.at == parser.tokens.len()).then_some(formula)
    }

    /// Every column this can read, given that a relative reference moves with
    /// the cell it is evaluated for.
    fn columns(&self, width: u32, out: &mut HashSet<u32>) {
        match self {
            Formula::Cell {
                col, col_absolute, ..
            } => {
                if *col_absolute {
                    if *col >= 0 {
                        out.insert(*col as u32);
                    }
                } else {
                    for step in 0..width as i64 {
                        if *col + step >= 0 {
                            out.insert((*col + step) as u32);
                        }
                    }
                }
            }
            Formula::Call(_, args) => {
                for arg in args {
                    arg.columns(width, out);
                }
            }
            Formula::Negate(inner) => inner.columns(width, out),
            Formula::Binary(_, left, right) => {
                left.columns(width, out);
                right.columns(width, out);
            }
            _ => {}
        }
    }

    fn eval(&self, context: &Context) -> Option<Value> {
        match self {
            Formula::Number(value) => Some(Value::Number(*value)),
            Formula::Text(text) => Some(Value::Text(text.clone())),
            Formula::Bool(value) => Some(Value::Bool(*value)),
            Formula::Cell {
                row,
                col,
                row_absolute,
                col_absolute,
            } => {
                let row = address(*row, *row_absolute, context.row, context.anchor_row)?;
                let col = address(*col, *col_absolute, context.col, context.anchor_col)?;
                Some(
                    context
                        .values
                        .get(&(row, col))
                        .cloned()
                        .unwrap_or(Value::Blank),
                )
            }
            Formula::Negate(inner) => Some(Value::Number(-number(&inner.eval(context)?)?)),
            Formula::Binary(op, left, right) => binary(*op, left.eval(context)?, right.eval(context)?),
            Formula::Call(name, args) => call(name, args, context),
        }
    }
}

fn address(reference: i64, absolute: bool, at: u32, anchor: u32) -> Option<u32> {
    let value = if absolute {
        reference
    } else {
        at as i64 + reference - anchor as i64
    };
    u32::try_from(value).ok()
}

fn binary(op: Op, a: Value, b: Value) -> Option<Value> {
    match op {
        Op::Add | Op::Subtract | Op::Multiply | Op::Divide | Op::Power => {
            let (a, b) = (number(&a)?, number(&b)?);
            let value = match op {
                Op::Add => a + b,
                Op::Subtract => a - b,
                Op::Multiply => a * b,
                Op::Divide => a / b,
                _ => a.powf(b),
            };
            value.is_finite().then_some(Value::Number(value))
        }
        Op::Concat => Some(Value::Text(format!("{}{}", text(&a), text(&b)))),
        Op::Equal => Some(Value::Bool(compare(&a, &b).is_some_and(|order| order.is_eq()))),
        Op::NotEqual => Some(Value::Bool(!compare(&a, &b).is_some_and(|order| order.is_eq()))),
        // Two values with no order between them — text against a number — leave
        // the rule undecided, and a rule that cannot be decided does not fire.
        Op::Less | Op::LessOrEqual | Op::Greater | Op::GreaterOrEqual => {
            let order = compare(&a, &b)?;
            Some(Value::Bool(match op {
                Op::Less => order.is_lt(),
                Op::LessOrEqual => order.is_le(),
                Op::Greater => order.is_gt(),
                _ => order.is_ge(),
            }))
        }
    }
}

fn call(name: &str, args: &[Formula], context: &Context) -> Option<Value> {
    let arg = |index: usize| -> Option<Value> { args.get(index)?.eval(context) };
    let count = |index: usize| -> Option<f64> { number(&arg(index)?) };
    match name {
        "AND" | "OR" => {
            let all = name == "AND";
            let mut answer = all;
            let mut any = false;
            for arg in args {
                any = true;
                let value = truth(&arg.eval(context)?)?;
                if all && !value {
                    answer = false;
                }
                if !all && value {
                    answer = true;
                }
            }
            any.then_some(Value::Bool(answer))
        }
        "NOT" => Some(Value::Bool(!truth(&arg(0)?)?)),
        // The branch not taken is never evaluated, so `IF(A1=0,0,1/A1)` is safe.
        "IF" => {
            if truth(&arg(0)?)? {
                Some(arg(1)?)
            } else if args.len() > 2 {
                Some(arg(2)?)
            } else {
                Some(Value::Bool(false))
            }
        }
        "ISBLANK" => Some(Value::Bool(arg(0)? == Value::Blank)),
        "ISNUMBER" => Some(Value::Bool(matches!(arg(0)?, Value::Number(_)))),
        "ISTEXT" => Some(Value::Bool(matches!(arg(0)?, Value::Text(_)))),
        "MOD" => {
            let (a, b) = (count(0)?, count(1)?);
            (b != 0.0).then(|| Value::Number(a - b * (a / b).floor()))
        }
        "ROW" => Some(Value::Number(context.row as f64 + 1.0)),
        "COLUMN" => Some(Value::Number(context.col as f64 + 1.0)),
        "ABS" => Some(Value::Number(count(0)?.abs())),
        "INT" => Some(Value::Number(count(0)?.floor())),
        "ROUND" => {
            let digits = if args.len() > 1 { count(1)? } else { 0.0 };
            let scale = 10f64.powf(digits);
            Some(Value::Number((count(0)? * scale).round() / scale))
        }
        "LEN" => Some(Value::Number(text(&arg(0)?).chars().count() as f64)),
        "TRIM" => Some(Value::Text(text(&arg(0)?).trim().to_owned())),
        "UPPER" => Some(Value::Text(text(&arg(0)?).to_uppercase())),
        "LOWER" => Some(Value::Text(text(&arg(0)?).to_lowercase())),
        "LEFT" | "RIGHT" => {
            let length = if args.len() > 1 { count(1)? } else { 1.0 };
            let want = length.max(0.0) as usize;
            let value = text(&arg(0)?);
            let taken: String = if name == "LEFT" {
                value.chars().take(want).collect()
            } else {
                let skip = value.chars().count().saturating_sub(want);
                value.chars().skip(skip).collect()
            };
            Some(Value::Text(taken))
        }
        "MID" => {
            let value = text(&arg(0)?);
            let start = count(1)?.max(1.0) as usize - 1;
            let want = count(2)?.max(0.0) as usize;
            Some(Value::Text(value.chars().skip(start).take(want).collect()))
        }
        "CONCATENATE" => {
            let mut out = String::new();
            for arg in args {
                out.push_str(&text(&arg.eval(context)?));
            }
            Some(Value::Text(out))
        }
        // A function this cannot answer is one it must not guess at.
        _ => None,
    }
}

fn truth(value: &Value) -> Option<bool> {
    match value {
        Value::Bool(value) => Some(*value),
        Value::Number(value) => Some(*value != 0.0),
        Value::Blank => Some(false),
        Value::Text(text) => match text.to_uppercase().as_str() {
            "TRUE" => Some(true),
            "FALSE" => Some(false),
            _ => None,
        },
    }
}

/// Excel's comparison looks at the two values' kinds before their contents: text
/// is never equal to a number, and an empty cell counts as `""` against text and
/// as zero against a number.
fn compare(a: &Value, b: &Value) -> Option<std::cmp::Ordering> {
    match (a, b) {
        (Value::Text(a), Value::Text(b)) => Some(a.to_uppercase().cmp(&b.to_uppercase())),
        (Value::Bool(a), Value::Bool(b)) => Some(a.cmp(b)),
        (Value::Bool(a), Value::Number(b)) | (Value::Number(b), Value::Bool(a)) => {
            (*a as u8 as f64).partial_cmp(b)
        }
        (Value::Number(a), Value::Number(b)) => a.partial_cmp(b),
        (Value::Blank, Value::Blank) => Some(std::cmp::Ordering::Equal),
        (Value::Blank, Value::Text(b)) => Some("".cmp(b.as_str())),
        (Value::Text(a), Value::Blank) => Some(a.as_str().cmp("")),
        (Value::Blank, Value::Number(b)) => 0f64.partial_cmp(b),
        (Value::Number(a), Value::Blank) => a.partial_cmp(&0f64),
        // Text against a number or a boolean has no equality and no order; an
        // empty cell against a boolean is not `FALSE` either.
        _ => None,
    }
}

/// A value as text would spell it. Excel writes a whole number without a
/// trailing `.0`, and that is what `&` and `CONCATENATE` produce.
fn text(value: &Value) -> String {
    match value {
        Value::Text(text) => text.clone(),
        Value::Number(value) => {
            if value.fract() == 0.0 && value.abs() < 1e15 {
                format!("{}", *value as i64)
            } else {
                format!("{value}")
            }
        }
        Value::Bool(value) => value.to_string().to_uppercase(),
        Value::Blank => String::new(),
    }
}

fn number(value: &Value) -> Option<f64> {
    match value {
        Value::Number(value) => Some(*value),
        Value::Bool(value) => Some(if *value { 1.0 } else { 0.0 }),
        Value::Blank => Some(0.0),
        // Excel refuses a text-to-number coercion outright where it matters, and
        // inventing one would invent an answer with it.
        Value::Text(_) => None,
    }
}

// ---------------------------------------------------------------------------
// Tokenizer and parser
// ---------------------------------------------------------------------------

#[derive(Clone, Debug, PartialEq)]
enum Token {
    Number(f64),
    Text(String),
    Cell {
        row: i64,
        col: i64,
        row_absolute: bool,
        col_absolute: bool,
    },
    Name(String),
    Op(Op),
    Open,
    Close,
    Comma,
}

/// Everything a rule's formula may be made of. Anything else — another sheet
/// (`Sheet1!A1`), a range, an array — gives up here, which drops the rule.
fn tokenize(source: &str) -> Option<Vec<Token>> {
    let chars: Vec<char> = source.trim().chars().collect();
    // Excel's dialog box shows a formula with a leading `=`, which the file
    // itself does not write.
    let mut at = usize::from(chars.first() == Some(&'='));
    let mut out = Vec::new();

    while at < chars.len() {
        let ch = chars[at];
        if ch.is_whitespace() {
            at += 1;
            continue;
        }
        if ch == '"' {
            let mut text = String::new();
            at += 1;
            loop {
                let ch = *chars.get(at)?;
                at += 1;
                if ch == '"' {
                    // A doubled quote is one quote inside the string.
                    if chars.get(at) == Some(&'"') {
                        text.push('"');
                        at += 1;
                        continue;
                    }
                    break;
                }
                text.push(ch);
            }
            out.push(Token::Text(text));
            continue;
        }
        if let Some((token, next)) = reference(&chars, at) {
            out.push(token);
            at = next;
            continue;
        }
        if ch.is_ascii_digit() || ch == '.' {
            let start = at;
            while chars.get(at).is_some_and(|ch| ch.is_ascii_digit() || *ch == '.') {
                at += 1;
            }
            if matches!(chars.get(at), Some('e' | 'E')) {
                at += 1;
                if matches!(chars.get(at), Some('+' | '-')) {
                    at += 1;
                }
                while chars.get(at).is_some_and(|ch| ch.is_ascii_digit()) {
                    at += 1;
                }
            }
            let text: String = chars[start..at].iter().collect();
            out.push(Token::Number(text.parse().ok()?));
            continue;
        }
        if ch.is_ascii_alphabetic() || ch == '_' {
            let start = at;
            while chars
                .get(at)
                .is_some_and(|ch| ch.is_ascii_alphanumeric() || *ch == '_')
            {
                at += 1;
            }
            out.push(Token::Name(chars[start..at].iter().collect::<String>().to_uppercase()));
            continue;
        }
        at += 1;
        match ch {
            '(' => out.push(Token::Open),
            ')' => out.push(Token::Close),
            ',' => out.push(Token::Comma),
            '<' => {
                let op = match chars.get(at) {
                    Some('>') => {
                        at += 1;
                        Op::NotEqual
                    }
                    Some('=') => {
                        at += 1;
                        Op::LessOrEqual
                    }
                    _ => Op::Less,
                };
                out.push(Token::Op(op));
            }
            '>' => {
                let op = match chars.get(at) {
                    Some('=') => {
                        at += 1;
                        Op::GreaterOrEqual
                    }
                    _ => Op::Greater,
                };
                out.push(Token::Op(op));
            }
            '=' => out.push(Token::Op(Op::Equal)),
            '+' => out.push(Token::Op(Op::Add)),
            '-' => out.push(Token::Op(Op::Subtract)),
            '*' => out.push(Token::Op(Op::Multiply)),
            '/' => out.push(Token::Op(Op::Divide)),
            '^' => out.push(Token::Op(Op::Power)),
            '&' => out.push(Token::Op(Op::Concat)),
            _ => return None,
        }
    }
    Some(out)
}

/// Reads `$G8`, `G$8`, `$G$8` or `G8` at `at`, and says where it ended.
///
/// Letters alone are a name — `MOD`, `TRUE` — so a cell is only a cell when a
/// row number follows it and no `(` does.
fn reference(chars: &[char], at: usize) -> Option<(Token, usize)> {
    let mut cursor = at;
    let col_absolute = chars.get(cursor) == Some(&'$');
    if col_absolute {
        cursor += 1;
    }
    let letters = cursor;
    while chars.get(cursor).is_some_and(|ch| ch.is_ascii_alphabetic()) {
        cursor += 1;
    }
    if cursor == letters || cursor - letters > 3 {
        return None;
    }
    let row_absolute = chars.get(cursor) == Some(&'$');
    if row_absolute {
        cursor += 1;
    }
    let digits = cursor;
    while chars.get(cursor).is_some_and(|ch| ch.is_ascii_digit()) {
        cursor += 1;
    }
    // No row number: a name, unless a function call follows.
    if cursor == digits || chars.get(cursor) == Some(&'(') {
        return None;
    }
    let letters: String = chars[letters..if row_absolute { digits - 1 } else { digits }].iter().collect();
    let row: i64 = chars[digits..cursor].iter().collect::<String>().parse().ok()?;
    if row < 1 {
        return None;
    }
    Some((
        Token::Cell {
            // Stored zero based, so an absolute reference is already the sheet's
            // own coordinate and a relative one is already an offset.
            row: row - 1,
            col: column_index(&letters)?,
            row_absolute,
            col_absolute,
        },
        cursor,
    ))
}

fn column_index(letters: &str) -> Option<i64> {
    let mut index: i64 = 0;
    for ch in letters.chars() {
        let letter = ch.to_ascii_uppercase() as u8;
        if !letter.is_ascii_uppercase() {
            return None;
        }
        index = index * 26 + (letter - b'A' + 1) as i64;
    }
    (index > 0).then_some(index - 1)
}

struct Parser {
    tokens: Vec<Token>,
    at: usize,
}

impl Parser {
    fn peek(&self) -> Option<&Token> {
        self.tokens.get(self.at)
    }

    fn eat(&mut self, token: &Token) -> bool {
        if self.peek() == Some(token) {
            self.at += 1;
            return true;
        }
        false
    }

    fn is_op(&self, wanted: &[Op]) -> Option<Op> {
        match self.peek() {
            Some(Token::Op(op)) if wanted.contains(op) => Some(*op),
            _ => None,
        }
    }

    /// The lowest precedence first, as Excel's own operator table has it.
    fn expression(&mut self) -> Option<Formula> {
        let mut left = self.comparison()?;
        while let Some(op) = self.is_op(&[Op::Concat]) {
            self.at += 1;
            let right = self.comparison()?;
            left = Formula::Binary(op, Box::new(left), Box::new(right));
        }
        Some(left)
    }

    fn comparison(&mut self) -> Option<Formula> {
        const OPS: [Op; 6] = [
            Op::Equal,
            Op::NotEqual,
            Op::Less,
            Op::LessOrEqual,
            Op::Greater,
            Op::GreaterOrEqual,
        ];
        let mut left = self.sum()?;
        while let Some(op) = self.is_op(&OPS) {
            self.at += 1;
            let right = self.sum()?;
            left = Formula::Binary(op, Box::new(left), Box::new(right));
        }
        Some(left)
    }

    fn sum(&mut self) -> Option<Formula> {
        let mut left = self.product()?;
        while let Some(op) = self.is_op(&[Op::Add, Op::Subtract]) {
            self.at += 1;
            let right = self.product()?;
            left = Formula::Binary(op, Box::new(left), Box::new(right));
        }
        Some(left)
    }

    fn product(&mut self) -> Option<Formula> {
        let mut left = self.power()?;
        while let Some(op) = self.is_op(&[Op::Multiply, Op::Divide]) {
            self.at += 1;
            let right = self.power()?;
            left = Formula::Binary(op, Box::new(left), Box::new(right));
        }
        Some(left)
    }

    /// Right associative, which is how Excel reads `2^3^2`.
    fn power(&mut self) -> Option<Formula> {
        let left = self.unary()?;
        if self.eat(&Token::Op(Op::Power)) {
            let right = self.power()?;
            return Some(Formula::Binary(Op::Power, Box::new(left), Box::new(right)));
        }
        Some(left)
    }

    fn unary(&mut self) -> Option<Formula> {
        if self.eat(&Token::Op(Op::Subtract)) {
            return Some(Formula::Negate(Box::new(self.unary()?)));
        }
        if self.eat(&Token::Op(Op::Add)) {
            return self.unary();
        }
        self.primary()
    }

    fn primary(&mut self) -> Option<Formula> {
        let token = self.peek()?.clone();
        self.at += 1;
        match token {
            Token::Number(value) => Some(Formula::Number(value)),
            Token::Text(text) => Some(Formula::Text(text)),
            Token::Cell {
                row,
                col,
                row_absolute,
                col_absolute,
            } => Some(Formula::Cell {
                row,
                col,
                row_absolute,
                col_absolute,
            }),
            Token::Name(name) if name == "TRUE" => Some(Formula::Bool(true)),
            Token::Name(name) if name == "FALSE" => Some(Formula::Bool(false)),
            Token::Name(name) => {
                self.eat(&Token::Open).then_some(())?;
                let mut args = Vec::new();
                if !self.eat(&Token::Close) {
                    loop {
                        args.push(self.expression()?);
                        if self.eat(&Token::Comma) {
                            continue;
                        }
                        self.eat(&Token::Close).then_some(())?;
                        break;
                    }
                }
                Some(Formula::Call(name, args))
            }
            Token::Open => {
                let inner = self.expression()?;
                self.eat(&Token::Close).then_some(inner)
            }
            _ => None,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn patch() -> Patch {
        Patch {
            fill: Some("#BFBFBF".to_owned()),
            ..Default::default()
        }
    }

    /// Evaluates a formula with the range's corner on the cell itself, which is
    /// all a formula without references needs.
    fn eval(formula: &str, row: u32, col: u32, cells: &[((u32, u32), Value)]) -> Option<Value> {
        eval_at(formula, row, col, 0, 0, cells)
    }

    fn eval_at(
        formula: &str,
        row: u32,
        col: u32,
        anchor_row: u32,
        anchor_col: u32,
        cells: &[((u32, u32), Value)],
    ) -> Option<Value> {
        let values: HashMap<(u32, u32), Value> = cells.iter().cloned().collect();
        let context = Context {
            row,
            col,
            anchor_row,
            anchor_col,
            values: &values,
        };
        Formula::parse(formula)?.eval(&context)
    }

    fn text(value: &str) -> Value {
        Value::Text(value.to_owned())
    }

    /// The rule this whole feature was written for: a row greyed out because its
    /// judgement column says the row is out of scope. `B8:N161`, so the corner a
    /// relative reference is measured from is B8.
    fn judgement_row(row: u32, cells: &[((u32, u32), Value)]) -> Option<Value> {
        eval_at(
            r#"OR(AND($G8="対象外",$H8<>"対象"),$H8="対象外")"#,
            row,
            1,
            7,
            1,
            cells,
        )
    }

    #[test]
    fn reads_a_differential_fill() {
        // The first is how the review checklists write it: a `bgColor` with the
        // pattern type left out, which Excel still paints solid. The second is
        // the shape Excel itself writes for a named style.
        let xml = r#"<styleSheet><dxfs count="3">
            <dxf><fill><patternFill><bgColor theme="0" tint="-0.249946592608417"/></patternFill></fill></dxf>
            <dxf><font><color rgb="FF9C0006"/><b/></font><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>
            <dxf><fill><patternFill patternType="none"><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>
        </dxfs></styleSheet>"#;

        let patches = dxfs(xml, &[]);
        assert_eq!(patches.len(), 3);
        // White darkened by a quarter, which is the grey Excel paints — and the
        // same number its object model reported for these cells. An empty theme
        // still lands on white, since that is what slot 0 means.
        assert_eq!(patches[0].fill.as_deref(), Some("#BFBFBF"));
        assert_eq!(patches[1].fill.as_deref(), Some("#FFC7CE"));
        assert_eq!(patches[1].color.as_deref(), Some("#9C0006"));
        assert_eq!(patches[1].bold, Some(true));
        // An explicit "no pattern" paints nothing at all.
        assert_eq!(patches[2].fill, None);
    }

    #[test]
    fn reads_a_differential_border() {
        let xml = r#"<styleSheet><dxfs count="1"><dxf><border><top style="thin"><color indexed="64"/></top></border></dxf></dxfs></styleSheet>"#;
        let patches = dxfs(xml, &[]);
        let border = patches[0].border.as_ref().expect("a border");
        assert_eq!(border.top.as_ref().map(|edge| edge.line), Some("solid"));
        assert_eq!(
            border.top.as_ref().and_then(|edge| edge.color.as_deref()),
            Some("#000000")
        );
        assert!(border.bottom.is_none());
    }

    #[test]
    fn reads_an_expression_rule() {
        let xml = r#"<worksheet><conditionalFormatting sqref="B8:N161"><cfRule type="expression" dxfId="0" priority="5" stopIfTrue="1"><formula>OR(AND($G8="対象外",$H8&lt;&gt;"対象"),$H8="対象外")</formula></cfRule></conditionalFormatting></worksheet>"#;
        let rules = rules(xml);
        assert_eq!(rules.len(), 1);
        let rule = &rules[0];
        assert_eq!(rule.patch, 0);
        assert_eq!(rule.priority, 5);
        assert!(rule.stop_if_true);
        assert_eq!(rule.ranges.len(), 1);
        let range = &rule.ranges[0];
        assert_eq!((range.row, range.col, range.rows, range.cols), (7, 1, 154, 13));
        // `G` and `H`, which is what the value pass has to go and find.
        let mut columns = HashSet::new();
        rule.columns(&mut columns);
        assert_eq!(columns, HashSet::from([6, 7]));
    }

    #[test]
    fn keeps_several_ranges_on_one_rule() {
        let xml = r#"<worksheet><conditionalFormatting sqref="A1:A5 C1:C5"><cfRule type="expression" dxfId="2" priority="1"><formula>$A1&gt;0</formula></cfRule></conditionalFormatting></worksheet>"#;
        let rules = rules(xml);
        assert_eq!(rules[0].ranges.len(), 2);
        assert_eq!(rules[0].ranges[1].col, 2);
    }

    /// A rule this cannot evaluate must not be guessed at: a colour scale is a
    /// drawing, and a formula that reads another sheet has an answer here that
    /// is not the right one.
    #[test]
    fn drops_rules_it_cannot_read() {
        let xml = r#"<worksheet><conditionalFormatting sqref="A1:A5">
            <cfRule type="colorScale" priority="1"><colorScale/></cfRule>
            <cfRule type="expression" dxfId="0" priority="2"><formula>Sheet2!A1&gt;0</formula></cfRule>
            <cfRule type="expression" dxfId="7" priority="3"><formula>COUNTIF($A:$A,$A1)&gt;1</formula></cfRule>
        </conditionalFormatting></worksheet>"#;
        assert!(rules(xml).is_empty());
    }

    #[test]
    fn relative_references_follow_the_cell_they_are_evaluated_for() {
        // `$G8` is written at the range's corner, B8, so every row below reads
        // its own G — and the column stays put.
        let row = [((7u32, 6u32), text("対象")), ((9, 6), text("対象外"))];
        assert_eq!(judgement_row(7, &row), Some(Value::Bool(false)));
        assert_eq!(judgement_row(9, &row), Some(Value::Bool(true)));
        // A blank H is not "対象", which is what the second half of the rule
        // leans on.
        assert_eq!(
            eval_at(r#"$H8<>"対象""#, 7, 1, 7, 1, &row),
            Some(Value::Bool(true))
        );
        // A reference that lands off the sheet has no answer, so the rule does
        // not fire rather than matching a cell that is not there.
        assert_eq!(eval_at("$A1", 3, 1, 7, 1, &row), None);
    }

    #[test]
    fn an_out_of_scope_row_fires_and_an_in_scope_one_does_not() {
        assert_eq!(
            judgement_row(7, &[((7, 6), text("対象外"))]),
            Some(Value::Bool(true))
        );
        assert_eq!(
            judgement_row(7, &[((7, 6), text("対象"))]),
            Some(Value::Bool(false))
        );
        // H can name the row out of scope on its own.
        assert_eq!(
            judgement_row(7, &[((7, 6), text("対象")), ((7, 7), text("対象外"))]),
            Some(Value::Bool(true))
        );
    }

    #[test]
    fn compares_the_way_excel_does() {
        let number = |value: f64| Value::Number(value);
        assert_eq!(binary(Op::Equal, number(1.0), number(1.0)), Some(Value::Bool(true)));
        // Text is never equal to a number, but it does have an order against
        // another string, and that order ignores case.
        assert_eq!(binary(Op::Equal, text("A"), number(1.0)), Some(Value::Bool(false)));
        assert_eq!(binary(Op::Equal, text("a"), text("A")), Some(Value::Bool(true)));
        assert_eq!(binary(Op::NotEqual, Value::Blank, text("x")), Some(Value::Bool(true)));
        // Undecidable is not false: an ordering across kinds cannot be answered.
        assert_eq!(binary(Op::Greater, text("a"), number(1.0)), None);
    }

    #[test]
    fn arithmetic_functions_and_concatenation() {
        assert_eq!(
            eval("MOD(ROW(),43)=5", 47, 1, &[]),
            Some(Value::Bool(true))
        );
        assert_eq!(eval("ROW()", 46, 1, &[]), Some(Value::Number(47.0)));
        assert_eq!(eval(r#""a"&1"#, 0, 0, &[]), Some(text("a1")));
        assert_eq!(eval(r#"UPPER("abc")&"-2""#, 0, 0, &[]), Some(text("ABC-2")));
        assert_eq!(eval("ISBLANK($A1)", 3, 0, &[]), Some(Value::Bool(true)));
        assert_eq!(eval("-1+2*3", 0, 0, &[]), Some(Value::Number(5.0)));
        assert_eq!(eval("2^3^2", 0, 0, &[]), Some(Value::Number(512.0)));
        // An unknown function is undecidable rather than false.
        assert_eq!(eval("SEARCH(\"a\",$A1)", 3, 0, &[]), None);
    }

    #[test]
    fn splits_a_run_around_the_cells_a_rule_fires_on() {
        // One run of five columns, with two cells of it repainted.
        assert_eq!(
            rebuild_row(&[[0, 5, 7]], &[(1, 9), (3, 9)]),
            vec![[0, 0, 7], [1, 1, 9], [2, 2, 7], [3, 3, 9], [4, 5, 7]]
        );
        // A cell that had no style of its own still becomes a run, and two
        // neighbours that end up alike stay one.
        assert_eq!(
            rebuild_row(&[], &[(4, 9), (5, 9), (7, 8)]),
            vec![[4, 5, 9], [7, 7, 8]]
        );
    }

    #[test]
    fn folds_a_fill_onto_the_rows_a_rule_matches() {
        let mut layout = SheetLayout::empty();
        layout.rows = 4;
        layout.cols = 3;

        let rules = rules(
            r#"<conditionalFormatting sqref="A1:C4"><cfRule type="expression" dxfId="0" priority="1"><formula>$A1&gt;=2</formula></cfRule></conditionalFormatting>"#,
        );
        let values: HashMap<(u32, u32), Value> = (0..4)
            .map(|row| ((row, 0), Value::Number(row as f64 + 1.0)))
            .collect();
        fold(&mut layout, &[patch()], &rules, &values);

        let fill = |row: u32| {
            layout
                .style_index(row, 2)
                .and_then(|index| layout.styles.get(index as usize))
                .and_then(|style| style.fill.clone())
        };
        // The first row's value is below the threshold, so it keeps no style at
        // all rather than an empty one.
        assert_eq!(fill(0), None);
        assert_eq!(fill(1).as_deref(), Some("#BFBFBF"));
        assert_eq!(fill(3).as_deref(), Some("#BFBFBF"));
    }

    #[test]
    fn a_higher_priority_rule_keeps_the_property_it_sets() {
        let mut layout = SheetLayout::empty();
        layout.rows = 1;
        layout.cols = 1;

        // Both rules match. The later one is the lower priority, so the fill is
        // the one it never sets... and the font colour it does.
        let rules = rules(
            r#"<conditionalFormatting sqref="A1:A1">
                <cfRule type="expression" dxfId="0" priority="1"><formula>TRUE</formula></cfRule>
                <cfRule type="expression" dxfId="1" priority="2"><formula>TRUE</formula></cfRule>
            </conditionalFormatting>"#,
        );
        let patches = vec![
            Patch {
                fill: Some("#BFBFBF".to_owned()),
                ..Default::default()
            },
            Patch {
                fill: Some("#FFC7CE".to_owned()),
                color: Some("#9C0006".to_owned()),
                ..Default::default()
            },
        ];
        fold(&mut layout, &patches, &rules, &HashMap::new());

        let style = layout
            .style_index(0, 0)
            .and_then(|index| layout.styles.get(index as usize))
            .expect("a style");
        assert_eq!(style.fill.as_deref(), Some("#BFBFBF"));
        assert_eq!(style.color.as_deref(), Some("#9C0006"));
    }

    #[test]
    fn a_rule_that_stops_leaves_no_room_for_the_next_one() {
        let mut layout = SheetLayout::empty();
        layout.rows = 1;
        layout.cols = 1;

        let rules = rules(
            r#"<conditionalFormatting sqref="A1:A1">
                <cfRule type="expression" dxfId="0" priority="1" stopIfTrue="1"><formula>TRUE</formula></cfRule>
                <cfRule type="expression" dxfId="1" priority="2"><formula>TRUE</formula></cfRule>
            </conditionalFormatting>"#,
        );
        let patches = vec![
            Patch {
                fill: Some("#BFBFBF".to_owned()),
                ..Default::default()
            },
            Patch {
                color: Some("#9C0006".to_owned()),
                ..Default::default()
            },
        ];
        fold(&mut layout, &patches, &rules, &HashMap::new());

        let style = layout
            .style_index(0, 0)
            .and_then(|index| layout.styles.get(index as usize))
            .expect("a style");
        assert_eq!(style.fill.as_deref(), Some("#BFBFBF"));
        assert_eq!(style.color, None);
    }

    /// The workbook comes from `SHEETLENS_XLSX` rather than a literal: a dump
    /// that hardcodes a path ends up reading a different file than the one
    /// anyone is looking at.
    #[test]
    #[ignore = "reads a workbook from the developer's machine"]
    fn dump_conditional_formatting() {
        let Some(path) = std::env::var("SHEETLENS_XLSX").ok() else {
            return;
        };
        let book = crate::excel::load(&path).expect("open");
        let mut cache = crate::layout::Cache::default();
        for sheet in &book.info.sheets {
            let Ok(layout) = crate::layout::read_cached(&mut cache, &path, &sheet.name) else {
                continue;
            };
            // The fill of one column, as runs — the rows a rule covers read as
            // a handful of lines rather than one per row.
            let mut runs: Vec<(u32, u32, Option<String>)> = Vec::new();
            for row in 0..layout.rows {
                let fill = layout
                    .style_index(row, 1)
                    .and_then(|index| layout.styles.get(index as usize))
                    .and_then(|style| style.fill.clone());
                match runs.last_mut() {
                    Some(last) if last.2 == fill => last.1 = row,
                    _ => runs.push((row, row, fill)),
                }
            }
            if runs.iter().all(|run| run.2.is_none()) {
                continue;
            }
            println!("--- {} : {} runs in column B", sheet.name, runs.len());
            for (from, to, fill) in runs {
                println!(
                    "  rows {from}..{to}  fill {}",
                    fill.unwrap_or_else(|| "none".to_owned())
                );
            }
        }
    }
}
