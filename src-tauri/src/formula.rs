//! The formula language a workbook writes: tokens, a syntax tree, and enough of
//! an evaluator to answer with.
//!
//! Two callers ask questions of the same syntax. A conditional format's rule is
//! a formula whose references are measured from the rule's own anchor; a cell
//! whose value the file left for Excel to work out is a formula whose references
//! are the cells as written. They differ only in what `Context` answers, so the
//! parser, the tree and the evaluation live here and both of them share it.

use std::collections::HashSet;

use chrono::{Datelike, NaiveDate};

/// Excel's sheet bounds, which `OFFSET` has to stay inside.
const SHEET_ROWS: i64 = 1_048_576;
const SHEET_COLS: i64 = 16_384;

/// The value a cell holds, as far as a comparison is concerned. A cell the sheet
/// never mentions is `Blank` rather than missing: that is what an empty cell is
/// to a formula.
#[derive(Clone, Debug, PartialEq)]
pub(crate) enum Value {
    Number(f64),
    Text(String),
    Bool(bool),
    Blank,
    /// Something that has no value at all — a formula that reaches itself, say.
    /// Arithmetic over one has no answer, which is how the cell that carries it
    /// stays blank instead of being given a number the sheet never had.
    Error(&'static str),
}

/// What a formula is evaluated against: the cell it belongs to, the anchor its
/// relative references are measured from, and the values it can read.
pub(crate) trait Context {
    /// The cell being asked about — what `ROW()` and `COLUMN()` answer.
    fn here(&self) -> (u32, u32);
    /// Where a relative reference counts from. A rule's anchor is its own; a
    /// cell's formula counts from the cell, which makes `A1` mean A1.
    fn anchor(&self) -> (u32, u32);
    /// The value of a cell, on this formula's own sheet (`None`) or the one a
    /// reference names.
    fn value_in(&self, sheet: Option<&str>, row: u32, col: u32) -> Value;
    /// A rectangle of cells, row by row, both corners included. `from` and `to`
    /// are already in sheet order, whichever way the formula wrote them.
    fn area_in(&self, sheet: Option<&str>, from: (u32, u32), to: (u32, u32)) -> Vec<Value>;
    /// What `CELL("filename")` reports — the workbook's own path and the sheet
    /// named by the reference, or this formula's sheet when it names none.
    /// A call site with no workbook behind it has no name to give.
    fn filename(&self, _sheet: Option<&str>) -> Option<String> {
        None
    }
    /// How many rows and columns a sheet uses, when the call site can tell.
    /// A whole-column reference is clamped to it: the cells past the end are
    /// empty, and walking a million of them answers what walking none does.
    fn used_bounds(&self, _sheet: Option<&str>) -> Option<(u32, u32)> {
        None
    }
}

/// A cell with nothing to read beside it: the formula is evaluated for where it
/// stands, and every reference it takes comes back `Blank`. This is what a
/// call site that cannot resolve references yet uses — the conditional format
/// rules bring their own context, and the workbook's own values grow out of
/// this one once the references can be followed.
pub(crate) struct At {
    pub row: u32,
    pub col: u32,
}

impl Context for At {
    fn here(&self) -> (u32, u32) {
        (self.row, self.col)
    }
    fn anchor(&self) -> (u32, u32) {
        (self.row, self.col)
    }
    fn value_in(&self, _sheet: Option<&str>, _row: u32, _col: u32) -> Value {
        Value::Blank
    }
    fn area_in(&self, _sheet: Option<&str>, _from: (u32, u32), _to: (u32, u32)) -> Vec<Value> {
        // Nothing can be read, so a rectangle has no cells to offer: the range
        // functions then see an empty set, which is what a sheet with nothing
        // in it would give them.
        Vec::new()
    }
}

// ---------------------------------------------------------------------------
// The formula
// ---------------------------------------------------------------------------

/// One corner of a reference, as the formula spells it. An absolute row or
/// column addresses the sheet directly; a relative one is the distance from the
/// anchor the context gives.
#[derive(Clone, Copy, Debug, PartialEq)]
pub(crate) struct Ref {
    pub row: i64,
    pub col: i64,
    pub row_absolute: bool,
    pub col_absolute: bool,
}

#[derive(Clone, Debug)]
pub(crate) enum Formula {
    Number(f64),
    Text(String),
    Bool(bool),
    /// A cell reference, and the sheet it names when it names one.
    Cell(Option<Box<str>>, Ref),
    /// A rectangle: `A1:B2`, both corners as written, and the sheet it names.
    Area(Option<Box<str>>, Ref, Ref),
    /// An argument that is not there: `IF(a,b,)` names an empty third one,
    /// which Excel reads as zero.
    Missing,
    Call(String, Vec<Formula>),
    Negate(Box<Formula>),
    Binary(Op, Box<Formula>, Box<Formula>),
}

#[derive(Clone, Copy, PartialEq, Debug)]
pub(crate) enum Op {
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
    pub(crate) fn parse(text: &str) -> Option<Formula> {
        let tokens = tokenize(text)?;
        let mut parser = Parser { tokens, at: 0 };
        let formula = parser.expression()?;
        (parser.at == parser.tokens.len()).then_some(formula)
    }

    /// Every column this can read, given that a relative reference moves with
    /// the cell it is evaluated for. A rectangle is taken to reach every column
    /// there is: only the rules ask, and a rule spanning a range is rare enough
    /// that reading the whole sheet is the cheaper answer.
    pub(crate) fn columns(&self, width: u32, out: &mut HashSet<u32>) {
        match self {
            Formula::Cell(_, at) => add_columns(at, width, out),
            Formula::Area(..) => {
                for col in 0..width {
                    out.insert(col);
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

    /// Whether this reads another cell's value — what a formula cannot be
    /// answered without once its references can be followed.
    pub(crate) fn reads_a_cell(&self) -> bool {
        match self {
            Formula::Cell(..) | Formula::Area(..) => true,
            Formula::Call(_, args) => args.iter().any(Formula::reads_a_cell),
            Formula::Negate(inner) => inner.reads_a_cell(),
            Formula::Binary(_, left, right) => left.reads_a_cell() || right.reads_a_cell(),
            _ => false,
        }
    }

    pub(crate) fn eval(&self, context: &dyn Context) -> Option<Value> {
        match self {
            Formula::Number(value) => Some(Value::Number(*value)),
            Formula::Text(text) => Some(Value::Text(text.clone())),
            Formula::Bool(value) => Some(Value::Bool(*value)),
            // An omitted argument is a zero, as it is in Excel.
            Formula::Missing => Some(Value::Number(0.0)),
            Formula::Cell(sheet, at) => {
                let (row, col) = locate(at, context)?;
                Some(context.value_in(sheet.as_deref(), row, col))
            }
            // A rectangle is not a value: only a function takes one, and it asks
            // for the cells rather than for this.
            Formula::Area(..) => None,
            Formula::Negate(inner) => Some(Value::Number(-number(&inner.eval(context)?)?)),
            Formula::Binary(op, left, right) => binary(*op, left.eval(context)?, right.eval(context)?),
            Formula::Call(name, args) => call(name, args, context),
        }
    }
}

fn add_columns(at: &Ref, width: u32, out: &mut HashSet<u32>) {
    if at.col_absolute {
        if at.col >= 0 {
            out.insert(at.col as u32);
        }
    } else {
        for step in 0..width as i64 {
            if at.col + step >= 0 {
                out.insert((at.col + step) as u32);
            }
        }
    }
}

/// Where a reference points, with a relative row and column laid on the anchor.
fn locate(at: &Ref, context: &dyn Context) -> Option<(u32, u32)> {
    let (here_row, here_col) = context.here();
    let (anchor_row, anchor_col) = context.anchor();
    Some((
        address(at.row, at.row_absolute, here_row, anchor_row)?,
        address(at.col, at.col_absolute, here_col, anchor_col)?,
    ))
}

/// The values an argument stands for: a rectangle's cells, or the single value
/// anything else evaluates to.
fn spread(args: &[Formula], index: usize, context: &dyn Context) -> Option<Vec<Value>> {
    match args.get(index)? {
        Formula::Area(sheet, from, to) => {
            let (r0, c0) = locate(from, context)?;
            let (r1, c1) = locate(to, context)?;
            Some(context.area_in(sheet.as_deref(), (r0.min(r1), c0.min(c1)), (r0.max(r1), c0.max(c1))))
        }
        other => Some(vec![other.eval(context)?]),
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
    // An error on either side is the answer: nothing can be worked out of it.
    if let Value::Error(code) = &a {
        return Some(Value::Error(code));
    }
    if let Value::Error(code) = &b {
        return Some(Value::Error(code));
    }
    match op {
        Op::Add | Op::Subtract | Op::Multiply | Op::Divide | Op::Power => {
            let (a, b) = (number(&a)?, number(&b)?);
            if op == Op::Divide && b == 0.0 {
                return Some(Value::Error("#DIV/0!"));
            }
            if op == Op::Power && a == 0.0 && b < 0.0 {
                return Some(Value::Error("#DIV/0!"));
            }
            let value = match op {
                Op::Add => a + b,
                Op::Subtract => a - b,
                Op::Multiply => a * b,
                Op::Divide => a / b,
                _ => a.powf(b),
            };
            // A result that is not a number is Excel's error, and `ISERROR`
            // has to see it as one.
            if value.is_finite() {
                Some(Value::Number(value))
            } else {
                Some(Value::Error("#NUM!"))
            }
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

fn call(name: &str, args: &[Formula], context: &dyn Context) -> Option<Value> {
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
        "ISERROR" => Some(Value::Bool(matches!(arg(0)?, Value::Error(_)))),
        // `TYPE` answers what kind of value it was given. A blank cell counts
        // as a number, and a range is refused as Excel refuses one.
        "TYPE" => {
            if matches!(args.first(), Some(Formula::Area(..))) {
                return Some(Value::Number(16.0));
            }
            Some(Value::Number(match arg(0)? {
                Value::Number(_) | Value::Blank => 1.0,
                Value::Text(_) => 2.0,
                Value::Bool(_) => 4.0,
                Value::Error(_) => 16.0,
            }))
        }
        // `OFFSET` answers a reference rather than a value: the cell it names
        // is read here, so arithmetic sees what is in it and `TYPE` sees its
        // kind. A rectangle (height or width beyond one) has no value to carry
        // here and is left unanswered.
        "OFFSET" => {
            let (sheet, base) = match args.first()? {
                Formula::Cell(sheet, reference) => {
                    let Some(base) = locate(reference, context) else {
                        return Some(Value::Error("#REF!"));
                    };
                    (sheet.as_deref(), base)
                }
                Formula::Area(..) => return None,
                _ => return Some(Value::Error("#VALUE!")),
            };
            let down = count(1)?.trunc() as i64;
            let right = count(2)?.trunc() as i64;
            if args.len() > 3 {
                let height = count(3)?.trunc() as i64;
                let width = if args.len() > 4 { count(4)?.trunc() as i64 } else { 1 };
                if height < 1 || width < 1 {
                    return Some(Value::Error("#REF!"));
                }
                if height != 1 || width != 1 {
                    return None;
                }
            }
            let row = i64::from(base.0) + down;
            let col = i64::from(base.1) + right;
            if !(0..SHEET_ROWS).contains(&row) || !(0..SHEET_COLS).contains(&col) {
                return Some(Value::Error("#REF!"));
            }
            Some(context.value_in(sheet, row as u32, col as u32))
        }
        // `VLOOKUP` reads one column of a rectangle by finding the lookup
        // value in its first one. The correlation sheets ask only for exact
        // matches; approximate matching is modelled as the last row that is
        // not greater than the value, which is what Excel answers on the
        // sorted columns it is meant for.
        "VLOOKUP" => {
            let wanted = arg(0)?;
            let (sheet, from, to) = match args.get(1)? {
                Formula::Area(sheet, from, to) => {
                    let (Some(from), Some(to)) = (locate(from, context), locate(to, context)) else {
                        return Some(Value::Error("#REF!"));
                    };
                    (sheet.as_deref(), from, to)
                }
                Formula::Cell(sheet, at) => {
                    let Some(at) = locate(at, context) else {
                        return Some(Value::Error("#REF!"));
                    };
                    (sheet.as_deref(), at, at)
                }
                _ => return Some(Value::Error("#VALUE!")),
            };
            let (r0, r1) = (from.0.min(to.0), from.0.max(to.0));
            let (c0, c1) = (from.1.min(to.1), from.1.max(to.1));
            // A whole-column table reaches as far as the sheet does, not as
            // far as the reference says.
            let bounds = context.used_bounds(sheet);
            let (r1, c1) = match bounds {
                Some((rows, cols)) => (r1.min(rows.saturating_sub(1)), c1.min(cols.saturating_sub(1))),
                None => (r1, c1),
            };
            // With no sheet to measure, a whole-column table would be a walk
            // over a million empty cells; leave it unanswered instead.
            if bounds.is_none() && u64::from(r1 - r0 + 1) > 200_000 {
                return None;
            }
            if r0 > r1 || c0 > c1 {
                return Some(Value::Error("#N/A"));
            }
            let index = count(2)?.trunc();
            let columns = f64::from(c1 - c0 + 1);
            if index < 1.0 {
                return Some(Value::Error("#VALUE!"));
            }
            if index > columns {
                return Some(Value::Error("#REF!"));
            }
            let exact = if args.len() > 3 { !truth(&arg(3)?)? } else { false };
            let target = c0 + index as u32 - 1;
            let mut best: Option<u32> = None;
            for row in r0..=r1 {
                let key = context.value_in(sheet, row, c0);
                if exact {
                    if lookup_matches(&key, &wanted) {
                        return Some(context.value_in(sheet, row, target));
                    }
                } else if compare(&key, &wanted).is_some_and(|order| !order.is_gt()) {
                    best = Some(row);
                }
            }
            Some(match best {
                Some(row) => context.value_in(sheet, row, target),
                None => Value::Error("#N/A"),
            })
        }
        // `IFERROR` answers its value, or the second argument when that value
        // is an error. A function this cannot answer leaves nothing to catch.
        "IFERROR" => match args.first().and_then(|value| value.eval(context)) {
            Some(Value::Error(_)) => arg(1),
            Some(value) => Some(value),
            None => None,
        },
        // `CHAR` spells one character by its code — the documents use it to put
        // a line break inside a concatenation.
        "CHAR" => {
            let code = count(0)?.trunc();
            if !(1.0..=255.0).contains(&code) {
                return Some(Value::Error("#VALUE!"));
            }
            Some(match char::from_u32(code as u32) {
                Some(ch) => Value::Text(ch.to_string()),
                None => Value::Error("#VALUE!"),
            })
        }
        // `COUNTIFS` counts the rows where every range's own criterion holds.
        // The ranges have to be the same size; Excel refuses the call when
        // they are not.
        "COUNTIFS" => {
            if args.is_empty() || args.len() % 2 != 0 {
                return None;
            }
            let mut columns: Vec<Vec<Value>> = Vec::new();
            let mut criteria: Vec<Value> = Vec::new();
            for pair in args.chunks(2) {
                let index = columns.len() * 2;
                columns.push(spread(args, index, context)?);
                criteria.push(pair[1].eval(context)?);
            }
            if columns.iter().any(|column| column.len() != columns[0].len()) {
                return Some(Value::Error("#VALUE!"));
            }
            let hits = (0..columns[0].len())
                .filter(|&row| {
                    columns
                        .iter()
                        .zip(&criteria)
                        .all(|(column, rule)| criteria_holds(&column[row], rule))
                })
                .count();
            Some(Value::Number(hits as f64))
        }
        // `COUNTBLANK` counts what shows nothing: an empty cell, or text that
        // spells nothing.
        "COUNTBLANK" => {
            let blanks = spread(args, 0, context)?
                .iter()
                .filter(|value| match value {
                    Value::Blank => true,
                    Value::Text(text) => text.is_empty(),
                    _ => false,
                })
                .count();
            Some(Value::Number(blanks as f64))
        }
        // `SUBSTITUTE` replaces text exactly: every occurrence, or the one at
        // the given position counting from one.
        "SUBSTITUTE" => {
            let source = text(&arg(0)?);
            let old = text(&arg(1)?);
            let new = text(&arg(2)?);
            if old.is_empty() {
                return Some(Value::Text(source));
            }
            if args.len() > 3 {
                let wanted = count(3)?.trunc();
                if wanted < 1.0 {
                    return Some(Value::Error("#VALUE!"));
                }
                let mut out = String::new();
                let mut rest = source.as_str();
                let mut seen = 1i64;
                while let Some(at) = rest.find(&old) {
                    if seen == wanted as i64 {
                        out.push_str(&rest[..at]);
                        out.push_str(&new);
                        out.push_str(&rest[at + old.len()..]);
                        return Some(Value::Text(out));
                    }
                    out.push_str(&rest[..at + old.len()]);
                    rest = &rest[at + old.len()..];
                    seen += 1;
                }
                Some(Value::Text(source))
            } else {
                Some(Value::Text(source.replace(&old, &new)))
            }
        }
        "MOD" => {
            let (a, b) = (count(0)?, count(1)?);
            (b != 0.0).then(|| Value::Number(a - b * (a / b).floor()))
        }
        "ROW" => Some(Value::Number(context.here().0 as f64 + 1.0)),
        "COLUMN" => Some(Value::Number(context.here().1 as f64 + 1.0)),
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
        // `SEARCH` looks for text without caring about case and with Excel's
        // wildcards; `FIND` matches the text exactly. Both answer a 1-based
        // position, and `#VALUE!` when there is none.
        "SEARCH" | "FIND" => {
            let needle: Vec<char> = text(&arg(0)?).chars().collect();
            let haystack: Vec<char> = text(&arg(1)?).chars().collect();
            let start = if args.len() > 2 { count(2)? } else { 1.0 };
            let start = start.trunc();
            if !start.is_finite() || start < 1.0 {
                return Some(Value::Error("#VALUE!"));
            }
            let from = start as usize - 1;
            let found = (from..=haystack.len()).find(|&at| {
                if name == "SEARCH" {
                    search_at(&needle, &haystack[at..])
                } else {
                    haystack[at..].starts_with(&needle)
                }
            });
            Some(match found {
                Some(at) => Value::Number(at as f64 + 1.0),
                None => Value::Error("#VALUE!"),
            })
        }
        "DATE" => {
            let serial = date_serial(count(0)?.trunc(), count(1)?.trunc(), count(2)?.trunc());
            Some(match serial {
                Some(serial) => Value::Number(serial),
                None => Value::Error("#NUM!"),
            })
        }
        // The documents use `CELL` only to learn where the workbook lives, so
        // the other info types stay unanswered rather than guessed at.
        "CELL" => {
            if !text(&arg(0)?).eq_ignore_ascii_case("filename") {
                return None;
            }
            let sheet = match args.get(1) {
                Some(Formula::Cell(sheet, _)) => sheet.as_deref(),
                _ => None,
            };
            Some(Value::Text(context.filename(sheet)?))
        }
        "CONCATENATE" => {
            let mut out = String::new();
            for arg in args {
                out.push_str(&text(&arg.eval(context)?));
            }
            Some(Value::Text(out))
        }
        "SUM" | "MAX" | "MIN" => {
            let numbers = numbers_in(args, context)?;
            // Excel answers zero for a maximum or minimum over nothing, but a
            // session of positive numbers must not be folded from zero or the
            // smallest would always be that zero.
            let value = if numbers.is_empty() {
                0.0
            } else {
                match name {
                    "SUM" => numbers.iter().sum(),
                    "MAX" => numbers.iter().copied().fold(f64::NEG_INFINITY, f64::max),
                    _ => numbers.iter().copied().fold(f64::INFINITY, f64::min),
                }
            };
            Some(Value::Number(value))
        }
        "COUNTA" => {
            let mut count = 0.0;
            for index in 0..args.len() {
                for value in spread(args, index, context)? {
                    if value != Value::Blank {
                        count += 1.0;
                    }
                }
            }
            Some(Value::Number(count))
        }
        "COUNTIF" => {
            let values = spread(args, 0, context)?;
            let criteria = arg(1)?;
            let hits = values.iter().filter(|value| criteria_holds(value, &criteria)).count();
            Some(Value::Number(hits as f64))
        }
        "TEXT" => {
            let value = arg(0)?;
            let code = text(&arg(1)?);
            Some(Value::Text(format_as(&value, &code)?))
        }
        "INDIRECT" => {
            // The address is built at run time — `"A"&ROW()-1` is how a
            // generated sheet walks up its own column — so it is parsed here
            // rather than with the rest of the formula.
            match Formula::parse(&text(&arg(0)?))? {
                Formula::Cell(sheet, at) => {
                    let (row, col) = locate(&at, context)?;
                    Some(context.value_in(sheet.as_deref(), row, col))
                }
                _ => None,
            }
        }
        // A function this cannot answer is one it must not guess at.
        _ => None,
    }
}

/// The numbers a set of arguments holds: a rectangle contributes the numbers
/// among its cells, anything else its own value. Text is skipped rather than
/// refused, which is what Excel does inside SUM and MAX.
fn numbers_in(args: &[Formula], context: &dyn Context) -> Option<Vec<f64>> {
    let mut out = Vec::new();
    for index in 0..args.len() {
        for value in spread(args, index, context)? {
            if let Value::Number(number) = value {
                out.push(number);
            }
        }
    }
    Some(out)
}

/// Whether a value answers `COUNTIF`'s criterion: a value, `">=5"`, or text with
/// Excel's `*` and `?` wildcards in it.
///
/// COUNTIF does not count blanks the way a formula compares them: a numeric
/// criterion skips an empty cell even though `A1=0` is true for one, while
/// `""` counts it and an explicit `"="` does not. A criterion cell that is
/// itself empty is read as zero.
fn criteria_holds(value: &Value, criteria: &Value) -> bool {
    let zero = Value::Number(0.0);
    let criteria = match criteria {
        Value::Blank => &zero,
        other => other,
    };
    let Value::Text(rule) = criteria else {
        return match (value, criteria) {
            (Value::Blank, Value::Number(_)) => false,
            _ => compare(value, criteria).is_some_and(|order| order.is_eq()),
        };
    };
    const SIGNS: [&str; 6] = [">=", "<=", "<>", ">", "<", "="];
    let found = SIGNS
        .iter()
        .find_map(|sign| rule.strip_prefix(sign).map(|rest| (*sign, rest)));
    let (test, rest) = found.unwrap_or(("=", rule.as_str()));
    let wanted = match number_text(rest) {
        Some(wanted) => Value::Number(wanted),
        None => Value::Text(rest.to_owned()),
    };
    if matches!(value, Value::Blank) {
        return match (&wanted, test) {
            // An order or an equality against a number leaves an empty cell
            // out; "not equal" is the one criterion that counts it.
            (Value::Number(_), "<>") => true,
            (Value::Number(_), _) => false,
            // `""` counts empty cells; an explicit `"="` counts only text
            // that spells nothing.
            (Value::Text(text), "=") if text.is_empty() => found.is_none(),
            _ => compare(value, &wanted).is_some_and(|order| matches_order(test, order)),
        };
    }
    if test == "=" && matches!(&wanted, Value::Text(text) if text.contains(['*', '?'])) {
        let Value::Text(have) = value else {
            return false;
        };
        return wildcard(&text_string(&wanted), have);
    }
    match compare(value, &wanted) {
        Some(order) => matches_order(test, order),
        // Across kinds there is no order; only "not equal" can be answered.
        None => test == "<>",
    }
}

fn matches_order(test: &str, order: std::cmp::Ordering) -> bool {
    match test {
        "=" => order.is_eq(),
        "<>" => !order.is_eq(),
        ">" => order.is_gt(),
        ">=" => order.is_ge(),
        "<" => order.is_lt(),
        _ => order.is_le(),
    }
}

fn text_string(value: &Value) -> String {
    match value {
        Value::Text(text) => text.clone(),
        other => text(other),
    }
}

/// A number written out in the criterion, or `None` when it is text.
fn number_text(text: &str) -> Option<f64> {
    let trimmed = text.trim();
    (!trimmed.is_empty()).then(|| trimmed.parse().ok()).flatten()
}

/// Excel's own wildcards: `*` for a run of characters, `?` for one, both
/// ignoring case.
fn wildcard(pattern: &str, text: &str) -> bool {
    let pattern: Vec<char> = pattern.to_uppercase().chars().collect();
    let text: Vec<char> = text.to_uppercase().chars().collect();
    fn go(pattern: &[char], text: &[char]) -> bool {
        match pattern.first() {
            None => text.is_empty(),
            Some('*') => (0..=text.len()).any(|skip| go(&pattern[1..], &text[skip..])),
            Some('?') => !text.is_empty() && go(&pattern[1..], &text[1..]),
            Some(wanted) => text.first() == Some(wanted) && go(&pattern[1..], &text[1..]),
        }
    }
    go(&pattern, &text)
}

/// Exact `VLOOKUP` matching: text against text ignores case and honors the
/// same wildcards `COUNTIF` does; everything else compares by value.
fn lookup_matches(key: &Value, wanted: &Value) -> bool {
    if let (Value::Text(key), Value::Text(wanted)) = (key, wanted) {
        if wanted.contains(['*', '?']) {
            return wildcard(wanted, key);
        }
    }
    compare(key, wanted).is_some_and(|order| order.is_eq())
}

/// Whether `SEARCH`'s pattern matches at the start of this text: `*` for a run
/// of characters, `?` for one, `~` for the next character itself, and no case
/// to any of it. The rest of the text may be anything — the caller is looking
/// for a position, not for the whole string to match.
fn search_at(pattern: &[char], text: &[char]) -> bool {
    match pattern.first() {
        None => true,
        Some('~') => match pattern.get(1) {
            Some(literal) => {
                text.first().is_some_and(|ch| ch.eq_ignore_ascii_case(literal)) && search_at(&pattern[2..], &text[1..])
            }
            None => text.first().is_some_and(|ch| *ch == '~') && search_at(&pattern[1..], &text[1..]),
        },
        Some('*') => (0..=text.len()).any(|skip| search_at(&pattern[1..], &text[skip..])),
        Some('?') => !text.is_empty() && search_at(&pattern[1..], &text[1..]),
        Some(wanted) => {
            text.first().is_some_and(|ch| ch.eq_ignore_ascii_case(wanted)) && search_at(&pattern[1..], &text[1..])
        }
    }
}

/// The serial number Excel gives a date, measured against the 1900 system it
/// ships: month and day run past their own bounds into the next month or year,
/// years below 1900 count inside it, and everything before 1900-03-01 is one
/// serial lower than the calendar says, because Excel keeps a 1900-02-29 that
/// never happened.
fn date_serial(year: f64, month: f64, day: f64) -> Option<f64> {
    if !year.is_finite() || !month.is_finite() || !day.is_finite() {
        return None;
    }
    let mut year = year as i64;
    if (0..1900).contains(&year) {
        year += 1900;
    }
    let month = month as i64;
    let day = day as i64;
    if !(1900..=9999).contains(&year) || !(-1_200_000..=1_200_000).contains(&month) {
        return None;
    }
    if !(-4_000_000..=4_000_000).contains(&day) {
        return None;
    }
    let months = year.checked_mul(12)?.checked_add(month - 1)?;
    let first = NaiveDate::from_ymd_opt(
        i32::try_from(months.div_euclid(12)).ok()?,
        months.rem_euclid(12) as u32 + 1,
        1,
    )?;
    let date = first.checked_add_signed(chrono::TimeDelta::days(day - 1))?;
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 30)?;
    let mut serial = i64::from(date.num_days_from_ce()) - i64::from(epoch.num_days_from_ce());
    if date < NaiveDate::from_ymd_opt(1900, 3, 1)? {
        serial -= 1;
    }
    (serial >= 1).then_some(serial as f64)
}

/// `TEXT(value, "yyyy/mm/dd")`: a number rendered through a format code. The
/// codes are the number formats the rest of the app already reads, so a date is
/// the same serial turned the same way.
fn format_as(value: &Value, code: &str) -> Option<String> {
    let serial = number(value)?;
    let when = calamine::ExcelDateTime::new(serial, calamine::ExcelDateTimeType::DateTime, false);
    crate::numfmt::apply(&calamine::Data::DateTime(when), code)
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
        // A value that is not there at all cannot be true or false.
        Value::Error(_) => None,
    }
}

/// Excel's comparison looks at the two values' kinds before their contents: text
/// is never equal to a number, and an empty cell counts as `""` against text and
/// as zero against a number.
pub(crate) fn compare(a: &Value, b: &Value) -> Option<std::cmp::Ordering> {
    match (a, b) {
        (Value::Text(a), Value::Text(b)) => Some(a.to_uppercase().cmp(&b.to_uppercase())),
        (Value::Bool(a), Value::Bool(b)) => Some(a.cmp(b)),
        (Value::Bool(a), Value::Number(b)) | (Value::Number(b), Value::Bool(a)) => (*a as u8 as f64).partial_cmp(b),
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
        // Excel carries the code through a concatenation as its text.
        Value::Error(code) => (*code).to_owned(),
    }
}

pub(crate) fn number(value: &Value) -> Option<f64> {
    match value {
        Value::Number(value) => Some(*value),
        Value::Bool(value) => Some(if *value { 1.0 } else { 0.0 }),
        Value::Blank => Some(0.0),
        // Excel refuses a text-to-number coercion outright where it matters, and
        // inventing one would invent an answer with it.
        Value::Text(_) | Value::Error(_) => None,
    }
}

// ---------------------------------------------------------------------------
// Tokenizer and parser
// ---------------------------------------------------------------------------

#[derive(Clone, Debug, PartialEq)]
enum Token {
    Number(f64),
    Text(String),
    /// A cell, and the sheet it names when it names one.
    Cell(Option<String>, Ref),
    /// A rectangle, `A1:B2`, both corners as written.
    Area(Option<String>, Ref, Ref),
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
        // A sheet prefix belongs to the reference that follows it, and a
        // reference followed by `:` takes a second one into a rectangle.
        let prefix = sheet_name(&chars, at);
        let start = prefix.as_ref().map_or(at, |(_, next)| *next);
        if let Some((from, next)) = reference(&chars, start) {
            let sheet = prefix.map(|(name, _)| name);
            if chars.get(next) == Some(&':') {
                if let Some((to, end)) = reference(&chars, next + 1) {
                    out.push(Token::Area(sheet, from, to));
                    at = end;
                    continue;
                }
            }
            out.push(Token::Cell(sheet, from));
            at = next;
            continue;
        }
        // `C:C` names whole columns: the rectangle is as tall as the sheet,
        // and a caller that reads it clamps it to what the sheet uses.
        if let Some((from, to, end)) = column_area(&chars, start) {
            out.push(Token::Area(prefix.map(|(name, _)| name), from, to));
            at = end;
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
            while chars.get(at).is_some_and(|ch| ch.is_ascii_alphanumeric() || *ch == '_') {
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

/// A run of whole columns as a rectangle — `$C:$C`, `C:D` — reaching the
/// sheet's full height. Whoever reads the cells clamps it to the used range.
fn column_area(chars: &[char], at: usize) -> Option<(Ref, Ref, usize)> {
    let (from, next) = column_only(chars, at)?;
    if chars.get(next) != Some(&':') {
        return None;
    }
    let (mut to, end) = column_only(chars, next + 1)?;
    // The second corner is the sheet's last row, not its first.
    to.row = SHEET_ROWS - 1;
    Some((from, to, end))
}

/// One column name as a reference spells it without a row: `C`, `$C`.
fn column_only(chars: &[char], at: usize) -> Option<(Ref, usize)> {
    let mut cursor = at;
    let col_absolute = chars.get(cursor) == Some(&'$');
    if col_absolute {
        cursor += 1;
    }
    let letters = cursor;
    while chars.get(cursor).is_some_and(|ch| ch.is_ascii_alphabetic()) {
        cursor += 1;
        if cursor - letters > 3 {
            return None;
        }
    }
    if cursor == letters {
        return None;
    }
    // A digit or a `$` after the letters makes it a cell, which `reference`
    // has already read.
    if chars.get(cursor).is_some_and(|ch| ch.is_ascii_digit() || *ch == '$') {
        return None;
    }
    let name: String = chars[letters..cursor].iter().collect();
    let col = column_index(&name)?;
    if col >= SHEET_COLS {
        return None;
    }
    Some((
        Ref {
            row: 0,
            col,
            row_absolute: true,
            col_absolute,
        },
        cursor,
    ))
}

/// Reads `$G8`, `G$8`, `$G$8` or `G8` at `at`, and says where it ended.
///
/// Letters alone are a name — `MOD`, `TRUE` — so a cell is only a cell when a
/// row number follows it and no `(` does.
fn reference(chars: &[char], at: usize) -> Option<(Ref, usize)> {
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
    let letters: String = chars[letters..if row_absolute { digits - 1 } else { digits }]
        .iter()
        .collect();
    let row: i64 = chars[digits..cursor].iter().collect::<String>().parse().ok()?;
    if row < 1 {
        return None;
    }
    Some((
        Ref {
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

/// A sheet name followed by `!`, as `Sheet1!` or `'A sheet'!` — the doubled
/// quote inside the quoted form standing for one quote.
pub(crate) fn sheet_name(chars: &[char], at: usize) -> Option<(String, usize)> {
    let (name, cursor) = if chars.get(at) == Some(&'\'') {
        let mut cursor = at + 1;
        let mut name = String::new();
        loop {
            let ch = *chars.get(cursor)?;
            cursor += 1;
            if ch == '\'' {
                if chars.get(cursor) == Some(&'\'') {
                    name.push('\'');
                    cursor += 1;
                    continue;
                }
                break;
            }
            name.push(ch);
        }
        (name, cursor)
    } else {
        let mut cursor = at;
        while chars
            .get(cursor)
            .is_some_and(|ch| ch.is_alphanumeric() || matches!(ch, '_' | '.' | '-'))
        {
            cursor += 1;
        }
        (chars[at..cursor].iter().collect(), cursor)
    };
    if name.is_empty() || chars.get(cursor) != Some(&'!') {
        return None;
    }
    // Only a reference can follow the `!`: `A1!=1` is a comparison, not a sheet.
    let after = *chars.get(cursor + 1)?;
    (after.is_ascii_alphabetic() || after == '$').then_some((name, cursor + 1))
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
            Token::Cell(sheet, at) => Some(Formula::Cell(sheet.map(String::into_boxed_str), at)),
            Token::Area(sheet, from, to) => Some(Formula::Area(sheet.map(String::into_boxed_str), from, to)),
            Token::Name(name) if name == "TRUE" => Some(Formula::Bool(true)),
            Token::Name(name) if name == "FALSE" => Some(Formula::Bool(false)),
            Token::Name(name) => {
                self.eat(&Token::Open).then_some(())?;
                let mut args = Vec::new();
                if !self.eat(&Token::Close) {
                    loop {
                        // An argument may be left out: `IF(a,b,)` has three,
                        // the last of them empty.
                        if matches!(self.peek(), Some(Token::Comma) | Some(Token::Close)) {
                            args.push(Formula::Missing);
                        } else {
                            args.push(self.expression()?);
                        }
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

    fn number(value: f64) -> Value {
        Value::Number(value)
    }

    fn string(value: &str) -> Value {
        Value::Text(value.to_owned())
    }

    #[test]
    fn compares_the_way_excel_does() {
        assert_eq!(binary(Op::Equal, number(1.0), number(1.0)), Some(Value::Bool(true)));
        // Text is never equal to a number, but it does have an order against
        // another string, and that order ignores case.
        assert_eq!(binary(Op::Equal, string("A"), number(1.0)), Some(Value::Bool(false)));
        assert_eq!(binary(Op::Equal, string("a"), string("A")), Some(Value::Bool(true)));
        assert_eq!(binary(Op::NotEqual, Value::Blank, string("x")), Some(Value::Bool(true)));
        // Undecidable is not false: an ordering across kinds cannot be answered.
        assert_eq!(binary(Op::Greater, string("a"), number(1.0)), None);
    }

    #[test]
    fn minimums_and_maximums_do_not_start_from_zero() {
        assert_eq!(answer("MIN(5,3)"), Some(number(3.0)));
        assert_eq!(answer("MIN(5)"), Some(number(5.0)));
        assert_eq!(answer("MAX(-1,-2)"), Some(number(-1.0)));
        // Over nothing Excel answers zero.
        assert_eq!(answer("MIN(A1:A2)"), Some(number(0.0)));
    }

    /// Measured from Excel (`probe-char.ps1`).
    #[test]
    fn reads_characters_by_code_the_way_excel_does() {
        assert_eq!(answer("CHAR(10)"), Some(string("\n")));
        assert_eq!(answer("CHAR(65)"), Some(string("A")));
        assert_eq!(answer("CHAR(13)"), Some(string("\r")));
        assert_eq!(answer("CHAR(10.9)"), Some(string("\n")));
        assert_eq!(answer("CHAR(0)"), Some(Value::Error("#VALUE!")));
        assert_eq!(answer("CHAR(256)"), Some(Value::Error("#VALUE!")));
    }

    /// What a formula answers with nothing around it to read.
    fn answer(source: &str) -> Option<Value> {
        Formula::parse(source)?.eval(&At { row: 0, col: 0 })
    }

    /// These answers were measured from Excel: `probe-functions.ps1` writes the
    /// formulas into a workbook, recalculates it and reads back `.Value2`.
    #[test]
    fn searches_and_finds_the_way_excel_does() {
        assert_eq!(answer(r#"SEARCH("b","abc")"#), Some(number(2.0)));
        assert_eq!(answer(r#"SEARCH("B","abc")"#), Some(number(2.0)));
        assert_eq!(answer(r#"SEARCH("z","abc")"#), Some(Value::Error("#VALUE!")));
        assert_eq!(answer(r#"SEARCH("","abc")"#), Some(number(1.0)));
        assert_eq!(answer(r#"SEARCH("","abc",2)"#), Some(number(2.0)));
        assert_eq!(answer(r#"SEARCH("?b","abc")"#), Some(number(1.0)));
        assert_eq!(answer(r#"SEARCH("~?","a?b")"#), Some(number(2.0)));
        assert_eq!(answer(r#"SEARCH("*","abc")"#), Some(number(1.0)));
        assert_eq!(answer(r#"SEARCH("b","abc",4)"#), Some(Value::Error("#VALUE!")));
        assert_eq!(answer(r#"SEARCH("*b","abc")"#), Some(number(1.0)));
        assert_eq!(answer(r#"SEARCH("b*","abc")"#), Some(number(2.0)));
        // A fractional start counts from its whole part.
        assert_eq!(answer(r#"SEARCH("b","abc",2.9)"#), Some(number(2.0)));

        assert_eq!(answer(r#"FIND("b","abc")"#), Some(number(2.0)));
        assert_eq!(answer(r#"FIND("B","abc")"#), Some(Value::Error("#VALUE!")));
        assert_eq!(answer(r#"FIND("b","abc",4)"#), Some(Value::Error("#VALUE!")));
        assert_eq!(answer(r#"FIND("","abc")"#), Some(number(1.0)));
        // An empty needle may start one past the end.
        assert_eq!(answer(r#"FIND("","abc",4)"#), Some(number(4.0)));
        assert_eq!(answer(r#"FIND("b","abc",2.9)"#), Some(number(2.0)));
    }

    #[test]
    fn recognises_an_error_as_an_error() {
        assert_eq!(answer(r#"ISERROR(SEARCH("z","abc"))"#), Some(Value::Bool(true)));
        assert_eq!(answer(r#"ISERROR("abc")"#), Some(Value::Bool(false)));
        assert_eq!(answer("ISERROR(1/0)"), Some(Value::Bool(true)));
    }

    #[test]
    fn dates_become_the_serials_excel_shows() {
        let date = |y: f64, m: f64, d: f64| answer(&format!("DATE({y},{m},{d})"));
        assert_eq!(date(2026.0, 11.0, 25.0), Some(number(46351.0)));
        assert_eq!(date(2020.0, 13.0, 1.0), Some(number(44197.0)));
        assert_eq!(date(2020.0, 1.0, 0.0), Some(number(43830.0)));
        assert_eq!(date(119.0, 1.0, 1.0), Some(number(43466.0)));
        assert_eq!(date(1900.0, 1.0, 1.0), Some(number(1.0)));
        assert_eq!(date(1900.0, 2.0, 28.0), Some(number(59.0)));
        assert_eq!(date(1900.0, 3.0, 1.0), Some(number(61.0)));
        assert_eq!(date(1970.0, 1.0, 1.0), Some(number(25569.0)));
        // The arguments are cut to whole numbers, as Excel cuts them.
        assert_eq!(date(2020.9, 1.9, 1.9), Some(number(43831.0)));
    }

    #[test]
    fn cell_filename_reports_the_workbook_and_sheet() {
        // Excel writes `C:\dir\[book.xlsx]Sheet1`, measured on a saved file.
        struct Named;
        impl Context for Named {
            fn here(&self) -> (u32, u32) {
                (0, 0)
            }
            fn anchor(&self) -> (u32, u32) {
                (0, 0)
            }
            fn value_in(&self, _sheet: Option<&str>, _row: u32, _col: u32) -> Value {
                Value::Blank
            }
            fn area_in(&self, _sheet: Option<&str>, _from: (u32, u32), _to: (u32, u32)) -> Vec<Value> {
                Vec::new()
            }
            fn filename(&self, sheet: Option<&str>) -> Option<String> {
                Some(format!(r"C:\work\[book.xlsx]{}", sheet.unwrap_or("Sheet1")))
            }
        }
        let formula = Formula::parse(r#"CELL("filename",A1)"#).expect("parse");
        assert_eq!(formula.eval(&Named), Some(string(r"C:\work\[book.xlsx]Sheet1")));
        let formula = Formula::parse(r#"FIND("]",CELL("filename",A1))"#).expect("parse");
        assert_eq!(formula.eval(&Named), Some(number(19.0)));
    }

    /// A handful of cells a formula can read, on the formula's own sheet or on
    /// one it names.
    struct Grid {
        here: (u32, u32),
        cells: Vec<(Option<&'static str>, u32, u32, Value)>,
    }

    impl Context for Grid {
        fn here(&self) -> (u32, u32) {
            self.here
        }
        fn anchor(&self) -> (u32, u32) {
            self.here
        }
        fn value_in(&self, sheet: Option<&str>, row: u32, col: u32) -> Value {
            self.cells
                .iter()
                .find(|(at, r, c, _)| *at == sheet && *r == row && *c == col)
                .map(|(_, _, _, value)| value.clone())
                .unwrap_or(Value::Blank)
        }
        fn area_in(&self, sheet: Option<&str>, from: (u32, u32), to: (u32, u32)) -> Vec<Value> {
            let to = match self.used_bounds(sheet) {
                Some((rows, cols)) => (to.0.min(rows.saturating_sub(1)), to.1.min(cols.saturating_sub(1))),
                None => to,
            };
            if from.0 > to.0 || from.1 > to.1 {
                return Vec::new();
            }
            let mut out = Vec::new();
            for row in from.0..=to.0 {
                for col in from.1..=to.1 {
                    out.push(self.value_in(sheet, row, col));
                }
            }
            out
        }
        fn used_bounds(&self, _sheet: Option<&str>) -> Option<(u32, u32)> {
            let rows = self.cells.iter().map(|(_, row, _, _)| *row + 1).max().unwrap_or(0);
            let cols = self.cells.iter().map(|(_, _, col, _)| *col + 1).max().unwrap_or(0);
            Some((rows, cols))
        }
    }

    /// Measured from Excel (`probe-offset.ps1`): A1=10, A2=20, B1="x", the
    /// formula sits at B9 with 5 in B8, and Sheet2!A2 holds 6.
    #[test]
    fn offsets_and_types_the_way_excel_does() {
        let grid = Grid {
            here: (8, 1),
            cells: vec![
                (None, 0, 0, number(10.0)),
                (None, 1, 0, number(20.0)),
                (None, 0, 1, string("x")),
                (None, 7, 1, number(5.0)),
                (Some("Sheet2"), 1, 0, number(6.0)),
            ],
        };
        let eval = |source: &str| Formula::parse(source).and_then(|formula| formula.eval(&grid));

        assert_eq!(eval("TYPE(1)"), Some(number(1.0)));
        assert_eq!(eval("TYPE(\"5\")"), Some(number(2.0)));
        assert_eq!(eval("TYPE(TRUE)"), Some(number(4.0)));
        assert_eq!(eval("TYPE(1/0)"), Some(number(16.0)));
        assert_eq!(eval("TYPE($A$3)"), Some(number(1.0)), "an empty cell is a number");

        assert_eq!(eval("OFFSET($A$1,1,0)"), Some(number(20.0)));
        assert_eq!(
            eval("OFFSET($A$1,1.9,0)"),
            Some(number(20.0)),
            "a fractional offset is cut to its whole part"
        );
        assert_eq!(eval("OFFSET($A$1,0,1)"), Some(string("x")));
        assert_eq!(eval("OFFSET(Sheet2!$A$1,1,0)"), Some(number(6.0)));
        assert_eq!(eval("OFFSET($A$1,-1,0)"), Some(Value::Error("#REF!")));
        assert_eq!(eval("TYPE(OFFSET($A$1,-1,0))"), Some(number(16.0)));
        // The corpus's counting pattern: the cell above, plus one.
        assert_eq!(
            eval("IF(TYPE(OFFSET(B9,-1,0))=1,OFFSET(B9,-1,0)+1,0)"),
            Some(number(6.0))
        );
        assert_eq!(eval("TYPE(OFFSET(B9,0,0))"), Some(number(1.0)));
        assert_eq!(eval("TYPE(A1:A2)"), Some(number(16.0)), "a range is refused");
    }

    /// Measured from Excel (`probe-vlookup.ps1`): A1:B3 is a/b/c against
    /// 10/20/30, D1:E3 is 1/2/3 against 100/200/300.
    #[test]
    fn looks_up_and_catches_errors_the_way_excel_does() {
        let grid = Grid {
            here: (0, 0),
            cells: vec![
                (None, 0, 0, string("a")),
                (None, 0, 1, number(10.0)),
                (None, 1, 0, string("b")),
                (None, 1, 1, number(20.0)),
                (None, 2, 0, string("c")),
                (None, 2, 1, number(30.0)),
                (None, 0, 3, number(1.0)),
                (None, 0, 4, number(100.0)),
                (None, 1, 3, number(2.0)),
                (None, 1, 4, number(200.0)),
                (None, 2, 3, number(3.0)),
                (None, 2, 4, number(300.0)),
            ],
        };
        let eval = |source: &str| Formula::parse(source).and_then(|formula| formula.eval(&grid));

        assert_eq!(eval(r#"VLOOKUP("b",$A$1:$B$3,2,FALSE)"#), Some(number(20.0)));
        assert_eq!(eval(r#"VLOOKUP("B",$A$1:$B$3,2,FALSE)"#), Some(number(20.0)));
        assert_eq!(eval(r#"VLOOKUP("z",$A$1:$B$3,2,FALSE)"#), Some(Value::Error("#N/A")));
        assert_eq!(eval(r#"VLOOKUP("b",$A$1:$B$3,0,FALSE)"#), Some(Value::Error("#VALUE!")));
        assert_eq!(eval(r#"VLOOKUP("b",$A$1:$B$3,3,FALSE)"#), Some(Value::Error("#REF!")));
        assert_eq!(eval(r#"VLOOKUP("b*",$A$1:$B$3,2,FALSE)"#), Some(number(20.0)));
        // A whole-column table reaches only as far as the sheet uses.
        assert_eq!(eval(r#"VLOOKUP("b",$A:$B,2,FALSE)"#), Some(number(20.0)));
        assert_eq!(eval(r#"VLOOKUP("a"&"",$A$1:$B$3,2,FALSE)"#), Some(number(10.0)));
        assert_eq!(eval(r#"VLOOKUP(2,$D$1:$E$3,2,FALSE)"#), Some(number(200.0)));
        assert_eq!(
            eval(r#"VLOOKUP("b",$A$1:$B$3,2.9,FALSE)"#),
            Some(number(20.0)),
            "a fractional column is cut to its whole part"
        );
        // Sorted columns answer the last value at or below the key.
        assert_eq!(eval(r#"VLOOKUP("bb",$A$1:$B$3,2,TRUE)"#), Some(number(20.0)));
        assert_eq!(eval(r#"VLOOKUP("b",$A$1:$B$3,2)"#), Some(number(20.0)));

        assert_eq!(
            eval(r#"IFERROR(VLOOKUP("z",$A$1:$B$3,2,FALSE),"none")"#),
            Some(string("none"))
        );
        assert_eq!(eval(r#"IFERROR(1/0,"x")"#), Some(string("x")));
        assert_eq!(eval(r#"IFERROR("ok","x")"#), Some(string("ok")));

        // A whole column in a range function reaches the used cells only.
        assert_eq!(eval("MAX($E:$E)"), Some(number(300.0)));
    }

    /// Measured from Excel (`probe-countifs.ps1`): A is 1/2/3, B is x/y/x,
    /// C is a/b/a, D1 is empty, D2 is 1, D3 is `=""`.
    #[test]
    fn counts_conditionally_and_substitutes_the_way_excel_does() {
        let grid = Grid {
            here: (0, 0),
            cells: vec![
                (None, 0, 0, number(1.0)),
                (None, 1, 0, number(2.0)),
                (None, 2, 0, number(3.0)),
                (None, 0, 1, string("x")),
                (None, 1, 1, string("y")),
                (None, 2, 1, string("x")),
                (None, 0, 2, string("a")),
                (None, 1, 2, string("b")),
                (None, 2, 2, string("a")),
                (None, 1, 3, number(1.0)),
                (None, 2, 3, string("")),
            ],
        };
        let eval = |source: &str| Formula::parse(source).and_then(|formula| formula.eval(&grid));

        assert_eq!(eval(r#"COUNTIFS($A$1:$A$3,">1")"#), Some(number(2.0)));
        assert_eq!(eval(r#"COUNTIFS($A$1:$A$3,">1",$B$1:$B$3,"x")"#), Some(number(1.0)));
        assert_eq!(eval(r#"COUNTIFS($A$1:$A$3,">=2",$B$1:$B$3,"<>y")"#), Some(number(1.0)));
        assert_eq!(eval(r#"COUNTIFS($B$1:$B$3,"x*")"#), Some(number(2.0)));
        assert_eq!(
            eval(r#"COUNTIFS($A$1:$A$2,">0",$B$1:$B$3,"x")"#),
            Some(Value::Error("#VALUE!"))
        );
        assert_eq!(eval(r#"COUNTIFS($A:$A,">0")"#), Some(number(3.0)));
        assert_eq!(eval(r#"COUNTIFS($B:$B,"x",$C:$C,"a")"#), Some(number(2.0)));
        assert_eq!(
            eval(r#"COUNTIFS($A$1:$A$3,">1",$B$1:$B$3,"x",$C$1:$C$3,"a")"#),
            Some(number(1.0))
        );
        assert_eq!(eval(r#"COUNTIFS($A$1:$A$3,">1",$B$1:$B$3,"")"#), Some(number(0.0)));

        assert_eq!(eval(r#"SUBSTITUTE("abcabc","b","X")"#), Some(string("aXcaXc")));
        assert_eq!(eval(r#"SUBSTITUTE("abcabc","b","X",2)"#), Some(string("abcaXc")));
        assert_eq!(eval(r#"SUBSTITUTE("abc","","X")"#), Some(string("abc")));
        assert_eq!(eval(r#"SUBSTITUTE("abc","d","X")"#), Some(string("abc")));
        assert_eq!(eval(r#"SUBSTITUTE("abcabc","b","X",3)"#), Some(string("abcabc")));
        assert_eq!(eval(r#"SUBSTITUTE("abc","a","X",0)"#), Some(Value::Error("#VALUE!")));
        assert_eq!(eval(r#"SUBSTITUTE("ABCabc","A","x")"#), Some(string("xBCabc")));
        assert_eq!(eval(r#"SUBSTITUTE(12345,3,9)"#), Some(string("12945")));
        assert_eq!(eval(r#"SUBSTITUTE("a-b-c","-","")"#), Some(string("abc")));

        assert_eq!(eval(r#"COUNTBLANK($D$1:$D$3)"#), Some(number(2.0)));
    }

    /// Measured from Excel (`probe-criteria.ps1`): over blank, "", 0, 1 and
    /// "x", these are the counts `COUNTIF` gives.
    #[test]
    fn counts_blanks_the_way_excel_does() {
        let values = vec![Value::Blank, string(""), number(0.0), number(1.0), string("x")];
        let count = |criteria: Value| values.iter().filter(|value| criteria_holds(value, &criteria)).count();
        assert_eq!(count(number(0.0)), 1, "a numeric criterion skips blanks");
        assert_eq!(count(Value::Blank), 1, "an empty criterion cell is zero");
        assert_eq!(count(string("")), 2, "\"\" counts empty text and blanks");
        assert_eq!(count(string("<>x")), 4);
        assert_eq!(count(string(">0")), 1);
        assert_eq!(count(string("<1")), 1);
        assert_eq!(count(string("*")), 2, "a wildcard matches text only");
        assert_eq!(count(string("=")), 1, "an explicit = counts only empty text");
    }
}
