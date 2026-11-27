//! The formula language a workbook writes: tokens, a syntax tree, and enough of
//! an evaluator to answer with.
//!
//! Two callers ask questions of the same syntax. A conditional format's rule is
//! a formula whose references are measured from the rule's own anchor; a cell
//! whose value the file left for Excel to work out is a formula whose references
//! are the cells as written. They differ only in what `Context` answers, so the
//! parser, the tree and the evaluation live here and both of them share it.

use std::collections::HashSet;

/// The value a cell holds, as far as a comparison is concerned. A cell the sheet
/// never mentions is `Blank` rather than missing: that is what an empty cell is
/// to a formula.
#[derive(Clone, Debug, PartialEq)]
pub(crate) enum Value {
    Number(f64),
    Text(String),
    Bool(bool),
    Blank,
}

/// What a formula is evaluated against: the cell it belongs to, the anchor its
/// relative references are measured from, and the values it can read.
pub(crate) trait Context {
    /// The cell being asked about — what `ROW()` and `COLUMN()` answer.
    fn here(&self) -> (u32, u32);
    /// Where a relative reference counts from. A rule's anchor is its own; a
    /// cell's formula counts from the cell, which makes `A1` mean A1.
    fn anchor(&self) -> (u32, u32);
    /// The value of a cell. One the sheet never mentions is `Blank`.
    fn value_at(&self, row: u32, col: u32) -> Value;
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
    fn value_at(&self, _row: u32, _col: u32) -> Value {
        Value::Blank
    }
}

// ---------------------------------------------------------------------------
// The formula
// ---------------------------------------------------------------------------

#[derive(Clone, Debug)]
pub(crate) enum Formula {
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
    /// the cell it is evaluated for.
    pub(crate) fn columns(&self, width: u32, out: &mut HashSet<u32>) {
        match self {
            Formula::Cell { col, col_absolute, .. } => {
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

    /// Whether this reads another cell's value, which nothing here can answer
    /// yet: a formula that only knows where it stands can be worked out, one
    /// that reaches for a neighbour is left as blank as the file left it.
    pub(crate) fn reads_a_cell(&self) -> bool {
        match self {
            Formula::Cell { .. } => true,
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
            Formula::Cell {
                row,
                col,
                row_absolute,
                col_absolute,
            } => {
                let (here_row, here_col) = context.here();
                let (anchor_row, anchor_col) = context.anchor();
                let row = address(*row, *row_absolute, here_row, anchor_row)?;
                let col = address(*col, *col_absolute, here_col, anchor_col)?;
                Some(context.value_at(row, col))
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
    }
}

pub(crate) fn number(value: &Value) -> Option<f64> {
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
    let letters: String = chars[letters..if row_absolute { digits - 1 } else { digits }]
        .iter()
        .collect();
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
}
