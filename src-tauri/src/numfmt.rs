//! Renders a formatted cell the way Excel would.
//!
//! Cells hold a raw value plus a format code. A date code turns a serial into
//! `2025年05月30日`; a number code turns it into `1,234.50` or `0 件`. A code
//! this cannot read returns `None`, so the caller falls back to a plain
//! rendering instead of showing something wrong.

use calamine::{Data, ExcelDateTime};
use chrono::{Datelike, NaiveDateTime, Timelike, Weekday};

/// Renders `value` with `code`, or `None` when the code is one this cannot read.
pub fn apply(value: &Data, code: &str) -> Option<String> {
    match value {
        Data::DateTime(datetime) => {
            if let Some(text) = date(datetime, code) {
                return Some(text);
            }
            // A date-typed cell with a number code shows the serial number.
            number(datetime.as_f64(), code)
        }
        Data::Float(value) => number(*value, code),
        Data::Int(value) => number(*value as f64, code),
        _ => None,
    }
}

/// Whether a code renders a date or a time, so a caller can tell the two
/// kinds of formatted cell apart.
pub fn is_date_code(code: &str) -> bool {
    has_date_field(first_section(code))
}

fn date(datetime: &ExcelDateTime, code: &str) -> Option<String> {
    let section = first_section(code);
    if !has_date_field(section) {
        return None;
    }
    let serial = datetime.as_f64();
    // Excel has no date before 1900; it shows an error, not a serial.
    if serial < 0.0 {
        return None;
    }
    // Serial zero is Excel's 1900-01-00, a day before the calendar starts.
    if serial < 1.0 {
        return Some(render_day_zero(serial, section));
    }
    let naive = datetime.as_datetime()?;
    Some(render(&naive, section, serial))
}

/// The part before the first unquoted `;`, which is the one dates use.
fn first_section(code: &str) -> &str {
    let mut quoted = false;
    for (index, ch) in code.char_indices() {
        match ch {
            '"' => quoted = !quoted,
            ';' if !quoted => return &code[..index],
            _ => {}
        }
    }
    code
}

/// Whether a section names a date or time field. Quoted text, escapes and
/// brackets say nothing about the kind of format.
fn has_date_field(section: &str) -> bool {
    let chars: Vec<char> = section.chars().collect();
    let mut index = 0;
    while index < chars.len() {
        match chars[index] {
            '"' => {
                index += 1;
                while index < chars.len() && chars[index] != '"' {
                    index += 1;
                }
                index += 1;
            }
            '\\' | '_' | '*' => index += 2,
            '[' => {
                let start = index + 1;
                let mut end = start;
                while end < chars.len() && chars[end] != ']' {
                    end += 1;
                }
                let inside: String = chars[start..end].iter().collect();
                if elapsed_spec(&inside).is_some() {
                    return true;
                }
                index = end + 1;
            }
            ch if "ymdhs".contains(ch.to_ascii_lowercase()) => return true,
            _ => index += 1,
        }
    }
    false
}

/// The fields a date format can print.
struct Parts {
    year: i32,
    month: u32,
    day: u32,
    hour: u32,
    minute: u32,
    second: u32,
    weekday: Weekday,
    /// Whole seconds since serial zero, which `[h]`, `[m]` and `[s]` count.
    elapsed: u64,
}

fn parts_of(value: &NaiveDateTime, serial: f64) -> Parts {
    Parts {
        year: value.year(),
        month: value.month(),
        day: value.day(),
        hour: value.hour(),
        minute: value.minute(),
        second: value.second(),
        weekday: value.weekday(),
        elapsed: elapsed_seconds(serial),
    }
}

fn elapsed_seconds(serial: f64) -> u64 {
    (serial * 86_400.0 + 0.5).max(0.0) as u64
}

fn render_day_zero(serial: f64, section: &str) -> String {
    let seconds = elapsed_seconds(serial);
    render_parts(
        &Parts {
            year: 1900,
            month: 1,
            day: 0,
            hour: ((seconds / 3_600) % 24) as u32,
            minute: ((seconds / 60) % 60) as u32,
            second: (seconds % 60) as u32,
            // Excel's serial zero falls on the Sunday before its day one.
            weekday: Weekday::Sun,
            elapsed: seconds,
        },
        section,
    )
}

fn render(value: &NaiveDateTime, section: &str, serial: f64) -> String {
    render_parts(&parts_of(value, serial), section)
}

fn render_parts(value: &Parts, section: &str) -> String {
    let chars: Vec<char> = section.chars().collect();
    let twelve_hour = contains_ignore_case(&chars, "AM/PM") || contains_ignore_case(&chars, "A/P");
    let mut out = String::new();
    let mut index = 0;
    let mut seen_hour = false;

    while index < chars.len() {
        let ch = chars[index];
        match ch {
            '"' => {
                index += 1;
                while index < chars.len() && chars[index] != '"' {
                    out.push(chars[index]);
                    index += 1;
                }
                index += 1;
            }
            '\\' => {
                if index + 1 < chars.len() {
                    out.push(chars[index + 1]);
                }
                index += 2;
            }
            // `_x` reserves a space and `*x` repeats a fill; neither matters here.
            '_' | '*' => index += 2,
            '[' => {
                let before = index;
                while index < chars.len() && chars[index] != ']' {
                    index += 1;
                }
                let inside: String = chars[before + 1..index].iter().collect();
                if let Some((letter, width)) = elapsed_spec(&inside) {
                    let total = value.elapsed;
                    let unit = match letter {
                        'h' => total / 3_600,
                        'm' => total / 60,
                        _ => total,
                    };
                    if width >= 2 {
                        out.push_str(&format!("{unit:02}"));
                    } else {
                        out.push_str(&unit.to_string());
                    }
                    // An `m` after an elapsed hour counts minutes, as it does
                    // after a clock hour.
                    if letter == 'h' {
                        seen_hour = true;
                    }
                }
                index += 1;
            }
            'y' | 'd' | 'h' | 's' | 'm' => {
                let width = run_length(&chars, index, ch);
                // `m` is minutes when an hour came before it or seconds follow,
                // and months otherwise.
                let minutes = ch == 'm' && (seen_hour || seconds_follow(&chars, index + width));
                out.push_str(&field(value, ch, width, minutes, twelve_hour));
                if ch == 'h' {
                    seen_hour = true;
                }
                index += width;
            }
            'A' | 'a' => {
                let rest: String = chars[index..].iter().collect();
                let upper = rest.to_uppercase();
                if upper.starts_with("AM/PM") {
                    out.push_str(if value.hour < 12 { "AM" } else { "PM" });
                    index += 5;
                } else if upper.starts_with("A/P") {
                    out.push_str(if value.hour < 12 { "A" } else { "P" });
                    index += 3;
                } else {
                    out.push(ch);
                    index += 1;
                }
            }
            other => {
                out.push(other);
                index += 1;
            }
        }
    }
    out
}

fn run_length(chars: &[char], start: usize, ch: char) -> usize {
    let mut width = 0;
    while start + width < chars.len() && chars[start + width] == ch {
        width += 1;
    }
    width
}

/// `[h]`, `[mm]`, `[ss]`: a total rather than a clock field. Only an all-letter
/// bracket is one; `[Red]` and `[Magenta]` name colours.
fn elapsed_spec(inside: &str) -> Option<(char, usize)> {
    let lower = inside.to_ascii_lowercase();
    let mut chars = lower.chars();
    let letter = chars.next()?;
    if !"hms".contains(letter) || !chars.all(|ch| ch == letter) {
        return None;
    }
    Some((letter, inside.chars().count()))
}

fn seconds_follow(chars: &[char], mut index: usize) -> bool {
    while index < chars.len() {
        match chars[index] {
            's' => return true,
            'h' | 'd' | 'y' | 'm' => return false,
            _ => index += 1,
        }
    }
    false
}

fn contains_ignore_case(chars: &[char], needle: &str) -> bool {
    let text: String = chars.iter().collect::<String>().to_uppercase();
    text.contains(needle)
}

fn field(value: &Parts, letter: char, width: usize, minutes: bool, twelve_hour: bool) -> String {
    match letter {
        'y' => {
            if width <= 2 {
                format!("{:02}", value.year.rem_euclid(100))
            } else {
                format!("{:04}", value.year)
            }
        }
        'd' => match width {
            1 => value.day.to_string(),
            2 => format!("{:02}", value.day),
            3 => weekday(value.weekday, true),
            _ => weekday(value.weekday, false),
        },
        'h' => {
            let hour = if twelve_hour {
                let hour = value.hour % 12;
                if hour == 0 {
                    12
                } else {
                    hour
                }
            } else {
                value.hour
            };
            if width >= 2 {
                format!("{hour:02}")
            } else {
                hour.to_string()
            }
        }
        's' => {
            if width >= 2 {
                format!("{:02}", value.second)
            } else {
                value.second.to_string()
            }
        }
        'm' if minutes => {
            if width >= 2 {
                format!("{:02}", value.minute)
            } else {
                value.minute.to_string()
            }
        }
        'm' => match width {
            1 => value.month.to_string(),
            2 => format!("{:02}", value.month),
            3 => month(value.month, true),
            _ => month(value.month, false),
        },
        _ => String::new(),
    }
}

fn month(month: u32, short: bool) -> String {
    const FULL: [&str; 12] = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    ];
    let name = FULL[(month.clamp(1, 12) - 1) as usize];
    if short {
        name.chars().take(3).collect()
    } else {
        name.to_owned()
    }
}

fn weekday(day: Weekday, short: bool) -> String {
    const NAMES: [&str; 7] = [
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
        "Sunday",
    ];
    let name = NAMES[day.num_days_from_monday() as usize];
    if short {
        name.chars().take(3).collect()
    } else {
        name.to_owned()
    }
}

// ---------------------------------------------------------------------------
// Number formats
// ---------------------------------------------------------------------------

/// A piece of a number format, in the order it is written.
#[derive(Clone, Debug)]
enum Token {
    Text(String),
    Digit(char),
    Dot,
    Comma,
    Slash,
    Percent,
    General,
    /// `E+00` or `E-0`: the sign to write, and how many exponent digits.
    Exponent {
        plus: bool,
        digits: usize,
    },
}

/// Renders a number through a format code, or `None` when the code is a date
/// or a construct this does not model.
pub fn number(value: f64, code: &str) -> Option<String> {
    if !value.is_finite() {
        return None;
    }
    let sections = split_sections(code);
    if sections.is_empty() || sections.iter().any(|section| has_date_field(section)) {
        return None;
    }
    let (section, auto_minus) = choose(&sections, value)?;
    render_number(value, section, auto_minus)
}

/// The sections separated by an unquoted `;`.
fn split_sections(code: &str) -> Vec<&str> {
    let mut out = Vec::new();
    let mut start = 0;
    let mut chars = code.char_indices().peekable();
    while let Some((index, ch)) = chars.next() {
        match ch {
            '"' => {
                while let Some((_, inside)) = chars.next() {
                    if inside == '"' {
                        if chars.peek().is_some_and(|(_, next)| *next == '"') {
                            chars.next();
                        } else {
                            break;
                        }
                    }
                }
            }
            '\\' | '_' | '*' => {
                chars.next();
            }
            ';' => {
                out.push(&code[start..index]);
                start = index + 1;
            }
            _ => {}
        }
    }
    out.push(&code[start..]);
    out
}

/// The section a value belongs in, and whether a minus sign is still owed.
fn choose<'a>(sections: &'a [&'a str], value: f64) -> Option<(&'a str, bool)> {
    let conditions: Vec<Option<(&'static str, f64)>> = sections.iter().map(|section| condition(section)).collect();
    if conditions.iter().any(Option::is_some) {
        for (index, test) in conditions.iter().enumerate() {
            if let Some((op, number)) = test {
                if holds(op, value, *number) {
                    return Some((sections[index], value < 0.0));
                }
            }
        }
        // The last section without a test is the one that catches the rest.
        for (index, test) in conditions.iter().enumerate().rev() {
            if test.is_none() {
                return Some((sections[index], value < 0.0));
            }
        }
        return Some(("", false));
    }
    let mut chosen = 0;
    let mut auto_minus = value < 0.0;
    match sections.len() {
        0 => return None,
        1 => {}
        2 => {
            if value < 0.0 {
                chosen = 1;
                auto_minus = false;
            }
        }
        _ => {
            if value < 0.0 {
                chosen = 1;
                auto_minus = false;
            } else if value == 0.0 {
                chosen = 2;
            }
        }
    }
    Some((sections[chosen], auto_minus))
}

/// A `[=0]`-style test at the front of a section.
fn condition(section: &str) -> Option<(&'static str, f64)> {
    let rest = section.trim_start().strip_prefix('[')?;
    let end = rest.find(']')?;
    let inside = &rest[..end];
    for op in ["<=", ">=", "<>", "=", "<", ">"] {
        if let Some(number) = inside.strip_prefix(op) {
            if let Ok(number) = number.trim().parse::<f64>() {
                return Some((op, number));
            }
        }
    }
    None
}

fn holds(op: &str, value: f64, number: f64) -> bool {
    match op {
        "<=" => value <= number,
        ">=" => value >= number,
        "<>" => value != number,
        "=" => value == number,
        "<" => value < number,
        _ => value > number,
    }
}

/// Splits a section into the pieces that are written out and the digits that
/// are filled in.
fn tokens(section: &str) -> Vec<Token> {
    let chars: Vec<char> = section.chars().collect();
    let mut out = Vec::new();
    let mut index = 0;
    while index < chars.len() {
        match chars[index] {
            '"' => {
                index += 1;
                let mut text = String::new();
                while index < chars.len() {
                    if chars[index] == '"' {
                        if chars.get(index + 1) == Some(&'"') {
                            text.push('"');
                            index += 2;
                            continue;
                        }
                        index += 1;
                        break;
                    }
                    text.push(chars[index]);
                    index += 1;
                }
                out.push(Token::Text(text));
            }
            '\\' => {
                if let Some(ch) = chars.get(index + 1) {
                    out.push(Token::Text(ch.to_string()));
                }
                index += 2;
            }
            // `_x` reserves the width of x: a space is the closest this can get.
            '_' => {
                out.push(Token::Text(" ".to_owned()));
                index += 2;
            }
            // `*x` fills the cell with x; the grid draws its own text, so the
            // fill is left out rather than rendered once and wrong.
            '*' => index += 2,
            '[' => {
                let start = index + 1;
                let mut end = start;
                while end < chars.len() && chars[end] != ']' {
                    end += 1;
                }
                let inside: String = chars[start..end].iter().collect();
                if let Some(symbol) = currency_symbol(&inside) {
                    if !symbol.is_empty() {
                        out.push(Token::Text(symbol));
                    }
                }
                index = end + 1;
            }
            '0' | '#' | '?' => {
                out.push(Token::Digit(chars[index]));
                index += 1;
            }
            '.' => {
                out.push(Token::Dot);
                index += 1;
            }
            ',' => {
                out.push(Token::Comma);
                index += 1;
            }
            '%' => {
                out.push(Token::Percent);
                index += 1;
            }
            '/' => {
                out.push(Token::Slash);
                index += 1;
            }
            'E' | 'e' => match chars.get(index + 1) {
                Some('+') | Some('-') => {
                    let plus = chars[index + 1] == '+';
                    let start = index + 2;
                    let mut end = start;
                    while chars.get(end) == Some(&'0') {
                        end += 1;
                    }
                    if end > start {
                        out.push(Token::Exponent {
                            plus,
                            digits: end - start,
                        });
                        index = end;
                    } else {
                        out.push(Token::Text(chars[index].to_string()));
                        index += 1;
                    }
                }
                _ => {
                    out.push(Token::Text(chars[index].to_string()));
                    index += 1;
                }
            },
            ch if ch.eq_ignore_ascii_case(&'g') => {
                let rest: String = chars[index..].iter().collect();
                if rest.to_ascii_lowercase().starts_with("general") {
                    out.push(Token::General);
                    index += 7;
                } else {
                    out.push(Token::Text(ch.to_string()));
                    index += 1;
                }
            }
            other => {
                out.push(Token::Text(other.to_string()));
                index += 1;
            }
        }
    }
    out
}

/// `[$¥-411]` names a currency symbol; `[$-411]` and `[Red]` name none.
fn currency_symbol(inside: &str) -> Option<String> {
    let rest = inside.strip_prefix('$')?;
    let symbol = match rest.rfind('-') {
        Some(at) => &rest[..at],
        None => rest,
    };
    Some(symbol.to_owned())
}

fn render_number(value: f64, section: &str, auto_minus: bool) -> Option<String> {
    let items = tokens(section);
    let percents = items.iter().filter(|item| matches!(item, Token::Percent)).count();
    let scaled = value * 100f64.powi(percents as i32);
    // "General" is the number's own spelling, with whatever text surrounds it.
    if items.iter().any(|item| matches!(item, Token::General)) {
        let mut out = String::new();
        if auto_minus {
            out.push('-');
        }
        for item in &items {
            match item {
                Token::Text(text) => out.push_str(text),
                Token::Percent => out.push('%'),
                Token::General => out.push_str(&general_text(scaled)),
                _ => {}
            }
        }
        return Some(out);
    }
    if let Some(at) = items.iter().position(|item| matches!(item, Token::Exponent { .. })) {
        return render_scientific(scaled, &items, at, auto_minus);
    }
    if let Some(at) = items.iter().position(|item| matches!(item, Token::Slash)) {
        return render_fraction(scaled, &items, at, auto_minus);
    }
    let has_region = items
        .iter()
        .any(|item| matches!(item, Token::Digit(_) | Token::Dot | Token::Comma));
    let has_digit = items.iter().any(|item| matches!(item, Token::Digit(_)));
    if !has_region || !has_digit {
        // No digits: the section is text alone, such as the "-" of a zero
        // section.
        let mut out = String::new();
        for item in &items {
            match item {
                Token::Text(text) => out.push_str(text),
                Token::Percent => out.push('%'),
                _ => {}
            }
        }
        return Some(out);
    }
    let first = items
        .iter()
        .position(|item| matches!(item, Token::Digit(_) | Token::Dot | Token::Comma))?;
    let last = items
        .iter()
        .rposition(|item| matches!(item, Token::Digit(_) | Token::Dot | Token::Comma))?;

    let prefix = &items[..first];
    let pattern = &items[first..=last];
    let suffix = &items[last + 1..];

    let mut integer: Vec<char> = Vec::new();
    let mut fraction: Vec<char> = Vec::new();
    let mut after_dot = false;
    let mut grouping = false;
    let mut scale = 0usize;
    let last_digit = pattern.iter().rposition(|item| matches!(item, Token::Digit(_)))?;
    for (index, item) in pattern.iter().enumerate() {
        match item {
            Token::Digit(ch) => {
                if after_dot {
                    fraction.push(*ch);
                } else {
                    integer.push(*ch);
                }
            }
            Token::Dot => after_dot = true,
            Token::Comma => {
                if after_dot {
                    continue;
                }
                // A comma between digits groups them; one after the last digit
                // scales the value by a thousand.
                if index > last_digit {
                    scale += 1;
                } else {
                    grouping = true;
                }
            }
            _ => {}
        }
    }

    let magnitude = (scaled / 1000f64.powi(scale as i32)).abs();
    let places = fraction.len();
    let digits = round_scaled(magnitude, places);
    let split = digits.len() - places;
    let (integer_digits, fraction_digits) = (digits[..split].to_owned(), digits[split..].to_owned());

    let mandatory = integer.iter().filter(|ch| **ch == '0').count();
    let optional = integer.iter().filter(|ch| **ch == '?').count();
    let mut integer_text = integer_digits;
    if integer_text == "0" && mandatory == 0 {
        integer_text.clear();
    }
    while integer_text.len() < mandatory {
        integer_text.insert(0, '0');
    }
    if grouping && !integer_text.is_empty() {
        integer_text = group(&integer_text);
    }
    if optional > 0 {
        let wanted = mandatory + optional;
        let present = integer_text.chars().filter(|ch| ch.is_ascii_digit()).count();
        for _ in present..wanted {
            integer_text.insert(0, ' ');
        }
    }

    let fraction_text = if places > 0 {
        fraction_from(&fraction_digits, &fraction)
    } else {
        String::new()
    };

    let mut out = String::new();
    if auto_minus {
        out.push('-');
    }
    emit(prefix, &mut out);
    out.push_str(&integer_text);
    if places > 0 {
        out.push('.');
        out.push_str(&fraction_text);
    }
    emit(suffix, &mut out);
    Some(out)
}

/// `0.00E+00` and `##0.0E+0`: a mantissa and an exponent. The exponent is a
/// multiple of the integer placeholder count, which is what makes `##0.0` an
/// engineering format.
fn render_scientific(value: f64, items: &[Token], at: usize, auto_minus: bool) -> Option<String> {
    let Token::Exponent { plus, digits } = items[at] else {
        return None;
    };
    let before = &items[..at];
    let after = &items[at + 1..];
    let first = before.iter().position(|item| matches!(item, Token::Digit(_)))?;
    let prefix = &before[..first];
    let pattern = &before[first..];
    let mut integer = Vec::new();
    let mut fraction = Vec::new();
    let mut after_dot = false;
    for item in pattern {
        match item {
            Token::Digit(ch) => {
                if after_dot {
                    fraction.push(*ch);
                } else {
                    integer.push(*ch);
                }
            }
            Token::Dot => after_dot = true,
            _ => {}
        }
    }
    let slots = integer.len().max(1);
    let places = fraction.len();
    let magnitude = value.abs();
    let exponent = if magnitude == 0.0 {
        0
    } else {
        let floor = magnitude.log10().floor() as i64;
        floor.div_euclid(slots as i64) * slots as i64
    };
    let mantissa = magnitude / 10f64.powi(exponent as i32);
    let digits_text = round_scaled(mantissa, places);
    let split = digits_text.len() - places;
    let mut integer_text = digits_text[..split].to_owned();
    let mandatory = integer.iter().filter(|ch| **ch == '0').count();
    // Zero is written out to the full integer width: `000.0E+0`.
    let wanted = if magnitude == 0.0 { slots } else { mandatory };
    while integer_text.len() < wanted {
        integer_text.insert(0, '0');
    }
    let fraction_text = if places > 0 {
        fraction_from(&digits_text[split..], &fraction)
    } else {
        String::new()
    };
    let mut out = String::new();
    if auto_minus {
        out.push('-');
    }
    emit(prefix, &mut out);
    out.push_str(&integer_text);
    if places > 0 {
        out.push('.');
        out.push_str(&fraction_text);
    }
    out.push('E');
    if exponent < 0 {
        out.push('-');
    } else if plus {
        out.push('+');
    }
    let exponent_text = exponent.abs().to_string();
    for _ in exponent_text.len()..digits.max(1) {
        out.push('0');
    }
    out.push_str(&exponent_text);
    emit(after, &mut out);
    Some(out)
}

/// `# ?/?`, `# ??/??`, `?/??` or `# ?/8`: a best fraction with the
/// denominator the placeholders allow, or the one written out.
fn render_fraction(value: f64, items: &[Token], slash: usize, auto_minus: bool) -> Option<String> {
    let mut num_start = slash;
    while num_start > 0 && matches!(items[num_start - 1], Token::Digit(_)) {
        num_start -= 1;
    }
    if num_start == slash {
        return None;
    }
    let numerator_pattern: Vec<char> = items[num_start..slash]
        .iter()
        .filter_map(|item| match item {
            Token::Digit(ch) => Some(*ch),
            _ => None,
        })
        .collect();
    let mut den_end = slash + 1;
    let mut fixed: Option<u64> = None;
    let denominator_pattern: Vec<char>;
    match items.get(den_end) {
        Some(Token::Digit(_)) => {
            while matches!(items.get(den_end), Some(Token::Digit(_))) {
                den_end += 1;
            }
            denominator_pattern = items[slash + 1..den_end]
                .iter()
                .filter_map(|item| match item {
                    Token::Digit(ch) => Some(*ch),
                    _ => None,
                })
                .collect();
        }
        Some(Token::Text(text)) => {
            let digits: String = text.chars().filter(|ch| ch.is_ascii_digit()).collect();
            if digits.is_empty() || digits.len() != text.len() {
                return None;
            }
            fixed = digits.parse().ok().filter(|q| *q > 0);
            denominator_pattern = vec!['0'; digits.len()];
            den_end += 1;
        }
        _ => return None,
    }
    let integer_digits: Vec<(usize, char)> = items[..num_start]
        .iter()
        .enumerate()
        .filter_map(|(index, item)| match item {
            Token::Digit(ch) => Some((index, *ch)),
            _ => None,
        })
        .collect();
    let mixed = !integer_digits.is_empty();
    let first_numeric = if mixed { integer_digits[0].0 } else { num_start };
    let int_end = if mixed {
        integer_digits.last().map_or(num_start, |(index, _)| index + 1)
    } else {
        num_start
    };
    let prefix = &items[..first_numeric];
    let integer_pattern: Vec<char> = integer_digits.iter().map(|(_, ch)| *ch).collect();
    let between = &items[int_end..num_start];
    let suffix = &items[den_end..];

    let magnitude = value.abs();
    let mut whole = if mixed { magnitude.floor() } else { 0.0 };
    let remainder = if mixed { magnitude - whole } else { magnitude };
    let (mut p, q) = match fixed {
        Some(q) => (((remainder * q as f64).round() as u64), q),
        None => {
            let max_q = 10u64
                .saturating_pow(denominator_pattern.len() as u32)
                .saturating_sub(1)
                .max(1);
            let mut best = (0u64, 1u64, f64::INFINITY);
            for q in 1..=max_q {
                let p = (remainder * q as f64).round() as u64;
                let error = (remainder - p as f64 / q as f64).abs();
                if error < best.2 {
                    best = (p, q, error);
                }
            }
            (best.0, best.1)
        }
    };
    if mixed && p == q {
        // The fraction rounded up to a whole one: show it as the next integer.
        whole += 1.0;
        p = 0;
    }
    let blank = mixed && p == 0;

    let mut out = String::new();
    if auto_minus {
        out.push('-');
    }
    emit(prefix, &mut out);
    if mixed {
        let whole_text = format!("{}", whole as u64);
        let text = if whole == 0.0 && !blank {
            if integer_pattern.iter().any(|ch| *ch == '0') {
                "0".to_owned()
            } else if integer_pattern.iter().any(|ch| *ch == '?') {
                " ".repeat(integer_pattern.len())
            } else {
                String::new()
            }
        } else {
            pad_number(&whole_text, &integer_pattern, true)
        };
        out.push_str(&text);
    }
    if blank {
        let mut spacing = String::new();
        emit(between, &mut spacing);
        let width = spacing.chars().count() + numerator_pattern.len() + 1 + denominator_pattern.len();
        for _ in 0..width {
            out.push(' ');
        }
    } else {
        emit(between, &mut out);
        out.push_str(&pad_number(&p.to_string(), &numerator_pattern, true));
        out.push('/');
        out.push_str(&pad_number(&q.to_string(), &denominator_pattern, false));
    }
    emit(suffix, &mut out);
    Some(out)
}

/// The fractional digits a `0/#/?` pattern keeps: optional zeros fall away
/// from the right, `?` leaves a space where one would have been, `0` keeps it.
fn fraction_from(digits: &str, pattern: &[char]) -> String {
    let places = pattern.len();
    let bytes = digits.as_bytes();
    let mut keep = vec![true; places];
    let mut trimming = true;
    for index in (0..places).rev() {
        let digit = bytes[index] as char;
        if trimming && digit == '0' {
            match pattern[index] {
                '#' | '?' => keep[index] = false,
                _ => trimming = false,
            }
        } else {
            trimming = false;
        }
    }
    let mut out = String::new();
    for index in 0..places {
        if keep[index] {
            out.push(bytes[index] as char);
        } else if pattern[index] == '?' {
            out.push(' ');
        }
    }
    out
}

/// A number set to the width of a placeholder run: `0` pads with zeros, `?`
/// with spaces, `#` not at all.
fn pad_number(text: &str, pattern: &[char], right_align: bool) -> String {
    let width = pattern.len();
    let pad = if pattern.iter().any(|ch| *ch == '0') {
        '0'
    } else if pattern.iter().any(|ch| *ch == '?') {
        ' '
    } else {
        return text.to_owned();
    };
    let fill = width.saturating_sub(text.chars().count());
    let mut out = String::new();
    if right_align {
        for _ in 0..fill {
            out.push(pad);
        }
        out.push_str(text);
    } else {
        out.push_str(text);
        for _ in 0..fill {
            out.push(pad);
        }
    }
    out
}

fn emit(items: &[Token], out: &mut String) {
    for item in items {
        match item {
            Token::Text(text) => out.push_str(text),
            Token::Percent => out.push('%'),
            Token::Comma => out.push(','),
            Token::Dot => out.push('.'),
            _ => {}
        }
    }
}

/// The number rounded half away from zero at `places` decimals, as a string of
/// digits with the decimal point moved `places` to the right. The rounding
/// reads the number as its fifteen-decimal spelling — Excel rounds `1.005` up
/// to `1.01`, where the binary value alone is just under the half.
fn round_scaled(value: f64, places: usize) -> String {
    let text = format!("{:.15}", value.abs());
    let (int, frac) = text.split_once('.').unwrap_or((text.as_str(), ""));
    let mut digits: Vec<u8> = int.bytes().chain(frac.bytes().take(15)).collect();
    let cut = int.len() + places;
    while digits.len() <= cut {
        digits.push(b'0');
    }
    let round_up = digits[cut] >= b'5';
    let mut kept: Vec<u8> = digits[..cut].to_vec();
    if round_up {
        let mut index = kept.len();
        loop {
            if index == 0 {
                kept.insert(0, b'1');
                break;
            }
            index -= 1;
            if kept[index] == b'9' {
                kept[index] = b'0';
            } else {
                kept[index] += 1;
                break;
            }
        }
    }
    String::from_utf8(kept).unwrap_or_default()
}

fn group(digits: &str) -> String {
    let mut out = String::with_capacity(digits.len() + digits.len() / 3);
    for (index, ch) in digits.chars().enumerate() {
        if index > 0 && (digits.len() - index) % 3 == 0 {
            out.push(',');
        }
        out.push(ch);
    }
    out
}

fn general_text(value: f64) -> String {
    if value.fract() == 0.0 && value.abs() < 1e15 {
        format!("{}", value as i64)
    } else {
        format!("{value}")
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use calamine::ExcelDateTimeType;
    use chrono::NaiveDate;

    fn serial(year: i32, month: u32, day: u32) -> f64 {
        let epoch = NaiveDate::from_ymd_opt(1899, 12, 30).unwrap();
        NaiveDate::from_ymd_opt(year, month, day)
            .unwrap()
            .signed_duration_since(epoch)
            .num_days() as f64
    }

    fn at(value: f64) -> Data {
        Data::DateTime(ExcelDateTime::new(value, ExcelDateTimeType::DateTime, false))
    }

    #[test]
    fn renders_the_format_this_workbook_actually_uses() {
        // `<numFmt numFmtId="177" formatCode="yyyy&quot;年&quot;mm&quot;月&quot;dd&quot;日&quot;;@"/>`
        assert_eq!(
            apply(&at(serial(2025, 5, 30)), r#"yyyy"年"mm"月"dd"日";@"#).as_deref(),
            Some("2025年05月30日")
        );
    }

    #[test]
    fn splits_month_from_minute_by_position() {
        // Half a day past midnight is noon.
        let value = at(serial(2025, 5, 30) + 0.5);
        assert_eq!(apply(&value, "mm").as_deref(), Some("05"));
        assert_eq!(apply(&value, "hh:mm").as_deref(), Some("12:00"));
        assert_eq!(apply(&value, "h:mm:ss").as_deref(), Some("12:00:00"));
    }

    /// Measured from Excel (`probe-numfmt.ps1`); the `*` fill it pads with is
    /// not part of the text this renders.
    #[test]
    fn renders_numbers_the_way_excel_does() {
        let render = |value: f64, code: &str| number(value, code);
        assert_eq!(render(1234.5, "#,##0;-#,##0;\"-\"").as_deref(), Some("1,235"));
        assert_eq!(render(-1234.5, "#,##0;-#,##0;\"-\"").as_deref(), Some("-1,235"));
        assert_eq!(render(0.0, "#,##0;-#,##0;\"-\"").as_deref(), Some("-"));
        assert_eq!(render(1234.5, "0.000000000").as_deref(), Some("1234.500000000"));
        assert_eq!(render(1234.5, "0_ ").as_deref(), Some("1235 "));
        assert_eq!(render(-1234.5, "0_);(0)").as_deref(), Some("(1235)"));
        assert_eq!(render(1234.5, "0\" 件\"").as_deref(), Some("1235 件"));
        assert_eq!(
            render(1234.5, "\"￡\"#,##0.00;-\"￡\"#,##0.00").as_deref(),
            Some("￡1,234.50")
        );
        assert_eq!(
            render(-1234.5, "\"￡\"#,##0.00;-\"￡\"#,##0.00").as_deref(),
            Some("-￡1,234.50")
        );
        assert_eq!(render(1234.5, "0.00_)").as_deref(), Some("1234.50 "));
        assert_eq!(render(1234.5, "\"$\"#,##0_);(\"$\"#,##0)").as_deref(), Some("$1,235 "));
        assert_eq!(
            render(-1234.5, "\"$\"#,##0_);(\"$\"#,##0)").as_deref(),
            Some("($1,235)")
        );
        assert_eq!(render(1234.5, "#,##0.0\"人月\"").as_deref(), Some("1,234.5人月"));
        assert_eq!(render(0.0, "####").as_deref(), Some(""));
        assert_eq!(render(12.0, "####").as_deref(), Some("12"));
        assert_eq!(render(-0.5, "####").as_deref(), Some("-1"));
        assert_eq!(render(1234.5, "\"c× \"0%").as_deref(), Some("c× 123450%"));
        assert_eq!(render(-1234.5, "\"c× \"0%").as_deref(), Some("-c× 123450%"));
        assert_eq!(render(1234.5, "0.0\" 件\"").as_deref(), Some("1234.5 件"));
        assert_eq!(render(1234.5, "\"¥\"#,##0;-\"¥\"#,##0").as_deref(), Some("¥1,235"));
        assert_eq!(render(-1234.5, "#,##0;(#,##0)").as_deref(), Some("(1,235)"));
        assert_eq!(render(-1234.5, "0_ ;-0 ").as_deref(), Some("-1235 "));
        assert_eq!(render(1.005, "0.00").as_deref(), Some("1.01"));
        assert_eq!(render(0.4567, "0.00").as_deref(), Some("0.46"));
        assert_eq!(render(1234.5, "0").as_deref(), Some("1235"));
        assert_eq!(render(-0.5, "0").as_deref(), Some("-1"));
        assert_eq!(render(1234.5, "\"(\"General\")\"").as_deref(), Some("(1234.5)"));
        assert_eq!(render(0.0, "[=0]\"\";General").as_deref(), Some(""));
        assert_eq!(render(5.0, "[=0]\"\";General").as_deref(), Some("5"));
        assert_eq!(render(0.0, "[=0]\"\";0").as_deref(), Some(""));
        assert_eq!(render(5.0, "[Red]0").as_deref(), Some("5"));
        // Serial zero is Excel's day zero, 1900-01-00, not the epoch.
        assert_eq!(apply(&at(0.0), "mm/dd").as_deref(), Some("01/00"));
    }

    /// Measured from Excel (`probe-numfmt.ps1`).
    #[test]
    fn renders_scientific_fractions_and_elapsed_time() {
        let render = |value: f64, code: &str| number(value, code);
        assert_eq!(render(12345.0, "0.00E+00").as_deref(), Some("1.23E+04"));
        assert_eq!(render(0.000123, "0.00E+00").as_deref(), Some("1.23E-04"));
        assert_eq!(render(-12345.0, "0.00E+00").as_deref(), Some("-1.23E+04"));
        assert_eq!(render(0.0, "0.00E+00").as_deref(), Some("0.00E+00"));
        assert_eq!(render(1.5, "0.00E+00").as_deref(), Some("1.50E+00"));
        assert_eq!(render(12345.0, "##0.0E+0").as_deref(), Some("12.3E+3"));
        assert_eq!(render(0.000123, "##0.0E+0").as_deref(), Some("123.0E-6"));
        assert_eq!(render(0.0, "##0.0E+0").as_deref(), Some("000.0E+0"));
        assert_eq!(render(12345.0, "0.0E-00").as_deref(), Some("1.2E04"));

        assert_eq!(render(1.5, "# ?/?").as_deref(), Some("1 1/2"));
        assert_eq!(render(0.25, "# ?/?").as_deref(), Some(" 1/4"));
        assert_eq!(render(0.333, "# ?/?").as_deref(), Some(" 1/3"));
        assert_eq!(render(0.0, "# ?/?").as_deref(), Some("0    "));
        assert_eq!(render(0.95, "# ?/?").as_deref(), Some("1    "));
        assert_eq!(render(0.333, "# ??/??").as_deref(), Some("  1/3 "));
        assert_eq!(render(0.123, "# ??/??").as_deref(), Some("  8/65"));
        assert_eq!(render(0.0123, "# ??/??").as_deref(), Some("  1/81"));
        assert_eq!(render(1.5, "?/??").as_deref(), Some("3/2 "));
        assert_eq!(render(0.0, "?/??").as_deref(), Some("0/1 "));
        assert_eq!(render(1.5, "# ?/8").as_deref(), Some("1 4/8"));
        assert_eq!(render(0.333, "# ?/8").as_deref(), Some(" 3/8"));
        assert_eq!(render(0.1, "# ?/8").as_deref(), Some(" 1/8"));

        let elapsed = |value: f64, code: &str| apply(&at(value), code);
        assert_eq!(elapsed(1.5, "[h]:mm:ss").as_deref(), Some("36:00:00"));
        assert_eq!(elapsed(0.25, "[h]:mm:ss").as_deref(), Some("6:00:00"));
        assert_eq!(elapsed(2.75, "[h]:mm:ss").as_deref(), Some("66:00:00"));
        assert_eq!(elapsed(1.5, "[m]:ss").as_deref(), Some("2160:00"));
        assert_eq!(elapsed(0.5, "[s]").as_deref(), Some("43200"));
        assert_eq!(elapsed(1.5, "[hh]:mm").as_deref(), Some("36:00"));
        assert_eq!(elapsed(0.25, "[hh]:mm").as_deref(), Some("06:00"));
        assert_eq!(elapsed(0.0001, "[h]:mm:ss").as_deref(), Some("0:00:09"));
        assert_eq!(elapsed(0.99999, "[h]:mm:ss").as_deref(), Some("23:59:59"));
    }

    /// A date-typed cell carrying a number format shows the serial.
    #[test]
    fn reads_a_number_format_on_a_date_cell() {
        assert_eq!(
            apply(&at(serial(2025, 5, 30)), "#,##0.00").as_deref(),
            Some("45,807.00")
        );
    }
}
