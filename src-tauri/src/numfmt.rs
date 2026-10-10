//! Renders a date-formatted cell the way Excel would.
//!
//! Cells hold a raw value plus a format code, and the format is what turns the
//! serial number into `2025年05月30日`. Only date and time codes are handled
//! here; anything else returns `None` so the caller falls back to a plain
//! rendering instead of showing something wrong.

use calamine::Data;
use chrono::{Datelike, NaiveDateTime, Timelike};

/// Renders `value` with the first section of `code`, or `None` when the code
/// is not a date format this understands.
pub fn apply(value: &Data, code: &str) -> Option<String> {
    let Data::DateTime(datetime) = value else {
        return None;
    };
    let naive = datetime.as_datetime()?;
    let section = first_section(code);
    // A format with no date field in it is not a date format.
    if !section.chars().any(|ch| "ymdhs".contains(ch)) {
        return None;
    }
    Some(render(&naive, section))
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

fn render(value: &NaiveDateTime, section: &str) -> String {
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
                while index < chars.len() && chars[index] != ']' {
                    index += 1;
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
                    out.push_str(if value.hour() < 12 { "AM" } else { "PM" });
                    index += 5;
                } else if upper.starts_with("A/P") {
                    out.push_str(if value.hour() < 12 { "A" } else { "P" });
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

fn field(
    value: &NaiveDateTime,
    letter: char,
    width: usize,
    minutes: bool,
    twelve_hour: bool,
) -> String {
    match letter {
        'y' => {
            if width <= 2 {
                format!("{:02}", value.year().rem_euclid(100))
            } else {
                format!("{:04}", value.year())
            }
        }
        'd' => match width {
            1 => value.day().to_string(),
            2 => format!("{:02}", value.day()),
            3 => weekday(value, true),
            _ => weekday(value, false),
        },
        'h' => {
            let hour = if twelve_hour {
                let hour = value.hour() % 12;
                if hour == 0 {
                    12
                } else {
                    hour
                }
            } else {
                value.hour()
            };
            if width >= 2 {
                format!("{hour:02}")
            } else {
                hour.to_string()
            }
        }
        's' => {
            if width >= 2 {
                format!("{:02}", value.second())
            } else {
                value.second().to_string()
            }
        }
        'm' if minutes => {
            if width >= 2 {
                format!("{:02}", value.minute())
            } else {
                value.minute().to_string()
            }
        }
        'm' => match width {
            1 => value.month().to_string(),
            2 => format!("{:02}", value.month()),
            3 => month(value.month(), true),
            _ => month(value.month(), false),
        },
        _ => String::new(),
    }
}

fn month(month: u32, short: bool) -> String {
    const FULL: [&str; 12] = [
        "January", "February", "March", "April", "May", "June", "July", "August", "September",
        "October", "November", "December",
    ];
    let name = FULL[(month.clamp(1, 12) - 1) as usize];
    if short {
        name.chars().take(3).collect()
    } else {
        name.to_owned()
    }
}

fn weekday(value: &NaiveDateTime, short: bool) -> String {
    const NAMES: [&str; 7] = [
        "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday",
    ];
    let name = NAMES[value.weekday().num_days_from_monday() as usize];
    if short {
        name.chars().take(3).collect()
    } else {
        name.to_owned()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use calamine::{ExcelDateTime, ExcelDateTimeType};
    use chrono::NaiveDate;

    fn serial(year: i32, month: u32, day: u32) -> f64 {
        let epoch = NaiveDate::from_ymd_opt(1899, 12, 30).unwrap();
        NaiveDate::from_ymd_opt(year, month, day)
            .unwrap()
            .signed_duration_since(epoch)
            .num_days() as f64
    }

    fn at(value: f64) -> Data {
        Data::DateTime(ExcelDateTime::new(
            value,
            ExcelDateTimeType::DateTime,
            false,
        ))
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

    #[test]
    fn leaves_non_date_formats_alone() {
        assert_eq!(apply(&at(serial(2025, 5, 30)), "#,##0.00"), None);
        assert_eq!(apply(&Data::Float(1.5), r#"yyyy"年""#), None);
    }
}
