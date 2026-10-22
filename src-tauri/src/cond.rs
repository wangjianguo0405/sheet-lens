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
//! The rule types that paint through a `dxf` are read: an expression, a
//! comparison against the cell itself (`cellIs`), a text test, and duplicates.
//! Colour scales, data bars and icon sets are drawings rather than formats, and
//! a rule this cannot read paints nothing rather than something wrong.

use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};

use chrono::{Datelike, Duration, Local, NaiveDate};
use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;
use quick_xml::XmlVersion;

use crate::layout::{
    attr, color_edge, color_of, entity_text, intern, parse_cell, parse_range, set_edge, Cache,
    CellBorder, CellStyle, MergeRect, Palette, SheetLayout, Side,
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
    /// What the rule paints, when it fires.
    paint: Paint,
    /// The cells the rule covers. Each range is its own anchor: a rule's formula
    /// is written relative to the top-left of the range it sits on.
    ranges: Vec<MergeRect>,
    matcher: Matcher,
    /// The `x14:id` a base rule names, which ties it to the extension block's
    /// settings (a data bar's axis and negative colour live there).
    x14_id: Option<String>,
}

/// What a firing rule puts on the cell. A `dxf` is the same for every cell it
/// covers; a colour scale and a data bar are computed from the cell's own value
/// against the values in the rule's range.
enum Paint {
    Dxf(usize),
    ColorScale(ColorScale),
    DataBar(DataBar),
    IconSet(IconSet),
}

/// A colour scale: two or three stops, each a position in the range and the
/// colour the cell takes there.
struct ColorScale {
    points: Vec<ScalePoint>,
    colors: Vec<String>,
}

/// A data bar: the bar's colour, and the two points its length runs between.
struct DataBar {
    color: String,
    /// What a bar below the axis is drawn in, when that differs.
    negative_color: Option<String>,
    points: [ScalePoint; 2],
    axis: Axis,
    axis_color: Option<String>,
}

/// Where zero sits across the cell, which is what decides the direction a
/// negative bar runs in.
#[derive(Clone, Copy, PartialEq, Debug)]
enum Axis {
    /// No axis: bars scale from the left edge over the range.
    None,
    /// Zero at the middle of the cell.
    Middle,
    /// Zero where the range's own split puts it.
    Auto,
}

/// An icon set: a family of marks, one per level, drawn at the left of the
/// cell. The points are the levels' lower ends, the first being the floor.
struct IconSet {
    family: String,
    points: Vec<ScalePoint>,
    reverse: bool,
}

/// Where a stop sits. Excel writes the first and last as `min`/`max`, and a
/// middle stop as a `formula` — which for a fixed position is just the number.
#[derive(Clone, Debug)]
enum ScalePoint {
    Min,
    Max,
    Number(f64),
    Percent(f64),
    Percentile(f64),
}

/// What makes a rule fire for a cell.
///
/// The types that paint through a `dxf` are here, together with the drawing
/// types a cell's position decides. Icon sets are left out: their marks are
/// assets rather than a colour, and a rule this cannot read paints nothing
/// rather than something wrong.
enum Matcher {
    /// `expression`: the formula has to evaluate to TRUE.
    Expression(Formula),
    /// `cellIs`: the cell's own value against one or two operands. Excel writes
    /// the operands as `<formula>` too — usually a constant, sometimes a cell.
    CellIs {
        op: CellOp,
        first: Formula,
        second: Option<Formula>,
    },
    /// `containsText` and its relatives, against the cell's own text. Excel
    /// matches them without regard to case.
    Text { test: TextTest, needle: String },
    /// `duplicateValues` / `uniqueValues`: whether the value repeats among the
    /// rule's own ranges. Blanks never count either way.
    Duplicate { keep_duplicates: bool },
    /// `top10`: whether the cell's value is among the highest or lowest `rank`
    /// in the rule's range. `percent` makes `rank` a share of the count instead.
    Top {
        bottom: bool,
        rank: u32,
        percent: bool,
    },
    /// `aboveAverage` and its below/equal variants, optionally shifted by a
    /// number of standard deviations.
    Average {
        above: bool,
        equal: bool,
        deviations: u32,
    },
    /// A colour scale or a data bar covers its whole range.
    Always,
    /// `timePeriod`: the cell's date falls inside one of Excel's named windows,
    /// measured against the current date.
    TimePeriod(TimePeriod),
}

#[derive(Clone, Copy, PartialEq, Debug)]
enum CellOp {
    Equal,
    NotEqual,
    Greater,
    GreaterOrEqual,
    Less,
    LessOrEqual,
    Between,
    NotBetween,
}

#[derive(Clone, Copy, PartialEq, Debug)]
enum TextTest {
    Contains,
    NotContains,
    Begins,
    Ends,
}

fn cell_op(name: &str) -> Option<CellOp> {
    Some(match name {
        "equal" => CellOp::Equal,
        "notEqual" => CellOp::NotEqual,
        "greaterThan" => CellOp::Greater,
        "greaterThanOrEqual" => CellOp::GreaterOrEqual,
        "lessThan" => CellOp::Less,
        "lessThanOrEqual" => CellOp::LessOrEqual,
        "between" => CellOp::Between,
        "notBetween" => CellOp::NotBetween,
        _ => return None,
    })
}

/// Reads a sheet's `<conditionalFormatting>` blocks.
///
/// A rule whose operands cannot be read is dropped here rather than at every
/// cell: there is nothing to evaluate it with, and a rule that never fires
/// paints nothing, which is what not being there looks like too.
pub fn rules(xml: &str, palette: Palette) -> Vec<Rule> {
    let mut reader = XmlReader::from_str(xml);
    let mut buf = Vec::new();
    let mut out = Vec::new();

    let mut sqref: Vec<MergeRect> = Vec::new();
    let mut current: Option<Pending> = None;
    let mut in_formula = false;
    let mut formula = String::new();
    // The `x14:id` that links a base rule to its extension settings.
    let mut in_id = false;
    let mut id_text = String::new();
    // Inside `<colorScale>`, `<dataBar>` or `<iconSet>`.
    let mut collecting: Option<Collecting> = None;

    loop {
        buf.clear();
        let event = reader.read_event_into(&mut buf);
        let empty = matches!(event, Ok(Event::Empty(_)));
        match event {
            Ok(Event::Start(element)) | Ok(Event::Empty(element)) => match element.local_name().as_ref() {
                b"conditionalFormatting" => {
                    sqref = sqref_of(attr(&element, b"sqref").as_deref());
                }
                b"cfRule" => {
                    let pending = pending_rule(&element);
                    // A rule with no children — `duplicateValues` is written that
                    // way — is complete the moment it opens.
                    match pending {
                        Some(pending) if empty => {
                            if let Some(rule) = finish(pending, &sqref) {
                                out.push(rule);
                            }
                        }
                        other => current = other,
                    }
                }
                b"colorScale" if current.is_some() => collecting = Some(Collecting::Scale),
                b"dataBar" if current.is_some() => collecting = Some(Collecting::Bar),
                b"iconSet" if current.is_some() => {
                    collecting = Some(Collecting::Icons);
                    if let Some(current) = current.as_mut() {
                        // The attribute's default, per the schema: a file that
                        // sets the family Excel calls the default writes none.
                        current.family = Some(
                            attr(&element, b"iconSet")
                                .unwrap_or_else(|| "3TrafficLights1".to_owned()),
                        );
                        current.reverse = flag(attr(&element, b"reverse"));
                    }
                }
                b"cfvo" if collecting.is_some() => {
                    if let (Some(point), Some(current)) =
                        (scale_point(&element), current.as_mut())
                    {
                        current.points.push(point);
                    }
                }
                b"color" if collecting.is_some() => {
                    if let (Some(color), Some(current)) =
                        (color_of(&element, palette), current.as_mut())
                    {
                        current.colors.push(color);
                    }
                }
                b"formula" if current.is_some() && !empty => {
                    in_formula = true;
                    formula.clear();
                }
                b"id" if current.is_some() && !empty => {
                    in_id = true;
                    id_text.clear();
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
            Ok(Event::Text(text)) if in_id => {
                if let Ok(text) = text.xml_content(XmlVersion::default()) {
                    id_text.push_str(&text);
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
                        current.formulas.push(std::mem::take(&mut formula));
                    }
                }
                b"id" if in_id => {
                    in_id = false;
                    if let Some(current) = current.as_mut() {
                        current.x14_id = Some(std::mem::take(&mut id_text).trim().to_owned());
                    }
                }
                b"colorScale" | b"dataBar" | b"iconSet" => collecting = None,
                b"cfRule" => {
                    collecting = None;
                    if let Some(pending) = current.take() {
                        if let Some(rule) = finish(pending, &sqref) {
                            out.push(rule);
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

/// A `<cfRule>` opened but not yet finished: which type it is, and the children
/// read so far.
struct Pending {
    priority: u32,
    stop_if_true: bool,
    patch: Option<usize>,
    kind: Kind,
    formulas: Vec<String>,
    points: Vec<ScalePoint>,
    colors: Vec<String>,
    family: Option<String>,
    reverse: bool,
    x14_id: Option<String>,
}

/// Which kind of drawing element the rule's children belong to.
#[derive(Clone, Copy, PartialEq, Debug)]
enum Collecting {
    Scale,
    Bar,
    Icons,
}

enum Kind {
    Expression,
    CellIs(CellOp),
    Text(TextTest, String),
    Duplicate(bool),
    Top {
        bottom: bool,
        rank: u32,
        percent: bool,
    },
    Average {
        above: bool,
        equal: bool,
        deviations: u32,
    },
    TimePeriod(TimePeriod),
    ColorScale,
    DataBar,
    IconSet,
}

/// The named date windows a `timePeriod` rule tests.
#[derive(Clone, Copy, PartialEq, Debug)]
enum TimePeriod {
    Yesterday,
    Today,
    Tomorrow,
    Last7Days,
    LastWeek,
    ThisWeek,
    NextWeek,
    LastMonth,
    ThisMonth,
    NextMonth,
}

impl TimePeriod {
    fn parse(value: &str) -> Option<Self> {
        Some(match value {
            "yesterday" => Self::Yesterday,
            "today" => Self::Today,
            "tomorrow" => Self::Tomorrow,
            "last7Days" => Self::Last7Days,
            "lastWeek" => Self::LastWeek,
            "thisWeek" => Self::ThisWeek,
            "nextWeek" => Self::NextWeek,
            "lastMonth" => Self::LastMonth,
            "thisMonth" => Self::ThisMonth,
            "nextMonth" => Self::NextMonth,
            _ => return None,
        })
    }

    /// Measured against Excel (2026-10-22): it compares the value's **date
    /// part** — a time of day does not matter, and neither does the cell's
    /// number format — and its weeks run Sunday to Saturday.
    fn matches(self, date: NaiveDate, today: NaiveDate) -> bool {
        let week = |start: NaiveDate| date >= start && date < start + Duration::days(7);
        let this_week = today - Duration::days(today.weekday().num_days_from_sunday() as i64);
        let first_of_month = NaiveDate::from_ymd_opt(today.year(), today.month(), 1).unwrap();
        let before = first_of_month - Duration::days(1);
        let last_month = NaiveDate::from_ymd_opt(before.year(), before.month(), 1).unwrap();
        let after = first_of_month + Duration::days(32);
        let next_month = NaiveDate::from_ymd_opt(after.year(), after.month(), 1).unwrap();
        let month = |first: NaiveDate| {
            date.year() == first.year() && date.month() == first.month()
        };
        match self {
            Self::Yesterday => date == today - Duration::days(1),
            Self::Today => date == today,
            Self::Tomorrow => date == today + Duration::days(1),
            // "Last 7 Days" includes today and the six before it.
            Self::Last7Days => date >= today - Duration::days(6) && date <= today,
            Self::LastWeek => week(this_week - Duration::days(7)),
            Self::ThisWeek => week(this_week),
            Self::NextWeek => week(this_week + Duration::days(7)),
            Self::LastMonth => month(last_month),
            Self::ThisMonth => month(first_of_month),
            Self::NextMonth => month(next_month),
        }
    }
}

/// An Excel date serial as a calendar date. Serial 60 is the 1900 leap-year
/// bug's nonexistent 29 February and is rejected; the 1904 system is not
/// handled.
fn date_of_serial(serial: f64) -> Option<NaiveDate> {
    if !serial.is_finite() {
        return None;
    }
    let serial = serial.floor();
    if serial < 1.0 || serial == 60.0 {
        return None;
    }
    let days = if serial < 61.0 {
        serial - 25568.0
    } else {
        serial - 25569.0
    };
    NaiveDate::from_ymd_opt(1970, 1, 1)?.checked_add_signed(Duration::days(days as i64))
}

/// An optional OOXML boolean: absent means false.
fn flag(value: Option<String>) -> bool {
    value.is_some_and(|v| v == "1" || v == "true")
}

/// The rule a `<cfRule>` element names, if this reader has a way to evaluate it.
/// A rule with no `dxfId` paints no format unless it is a colour scale or a data
/// bar, whose paint is computed from the cell; icon sets stop here.
fn pending_rule(element: &quick_xml::events::BytesStart) -> Option<Pending> {
    let patch = attr(element, b"dxfId").and_then(|v| v.parse::<usize>().ok());
    let kind = match attr(element, b"type").as_deref() {
        Some("expression") => Kind::Expression,
        Some("cellIs") => Kind::CellIs(cell_op(&attr(element, b"operator")?)?),
        Some("containsText") => Kind::Text(TextTest::Contains, attr(element, b"text")?),
        Some("notContainsText") => Kind::Text(TextTest::NotContains, attr(element, b"text")?),
        Some("beginsWith") => Kind::Text(TextTest::Begins, attr(element, b"text")?),
        Some("endsWith") => Kind::Text(TextTest::Ends, attr(element, b"text")?),
        Some("duplicateValues") => Kind::Duplicate(true),
        Some("uniqueValues") => Kind::Duplicate(false),
        Some("top10") => Kind::Top {
            bottom: flag(attr(element, b"bottom")),
            rank: attr(element, b"rank")
                .and_then(|v| v.parse().ok())
                .unwrap_or(10),
            percent: flag(attr(element, b"percent")),
        },
        Some("aboveAverage") => Kind::Average {
            // Absent means above, which is what Excel writes for a rule it
            // created above the average.
            above: !attr(element, b"aboveAverage")
                .is_some_and(|v| v == "0" || v == "false"),
            equal: flag(attr(element, b"equalAverage")),
            deviations: attr(element, b"stdDev")
                .and_then(|v| v.parse::<i32>().ok())
                .map(i32::unsigned_abs)
                .unwrap_or(0),
        },
        Some("timePeriod") => Kind::TimePeriod(TimePeriod::parse(
            &attr(element, b"timePeriod")?,
        )?),
        Some("colorScale") => Kind::ColorScale,
        Some("dataBar") => Kind::DataBar,
        Some("iconSet") => Kind::IconSet,
        _ => return None,
    };
    if patch.is_none()
        && !matches!(
            kind,
            Kind::ColorScale | Kind::DataBar | Kind::IconSet
        )
    {
        return None;
    }
    Some(Pending {
        priority: attr(element, b"priority")
            .and_then(|v| v.parse::<u32>().ok())
            .unwrap_or(u32::MAX),
        stop_if_true: attr(element, b"stopIfTrue")
            .is_some_and(|v| v == "1" || v == "true"),
        patch,
        kind,
        formulas: Vec::new(),
        points: Vec::new(),
        colors: Vec::new(),
        family: None,
        reverse: false,
        x14_id: None,
    })
}

/// One `<cfvo>`: where a colour-scale or data-bar stop sits.
fn scale_point(element: &quick_xml::events::BytesStart) -> Option<ScalePoint> {
    let value = attr(element, b"val").and_then(|v| v.parse::<f64>().ok());
    Some(match attr(element, b"type").as_deref()? {
        "min" => ScalePoint::Min,
        "max" => ScalePoint::Max,
        "num" => ScalePoint::Number(value?),
        "percent" => ScalePoint::Percent(value?),
        "percentile" => ScalePoint::Percentile(value?),
        // Excel normalises a middle stop of a colour scale to a formula, and a
        // fixed position there is simply its number.
        "formula" => ScalePoint::Number(value?),
        _ => return None,
    })
}

fn finish(pending: Pending, sqref: &[MergeRect]) -> Option<Rule> {
    let matcher = match &pending.kind {
        Kind::Expression => Matcher::Expression(Formula::parse(pending.formulas.first()?)?),
        Kind::CellIs(op) => {
            let first = Formula::parse(pending.formulas.first()?)?;
            let second = match op {
                CellOp::Between | CellOp::NotBetween => {
                    Some(Formula::parse(pending.formulas.get(1)?)?)
                }
                _ => None,
            };
            Matcher::CellIs {
                op: *op,
                first,
                second,
            }
        }
        Kind::Text(test, needle) => Matcher::Text {
            test: *test,
            needle: needle.clone(),
        },
        Kind::Duplicate(keep_duplicates) => Matcher::Duplicate {
            keep_duplicates: *keep_duplicates,
        },
        Kind::Top {
            bottom,
            rank,
            percent,
        } => Matcher::Top {
            bottom: *bottom,
            rank: *rank,
            percent: *percent,
        },
        Kind::Average {
            above,
            equal,
            deviations,
        } => Matcher::Average {
            above: *above,
            equal: *equal,
            deviations: *deviations,
        },
        Kind::TimePeriod(period) => Matcher::TimePeriod(*period),
        Kind::ColorScale | Kind::DataBar | Kind::IconSet => Matcher::Always,
    };
    let paint = match pending.kind {
        Kind::ColorScale => {
            if pending.points.len() < 2 || pending.points.len() != pending.colors.len() {
                return None;
            }
            Paint::ColorScale(ColorScale {
                points: pending.points,
                colors: pending.colors,
            })
        }
        Kind::DataBar => {
            let [min, max]: [ScalePoint; 2] =
                pending.points.clone().try_into().ok()?;
            Paint::DataBar(DataBar {
                color: pending.colors.first()?.clone(),
                negative_color: None,
                points: [min, max],
                axis: Axis::None,
                axis_color: None,
            })
        }
        Kind::IconSet => {
            if pending.points.len() < 2 {
                return None;
            }
            Paint::IconSet(IconSet {
                family: pending.family?,
                points: pending.points,
                reverse: pending.reverse,
            })
        }
        _ => Paint::Dxf(pending.patch?),
    };
    Some(Rule {
        priority: pending.priority,
        stop_if_true: pending.stop_if_true,
        paint,
        ranges: sqref.to_vec(),
        matcher,
        x14_id: pending.x14_id,
    })
}

/// `sqref` holds one or more A1 ranges, space separated.
fn sqref_of(sqref: Option<&str>) -> Vec<MergeRect> {
    sqref
        .unwrap_or_default()
        .split_whitespace()
        .filter_map(parse_range)
        .collect()
}

/// What an `<x14:dataBar>` adds to a bar: where zero sits, what a negative bar
/// is coloured, and the axis line. Modern Excel keeps these in the sheet's
/// extension block beside a compatibility rule, which is what ties the two
/// together.
#[derive(Clone, Default)]
struct X14Settings {
    /// Absent in the file means automatic, which is the schema default.
    axis: Option<Axis>,
    negative_same_as_positive: bool,
    negative_color: Option<String>,
    axis_color: Option<String>,
}

/// The extension block's data bars, in file order, each with the GUID a base
/// rule names.
fn x14_bars(xml: &str, palette: Palette) -> Vec<(String, X14Settings)> {
    let mut reader = XmlReader::from_str(xml);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    let mut inside = false;
    let mut id: Option<String> = None;
    let mut is_bar = false;
    let mut settings = X14Settings::default();

    loop {
        buf.clear();
        let event = reader.read_event_into(&mut buf);
        match event {
            Ok(Event::Start(element)) | Ok(Event::Empty(element)) => {
                match element.local_name().as_ref() {
                    b"conditionalFormattings" => inside = true,
                    b"conditionalFormatting" if inside => {
                        id = None;
                        is_bar = false;
                        settings = X14Settings::default();
                    }
                    b"cfRule" if inside => {
                        id = attr(&element, b"id");
                        is_bar = attr(&element, b"type").as_deref() == Some("dataBar");
                    }
                    b"dataBar" if inside && is_bar => {
                        settings.axis = Some(match attr(&element, b"axisPosition").as_deref() {
                            Some("none") => Axis::None,
                            Some("middle") => Axis::Middle,
                            _ => Axis::Auto,
                        });
                        settings.negative_same_as_positive =
                            flag(attr(&element, b"negativeBarColorSameAsPositive"));
                    }
                    b"negativeFillColor" if inside && is_bar => {
                        settings.negative_color = color_of(&element, palette);
                    }
                    b"axisColor" if inside && is_bar => {
                        settings.axis_color = color_of(&element, palette);
                    }
                    _ => {}
                }
            }
            Ok(Event::End(element)) => match element.local_name().as_ref() {
                b"conditionalFormatting" if inside => {
                    if is_bar {
                        if let Some(id) = id.take() {
                            out.push((id, std::mem::take(&mut settings)));
                        }
                    }
                }
                b"conditionalFormattings" => inside = false,
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    out
}

/// Ties the extension block's settings to the base rules that name them.
///
/// A bar written only in the extension — no compatibility rule beside it — is
/// left alone, because Excel leaves it alone: a hand-edited file whose base
/// rules were removed exports a picture with no bars at all. The extension is
/// an addition to a rule, not a rule of its own.
fn attach_x14(rules: &mut [Rule], xml: &str, palette: Palette) {
    let bars = x14_bars(xml, palette);
    if bars.is_empty() {
        return;
    }
    for rule in rules.iter_mut() {
        let Some(id) = rule.x14_id.as_ref() else {
            continue;
        };
        let Some((_, settings)) = bars.iter().find(|(bar_id, _)| bar_id == id) else {
            continue;
        };
        if let Paint::DataBar(bar) = &mut rule.paint {
            bar.axis = settings.axis.unwrap_or(Axis::None);
            bar.axis_color = settings.axis_color.clone();
            if !settings.negative_same_as_positive {
                bar.negative_color = settings.negative_color.clone();
            }
        }
    }
}

impl Rule {
    /// Every column the rule can read, so the value pass knows which ones to
    /// keep. A formula's relative reference reaches a column for every cell of
    /// its range, which is the range's own width either side of where it is
    /// written; a rule that compares the cell's own value needs the range's
    /// columns as well.
    fn columns(&self, out: &mut HashSet<u32>) {
        match &self.matcher {
            Matcher::Expression(formula) => {
                for rect in &self.ranges {
                    formula.columns(rect.cols.max(1), out);
                }
            }
            Matcher::CellIs { first, second, .. } => {
                for rect in &self.ranges {
                    first.columns(rect.cols.max(1), out);
                    if let Some(second) = second {
                        second.columns(rect.cols.max(1), out);
                    }
                }
                self.range_columns(out);
            }
            Matcher::Text { .. }
            | Matcher::Duplicate { .. }
            | Matcher::Top { .. }
            | Matcher::Average { .. }
            | Matcher::TimePeriod(_)
            | Matcher::Always => self.range_columns(out),
        }
    }

    /// The columns its own ranges cover.
    fn range_columns(&self, out: &mut HashSet<u32>) {
        for rect in &self.ranges {
            for col in rect.col..rect.col.saturating_add(rect.cols) {
                out.insert(col);
            }
        }
    }

    /// How often each value appears across the rule's own ranges, for the
    /// duplicate and unique rules. Blank cells are left out: Excel counts
    /// neither, and an empty cell is in neither set.
    fn value_counts(
        &self,
        values: &HashMap<(u32, u32), Value>,
        rows: u32,
        cols: u32,
    ) -> HashMap<ValueKey, usize> {
        let mut out: HashMap<ValueKey, usize> = HashMap::new();
        for rect in &self.ranges {
            let end_row = rect.row.saturating_add(rect.rows).min(rows);
            let end_col = rect.col.saturating_add(rect.cols).min(cols);
            for row in rect.row..end_row {
                for col in rect.col..end_col {
                    let Some(value) = values.get(&(row, col)) else {
                        continue;
                    };
                    if let Some(key) = ValueKey::of(value) {
                        *out.entry(key).or_default() += 1;
                    }
                }
            }
        }
        out
    }

    /// The numbers in the rule's own ranges, ascending. A colour scale and a
    /// data bar both measure a cell against this: the ends of the range, a
    /// percentile of it, or a fixed position in it.
    fn numbers(
        &self,
        values: &HashMap<(u32, u32), Value>,
        rows: u32,
        cols: u32,
    ) -> Vec<f64> {
        let mut out = Vec::new();
        for rect in &self.ranges {
            let end_row = rect.row.saturating_add(rect.rows).min(rows);
            let end_col = rect.col.saturating_add(rect.cols).min(cols);
            for row in rect.row..end_row {
                for col in rect.col..end_col {
                    if let Some(Value::Number(number)) = values.get(&(row, col)) {
                        out.push(*number);
                    }
                }
            }
        }
        out.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
        out
    }
}

impl ColorScale {
    /// The colour a value takes on this scale. A value outside the stops takes
    /// the end colour, which is what Excel does with a clipped scale.
    fn color(&self, context: &Context, numbers: &[f64]) -> Option<String> {
        let stops: Vec<f64> = self
            .points
            .iter()
            .map(|point| point.value(numbers))
            .collect::<Option<Vec<f64>>>()?;
        let value = number(&context.value())?;
        for index in 1..stops.len() {
            if value <= stops[index] || index == stops.len() - 1 {
                let span = stops[index] - stops[index - 1];
                let t = if span == 0.0 {
                    0.0
                } else {
                    ((value - stops[index - 1]) / span).clamp(0.0, 1.0)
                };
                return Some(mix(&self.colors[index - 1], &self.colors[index], t));
            }
        }
        self.colors.first().cloned()
    }
}

/// One cell's bar, measured across the cell. `from` is where the bar starts
/// and `to` where it ends, both as shares of the cell's width.
struct Bar {
    from: f64,
    to: f64,
    negative: bool,
    color: String,
    axis: Option<f64>,
    axis_color: Option<String>,
}

impl DataBar {
    /// The bar a value draws, or None when the value is not a number or the
    /// range has no width.
    ///
    /// Excel's own geometry: with no axis the bars scale from the left edge
    /// over the whole range; with an axis, a positive value runs right from it
    /// and a negative one runs left, each side scaled to its own extreme. The
    /// automatic axis sits where the range's two sides meet —
    /// `-min / (max - min)` — and the middle one at half the cell.
    fn cell(&self, context: &Context, numbers: &[f64]) -> Option<Bar> {
        let min = self.points[0].value(numbers)?;
        let max = self.points[1].value(numbers)?;
        if max <= min {
            return None;
        }
        let value = number(&context.value())?;
        let axis = match self.axis {
            Axis::None => 0.0,
            Axis::Middle => 0.5,
            Axis::Auto => {
                if min < 0.0 && max > 0.0 {
                    -min / (max - min)
                } else {
                    0.0
                }
            }
        };
        let (from, to, negative) = if self.axis == Axis::None || min >= 0.0 {
            (0.0, ((value - min) / (max - min)).clamp(0.0, 1.0), false)
        } else if value >= 0.0 {
            (axis, axis + (value / max) * (1.0 - axis), false)
        } else {
            (axis - (-value / -min) * axis, axis, true)
        };
        let color = if negative {
            self.negative_color
                .clone()
                .unwrap_or_else(|| self.color.clone())
        } else {
            self.color.clone()
        };
        Some(Bar {
            from,
            to,
            negative,
            color,
            axis: (self.axis != Axis::None && min < 0.0 && max > 0.0).then_some(axis),
            axis_color: self.axis_color.clone(),
        })
    }
}

impl IconSet {
    /// Which of the family's marks a value takes. A value below the floor still
    /// takes the lowest mark, and `reverse` flips the family end for end.
    fn level(&self, context: &Context, numbers: &[f64]) -> Option<u8> {
        let value = number(&context.value())?;
        let mut level = 0;
        for (index, point) in self.points.iter().enumerate() {
            if value >= point.value(numbers)? {
                level = index;
            }
        }
        let level = level.min(self.points.len() - 1);
        let level = if self.reverse {
            self.points.len() - 1 - level
        } else {
            level
        };
        Some(level as u8)
    }
}

impl ScalePoint {
    fn value(&self, numbers: &[f64]) -> Option<f64> {
        match self {
            ScalePoint::Min => numbers.first().copied(),
            ScalePoint::Max => numbers.last().copied(),
            ScalePoint::Number(value) => Some(*value),
            ScalePoint::Percent(percent) => {
                let (min, max) = (numbers.first()?, numbers.last()?);
                Some(min + (max - min) * percent / 100.0)
            }
            // The nearest value at or above the percentile — Excel's own reading
            // of a percentile stop, measured against a colour scale: with ten
            // values the 50th lands on the fifth.
            ScalePoint::Percentile(percent) => {
                let count = numbers.len();
                if count == 0 {
                    return None;
                }
                let rank = (percent / 100.0 * count as f64).ceil() as usize;
                let index = rank.saturating_sub(1).min(count - 1);
                numbers.get(index).copied()
            }
        }
    }
}

/// Interpolates two `#RRGGBB` colours channel by channel, which is how Excel
/// moves between a scale's stops.
fn mix(from: &str, to: &str, t: f64) -> String {
    let channel = |hex: &str, range: std::ops::Range<usize>| -> f64 {
        if hex.len() != 7 {
            return 0.0;
        }
        u8::from_str_radix(&hex[range], 16).unwrap_or(0) as f64
    };
    let blend = |range: std::ops::Range<usize>| -> u8 {
        let a = channel(from, range.clone());
        let b = channel(to, range);
        (a + (b - a) * t).round().clamp(0.0, 255.0) as u8
    };
    format!(
        "#{:02X}{:02X}{:02X}",
        blend(1..3),
        blend(3..5),
        blend(5..7)
    )
}

/// Reads the workbook's `<dxfs>`.
pub fn dxfs(xml: &str, palette: Palette) -> Vec<Patch> {
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
                        color_edge(&element, border, edge, palette);
                    }
                    b"color" if part == Part::Font => {
                        if let Some(patch) = patch.as_mut() {
                            patch.color = color_of(&element, palette);
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
                    b"fgColor" if part == Part::Fill => fg = color_of(&element, palette),
                    b"bgColor" if part == Part::Fill => bg = color_of(&element, palette),
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
    let palette = cache.palette();
    let mut rules = rules(xml, palette);
    attach_x14(&mut rules, xml, palette);
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
    let mut ordered: Vec<&Rule> = rules
        .iter()
        .filter(|rule| match rule.paint {
            Paint::Dxf(index) => index < patches.len(),
            _ => true,
        })
        .collect();
    ordered.sort_by_key(|rule| rule.priority);
    if ordered.is_empty() {
        return;
    }
    let today = Local::now().date_naive();

    // The multiset a duplicate or unique rule needs, built once for its whole
    // range rather than once per cell it asks about.
    let counts: Vec<Option<HashMap<ValueKey, usize>>> = ordered
        .iter()
        .map(|rule| match &rule.matcher {
            Matcher::Duplicate { .. } => Some(rule.value_counts(values, layout.rows, layout.cols)),
            _ => None,
        })
        .collect();
    // The numbers a colour scale or data bar measures against, likewise once
    // per rule.
    let numbers: Vec<Option<Vec<f64>>> = ordered
        .iter()
        .map(|rule| match &rule.matcher {
            // The rules that measure a cell against its whole range.
            Matcher::Top { .. } | Matcher::Average { .. } => {
                Some(rule.numbers(values, layout.rows, layout.cols))
            }
            _ => match &rule.paint {
                Paint::ColorScale(_) | Paint::DataBar(_) | Paint::IconSet(_) => {
                    Some(rule.numbers(values, layout.rows, layout.cols))
                }
                _ => None,
            },
        })
        .collect();

    // Which rules reach which row, in the order they are to be looked at. Built
    // up first so that a rule over a range the sheet never reaches costs nothing.
    let mut spans: BTreeMap<u32, Vec<Span<'_>>> = BTreeMap::new();
    let mut budget = MAX_CONDITIONAL_CELLS;
    'rules: for (index, rule) in ordered.iter().enumerate() {
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
                    index,
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
    let mut bars: Vec<crate::layout::DataBarCell> = Vec::new();
    let mut icons: Vec<crate::layout::IconCell> = Vec::new();

    for (&row, spans) in &spans {
        // Every cell a rule fires on. Properties accumulate in priority order,
        // keeping the first rule to set each one: Excel lets a lower-priority
        // rule supply the property a higher-priority one leaves alone. A rule
        // that stops when it fires ends the matter for that cell.
        let mut fired: BTreeMap<u32, Patch> = BTreeMap::new();
        let mut barred: BTreeMap<u32, Bar> = BTreeMap::new();
        let mut marked: BTreeMap<u32, (String, u8)> = BTreeMap::new();
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
                    today,
                };
                let numbers = numbers[span.index].as_deref().unwrap_or(&[]);
                if !span
                    .rule
                    .matcher
                    .matches(&context, counts[span.index].as_ref(), numbers)
                {
                    continue;
                }
                match &span.rule.paint {
                    Paint::Dxf(index) => {
                        merge(fired.entry(col).or_default(), &patches[*index]);
                    }
                    Paint::ColorScale(scale) => {
                        if let Some(fill) = scale.color(&context, numbers) {
                            merge(
                                fired.entry(col).or_default(),
                                &Patch {
                                    fill: Some(fill),
                                    ..Default::default()
                                },
                            );
                        }
                    }
                    Paint::DataBar(bar) => {
                        // Only the first bar to reach a cell is drawn; the rules
                        // are already in priority order.
                        if !barred.contains_key(&col) {
                            if let Some(cell) = bar.cell(&context, numbers) {
                                barred.insert(col, cell);
                            }
                        }
                    }
                    Paint::IconSet(set) => {
                        // Likewise, only the first icon set to reach a cell.
                        if !marked.contains_key(&col) {
                            if let Some(level) = set.level(&context, numbers) {
                                marked.insert(col, (set.family.clone(), level));
                            }
                        }
                    }
                }
                if span.rule.stop_if_true {
                    stopped.insert(col);
                }
            }
        }
        for (&col, bar) in &barred {
            bars.push(crate::layout::DataBarCell {
                row,
                col,
                color: bar.color.clone(),
                from: bar.from,
                to: bar.to,
                negative: bar.negative,
                axis: bar.axis,
                axis_color: bar.axis_color.clone(),
            });
        }
        for (&col, (family, level)) in &marked {
            icons.push(crate::layout::IconCell {
                row,
                col,
                family: family.clone(),
                level: *level,
            });
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
    layout.data_bars = bars;
    layout.icons = icons;
}

/// One rule's range, as it applies to a single row.
struct Span<'a> {
    rule: &'a Rule,
    /// Where the rule sits in the priority-ordered list, for its own precomputed
    /// range values.
    index: usize,
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
    /// The local date the sheet is being viewed on; `timePeriod` measures
    /// against it.
    today: NaiveDate,
}

impl Context<'_> {
    /// The value of the cell the rule is being asked about.
    fn value(&self) -> Value {
        self.values
            .get(&(self.row, self.col))
            .cloned()
            .unwrap_or(Value::Blank)
    }
}

/// A value as a multiset key. Text compares without regard to case, which is
/// how Excel counts duplicates; a blank has no key at all, since it counts as
/// no value rather than as an empty one.
#[derive(PartialEq, Eq, Hash)]
enum ValueKey {
    Number(u64),
    Text(String),
    Bool(bool),
}

impl ValueKey {
    fn of(value: &Value) -> Option<ValueKey> {
        Some(match value {
            Value::Blank => return None,
            // `-0` and `0` are the same value to a rule.
            Value::Number(number) => ValueKey::Number(if *number == 0.0 {
                0f64.to_bits()
            } else {
                number.to_bits()
            }),
            Value::Text(text) => ValueKey::Text(text.to_uppercase()),
            Value::Bool(flag) => ValueKey::Bool(*flag),
        })
    }
}

impl Matcher {
    /// Whether the rule fires for the cell `context` names. `duplicates` is the
    /// multiset of the rule's own range, built once per rule.
    fn matches(
        &self,
        context: &Context,
        duplicates: Option<&HashMap<ValueKey, usize>>,
        numbers: &[f64],
    ) -> bool {
        match self {
            Matcher::Expression(formula) => formula.eval(context) == Some(Value::Bool(true)),
            Matcher::CellIs { op, first, second } => {
                let Some(first) = first.eval(context) else {
                    return false;
                };
                let value = context.value();
                let order = |other: &Value| compare(&value, other);
                match op {
                    CellOp::Equal => order(&first).is_some_and(|order| order.is_eq()),
                    // Excel's `<>` is true for values of different kinds, which is
                    // the same reading its own formula language gives it.
                    CellOp::NotEqual => !order(&first).is_some_and(|order| order.is_eq()),
                    CellOp::Greater => order(&first).is_some_and(|order| order.is_gt()),
                    CellOp::GreaterOrEqual => order(&first).is_some_and(|order| order.is_ge()),
                    CellOp::Less => order(&first).is_some_and(|order| order.is_lt()),
                    CellOp::LessOrEqual => order(&first).is_some_and(|order| order.is_le()),
                    CellOp::Between | CellOp::NotBetween => {
                        let Some(second) = second.as_ref().and_then(|second| second.eval(context))
                        else {
                            return false;
                        };
                        let inside = order(&first).is_some_and(|order| order.is_ge())
                            && compare(&value, &second).is_some_and(|order| order.is_le());
                        if *op == CellOp::Between {
                            inside
                        } else {
                            !inside
                        }
                    }
                }
            }
            Matcher::Text { test, needle } => {
                let Value::Text(value) = context.value() else {
                    return false;
                };
                let (value, needle) = (value.to_uppercase(), needle.to_uppercase());
                match test {
                    TextTest::Contains => value.contains(&needle),
                    TextTest::NotContains => !value.contains(&needle),
                    TextTest::Begins => value.starts_with(&needle),
                    TextTest::Ends => value.ends_with(&needle),
                }
            }
            Matcher::Duplicate { keep_duplicates } => {
                let value = context.value();
                let Some(key) = ValueKey::of(&value) else {
                    return false;
                };
                let count = duplicates
                    .and_then(|counts| counts.get(&key))
                    .copied()
                    .unwrap_or(0);
                if *keep_duplicates {
                    count > 1
                } else {
                    count == 1
                }
            }
            Matcher::Top {
                bottom,
                rank,
                percent,
            } => {
                let Value::Number(value) = context.value() else {
                    return false;
                };
                let count = numbers.len();
                if count == 0 {
                    return false;
                }
                let take = if *percent {
                    ((count as f64 * *rank as f64 / 100.0).ceil() as usize).max(1)
                } else {
                    (*rank as usize).max(1)
                }
                .min(count);
                // Ties at the cut are all in, which is how Excel highlights a
                // rank whose boundary value repeats.
                if *bottom {
                    value <= numbers[take - 1]
                } else {
                    value >= numbers[count - take]
                }
            }
            Matcher::Average {
                above,
                equal,
                deviations,
            } => {
                let Value::Number(value) = context.value() else {
                    return false;
                };
                if numbers.is_empty() {
                    return false;
                }
                let count = numbers.len() as f64;
                let mean = numbers.iter().sum::<f64>() / count;
                let variance = numbers
                    .iter()
                    .map(|number| (number - mean) * (number - mean))
                    .sum::<f64>()
                    / count;
                let spread = *deviations as f64 * variance.sqrt();
                let threshold = if *above { mean + spread } else { mean - spread };
                if *above {
                    if *equal {
                        value >= threshold
                    } else {
                        value > threshold
                    }
                } else if *equal {
                    value <= threshold
                } else {
                    value < threshold
                }
            }
            Matcher::TimePeriod(period) => {
                let Value::Number(serial) = context.value() else {
                    return false;
                };
                // Excel compares the date part only; a General-format serial
                // fires just as a date-formatted cell does.
                date_of_serial(serial).is_some_and(|date| period.matches(date, context.today))
            }
            Matcher::Always => true,
        }
    }
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

    #[test]
    fn time_periods_match_excels_windows() {
        // 2026-10-22 is a Thursday; Excel's weeks run Sunday to Saturday.
        let today = NaiveDate::from_ymd_opt(2026, 10, 22).unwrap();
        let date = |month, day| NaiveDate::from_ymd_opt(2026, month, day).unwrap();

        assert!(TimePeriod::Yesterday.matches(date(10, 21), today));
        assert!(TimePeriod::Today.matches(date(10, 22), today));
        assert!(TimePeriod::Tomorrow.matches(date(10, 23), today));

        assert!(TimePeriod::Last7Days.matches(date(10, 16), today));
        assert!(!TimePeriod::Last7Days.matches(date(10, 15), today));

        assert!(TimePeriod::ThisWeek.matches(date(10, 18), today));
        assert!(TimePeriod::ThisWeek.matches(date(10, 24), today));
        assert!(!TimePeriod::ThisWeek.matches(date(10, 25), today));
        assert!(TimePeriod::LastWeek.matches(date(10, 11), today));
        assert!(TimePeriod::LastWeek.matches(date(10, 17), today));
        assert!(!TimePeriod::LastWeek.matches(date(10, 18), today));
        assert!(TimePeriod::NextWeek.matches(date(10, 25), today));
        assert!(TimePeriod::NextWeek.matches(date(10, 31), today));

        assert!(TimePeriod::LastMonth.matches(date(9, 1), today));
        assert!(TimePeriod::LastMonth.matches(date(9, 30), today));
        assert!(!TimePeriod::LastMonth.matches(date(10, 1), today));
        assert!(TimePeriod::ThisMonth.matches(date(10, 1), today));
        assert!(TimePeriod::ThisMonth.matches(date(10, 31), today));
        assert!(TimePeriod::NextMonth.matches(date(11, 1), today));
        assert!(TimePeriod::NextMonth.matches(date(11, 30), today));
    }

    #[test]
    fn excel_serials_keep_the_date_part() {
        let today = NaiveDate::from_ymd_opt(2026, 10, 22).unwrap();
        assert_eq!(date_of_serial(46317.0), Some(today));
        assert_eq!(date_of_serial(46317.752), Some(today));
        assert_eq!(date_of_serial(60.0), None);
        assert_eq!(date_of_serial(0.0), None);
    }

    #[test]
    fn time_period_rule_is_read() {
        let xml = r#"<worksheet><conditionalFormatting sqref="A1:A3"><cfRule type="timePeriod" timePeriod="thisWeek" dxfId="0" priority="1"/></conditionalFormatting></worksheet>"#;
        let rules = rules(xml, Palette::EMPTY);
        assert_eq!(rules.len(), 1);
        assert!(matches!(
            rules[0].matcher,
            Matcher::TimePeriod(TimePeriod::ThisWeek)
        ));
    }

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
            today: NaiveDate::from_ymd_opt(2026, 10, 22).unwrap(),
        };
        Formula::parse(formula)?.eval(&context)
    }

    fn text(value: &str) -> Value {
        Value::Text(value.to_owned())
    }

    /// Whether a rule fires for one cell, with the range's values laid out as a
    /// small sheet would.
    fn fires(rule: &Rule, row: u32, col: u32, cells: &[((u32, u32), Value)]) -> bool {
        fires_with(rule, row, col, cells, &[])
    }

    fn fires_with(
        rule: &Rule,
        row: u32,
        col: u32,
        cells: &[((u32, u32), Value)],
        numbers: &[f64],
    ) -> bool {
        let values: HashMap<(u32, u32), Value> = cells.iter().cloned().collect();
        let context = Context {
            row,
            col,
            anchor_row: 0,
            anchor_col: 0,
            values: &values,
            today: NaiveDate::from_ymd_opt(2026, 10, 22).unwrap(),
        };
        let counts = match &rule.matcher {
            Matcher::Duplicate { .. } => Some(rule.value_counts(&values, 8, 8)),
            _ => None,
        };
        rule.matcher.matches(&context, counts.as_ref(), numbers)
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

        let patches = dxfs(xml, Palette::EMPTY);
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
        let patches = dxfs(xml, Palette::EMPTY);
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
        let rules = rules(xml, Palette::EMPTY);
        assert_eq!(rules.len(), 1);
        let rule = &rules[0];
        assert!(matches!(rule.paint, Paint::Dxf(0)));
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
    fn reads_a_cell_is_rule_and_compares_the_cell() {
        let xml = r#"<worksheet><conditionalFormatting sqref="B2:B9"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThanOrEqual"><formula>50</formula></cfRule></conditionalFormatting></worksheet>"#;
        let rules = rules(xml, Palette::EMPTY);
        assert_eq!(rules.len(), 1);
        let rule = &rules[0];
        // B, the column the rule has to read to know the cell's own value.
        let mut columns = HashSet::new();
        rule.columns(&mut columns);
        assert_eq!(columns, HashSet::from([1]));
        assert!(fires(rule, 1, 1, &[((1, 1), Value::Number(60.0))]));
        assert!(fires(rule, 1, 1, &[((1, 1), Value::Number(50.0))]));
        assert!(!fires(rule, 1, 1, &[((1, 1), Value::Number(40.0))]));
        // An empty cell counts as zero, which is below the bound.
        assert!(!fires(rule, 1, 1, &[]));
    }

    #[test]
    fn a_cell_is_rule_can_ask_for_a_between() {
        let xml = r#"<worksheet><conditionalFormatting sqref="A1:A4"><cfRule type="cellIs" dxfId="0" priority="1" operator="between"><formula>10</formula><formula>20</formula></cfRule></conditionalFormatting></worksheet>"#;
        let rules = rules(xml, Palette::EMPTY);
        assert_eq!(rules.len(), 1);
        assert!(fires(&rules[0], 0, 0, &[((0, 0), Value::Number(15.0))]));
        assert!(!fires(&rules[0], 0, 0, &[((0, 0), Value::Number(25.0))]));
    }

    /// Excel writes a `<formula>` for its text tests too, but the `text`
    /// attribute says the same thing more directly, and the test is the one the
    /// grid can evaluate without a whole formula language.
    #[test]
    fn reads_the_text_rules() {
        let xml = r#"<worksheet>
            <conditionalFormatting sqref="A1:A9"><cfRule type="containsText" dxfId="0" priority="1" operator="containsText" text="abc"><formula>NOT(ISERROR(SEARCH("abc",A1)))</formula></cfRule></conditionalFormatting>
            <conditionalFormatting sqref="B1:B9"><cfRule type="beginsWith" dxfId="0" priority="2" operator="beginsWith" text="x"><formula>LEFT(B1,1)="x"</formula></cfRule></conditionalFormatting>
            <conditionalFormatting sqref="C1:C9"><cfRule type="endsWith" dxfId="0" priority="3" operator="endsWith" text="z"><formula>RIGHT(C1,1)="z"</formula></cfRule></conditionalFormatting>
            <conditionalFormatting sqref="D1:D9"><cfRule type="notContainsText" dxfId="0" priority="4" operator="notContainsText" text="q"><formula>ISERROR(SEARCH("q",D1))</formula></cfRule></conditionalFormatting>
        </worksheet>"#;
        let rules = rules(xml, Palette::EMPTY);
        assert_eq!(rules.len(), 4);
        // Excel's text tests ignore case.
        assert!(fires(&rules[0], 0, 0, &[((0, 0), text("xABCx"))]));
        assert!(!fires(&rules[0], 0, 0, &[((0, 0), text("xyz"))]));
        assert!(!fires(&rules[0], 0, 0, &[((0, 0), Value::Number(1.0))]));
        assert!(fires(&rules[1], 0, 1, &[((0, 1), text("Xylophone"))]));
        assert!(fires(&rules[2], 0, 2, &[((0, 2), text("fizz"))]));
        assert!(fires(&rules[3], 0, 3, &[((0, 3), text("abc"))]));
        assert!(!fires(&rules[3], 0, 3, &[((0, 3), text("aqb"))]));
    }

    #[test]
    fn reads_the_duplicate_and_unique_rules() {
        let xml = r#"<worksheet>
            <conditionalFormatting sqref="A1:A5"><cfRule type="duplicateValues" dxfId="0" priority="1"/></conditionalFormatting>
            <conditionalFormatting sqref="A1:A5"><cfRule type="uniqueValues" dxfId="0" priority="2"/></conditionalFormatting>
        </worksheet>"#;
        let rules = rules(xml, Palette::EMPTY);
        assert_eq!(rules.len(), 2, "a rule with no children is complete on its own");
        let cells = [
            ((0, 0), Value::Number(1.0)),
            ((1, 0), Value::Number(2.0)),
            ((2, 0), text("b")),
            ((3, 0), Value::Number(2.0)),
            ((4, 0), text("B")),
        ];
        // 1 is unique, 2 repeats, and "b"/"B" repeat regardless of case.
        assert!(!fires(&rules[0], 0, 0, &cells));
        assert!(fires(&rules[0], 1, 0, &cells));
        assert!(fires(&rules[0], 2, 0, &cells));
        assert!(fires(&rules[0], 4, 0, &cells));
        // The blank A6 is in neither set.
        assert!(!fires(&rules[0], 5, 0, &cells));
        assert!(fires(&rules[1], 0, 0, &cells));
        assert!(!fires(&rules[1], 1, 0, &cells));
        assert!(!fires(&rules[1], 5, 0, &cells));
    }

    #[test]
    fn reads_a_colour_scale_and_mixes_between_its_stops() {
        let xml = r#"<worksheet><conditionalFormatting sqref="B2:B11"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FF0000FF"/><color rgb="FFFF0000"/></colorScale></cfRule></conditionalFormatting></worksheet>"#;
        let rules = rules(xml, Palette::EMPTY);
        assert_eq!(rules.len(), 1);
        let Paint::ColorScale(scale) = &rules[0].paint else {
            panic!("a colour scale");
        };
        let numbers = [10.0, 20.0, 30.0, 100.0];
        let color = |value: f64| {
            let values = HashMap::from([((0u32, 0u32), Value::Number(value))]);
            let context = Context {
                row: 0,
                col: 0,
                anchor_row: 0,
                anchor_col: 0,
                values: &values,
                today: NaiveDate::from_ymd_opt(2026, 10, 22).unwrap(),
            };
            scale.color(&context, &numbers)
        };
        assert_eq!(color(10.0).as_deref(), Some("#0000FF"));
        assert_eq!(color(100.0).as_deref(), Some("#FF0000"));
        assert_eq!(color(55.0).as_deref(), Some("#800080"));
        // Outside the stops the end colour stands, as Excel's clipped scale does.
        assert_eq!(color(0.0).as_deref(), Some("#0000FF"));
    }

    #[test]
    fn reads_a_data_bar_and_measures_the_cell() {
        let xml = r#"<worksheet><conditionalFormatting sqref="D2:D11"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF0000FF"/></dataBar></cfRule></conditionalFormatting></worksheet>"#;
        let rules = rules(xml, Palette::EMPTY);
        assert_eq!(rules.len(), 1);
        let Paint::DataBar(bar) = &rules[0].paint else {
            panic!("a data bar");
        };
        assert_eq!(bar.color, "#0000FF");
        let numbers = [10.0, 100.0];
        let ends = |value: f64| {
            let values = HashMap::from([((0u32, 0u32), Value::Number(value))]);
            let context = Context {
                row: 0,
                col: 0,
                anchor_row: 0,
                anchor_col: 0,
                values: &values,
                today: NaiveDate::from_ymd_opt(2026, 10, 22).unwrap(),
            };
            bar.cell(&context, &numbers).map(|bar| (bar.from, bar.to))
        };
        assert_eq!(ends(10.0), Some((0.0, 0.0)));
        assert_eq!(ends(55.0), Some((0.0, 0.5)));
        assert_eq!(ends(100.0), Some((0.0, 1.0)));
        // A value below the range is clamped, not drawn backwards.
        assert_eq!(ends(5.0), Some((0.0, 0.0)));
        assert_eq!(ends(150.0), Some((0.0, 1.0)));
    }

    /// With an axis, a negative value runs left of it and a positive one right,
    /// each side scaled to its own extreme; the automatic axis sits where the
    /// range's two sides meet.
    #[test]
    fn a_data_bar_with_an_axis_runs_both_ways() {
        let numbers = [-20.0, 0.0, 50.0, 100.0];
        let bar = DataBar {
            color: "#0000FF".to_owned(),
            negative_color: None,
            points: [ScalePoint::Min, ScalePoint::Max],
            axis: Axis::Auto,
            axis_color: None,
        };
        let cell = |value: f64| {
            let values = HashMap::from([((0u32, 0u32), Value::Number(value))]);
            let context = Context {
                row: 0,
                col: 0,
                anchor_row: 0,
                anchor_col: 0,
                values: &values,
                today: NaiveDate::from_ymd_opt(2026, 10, 22).unwrap(),
            };
            bar.cell(&context, &numbers)
        };
        // The axis is at 20/120 of the cell.
        let axis = 20.0 / 120.0;
        let low = cell(-20.0).expect("a bar");
        assert!(low.negative);
        assert!((low.from - 0.0).abs() < 1e-9);
        assert!((low.to - axis).abs() < 1e-9);
        let high = cell(100.0).expect("a bar");
        assert!(!high.negative);
        assert!((high.from - axis).abs() < 1e-9);
        assert!((high.to - 1.0).abs() < 1e-9);
        // Zero draws no length but still marks the axis.
        let zero = cell(0.0).expect("a bar");
        assert!((zero.from - axis).abs() < 1e-9);
        assert!((zero.to - axis).abs() < 1e-9);
        assert_eq!(zero.axis, Some(axis));
    }

    #[test]
    fn reads_the_top_rank_rules() {
        let xml = r#"<worksheet>
            <conditionalFormatting sqref="A1:A10"><cfRule type="top10" dxfId="0" priority="1" rank="3"/></conditionalFormatting>
            <conditionalFormatting sqref="B1:B10"><cfRule type="top10" dxfId="0" priority="2" rank="2" bottom="1"/></conditionalFormatting>
            <conditionalFormatting sqref="C1:C10"><cfRule type="top10" dxfId="0" priority="3" rank="20" percent="1"/></conditionalFormatting>
        </worksheet>"#;
        let rules = rules(xml, Palette::EMPTY);
        assert_eq!(rules.len(), 3);
        let numbers: Vec<f64> = (1..=10).map(|n| n as f64 * 10.0).collect();
        let top3 = |value: f64| {
            fires_with(
                &rules[0],
                0,
                0,
                &[((0, 0), Value::Number(value))],
                &numbers,
            )
        };
        assert!(top3(100.0));
        assert!(top3(80.0));
        assert!(!top3(70.0));
        assert!(fires_with(
            &rules[1],
            0,
            1,
            &[((0, 1), Value::Number(20.0))],
            &numbers
        ));
        assert!(!fires_with(
            &rules[1],
            0,
            1,
            &[((0, 1), Value::Number(30.0))],
            &numbers
        ));
        // Twenty per cent of ten values is the top two.
        assert!(fires_with(
            &rules[2],
            0,
            2,
            &[((0, 2), Value::Number(90.0))],
            &numbers
        ));
        assert!(!fires_with(
            &rules[2],
            0,
            2,
            &[((0, 2), Value::Number(80.0))],
            &numbers
        ));
    }

    #[test]
    fn reads_the_average_rules() {
        let xml = r#"<worksheet>
            <conditionalFormatting sqref="A1:A10"><cfRule type="aboveAverage" dxfId="0" priority="1"/></conditionalFormatting>
            <conditionalFormatting sqref="B1:B10"><cfRule type="aboveAverage" dxfId="0" priority="2" aboveAverage="0" equalAverage="1"/></conditionalFormatting>
        </worksheet>"#;
        let rules = rules(xml, Palette::EMPTY);
        assert_eq!(rules.len(), 2);
        // 10 through 100: the mean is 55.
        let numbers: Vec<f64> = (1..=10).map(|n| n as f64 * 10.0).collect();
        let above = |value: f64| {
            fires_with(
                &rules[0],
                0,
                0,
                &[((0, 0), Value::Number(value))],
                &numbers,
            )
        };
        assert!(above(60.0));
        // Strictly above, unless `equalAverage` says otherwise.
        assert!(!above(55.0));
        assert!(fires_with(
            &rules[1],
            0,
            1,
            &[((0, 1), Value::Number(55.0))],
            &numbers
        ));
        assert!(!fires_with(
            &rules[1],
            0,
            1,
            &[((0, 1), Value::Number(60.0))],
            &numbers
        ));
    }

    #[test]
    fn reads_an_icon_set_and_picks_a_level() {
        let xml = r#"<worksheet>
            <conditionalFormatting sqref="A2:A11"><cfRule type="iconSet" priority="1"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule></conditionalFormatting>
            <conditionalFormatting sqref="B2:B11"><cfRule type="iconSet" priority="2"><iconSet iconSet="3TrafficLights1" reverse="1"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule></conditionalFormatting>
        </worksheet>"#;
        let rules = rules(xml, Palette::EMPTY);
        assert_eq!(rules.len(), 2);
        let Paint::IconSet(set) = &rules[0].paint else {
            panic!("an icon set");
        };
        assert_eq!(set.family, "3Arrows");
        let numbers = [10.0, 100.0];
        let level = |rule: &Rule, value: f64| {
            let Paint::IconSet(set) = &rule.paint else {
                panic!("an icon set");
            };
            let values = HashMap::from([((0u32, 0u32), Value::Number(value))]);
            let context = Context {
                row: 0,
                col: 0,
                anchor_row: 0,
                anchor_col: 0,
                values: &values,
                today: NaiveDate::from_ymd_opt(2026, 10, 22).unwrap(),
            };
            set.level(&context, &numbers)
        };
        assert_eq!(level(&rules[0], 10.0), Some(0));
        assert_eq!(level(&rules[0], 50.0), Some(1));
        assert_eq!(level(&rules[0], 100.0), Some(2));
        // The reversed family counts from the other end.
        assert_eq!(level(&rules[1], 100.0), Some(0));
        assert_eq!(level(&rules[1], 10.0), Some(2));
    }

    #[test]
    fn keeps_several_ranges_on_one_rule() {
        let xml = r#"<worksheet><conditionalFormatting sqref="A1:A5 C1:C5"><cfRule type="expression" dxfId="2" priority="1"><formula>$A1&gt;0</formula></cfRule></conditionalFormatting></worksheet>"#;
        let rules = rules(xml, Palette::EMPTY);
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
        assert!(rules(xml, Palette::EMPTY).is_empty());
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
            Palette::EMPTY,
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
            Palette::EMPTY,
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
            Palette::EMPTY,
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
