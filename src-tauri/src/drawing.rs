//! The shapes and pictures a sheet floats above its cells.
//!
//! Excel keeps these in a part of their own — `xl/drawings/drawingN.xml` — which
//! the sheet XML only points at by relationship id. Every shape is anchored to a
//! cell plus an offset in EMU, so placing one at all means asking the layout for
//! its own column widths and row heights; the file carries no pixel geometry.
//!
//! A shape's box is reported as *cell plus pixels into that cell* rather than an
//! absolute offset, because the grid compresses filtered-out rows and an
//! absolute offset would leave the shapes behind.

use std::collections::HashMap;

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader as XmlReader;
use quick_xml::XmlVersion;
use serde::Serialize;

use crate::layout::{attr, num_attr, tint, BLACK};

/// English Metric Units per CSS pixel: 914400 to the inch, 96 pixels to the inch.
const EMU_PER_PX: f64 = 9525.0;
/// A run with no `sz` renders at Excel's shape default, 18pt.
const DEFAULT_TEXT_PT: f64 = 18.0;
/// A line with no `w` renders at Excel's shape default, 0.75pt.
const DEFAULT_LINE_PX: f64 = 1.0;
/// Excel's own body insets, in EMU: 0.1in left and right, 0.05in top and bottom.
const DEFAULT_INSETS_EMU: [f64; 4] = [91440.0, 45720.0, 91440.0, 45720.0];

/// One shape, picture or connector.
#[derive(Serialize, Clone, Debug, Default)]
#[serde(rename_all = "camelCase")]
pub struct Shape {
    /// The preset geometry (`rect`, `ellipse`, `borderCallout2`, …), `picture`
    /// for an image, or `frame` for a chart or table, which we do not draw.
    pub kind: String,
    pub from_col: u32,
    pub from_x: f64,
    pub from_row: u32,
    pub from_y: f64,
    pub to_col: u32,
    pub to_x: f64,
    pub to_row: u32,
    pub to_y: f64,
    /// Degrees clockwise.
    pub rot: f64,
    pub flip_h: bool,
    pub flip_v: bool,
    /// Absent when the shape has no fill, which is most of them.
    pub fill: Option<Fill>,
    pub line: Option<Stroke>,
    pub text: Option<Text>,
    /// Pictures only: the media part inside the workbook, e.g. `xl/media/image3.png`.
    pub media: Option<String>,
    /// Preset adjustments in the order the preset declares them (`adj1`…), in
    /// the file's own units — hundred-thousandths. The renderer feeds them
    /// straight into the preset's formulas, which divide by 100000 themselves.
    pub adjust: Vec<f64>,
}

#[derive(Serialize, Clone, Debug, Default, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct Fill {
    /// A solid colour, or a gradient's first stop.
    pub color: String,
    /// Excel's `<a:alpha>`: 1.0 is opaque.
    pub alpha: f64,
    /// The stops of a gradient, in order. Empty for a solid fill.
    pub stops: Vec<Stop>,
    /// A linear gradient's direction in degrees, clockwise from pointing right.
    pub angle: f64,
}

#[derive(Serialize, Clone, Debug, Default, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct Stop {
    /// 0 to 1 across the shape.
    pub pos: f64,
    pub color: String,
    pub alpha: f64,
}

impl Fill {
    /// A plain colour, with no gradient behind it.
    fn solid(color: String, alpha: f64) -> Self {
        Self {
            color,
            alpha,
            ..Self::default()
        }
    }
}

#[derive(Serialize, Clone, Debug, Default)]
#[serde(rename_all = "camelCase")]
pub struct Stroke {
    pub color: String,
    pub alpha: f64,
    pub width: f64,
    /// The file's own preset name (`dash`, `lgDash`, `sysDot`, …); the renderer
    /// turns it into a dash pattern.
    pub dash: Option<String>,
    /// Arrowheads. Without these every arrow in a transition diagram is a bare
    /// line pointing nowhere.
    pub head: Option<Arrow>,
    pub tail: Option<Arrow>,
}

/// An arrowhead at one end of a line, as the file names it.
#[derive(Serialize, Clone, Debug, Default, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct Arrow {
    /// `triangle`, `stealth`, `diamond`, `oval` or `arrow`.
    pub kind: String,
    /// `sm`, `med` or `lg`.
    pub width: String,
    pub length: String,
}

/// A shape's text. Paragraphs are kept apart so the renderer can lay out lines;
/// the styling is shared, since real documents do not vary it within one shape.
#[derive(Serialize, Clone, Debug, Default)]
#[serde(rename_all = "camelCase")]
pub struct Text {
    pub paras: Vec<Vec<Run>>,
    /// Points, before the 96/72 conversion.
    pub size_pt: f64,
    pub color: String,
    /// `left`, `center` or `right`.
    pub align: String,
    /// `top`, `middle` or `bottom`.
    pub valign: String,
    /// Whether the text wraps at the shape's edge. Excel's `wrap="none"` lets it
    /// spill out, the way an unstyled cell does.
    pub wrap: bool,
    /// `lIns`, `tIns`, `rIns`, `bIns` in pixels.
    pub insets: [f64; 4],
}

#[derive(Serialize, Clone, Debug, Default)]
#[serde(rename_all = "camelCase")]
pub struct Run {
    pub text: String,
    pub bold: bool,
    pub italic: bool,
}

/// The layout arithmetic a drawing needs to turn EMU offsets into cell offsets.
pub struct Grid {
    /// Cumulative left edge of columns `0..=cols`, in pixels.
    col_starts: Vec<f64>,
    default_col_width: f64,
    row_heights: HashMap<u32, f64>,
    default_row_height: f64,
    rows: u32,
}

impl Grid {
    pub fn new(
        cols: u32,
        col_widths: &HashMap<u32, f64>,
        default_col_width: f64,
        rows: u32,
        row_heights: &HashMap<u32, f64>,
        default_row_height: f64,
    ) -> Self {
        let mut col_starts = Vec::with_capacity(cols as usize + 2);
        let mut x = 0.0;
        for col in 0..=cols {
            col_starts.push(x);
            x += col_widths.get(&col).copied().unwrap_or(default_col_width);
        }
        col_starts.push(x);
        Self {
            col_starts,
            default_col_width: default_col_width.max(1.0),
            row_heights: row_heights.clone(),
            default_row_height: default_row_height.max(1.0),
            rows,
        }
    }

    /// A shape can sit past the sheet's own extent — Excel keeps one anchored to
    /// a cell the used range never grew to — so both lookups extrapolate with the
    /// default size rather than clamping to the edge.
    fn col_start(&self, col: u32) -> f64 {
        match self.col_starts.get(col as usize) {
            Some(start) => *start,
            None => {
                let end = self.col_starts.last().copied().unwrap_or(0.0);
                let past = f64::from(col - (self.col_starts.len() as u32 - 1));
                end + self.default_col_width * past
            }
        }
    }

    fn row_height(&self, row: u32) -> f64 {
        self.row_heights
            .get(&row)
            .copied()
            .unwrap_or(self.default_row_height)
    }

    fn row_start(&self, row: u32) -> f64 {
        // Only ever asked about a row a shape is anchored near, and rows past the
        // declared heights are uniform, so the walk is short in practice.
        (0..row).map(|index| self.row_height(index)).sum()
    }

    /// The column a pixel offset falls in, plus the pixels into it.
    fn col_at(&self, x: f64) -> (u32, f64) {
        let last = self.col_starts.len().saturating_sub(2);
        let index = match self.col_starts[..=last]
            .binary_search_by(|start| start.partial_cmp(&x).unwrap_or(std::cmp::Ordering::Less))
        {
            Ok(index) => index.min(last),
            Err(0) => 0,
            Err(index) => (index - 1).min(last),
        };
        let mut col = index as u32;
        let mut start = self.col_starts[index];
        // Past the table every column takes the default width.
        while x >= start + self.default_col_width {
            start += self.default_col_width;
            col += 1;
        }
        (col, x - start)
    }

    /// The row a pixel offset falls in, plus the pixels into it. Rows are walked
    /// rather than precomputed: a sheet can claim a million of them, and only a
    /// handful of shapes are placed this way.
    fn row_at(&self, y: f64) -> (u32, f64) {
        let mut start = 0.0;
        let mut row = 0;
        while row < self.rows && y >= start + self.row_height(row) {
            start += self.row_height(row);
            row += 1;
        }
        while y >= start + self.default_row_height {
            start += self.default_row_height;
            row += 1;
        }
        (row, (y - start).max(0.0))
    }
}

/// Reads one sheet's drawing part into shapes.
pub fn parse(xml: &str, grid: &Grid, theme: &[Option<String>], ln_widths: &[f64]) -> Vec<Shape> {
    Parser::new(grid, theme, ln_widths).run(xml)
}

/// An anchor corner as the file writes it: a cell plus an EMU offset into it.
#[derive(Clone, Copy, Default)]
struct Point {
    col: u32,
    col_off: f64,
    row: u32,
    row_off: f64,
}

#[derive(Clone, Copy, PartialEq)]
enum AnchorKind {
    TwoCell,
    OneCell,
    Absolute,
}

/// Whether the file said nothing about a fill (`None`) or said "no fill"
/// (`Some(None)`), which is not the same thing when a style ref could supply one.
type ExplicitFill = Option<Option<Fill>>;

#[derive(Default)]
struct LineSpec {
    /// True once `<a:ln>` was seen at all, even with nothing inside it.
    declared: bool,
    none: bool,
    color: Option<Fill>,
    width: Option<f64>,
    dash: Option<String>,
    head: Option<Arrow>,
    tail: Option<Arrow>,
}

/// The `xdr:style` refs, which colour a shape that declares no colours itself.
#[derive(Default)]
struct StyleRefs {
    ln: Option<(u32, Fill)>,
    fill: Option<(u32, Fill)>,
    font: Option<Fill>,
}

/// Which slot the colour element now being read will fill.
#[derive(Clone, Copy, PartialEq)]
enum Target {
    Fill,
    Line,
    Run,
    LnRef,
    FillRef,
    FontRef,
    /// A `<a:gs>` stop inside a gradient, which is being collected rather than
    /// committed straight to the shape.
    GradientStop,
    /// `<a:effectRef>` carries a colour too, but it tints a shadow we do not
    /// draw. It needs a target of its own so it cannot be mistaken for a fill.
    EffectRef,
}

struct Parser<'a> {
    grid: &'a Grid,
    theme: &'a [Option<String>],
    ln_widths: &'a [f64],
    out: Vec<Shape>,
    /// The element names currently open, so a leaf can ask what encloses it.
    stack: Vec<Vec<u8>>,
    anchor: Option<AnchorKind>,
    from: Point,
    to: Option<Point>,
    size: (f64, f64),
    pos: (f64, f64),
    current: Option<Shape>,
    fill: ExplicitFill,
    line: LineSpec,
    refs: StyleRefs,
    /// The colour element being read, committed when it closes.
    pending: Option<(Fill, Target)>,
    /// The `idx` of the style ref now open, which selects its colour variant.
    ref_slot: Option<Target>,
    ref_index: u32,
    /// The gradient being collected, as its stops and its direction.
    gradient: Option<(Vec<Stop>, f64)>,
    paras: Vec<Vec<Run>>,
    run: Option<Run>,
    text_size: Option<f64>,
    text_align: Option<String>,
    text_valign: Option<String>,
    text_wrap: Option<bool>,
    text_insets: Option<[f64; 4]>,
    text_color: Option<Fill>,
}

impl<'a> Parser<'a> {
    fn new(grid: &'a Grid, theme: &'a [Option<String>], ln_widths: &'a [f64]) -> Self {
        Self {
            grid,
            theme,
            ln_widths,
            out: Vec::new(),
            stack: Vec::new(),
            anchor: None,
            from: Point::default(),
            to: None,
            size: (0.0, 0.0),
            pos: (0.0, 0.0),
            current: None,
            fill: None,
            line: LineSpec::default(),
            refs: StyleRefs::default(),
            pending: None,
            ref_slot: None,
            ref_index: 0,
            gradient: None,
            paras: Vec::new(),
            run: None,
            text_size: None,
            text_align: None,
            text_valign: None,
            text_wrap: None,
            text_insets: None,
            text_color: None,
        }
    }

    /// The local name of the element enclosing the one being read.
    fn parent(&self) -> &[u8] {
        self.stack.last().map(Vec::as_slice).unwrap_or(b"")
    }

    fn run(mut self, xml: &str) -> Vec<Shape> {
        let mut reader = XmlReader::from_str(xml);
        let mut buf = Vec::new();

        loop {
            buf.clear();
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(element)) => {
                    let name = element.local_name().as_ref().to_vec();
                    self.start(&element, &name);
                    self.stack.push(name);
                }
                Ok(Event::Empty(element)) => {
                    let name = element.local_name().as_ref().to_vec();
                    self.start(&element, &name);
                    // A self-closing colour element never gets an end event, and
                    // most colours in a drawing — every plain `<a:srgbClr/>` — are
                    // written that way.
                    self.flush(&name);
                }
                Ok(Event::Text(text)) => {
                    if let Ok(content) = text.xml_content(XmlVersion::default()) {
                        self.text(&content);
                    }
                }
                Ok(Event::End(element)) => {
                    let name = element.local_name().as_ref().to_vec();
                    self.end(&name);
                    if self.stack.last().map(Vec::as_slice) == Some(name.as_slice()) {
                        self.stack.pop();
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
        self.out
    }

    /// `a:col`, `a:colOff`, `a:row`, `a:rowOff` and `a:t` all put their value in
    /// the element's own text; nothing else in a drawing carries any.
    fn text(&mut self, content: &str) {
        match self.parent() {
            b"col" => {
                let value = content.trim().parse::<u32>().unwrap_or(0);
                self.corner().col = value;
            }
            b"colOff" => {
                let value = content.trim().parse::<f64>().unwrap_or(0.0);
                self.corner().col_off = value;
            }
            b"row" => {
                let value = content.trim().parse::<u32>().unwrap_or(0);
                self.corner().row = value;
            }
            b"rowOff" => {
                let value = content.trim().parse::<f64>().unwrap_or(0.0);
                self.corner().row_off = value;
            }
            b"t" => {
                if let Some(run) = self.run.as_mut() {
                    run.text.push_str(content);
                }
            }
            _ => {}
        }
    }

    /// The corner the enclosing element belongs to.
    fn corner(&mut self) -> &mut Point {
        match self.stack.get(self.stack.len().wrapping_sub(2)).map(Vec::as_slice) {
            Some(b"to") => self.to.get_or_insert_with(Point::default),
            _ => &mut self.from,
        }
    }

    fn start(&mut self, element: &BytesStart, name: &[u8]) {
        let parent = self.parent().to_vec();
        match name {
            b"twoCellAnchor" | b"oneCellAnchor" | b"absoluteAnchor" => {
                self.anchor = Some(match name {
                    b"twoCellAnchor" => AnchorKind::TwoCell,
                    b"oneCellAnchor" => AnchorKind::OneCell,
                    _ => AnchorKind::Absolute,
                });
                self.from = Point::default();
                self.to = None;
                self.size = (0.0, 0.0);
                self.pos = (0.0, 0.0);
            }
            b"to" => self.to = Some(Point::default()),
            // Only the anchor's own `ext` sizes the shape; the one inside `xfrm`
            // is a cached copy of a box we work out from the anchor anyway.
            b"ext" if parent == b"oneCellAnchor" || parent == b"absoluteAnchor" => {
                self.size = (
                    emu_to_px(num_attr(element, b"cx").unwrap_or(0.0)),
                    emu_to_px(num_attr(element, b"cy").unwrap_or(0.0)),
                );
            }
            b"pos" if parent == b"absoluteAnchor" => {
                self.pos = (
                    emu_to_px(num_attr(element, b"x").unwrap_or(0.0)),
                    emu_to_px(num_attr(element, b"y").unwrap_or(0.0)),
                );
            }
            b"sp" | b"cxnSp" | b"pic" | b"graphicFrame" => {
                self.current = Some(Shape {
                    kind: match name {
                        b"pic" => "picture".to_owned(),
                        b"graphicFrame" => "frame".to_owned(),
                        _ => String::new(),
                    },
                    ..Shape::default()
                });
                self.fill = None;
                self.line = LineSpec::default();
                self.refs = StyleRefs::default();
                self.paras.clear();
                self.run = None;
                self.text_size = None;
                self.text_align = None;
                self.text_valign = None;
                self.text_wrap = None;
                self.text_insets = None;
                self.text_color = None;
            }
            b"xfrm" if parent == b"spPr" || parent == b"txXfrm" => {
                if let Some(shape) = self.current.as_mut() {
                    shape.rot = num_attr(element, b"rot").unwrap_or(0.0) / 60_000.0;
                    shape.flip_h = is_true(element, b"flipH");
                    shape.flip_v = is_true(element, b"flipV");
                }
            }
            b"prstGeom" => {
                if let Some(shape) = self.current.as_mut() {
                    if shape.kind.is_empty() {
                        shape.kind = attr(element, b"prst").unwrap_or_else(|| "rect".to_owned());
                    }
                }
            }
            // `fmla="val 18750"` is thousandths of a percent of the shape size.
            b"gd" if parent == b"avLst" => {
                let value = attr(element, b"fmla")
                    .and_then(|formula| formula.strip_prefix("val ").map(str::to_owned))
                    .and_then(|value| value.trim().parse::<f64>().ok());
                if let (Some(shape), Some(value)) = (self.current.as_mut(), value) {
                    shape.adjust.push(value);
                }
            }
            b"noFill" => {
                if parent == b"ln" {
                    self.line.none = true;
                } else if self.fill.is_none() {
                    self.fill = Some(None);
                }
            }
            // A pattern fill is reported as no fill rather than as a wrong flat
            // one. Gradients are collected properly just below.
            b"pattFill" if self.parent_is_sp_pr(&parent) && self.fill.is_none() => {
                self.fill = Some(None);
            }
            b"gradFill" if self.parent_is_sp_pr(&parent) => {
                self.gradient = Some((Vec::new(), 0.0));
            }
            b"gs" => {
                if let Some((stops, _)) = self.gradient.as_mut() {
                    stops.push(Stop {
                        pos: num_attr(element, b"pos").unwrap_or(0.0) / 100_000.0,
                        ..Stop::default()
                    });
                }
            }
            b"lin" => {
                if let Some((_, angle)) = self.gradient.as_mut() {
                    // 60000ths of a degree, clockwise, zero pointing right.
                    *angle = num_attr(element, b"ang").unwrap_or(0.0) / 60_000.0;
                }
            }
            b"ln" => {
                self.line.declared = true;
                if let Some(width) = num_attr(element, b"w") {
                    self.line.width = Some(emu_to_px(width));
                }
            }
            b"prstDash" if parent == b"ln" => self.line.dash = attr(element, b"val"),
            b"headEnd" if parent == b"ln" => self.line.head = read_arrow(element),
            b"tailEnd" if parent == b"ln" => self.line.tail = read_arrow(element),
            b"lnRef" | b"fillRef" | b"fontRef" | b"effectRef" => {
                self.ref_slot = Some(match name {
                    b"lnRef" => Target::LnRef,
                    b"fillRef" => Target::FillRef,
                    b"fontRef" => Target::FontRef,
                    _ => Target::EffectRef,
                });
                self.ref_index = num_attr(element, b"idx").unwrap_or(0.0) as u32;
            }
            b"srgbClr" => {
                self.pending = attr(element, b"val")
                    .and_then(|hex| parse_hex(&hex))
                    .map(|color| (Fill::solid(color, 1.0), self.target()));
            }
            b"schemeClr" => {
                self.pending = attr(element, b"val")
                    .and_then(|name| scheme_slot(&name))
                    .and_then(|index| self.theme.get(index as usize).cloned().flatten())
                    .map(|color| (Fill::solid(color, 1.0), self.target()));
            }
            b"sysClr" => {
                self.pending = attr(element, b"lastClr")
                    .and_then(|hex| parse_hex(&hex))
                    .map(|color| (Fill::solid(color, 1.0), self.target()));
            }
            b"shade" | b"tint" | b"lumMod" | b"lumOff" | b"alpha" => {
                if let Some(percent) = num_attr(element, b"val") {
                    self.modify(name, percent / 100_000.0);
                }
            }
            b"bodyPr" => {
                self.text_valign = Some(
                    match attr(element, b"anchor").as_deref() {
                        Some("ctr") => "middle",
                        Some("b") => "bottom",
                        _ => "top",
                    }
                    .to_owned(),
                );
                self.text_wrap = Some(attr(element, b"wrap").as_deref() != Some("none"));
                let inset = |name: &[u8], index: usize| {
                    emu_to_px(num_attr(element, name).unwrap_or(DEFAULT_INSETS_EMU[index]))
                };
                self.text_insets = Some([
                    inset(b"lIns", 0),
                    inset(b"tIns", 1),
                    inset(b"rIns", 2),
                    inset(b"bIns", 3),
                ]);
            }
            b"p" => self.paras.push(Vec::new()),
            b"pPr" => {
                if let Some(value) = attr(element, b"algn") {
                    self.text_align = Some(align_of(&value));
                }
            }
            b"r" => self.run = Some(Run::default()),
            b"rPr" | b"defRPr" => {
                if let Some(run) = self.run.as_mut() {
                    run.bold = is_true(element, b"b");
                    run.italic = is_true(element, b"i");
                }
                if self.text_size.is_none() {
                    self.text_size = num_attr(element, b"sz").map(|size| size / 100.0);
                }
            }
            b"blip" => {
                if let Some(shape) = self.current.as_mut() {
                    // `r:embed` names the relationship to the image itself.
                    shape.media = attr(element, b"embed");
                }
            }
            _ => {}
        }
    }

    /// Whether a fill element belongs to the shape rather than to its outline or
    /// one of the style refs.
    fn parent_is_sp_pr(&self, parent: &[u8]) -> bool {
        parent == b"spPr" || parent.is_empty()
    }

    fn end(&mut self, name: &[u8]) {
        match name {
            b"twoCellAnchor" | b"oneCellAnchor" | b"absoluteAnchor" => self.anchor = None,
            b"srgbClr" | b"schemeClr" | b"sysClr" => {
                if let Some((fill, target)) = self.pending.take() {
                    self.commit(target, fill);
                }
            }
            b"lnRef" | b"fillRef" | b"fontRef" | b"effectRef" => self.ref_slot = None,
            b"gradFill" => {
                // A gradient is a fill in its own right; one that ends with no
                // stops is a fill that was never given any, so it stays absent.
                self.fill = match self.gradient.take() {
                    Some((stops, angle)) if !stops.is_empty() => Some(Some(Fill {
                        color: stops[0].color.clone(),
                        alpha: stops[0].alpha,
                        stops,
                        angle,
                    })),
                    Some(_) => Some(None),
                    None => self.fill.take(),
                };
            }
            b"r" => {
                if let Some(run) = self.run.take() {
                    if let Some(para) = self.paras.last_mut() {
                        if !run.text.is_empty() {
                            para.push(run);
                        }
                    }
                }
            }
            b"sp" | b"cxnSp" | b"pic" | b"graphicFrame" => self.finish(),
            _ => {}
        }
    }

    /// Settles an element that closed on itself and so never reported an end.
    fn flush(&mut self, name: &[u8]) {
        match name {
            b"srgbClr" | b"schemeClr" | b"sysClr" => {
                if let Some((fill, target)) = self.pending.take() {
                    self.commit(target, fill);
                }
            }
            b"lnRef" | b"fillRef" | b"fontRef" | b"effectRef" => self.ref_slot = None,
            _ => {}
        }
    }

    fn finish(&mut self) {
        let Some(mut shape) = self.current.take() else {
            return;
        };
        let (from, to) = self.resolve_box();
        shape.from_col = from.0;
        shape.from_x = from.1;
        shape.from_row = from.2;
        shape.from_y = from.3;
        shape.to_col = to.0;
        shape.to_x = to.1;
        shape.to_row = to.2;
        shape.to_y = to.3;

        if shape.kind.is_empty() {
            shape.kind = "rect".to_owned();
        }

        shape.fill = match self.fill.take() {
            Some(explicit) => explicit,
            // `idx="0"` means "no fill" in the style's own vocabulary.
            None => self.refs.fill.as_ref().and_then(|(index, fill)| {
                (*index > 0).then(|| fill.clone())
            }),
        };
        shape.line = self.effective_line();

        // Empty paragraphs are kept — an empty one is a blank line, and a
        // callout that opens with two of them starts its text two lines down,
        // which is exactly where Excel puts it. Dropping them pulled the text
        // up onto the numbered markers those blank lines were making room for.
        // Only a frame with no text at all is not text.
        let paras = std::mem::take(&mut self.paras);
        if paras.iter().any(|para| !para.is_empty()) {
            shape.text = Some(Text {
                size_pt: self.text_size.unwrap_or(DEFAULT_TEXT_PT),
                color: self
                    .text_color
                    .take()
                    .or_else(|| self.refs.font.clone())
                    .unwrap_or_else(|| Fill::solid(BLACK.to_owned(), 1.0))
                    .color,
                align: self.text_align.clone().unwrap_or_else(|| "left".to_owned()),
                valign: self.text_valign.clone().unwrap_or_else(|| "top".to_owned()),
                wrap: self.text_wrap.unwrap_or(true),
                insets: self
                    .text_insets
                    .unwrap_or(DEFAULT_INSETS_EMU.map(emu_to_px)),
                paras,
            });
        }

        self.out.push(shape);
    }

    /// The shape's box as (col, x, row, y) for each corner. Two-cell anchors name
    /// both corners; the others carry a size, which has to be walked back into a
    /// second cell.
    fn resolve_box(&self) -> ((u32, f64, u32, f64), (u32, f64, u32, f64)) {
        let (start_col, start_col_off, start_row, start_row_off) = match self.anchor {
            Some(AnchorKind::Absolute) => {
                let (col, x) = self.grid.col_at(self.pos.0);
                let (row, y) = self.grid.row_at(self.pos.1);
                (col, x, row, y)
            }
            _ => (
                self.from.col,
                emu_to_px(self.from.col_off),
                self.from.row,
                emu_to_px(self.from.row_off),
            ),
        };

        if let Some(to) = self.to {
            return (
                (start_col, start_col_off, start_row, start_row_off),
                (
                    to.col,
                    emu_to_px(to.col_off),
                    to.row,
                    emu_to_px(to.row_off),
                ),
            );
        }

        let right = self.grid.col_start(start_col) + start_col_off + self.size.0;
        let bottom = self.grid.row_start(start_row) + start_row_off + self.size.1;
        let (to_col, to_x) = self.grid.col_at(right);
        let (to_row, to_y) = self.grid.row_at(bottom);
        (
            (start_col, start_col_off, start_row, start_row_off),
            (to_col, to_x, to_row, to_y),
        )
    }

    fn effective_line(&self) -> Option<Stroke> {
        let spec = &self.line;
        if spec.none {
            return None;
        }
        // An explicit colour wins; failing that the style ref's, and failing that
        // a declared-but-uncoloured line is black, as DrawingML defines it.
        let (fill, from_ref) = match spec.color.clone() {
            Some(fill) => (fill, false),
            None => match self.refs.ln.as_ref() {
                Some((_, fill)) => (fill.clone(), true),
                None if spec.declared => (Fill::solid(BLACK.to_owned(), 1.0), false),
                None => return None,
            },
        };
        let width = spec.width.unwrap_or_else(|| {
            let index = match (from_ref, self.refs.ln.as_ref()) {
                (true, Some((index, _))) => *index,
                _ => 1,
            };
            self.ln_widths
                .get(index.saturating_sub(1) as usize)
                .copied()
                .unwrap_or(DEFAULT_LINE_PX)
        });
        Some(Stroke {
            color: fill.color,
            alpha: fill.alpha,
            width,
            dash: spec.dash.clone(),
            head: spec.head.clone(),
            tail: spec.tail.clone(),
        })
    }

    /// Which slot a colour element fills, from where it sits in the tree.
    fn target(&self) -> Target {
        if let Some(slot) = self.ref_slot {
            return slot;
        }
        if self.stack.iter().any(|name| name == b"gs") {
            Target::GradientStop
        } else if self.stack.iter().any(|name| name == b"ln") {
            Target::Line
        } else if self.stack.iter().any(|name| name == b"rPr" || name == b"defRPr") {
            Target::Run
        } else {
            Target::Fill
        }
    }

    fn commit(&mut self, target: Target, fill: Fill) {
        match target {
            Target::Fill => {
                if self.fill.is_none() {
                    self.fill = Some(Some(fill));
                }
            }
            Target::Line => self.line.color = Some(fill),
            Target::Run => {
                if self.text_color.is_none() {
                    self.text_color = Some(fill);
                }
            }
            Target::LnRef => self.refs.ln = Some((self.ref_index, fill)),
            Target::FillRef => self.refs.fill = Some((self.ref_index, fill)),
            Target::GradientStop => {
                if let Some((stops, _)) = self.gradient.as_mut() {
                    if let Some(stop) = stops.last_mut() {
                        *stop = Stop {
                            pos: stop.pos,
                            color: fill.color,
                            alpha: fill.alpha,
                        };
                    }
                }
            }
            Target::FontRef => {
                if self.refs.font.is_none() {
                    self.refs.font = Some(fill);
                }
            }
            // A shadow's tint is not drawn, so it goes nowhere.
            Target::EffectRef => {}
        }
    }

    /// Applies `<a:shade>`, `<a:tint>`, `<a:lumMod>`, `<a:lumOff>` and
    /// `<a:alpha>`, which a colour element carries as children.
    fn modify(&mut self, name: &[u8], value: f64) {
        let Some((fill, _)) = self.pending.as_mut() else {
            return;
        };
        fill.color = match name {
            b"alpha" => {
                fill.alpha *= value;
                return;
            }
            // `shade` multiplies the channels straight, the way Excel's does.
            b"shade" => scaled(&fill.color, value),
            b"tint" => tint(fill.color.clone(), Some(value.to_string())),
            b"lumMod" => with_luminance(&fill.color, |l| l * value),
            b"lumOff" => with_luminance(&fill.color, |l| l + value),
            // Excel's "subtle effect" gradients lean on `satMod` as heavily as
            // they do on `tint`, and leaving it out washes the colour out.
            b"satMod" => with_saturation(&fill.color, |s| s * value),
            b"satOff" => with_saturation(&fill.color, |s| s + value),
            _ => return,
        };
    }
}

/// Reads a `<a:headEnd>` or `<a:tailEnd>`. `type="none"` is the file saying
/// there is no arrowhead, which is not the same as saying nothing at all.
fn read_arrow(element: &BytesStart) -> Option<Arrow> {
    let kind = attr(element, b"type").filter(|kind| kind != "none")?;
    Some(Arrow {
        kind,
        width: attr(element, b"w").unwrap_or_else(|| "med".to_owned()),
        length: attr(element, b"len").unwrap_or_else(|| "med".to_owned()),
    })
}

fn is_true(element: &BytesStart, name: &[u8]) -> bool {    attr(element, name).is_some_and(|value| value == "1" || value == "true")
}

fn emu_to_px(emu: f64) -> f64 {
    emu / EMU_PER_PX
}

fn align_of(value: &str) -> String {
    match value {
        "ctr" => "center",
        "r" | "just" => "right",
        _ => "left",
    }
    .to_owned()
}

fn parse_hex(value: &str) -> Option<String> {
    let digits = value.get(..6)?;
    digits
        .chars()
        .all(|ch| ch.is_ascii_hexdigit())
        .then(|| format!("#{}", digits.to_uppercase()))
}

/// DrawingML names its theme slots, and `tx1`/`bg1` are aliases for `dk1`/`lt1`.
fn scheme_slot(name: &str) -> Option<u32> {
    Some(match name {
        "lt1" | "bg1" => 0,
        "dk1" | "tx1" => 1,
        "lt2" | "bg2" => 2,
        "dk2" | "tx2" => 3,
        "accent1" => 4,
        "accent2" => 5,
        "accent3" => 6,
        "accent4" => 7,
        "accent5" => 8,
        "accent6" => 9,
        "hlink" => 10,
        "folHlink" => 11,
        _ => return None,
    })
}

fn channel(hex: &str, index: usize) -> u8 {
    u8::from_str_radix(&hex[1 + index * 2..3 + index * 2], 16).unwrap_or(0)
}

/// Multiplies every channel, which is what `<a:shade>` does.
fn scaled(hex: &str, factor: f64) -> String {
    let byte = |index: usize| {
        (f64::from(channel(hex, index)) * factor)
            .round()
            .clamp(0.0, 255.0) as u8
    };
    format!("#{:02X}{:02X}{:02X}", byte(0), byte(1), byte(2))
}

fn with_luminance(hex: &str, change: impl FnOnce(f64) -> f64) -> String {
    let (h, s, l) = to_hsl(hex);
    from_hsl(h, s, change(l).clamp(0.0, 1.0))
}

fn with_saturation(hex: &str, change: impl FnOnce(f64) -> f64) -> String {
    let (h, s, l) = to_hsl(hex);
    from_hsl(h, change(s).clamp(0.0, 1.0), l)
}

fn to_hsl(hex: &str) -> (f64, f64, f64) {
    let (r, g, b) = (
        f64::from(channel(hex, 0)) / 255.0,
        f64::from(channel(hex, 1)) / 255.0,
        f64::from(channel(hex, 2)) / 255.0,
    );
    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let l = (max + min) / 2.0;
    let delta = max - min;
    let s = if delta == 0.0 {
        0.0
    } else {
        delta / (1.0 - (2.0 * l - 1.0).abs())
    };
    let h = if delta == 0.0 {
        0.0
    } else if max == r {
        60.0 * (((g - b) / delta) % 6.0)
    } else if max == g {
        60.0 * ((b - r) / delta + 2.0)
    } else {
        60.0 * ((r - g) / delta + 4.0)
    };
    (((h % 360.0) + 360.0) % 360.0, s, l)
}

fn from_hsl(h: f64, s: f64, l: f64) -> String {
    let c = (1.0 - (2.0 * l - 1.0).abs()) * s;
    let x = c * (1.0 - (((h / 60.0) % 2.0) - 1.0).abs());
    let m = l - c / 2.0;
    let (r, g, b) = match h as u32 / 60 {
        0 => (c, x, 0.0),
        1 => (x, c, 0.0),
        2 => (0.0, c, x),
        3 => (0.0, x, c),
        4 => (x, 0.0, c),
        _ => (c, 0.0, x),
    };
    let byte = |value: f64| ((value + m) * 255.0).round().clamp(0.0, 255.0) as u8;
    format!("#{:02X}{:02X}{:02X}", byte(r), byte(g), byte(b))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn solid(color: &str) -> Fill {
        Fill::solid(color.to_owned(), 1.0)
    }

    fn grid() -> Grid {
        let widths: HashMap<u32, f64> = [(0, 100.0), (1, 200.0)].into_iter().collect();
        let heights: HashMap<u32, f64> = [(0, 20.0)].into_iter().collect();
        Grid::new(2, &widths, 50.0, 10, &heights, 15.0)
    }

    /// The shape from the design documents: a red pentagon holding a step number.
    #[test]
    fn reads_a_plain_shape_with_an_explicit_fill() {
        let xml = r#"<xdr:wsDr xmlns:xdr="x" xmlns:a="a" xmlns:r="r"><xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>47625</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>95250</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>19050</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>28575</xdr:rowOff></xdr:to><xdr:sp><xdr:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="pentagon"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln><a:noFill/></a:ln></xdr:spPr><xdr:style><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef></xdr:style><xdr:txBody><a:bodyPr anchor="ctr"/><a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1100" b="1"/><a:t>1</a:t></a:r></a:p></xdr:txBody></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>"#;
        let shapes = parse(xml, &grid(), &[], &[]);
        assert_eq!(shapes.len(), 1);
        let shape = &shapes[0];
        assert_eq!(shape.kind, "pentagon");
        // 47625 EMU is 5px into column 1; 95250 is 10px into row 2.
        assert_eq!((shape.from_col, shape.from_x), (1, 5.0));
        assert_eq!((shape.from_row, shape.from_y), (2, 10.0));
        assert_eq!((shape.to_col, shape.to_x), (2, 2.0));
        assert_eq!((shape.to_row, shape.to_y), (3, 3.0));
        assert_eq!(
            shape.fill,
            Some(solid("#FF0000"))
        );
        // An explicit `noFill` line beats the style's `lnRef`.
        assert!(shape.line.is_none());
        let text = shape.text.as_ref().expect("text");
        assert_eq!(text.paras[0][0].text, "1");
        assert!(text.paras[0][0].bold);
        assert_eq!(text.align, "center");
        assert_eq!(text.valign, "middle");
        assert_eq!(text.size_pt, 11.0);
    }

    /// A callout declares no colours, so its whole appearance comes from the
    /// theme: `fillRef` gives the fill, `lnRef` the outline.
    #[test]
    fn falls_back_to_the_style_refs_and_the_theme() {
        let theme = vec![
            Some("#FFFFFF".to_owned()),
            Some("#000000".to_owned()),
            None,
            None,
            Some("#4472C4".to_owned()),
            Some("#ED7D31".to_owned()),
        ];
        // `accent1` at 50% shade is half its channels.
        let ln_widths = [0.67, 1.33, 2.0];
        // The `effectRef` names a different accent on purpose: its colour tints a
        // shadow, and reading it as the fill would paint the shape the wrong one.
        let xml = r#"<xdr:wsDr xmlns:xdr="x" xmlns:a="a"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:spPr><a:prstGeom prst="borderCallout2"><a:avLst><a:gd name="adj1" fmla="val 18750"/></a:avLst></a:prstGeom></xdr:spPr><xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent2"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></xdr:style><xdr:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="800"/><a:t>hi</a:t></a:r></a:p></xdr:txBody></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>"#;
        let shapes = parse(xml, &grid(), &theme, &ln_widths);
        let shape = &shapes[0];
        assert_eq!(shape.kind, "borderCallout2");
        // The file's own units, hundred-thousandths: the preset formulas divide.
        assert_eq!(shape.adjust, vec![18750.0]);
        assert_eq!(
            shape.fill,
            Some(solid("#4472C4"))
        );
        let line = shape.line.as_ref().expect("outline");
        assert_eq!(line.color, "#223962");
        // `lnRef idx="2"` takes the second entry of the theme's line widths.
        assert_eq!(line.width, 1.33);
        // The text colour comes from `fontRef`, which points at `lt1`.
        assert_eq!(shape.text.as_ref().unwrap().color, "#FFFFFF".to_owned());
    }

    /// A shape with no `to` corner is sized by `<xdr:ext>` instead.
    #[test]
    fn anchors_a_sized_shape_to_a_second_cell() {
        // 9525 EMU is one pixel: 20px into column 1 of a 100/200px pair lands at
        // 120px, and 80px wide ends 200px into column 1.
        let xml = r#"<xdr:wsDr xmlns:xdr="x" xmlns:a="a"><xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>190500</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="762000" cy="95250"/><xdr:pic><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill></xdr:pic></xdr:oneCellAnchor></xdr:wsDr>"#;
        let shapes = parse(xml, &grid(), &[], &[]);
        let shape = &shapes[0];
        assert_eq!(shape.kind, "picture");
        assert_eq!(shape.media.as_deref(), Some("rId1"));
        assert_eq!((shape.from_col, shape.from_x), (0, 20.0));
        // 20px + 80px = 100px, exactly column 1's left edge.
        assert_eq!((shape.to_col, shape.to_x), (1, 0.0));
        assert_eq!((shape.to_row, shape.to_y), (0, 10.0));
    }

    #[test]
    fn a_line_without_a_colour_takes_the_style_refs_black() {        let theme = vec![Some("#FFFFFF".to_owned()), Some("#0D0D0D".to_owned())];
        let xml = r#"<xdr:wsDr xmlns:xdr="x" xmlns:a="a"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:spPr><a:prstGeom prst="leftBrace"><a:avLst/></a:prstGeom><a:ln w="19050"/></xdr:spPr><xdr:style><a:lnRef idx="1"><a:schemeClr val="dk1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="dk1"/></a:fillRef></xdr:style></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>"#;
        let shapes = parse(xml, &grid(), &theme, &[0.5]);
        let shape = &shapes[0];
        assert_eq!(shape.kind, "leftBrace");
        // `fillRef idx="0"` is the style's way of saying "no fill".
        assert!(shape.fill.is_none());
        let line = shape.line.as_ref().expect("outline");
        assert_eq!(line.color, "#0D0D0D");
        assert_eq!(line.width, 2.0);
    }

    /// Excel writes a shape that only *looks* unrotated as a half turn plus both
    /// flips. Reading either half alone turns the shape — and its label — the
    /// wrong way round, which is what made every callout in the design documents
    /// come out upside down.
    #[test]
    fn reads_a_half_turn_and_its_cancelling_flips() {
        let xml = r#"<xdr:wsDr xmlns:xdr="x" xmlns:a="a"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:spPr><a:xfrm rot="10800000" flipH="1" flipV="1"><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm><a:prstGeom prst="borderCallout2"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>"#;
        let shapes = parse(xml, &grid(), &[], &[]);
        let shape = &shapes[0];
        assert_eq!(shape.rot, 180.0);
        assert!(shape.flip_h, "flipH");
        assert!(shape.flip_v, "flipV");
    }

    /// The callouts in the design documents are filled with a gradient built from
    /// three tints of the theme accent, not a flat colour. Reading `gradFill` as
    /// "no fill" left every one of them transparent.
    #[test]
    fn reads_a_gradient_fill_from_its_stops() {
        let theme = vec![
            Some("#FFFFFF".to_owned()),
            Some("#000000".to_owned()),
            None,
            None,
            Some("#4472C4".to_owned()),
        ];
        let xml = r#"<xdr:wsDr xmlns:xdr="x" xmlns:a="a"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="accent1"><a:tint val="50000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="accent1"/></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill></xdr:spPr></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>"#;
        let shapes = parse(xml, &grid(), &theme, &[]);
        let fill = shapes[0].fill.as_ref().expect("a gradient fill");
        assert_eq!(fill.stops.len(), 2);
        assert_eq!(fill.stops[0].pos, 0.0);
        assert_eq!(fill.stops[1].pos, 1.0);
        // The second stop is the accent untouched; the first is half a turn
        // towards white, so the two must differ.
        assert_eq!(fill.stops[1].color, "#4472C4");
        assert_ne!(fill.stops[0].color, "#4472C4");
        assert_eq!(fill.color, fill.stops[0].color);
        // 16200000 in 60000ths of a degree is 270, which points straight up.
        assert_eq!(fill.angle, 270.0);
    }

    /// Prints what every sheet's drawing part parsed into.
    ///
    /// The workbook comes from `SHEETLENS_XLSX` rather than a literal: an earlier
    /// dump like this hardcoded a path, and a session went into chasing a parser
    /// bug that was really the dump reading a different file.
    #[test]
    #[ignore = "reads a workbook from the developer's machine"]
    fn dump_drawings() {
        let Ok(path) = std::env::var("SHEETLENS_XLSX") else {
            return;
        };
        let book = crate::excel::load(&path).expect("open");
        let mut cache = crate::layout::Cache::default();
        for sheet in &book.info.sheets {
            let Ok(layout) = crate::layout::read_cached(&mut cache, &path, &sheet.name) else {
                continue;
            };
            if layout.drawings.is_empty() {
                continue;
            }
            println!("--- {} : {} shapes", sheet.name, layout.drawings.len());
            for shape in layout.drawings.iter().take(40) {
                let text = shape
                    .text
                    .as_ref()
                    .map(|text| {
                        format!(
                            "{}pt {} c{}r{} {:?}",
                            text.size_pt,
                            text.color,
                            text.align,
                            text.valign,
                            text.paras
                                .iter()
                                .map(|para| {
                                    para.iter().map(|run| run.text.as_str()).collect::<String>()
                                })
                                .collect::<Vec<_>>()
                                .join(" / ")
                        )
                    })
                    .unwrap_or_default();
                println!(
                    "  {:<15} c{}.{:>4.0} r{}.{:>4.0} → c{}.{:>4.0} r{}.{:>4.0}  rot={:.0}{}{} fill={:?} line={:?} adj={:?} media={:?} text=[{}]",
                    shape.kind,
                    shape.from_col,
                    shape.from_x,
                    shape.from_row,
                    shape.from_y,
                    shape.to_col,
                    shape.to_x,
                    shape.to_row,
                    shape.to_y,
                    shape.rot,
                    if shape.flip_h { " flipH" } else { "" },
                    if shape.flip_v { " flipV" } else { "" },
                    shape
                        .fill
                        .as_ref()
                        .map(|fill| format!("{} a{}", fill.color, fill.alpha)),
                    shape.line.as_ref().map(|line| format!(
                        "{} {}px {:?} head={:?} tail={:?}",
                        line.color,
                        line.width,
                        line.dash,
                        line.head.as_ref().map(|a| format!("{}/{}/{}", a.kind, a.width, a.length)),
                        line.tail.as_ref().map(|a| format!("{}/{}/{}", a.kind, a.width, a.length)),
                    )),
                    shape.adjust,
                    shape.media.as_deref().map(|media| media.rsplit('/').next().unwrap_or(media)),
                    text,
                );
            }
        }
    }
}
