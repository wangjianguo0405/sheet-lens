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
    /// Pictures only: the part of the image cut away, as fractions of it —
    /// `[left, top, right, bottom]`. Excel writes a negative side to extend the
    /// image past the box instead.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub crop: Option<[f64; 4]>,
    /// Set when the shape sits inside a group: its place as a fraction of the
    /// group's anchor box, which the renderer resolves with its own widths.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub group: Option<GroupFrame>,
    /// Set when the file carries its own geometry instead of naming a preset:
    /// a guide list and one or more paths, in the path's own coordinates.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub custom: Option<CustomGeometry>,
    /// Preset adjustments in the order the preset declares them (`adj1`…), in
    /// the file's own units — hundred-thousandths. The renderer feeds them
    /// straight into the preset's formulas, which divide by 100000 themselves.
    pub adjust: Vec<f64>,
}

#[derive(Serialize, Clone, Debug, Default)]
#[serde(rename_all = "camelCase")]
pub struct CustomGeometry {
    /// `gdLst`: named values the points are built from, in evaluation order.
    pub guides: Vec<CustomGuide>,
    pub paths: Vec<CustomPath>,
}

#[derive(Serialize, Clone, Debug, Default)]
#[serde(rename_all = "camelCase")]
pub struct CustomGuide {
    pub name: String,
    pub formula: String,
}

#[derive(Serialize, Clone, Debug, Default)]
#[serde(rename_all = "camelCase")]
pub struct CustomPath {
    /// The path's own coordinate space, scaled onto the shape's box.
    pub w: f64,
    pub h: f64,
    /// `norm`, `none`, `lighten`, `darken`…; absent means `norm`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill: Option<String>,
    /// Absent means stroked.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stroke: Option<bool>,
    pub commands: Vec<CustomCommand>,
}

#[derive(Serialize, Clone, Debug, Default)]
#[serde(rename_all = "camelCase")]
pub struct CustomCommand {
    /// `moveTo`, `lnTo`, `cubicBezTo`, `quadBezTo`, `arcTo` or `close`.
    pub kind: String,
    /// Coordinates as the file wrote them: a number or a guide's name.
    pub points: Vec<[String; 2]>,
    /// `arcTo`'s `wR hR stAng swAng`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub arc: Option<[String; 4]>,
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
    /// How the line joins its segments: `miter` (the default), `round` or
    /// `bevel`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub join: Option<String>,
    /// A miter join's limit, as a multiple of the line width. Excel draws with
    /// 8 when the file does not say, where SVG's own default is 4.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub miter_limit: Option<f64>,
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
    /// `left`, `center` or `right`. The shape's default, which the container
    /// uses; a paragraph that names its own is in `aligns`.
    pub align: String,
    /// One alignment per paragraph, when they are not all the default.
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub aligns: Vec<String>,
    /// `top`, `middle` or `bottom`.
    pub valign: String,
    /// Whether the text wraps at the shape's edge. Excel's `wrap="none"` lets it
    /// spill out, the way an unstyled cell does.
    pub wrap: bool,
    /// `lIns`, `tIns`, `rIns`, `bIns` in pixels.
    pub insets: [f64; 4],
    /// The face the shape's own `defRPr` names, for runs that name none.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub family: Option<String>,
    /// The face's own classification — `panose`, `charset` and `pitchFamily` —
    /// which says what to fall back to when the named face is not installed.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub panose: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub charset: Option<i32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pitch: Option<u32>,
    /// The east asian and complex-script faces beside the latin one. A browser
    /// picks a face per character, so listing them in this order is how Excel's
    /// own latin/ea split is reproduced.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub ea: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cs: Option<String>,
    /// One flag per paragraph: whether its base direction is right-to-left,
    /// which changes the bidi order of mixed text. Empty when all are left.
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub rtl: Vec<bool>,
}

#[derive(Serialize, Clone, Debug, Default)]
#[serde(rename_all = "camelCase")]
pub struct Run {
    pub text: String,
    pub bold: bool,
    pub italic: bool,
    /// `<a:latin typeface>`: the face this run is written in.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub family: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub panose: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub charset: Option<i32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pitch: Option<u32>,
    /// The east asian and complex-script faces beside the latin one. A browser
    /// picks a face per character, so listing them in this order is how Excel's
    /// own latin/ea split is reproduced.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub ea: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cs: Option<String>,
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
pub fn parse(
    xml: &str,
    grid: &Grid,
    theme: &[Option<String>],
    ln_widths: &[f64],
    fonts: &HashMap<String, String>,
) -> Vec<Shape> {
    Parser::new(grid, theme, ln_widths, fonts).run(xml)
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
    join: Option<String>,
    miter: Option<f64>,
}

/// The `xdr:style` refs, which colour a shape that declares no colours itself.
#[derive(Default)]
struct StyleRefs {
    ln: Option<(u32, Fill)>,
    fill: Option<(u32, Fill)>,
    font: Option<Fill>,
}

/// Where a group's children live. A group carries a child coordinate space
/// (`chOff`/`chExt`) that maps onto the group's own box, and Excel places that
/// box from the anchor — its `xfrm` is a cached copy that a resized sheet can
/// leave stale. The frame is the group's rect as a fraction of the anchor, which
/// is what lets the renderer resolve children against its own column widths
/// instead of the parser's approximate ones. Nested groups compose.
struct GroupXform {
    ch_off: (f64, f64),
    ch_ext: (f64, f64),
    /// `[left, top, width, height]` as fractions of the anchor's box.
    frame: [f64; 4],
    /// The innermost enclosing anchor, whose box the fractions are against.
    anchor: ((u32, f64, u32, f64), (u32, f64, u32, f64)),
    /// The group's own fill, which a child with `<a:grpFill/>` draws with.
    fill: Option<Fill>,
}

impl GroupXform {
    /// A child group's rect as fractions of the same anchor box this group's
    /// frame is measured in.
    fn child_frame(&self, off: (f64, f64), ext: (f64, f64)) -> [f64; 4] {
        let sx = if self.ch_ext.0 != 0.0 { 1.0 / self.ch_ext.0 } else { 0.0 };
        let sy = if self.ch_ext.1 != 0.0 { 1.0 / self.ch_ext.1 } else { 0.0 };
        [
            self.frame[0] + (off.0 - self.ch_off.0) * sx * self.frame[2],
            self.frame[1] + (off.1 - self.ch_off.1) * sy * self.frame[3],
            ext.0 * sx * self.frame[2],
            ext.1 * sy * self.frame[3],
        ]
    }
}

/// A shape's place inside its group, as a fraction of the group's anchor box.
#[derive(Serialize, Clone, Debug, Default)]
#[serde(rename_all = "camelCase")]
pub struct GroupFrame {
    pub from_col: u32,
    pub from_x: f64,
    pub from_row: u32,
    pub from_y: f64,
    pub to_col: u32,
    pub to_x: f64,
    pub to_row: u32,
    pub to_y: f64,
    pub fx: f64,
    pub fy: f64,
    pub fw: f64,
    pub fh: f64,
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
    /// The custom geometry being read, until its `custGeom` closes.
    custom: Option<CustomGeometry>,
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
    /// One base-direction flag per paragraph, parallel to `paras`.
    para_rtl: Vec<bool>,
    /// The direction `lstStyle` gives a paragraph that names none.
    list_style_rtl: Option<bool>,
    run: Option<Run>,
    text_size: Option<f64>,
    /// One alignment per paragraph, parallel to `paras`.
    para_align: Vec<String>,
    /// The alignment `lstStyle` gives a paragraph that names none.
    list_style_align: Option<String>,
    text_valign: Option<String>,
    text_wrap: Option<bool>,
    text_insets: Option<[f64; 4]>,
    text_color: Option<Fill>,
    /// The paragraph's own `b`/`i`, which a run that does not name them
    /// inherits. A shape's `lstStyle` is only the outermost default.
    text_bold: Option<bool>,
    text_italic: Option<bool>,
    /// The face the shape's own `defRPr` names, which its runs inherit.
    text_family: Option<String>,
    /// That face's classification, for runs that name none of their own.
    text_panose: Option<String>,
    text_charset: Option<i32>,
    text_pitch: Option<u32>,
    /// The east asian and complex-script faces the shape's own run properties
    /// name, which its runs inherit.
    text_ea: Option<String>,
    text_cs: Option<String>,
    /// The theme's own font scheme, which `+mn-ea` and its like index into.
    theme_fonts: &'a HashMap<String, String>,
    /// The group transforms now open, innermost last. A shape inside one is
    /// placed by its own `xfrm` — the anchor belongs to the group.
    groups: Vec<GroupXform>,
    /// A group's `xfrm`, in EMU, collected between its start and its end.
    group_off: (f64, f64),
    group_ext: (f64, f64),
    group_ch_off: (f64, f64),
    group_ch_ext: (f64, f64),
    /// The shape's own `xfrm` offset and size, in EMU. Unused unless the shape
    /// sits in a group; otherwise its anchor is what places it.
    sp_off: (f64, f64),
    sp_ext: (f64, f64),
    /// The fill the group being read declares (`grpSpPr`), for children that
    /// say `<a:grpFill/>`.
    group_fill: ExplicitFill,
    /// Whether the shape being read asked for its group's fill.
    shape_group_fill: bool,
}

impl<'a> Parser<'a> {
    fn new(
        grid: &'a Grid,
        theme: &'a [Option<String>],
        ln_widths: &'a [f64],
        fonts: &'a HashMap<String, String>,
    ) -> Self {
        Self {
            grid,
            theme,
            ln_widths,
            theme_fonts: fonts,
            out: Vec::new(),
            stack: Vec::new(),
            anchor: None,
            from: Point::default(),
            to: None,
            size: (0.0, 0.0),
            pos: (0.0, 0.0),
            current: None,
            custom: None,
            fill: None,
            line: LineSpec::default(),
            refs: StyleRefs::default(),
            pending: None,
            ref_slot: None,
            ref_index: 0,
            gradient: None,
            paras: Vec::new(),
            para_rtl: Vec::new(),
            list_style_rtl: None,
            run: None,
            text_size: None,
            para_align: Vec::new(),
            list_style_align: None,
            text_valign: None,
            text_wrap: None,
            text_insets: None,
            text_color: None,
            text_bold: None,
            text_italic: None,
            text_family: None,
            text_panose: None,
            text_charset: None,
            text_pitch: None,
            text_ea: None,
            text_cs: None,
            groups: Vec::new(),
            group_off: (0.0, 0.0),
            group_ext: (0.0, 0.0),
            group_ch_off: (0.0, 0.0),
            group_ch_ext: (0.0, 0.0),
            sp_off: (0.0, 0.0),
            sp_ext: (0.0, 0.0),
            group_fill: None,
            shape_group_fill: false,
        }
    }

    /// Opens a command on the custom path being read.
    fn custom_command(&mut self, kind: &[u8]) -> Option<&mut CustomCommand> {
        let path = self.custom.as_mut()?.paths.last_mut()?;
        path.commands.push(CustomCommand {
            kind: String::from_utf8_lossy(kind).into_owned(),
            points: Vec::new(),
            arc: None,
        });
        path.commands.last_mut()
    }

    /// The local name of the element enclosing the one being read.
    fn parent(&self) -> &[u8] {
        self.stack.last().map(Vec::as_slice).unwrap_or(b"")
    }

    /// Whether the `xfrm` being read is a group's rather than a shape's.
    fn in_group_xfrm(&self) -> bool {
        self.stack.iter().rev().nth(1).map(Vec::as_slice) == Some(b"grpSpPr")
    }

    /// Whether the `xfrm` being read belongs to the shape's own geometry, as
    /// opposed to a group's or a text body's.
    fn in_shape_xfrm(&self) -> bool {
        self.stack.iter().rev().nth(1).map(Vec::as_slice) == Some(b"spPr")
    }

    /// Whether the element being read sits in the shape's `lstStyle`, whose
    /// defaults are overridden by a paragraph's or a run's own formatting.
    fn in_list_style(&self) -> bool {
        self.stack.iter().any(|name| name == b"lstStyle")
    }

    /// Whether the element being read is a group's own properties — its fill,
    /// as opposed to a child's.
    fn in_group_props(&self) -> bool {
        self.stack.iter().any(|name| name == b"grpSpPr")
    }

    /// The face a `typeface` names, which may be a slot in the theme's font
    /// scheme rather than a font: `+mn-lt` is the minor scheme's latin face and
    /// `+mj-ea` the major scheme's east asian one. A slot the theme leaves
    /// empty has no face, and the run then draws in what it inherits.
    fn font_name(&self, typeface: Option<String>) -> Option<String> {
        let name = typeface?;
        let Some(slot) = name.strip_prefix('+') else {
            return Some(name);
        };
        self.theme_fonts.get(slot).cloned()
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
                // An entity is its own event rather than part of the text beside
                // it, so `&lt;` in a shape's wording is stitched back in here.
                Ok(Event::GeneralRef(reference)) => {
                    if let Some(ch) = crate::layout::entity_text(&reference) {
                        self.text(&ch.to_string());
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
                self.groups.clear();
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
                self.custom = None;
                self.fill = None;
                self.line = LineSpec::default();
                self.refs = StyleRefs::default();
                self.paras.clear();
                self.para_rtl.clear();
                self.list_style_rtl = None;
                self.run = None;
                self.text_size = None;
                self.para_align.clear();
                self.list_style_align = None;
                self.text_valign = None;
                self.text_wrap = None;
                self.text_insets = None;
                self.text_color = None;
                self.text_bold = None;
                self.text_italic = None;
                self.text_family = None;
                self.text_panose = None;
                self.text_charset = None;
                self.text_pitch = None;
                self.text_ea = None;
                self.text_cs = None;
                self.sp_off = (0.0, 0.0);
                self.sp_ext = (0.0, 0.0);
                self.shape_group_fill = false;
            }
            b"xfrm" if parent == b"spPr" || parent == b"txXfrm" => {
                if let Some(shape) = self.current.as_mut() {
                    shape.rot = num_attr(element, b"rot").unwrap_or(0.0) / 60_000.0;
                    shape.flip_h = is_true(element, b"flipH");
                    shape.flip_v = is_true(element, b"flipV");
                }
            }
            // A group's `xfrm` maps its children's coordinate space onto its own.
            b"grpSp" => self.group_fill = None,
            b"xfrm" if parent == b"grpSpPr" => {
                self.group_off = (0.0, 0.0);
                self.group_ext = (0.0, 0.0);
                self.group_ch_off = (0.0, 0.0);
                self.group_ch_ext = (0.0, 0.0);
            }
            b"off" if parent == b"xfrm" => {
                let at = (
                    num_attr(element, b"x").unwrap_or(0.0),
                    num_attr(element, b"y").unwrap_or(0.0),
                );
                if self.in_group_xfrm() {
                    self.group_off = at;
                } else if self.in_shape_xfrm() {
                    self.sp_off = at;
                }
            }
            b"ext" if parent == b"xfrm" => {
                let size = (
                    num_attr(element, b"cx").unwrap_or(0.0),
                    num_attr(element, b"cy").unwrap_or(0.0),
                );
                if self.in_group_xfrm() {
                    self.group_ext = size;
                } else if self.in_shape_xfrm() {
                    self.sp_ext = size;
                }
            }
            b"chOff" => {
                self.group_ch_off = (
                    num_attr(element, b"x").unwrap_or(0.0),
                    num_attr(element, b"y").unwrap_or(0.0),
                );
            }
            b"chExt" => {
                self.group_ch_ext = (
                    num_attr(element, b"cx").unwrap_or(0.0),
                    num_attr(element, b"cy").unwrap_or(0.0),
                );
            }
            b"prstGeom" => {
                if let Some(shape) = self.current.as_mut() {
                    if shape.kind.is_empty() {
                        shape.kind = attr(element, b"prst").unwrap_or_else(|| "rect".to_owned());
                    }
                }
            }
            b"custGeom" => self.custom = Some(CustomGeometry::default()),
            // The guide list belongs to a custom geometry; `avLst` has `gd`s of
            // its own, which are read as preset adjustments below.
            b"gd" if parent == b"gdLst" => {
                if let Some(custom) = self.custom.as_mut() {
                    custom.guides.push(CustomGuide {
                        name: attr(element, b"name").unwrap_or_default(),
                        formula: attr(element, b"fmla").unwrap_or_default(),
                    });
                }
            }
            b"path" if parent == b"pathLst" => {
                if let Some(custom) = self.custom.as_mut() {
                    custom.paths.push(CustomPath {
                        w: num_attr(element, b"w").unwrap_or(0.0),
                        h: num_attr(element, b"h").unwrap_or(0.0),
                        fill: attr(element, b"fill"),
                        stroke: attr(element, b"stroke")
                            .map(|value| value == "1" || value == "true"),
                        commands: Vec::new(),
                    });
                }
            }
            b"moveTo" | b"lnTo" | b"cubicBezTo" | b"quadBezTo" | b"close"
                if parent == b"path" =>
            {
                self.custom_command(name);
            }
            b"arcTo" if parent == b"path" => {
                let arc = [
                    attr(element, b"wR").unwrap_or_default(),
                    attr(element, b"hR").unwrap_or_default(),
                    attr(element, b"stAng").unwrap_or_default(),
                    attr(element, b"swAng").unwrap_or_default(),
                ];
                if let Some(command) = self.custom_command(name) {
                    command.arc = Some(arc);
                }
            }
            b"pt" if matches!(
                parent.as_slice(),
                b"moveTo" | b"lnTo" | b"cubicBezTo" | b"quadBezTo"
            ) =>
            {
                let point = [
                    attr(element, b"x").unwrap_or_else(|| "0".to_owned()),
                    attr(element, b"y").unwrap_or_else(|| "0".to_owned()),
                ];
                if let Some(command) = self
                    .custom
                    .as_mut()
                    .and_then(|custom| custom.paths.last_mut())
                    .and_then(|path| path.commands.last_mut())
                {
                    command.points.push(point);
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
                } else if self.in_group_props() {
                    self.group_fill = Some(None);
                } else if self.fill.is_none() {
                    self.fill = Some(None);
                }
            }
            // `<a:grpFill/>` is the child taking its group's own fill.
            b"grpFill" => self.shape_group_fill = true,
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
            // Join styles. `lim` is a percentage, so 800000 is eight times the
            // line width; a join the file does not describe is mitered.
            b"miter" if parent == b"ln" => {
                self.line.join = Some("miter".to_owned());
                if let Some(limit) = num_attr(element, b"lim") {
                    self.line.miter = Some(limit / 100_000.0);
                }
            }
            b"round" if parent == b"ln" => self.line.join = Some("round".to_owned()),
            b"bevel" if parent == b"ln" => self.line.join = Some("bevel".to_owned()),
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
            b"p" => {
                self.paras.push(Vec::new());
                self.para_rtl.push(self.list_style_rtl.unwrap_or(false));
                self.para_align.push(
                    self.list_style_align
                        .clone()
                        .unwrap_or_else(|| "left".to_owned()),
                );
            }
            b"pPr" => {
                if let Some(align) = attr(element, b"algn") {
                    if let Some(slot) = self.para_align.last_mut() {
                        *slot = align_of(&align);
                    }
                }
                if let Some(rtl) = bool_attr(element, b"rtl") {
                    if let Some(flag) = self.para_rtl.last_mut() {
                        *flag = rtl;
                    }
                }
            }
            // The level properties in `lstStyle` are the defaults a paragraph
            // that names no direction or alignment of its own takes.
            b"lvl1pPr" if self.in_list_style() => {
                if let Some(rtl) = bool_attr(element, b"rtl") {
                    self.list_style_rtl = Some(rtl);
                }
                if let Some(align) = attr(element, b"algn") {
                    self.list_style_align = Some(align_of(&align));
                }
            }
            b"r" => {
                self.run = Some(Run {
                    bold: self.text_bold.unwrap_or(false),
                    italic: self.text_italic.unwrap_or(false),
                    ..Run::default()
                })
            }
            // An explicit line break inside a paragraph. It is a run of its own
            // so the renderer's `pre-wrap` turns it into the break Excel draws.
            b"br" if parent == b"p" => {
                if let Some(para) = self.paras.last_mut() {
                    para.push(Run {
                        text: "\n".to_owned(),
                        ..Run::default()
                    });
                }
            }
            // A run's own `rPr` is the most specific text formatting; a
            // paragraph's `defRPr` sits under it, and a shape's `lstStyle` is
            // the outermost default — so a real level beats `lstStyle` even
            // when it comes later in the file, while an attribute a run omits
            // is inherited rather than turned off.
            b"rPr" | b"defRPr" => {
                let in_list_style = self.in_list_style();
                match self.run.as_mut() {
                    Some(run) => {
                        if let Some(bold) = bool_attr(element, b"b") {
                            run.bold = bold;
                        }
                        if let Some(italic) = bool_attr(element, b"i") {
                            run.italic = italic;
                        }
                    }
                    None => {
                        if let Some(bold) = bool_attr(element, b"b") {
                            if !in_list_style || self.text_bold.is_none() {
                                self.text_bold = Some(bold);
                            }
                        }
                        if let Some(italic) = bool_attr(element, b"i") {
                            if !in_list_style || self.text_italic.is_none() {
                                self.text_italic = Some(italic);
                            }
                        }
                    }
                }
                if let Some(size) = num_attr(element, b"sz") {
                    if self.run.is_some() || !in_list_style || self.text_size.is_none() {
                        self.text_size = Some(size / 100.0);
                    }
                }
            }
            // `<a:latin typeface>` names the run's font. In a `defRPr` there is
            // no run open and it becomes the shape's own face, which is what
            // the runs that name none inherit.
            b"latin" => {
                let face = self.font_name(attr(element, b"typeface"));
                let panose = attr(element, b"panose");
                let charset = num_attr(element, b"charset").map(|value| value as i32);
                let pitch = num_attr(element, b"pitchFamily")
                    .map(|value| value.clamp(0.0, 255.0) as u32);
                if let Some(run) = self.run.as_mut() {
                    run.family = face;
                    if panose.is_some() {
                        run.panose = panose;
                    }
                    if charset.is_some() {
                        run.charset = charset;
                    }
                    if pitch.is_some() {
                        run.pitch = pitch;
                    }
                } else {
                    let in_list_style = self.in_list_style();
                    if self.text_family.is_none() || !in_list_style {
                        self.text_family = face;
                    }
                    if panose.is_some() && (self.text_panose.is_none() || !in_list_style) {
                        self.text_panose = panose;
                    }
                    if charset.is_some() && (self.text_charset.is_none() || !in_list_style) {
                        self.text_charset = charset;
                    }
                    if pitch.is_some() && (self.text_pitch.is_none() || !in_list_style) {
                        self.text_pitch = pitch;
                    }
                }
            }
            // The east asian and complex-script faces sit beside the latin one
            // in the same run properties; a browser uses them per character.
            b"ea" | b"cs" => {
                let face = self.font_name(attr(element, b"typeface"));
                let east = name == b"ea";
                if let Some(run) = self.run.as_mut() {
                    if east {
                        run.ea = face;
                    } else {
                        run.cs = face;
                    }
                } else if east {
                    if self.text_ea.is_none() || !self.in_list_style() {
                        self.text_ea = face;
                    }
                } else if self.text_cs.is_none() || !self.in_list_style() {
                    self.text_cs = face;
                }
            }
            b"blip" => {
                if let Some(shape) = self.current.as_mut() {
                    // `r:embed` names the relationship to the image itself.
                    shape.media = attr(element, b"embed");
                }
            }
            // Thousandths of a percent cut off each side of the image; the
            // sides it does not name are not cropped.
            b"srcRect" => {
                if let Some(shape) = self.current.as_mut() {
                    let side = |name: &[u8]| num_attr(element, name).unwrap_or(0.0) / 100_000.0;
                    shape.crop = Some([side(b"l"), side(b"t"), side(b"r"), side(b"b")]);
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
            // A group's transform and fill are both known once its `grpSpPr`
            // has closed. Its box is the anchor's — Excel draws it from there
            // even when the `xfrm` disagrees — and the children map into it
            // from `chOff`/`chExt`.
            b"grpSpPr" => self.push_group(),
            b"grpSp" => {
                self.groups.pop();
            }
            b"sp" | b"cxnSp" | b"pic" | b"graphicFrame" => self.finish(),
            b"custGeom" => {
                if let Some(shape) = self.current.as_mut() {
                    if let Some(custom) = self.custom.take() {
                        if shape.kind.is_empty() {
                            shape.kind = "custom".to_owned();
                        }
                        shape.custom = Some(custom);
                    }
                }
            }
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
            b"grpSpPr" => self.push_group(),
            _ => {}
        }
    }

    /// Opens the transform and fill a group puts its children in, once its
    /// `grpSpPr` has been read.
    fn push_group(&mut self) {
        let (frame, anchor) = match self.groups.last() {
            Some(parent) => (
                parent.child_frame(self.group_off, self.group_ext),
                parent.anchor,
            ),
            None => ([0.0, 0.0, 1.0, 1.0], self.resolve_box()),
        };
        self.groups.push(GroupXform {
            ch_off: self.group_ch_off,
            ch_ext: self.group_ch_ext,
            frame,
            anchor,
            fill: match self.group_fill.take() {
                Some(Some(fill)) => Some(fill),
                _ => None,
            },
        });
    }

    fn finish(&mut self) {
        let Some(mut shape) = self.current.take() else {
            return;
        };
        // A shape inside a group is not anchored itself: its own `xfrm` says
        // where it sits in the group's child space, and the group's transform
        // says where that is on the sheet.
        let (from, to) = match self.groups.last() {
            Some(group) => {
                let frame = group.child_frame(self.sp_off, self.sp_ext);
                shape.group = Some(GroupFrame {
                    from_col: group.anchor.0.0,
                    from_x: group.anchor.0.1,
                    from_row: group.anchor.0.2,
                    from_y: group.anchor.0.3,
                    to_col: group.anchor.1.0,
                    to_x: group.anchor.1.1,
                    to_row: group.anchor.1.2,
                    to_y: group.anchor.1.3,
                    fx: frame[0],
                    fy: frame[1],
                    fw: frame[2],
                    fh: frame[3],
                });
                // Only a fallback: the renderer resolves `group` against the
                // real column widths, and these are the parser's approximation.
                let left = self.grid.col_start(group.anchor.0.0) + group.anchor.0.1;
                let top = self.grid.row_start(group.anchor.0.2) + group.anchor.0.3;
                let width = self.grid.col_start(group.anchor.1.0) + group.anchor.1.1 - left;
                let height = self.grid.row_start(group.anchor.1.2) + group.anchor.1.3 - top;
                let (from_col, from_x) = self.grid.col_at(left + frame[0] * width);
                let (from_row, from_y) = self.grid.row_at(top + frame[1] * height);
                let (to_col, to_x) = self.grid.col_at(left + (frame[0] + frame[2]) * width);
                let (to_row, to_y) = self.grid.row_at(top + (frame[1] + frame[3]) * height);
                (
                    (from_col, from_x, from_row, from_y),
                    (to_col, to_x, to_row, to_y),
                )
            }
            None => self.resolve_box(),
        };
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
        if self.shape_group_fill {
            shape.fill = self.groups.last().and_then(|group| group.fill.clone());
            self.shape_group_fill = false;
        }
        shape.line = self.effective_line();

        // Empty paragraphs are kept — an empty one is a blank line, and a
        // callout that opens with two of them starts its text two lines down,
        // which is exactly where Excel puts it. Dropping them pulled the text
        // up onto the numbered markers those blank lines were making room for.
        // Only a frame with no text at all is not text.
        let paras = std::mem::take(&mut self.paras);
        let mut rtl = std::mem::take(&mut self.para_rtl);
        if !rtl.iter().any(|&flag| flag) {
            rtl.clear();
        }
        // The shape's default is its first paragraph's; a shape whose
        // paragraphs agree needs only that one value.
        let mut aligns = std::mem::take(&mut self.para_align);
        let align = aligns
            .first()
            .cloned()
            .unwrap_or_else(|| "left".to_owned());
        if aligns.iter().all(|value| value == &align) {
            aligns.clear();
        }
        if paras.iter().any(|para| !para.is_empty()) {
            shape.text = Some(Text {
                size_pt: self.text_size.unwrap_or(DEFAULT_TEXT_PT),
                color: self
                    .text_color
                    .take()
                    .or_else(|| self.refs.font.clone())
                    .unwrap_or_else(|| Fill::solid(BLACK.to_owned(), 1.0))
                    .color,
                align,
                aligns,
                rtl,
                paras,
                valign: self.text_valign.clone().unwrap_or_else(|| "top".to_owned()),
                wrap: self.text_wrap.unwrap_or(true),
                insets: self
                    .text_insets
                    .unwrap_or(DEFAULT_INSETS_EMU.map(emu_to_px)),
                family: self.text_family.take(),
                panose: self.text_panose.take(),
                charset: self.text_charset.take(),
                pitch: self.text_pitch.take(),
                ea: self.text_ea.take(),
                cs: self.text_cs.take(),
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
            join: spec.join.clone(),
            miter_limit: spec.miter,
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
                if self.in_group_props() {
                    if self.group_fill.is_none() {
                        self.group_fill = Some(Some(fill));
                    }
                } else if self.fill.is_none() {
                    self.fill = Some(Some(fill));
                }
            }
            Target::Line => self.line.color = Some(fill),
            Target::Run => {
                // A run's own colour beats a paragraph default, which beats
                // `lstStyle` — the last one read wins in the second case.
                if self.run.is_some() || self.text_color.is_none() || !self.in_list_style() {
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

/// A boolean attribute, distinguishing "absent" from an explicit `0` — a run
/// that does not name `b` inherits the paragraph's, while `b="0"` turns it off.
fn bool_attr(element: &BytesStart, name: &[u8]) -> Option<bool> {
    attr(element, name).map(|value| value == "1" || value == "true")
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
        let shapes = parse(xml, &grid(), &[], &[], &HashMap::new());
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
        let shapes = parse(xml, &grid(), &theme, &ln_widths, &HashMap::new());
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
        let shapes = parse(xml, &grid(), &[], &[], &HashMap::new());
        let shape = &shapes[0];
        assert_eq!(shape.kind, "picture");
        assert_eq!(shape.media.as_deref(), Some("rId1"));
        assert_eq!((shape.from_col, shape.from_x), (0, 20.0));
        // 20px + 80px = 100px, exactly column 1's left edge.
        assert_eq!((shape.to_col, shape.to_x), (1, 0.0));
        assert_eq!((shape.to_row, shape.to_y), (0, 10.0));
    }

    /// `srcRect` cuts a fraction off each side of a picture before it is
    /// stretched into its anchor.
    #[test]
    fn reads_the_part_of_a_picture_that_is_cut_away() {
        let xml = r#"<xdr:wsDr xmlns:xdr="x" xmlns:a="a" xmlns:r="r"><xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="9525" cy="9525"/><xdr:pic><xdr:blipFill><a:blip r:embed="rId1"/><a:srcRect t="26517" b="-2"/><a:stretch/></xdr:blipFill></xdr:pic></xdr:oneCellAnchor></xdr:wsDr>"#;
        let shapes = parse(xml, &grid(), &[], &[], &HashMap::new());
        let crop = shapes[0].crop.expect("crop");
        assert!((crop[1] - 0.26517).abs() < 1e-9);
        assert!((crop[3] + 0.00002).abs() < 1e-9);
        assert_eq!((crop[0], crop[2]), (0.0, 0.0));
    }

    /// A shape inside a group is placed by the group's child coordinate space,
    /// as a fraction of the group's anchor box — nested groups compose, and the
    /// fractions are what the renderer resolves against the real column widths.
    #[test]
    fn places_a_grouped_shape_in_its_group() {
        // The group fills the anchor (col 0 to col 1 = 100px); its child space
        // is 200×100, and the ellipse sits at (150, 50) sized 50×50 in it.
        let xml = r#"<xdr:wsDr xmlns:xdr="x" xmlns:a="a"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:grpSp><xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="190500"/><a:chOff x="100" y="0"/><a:chExt cx="200" cy="100"/></a:xfrm></xdr:grpSpPr><xdr:sp><xdr:spPr><a:xfrm><a:off x="150" y="50"/><a:ext cx="50" cy="50"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp></xdr:grpSp></xdr:twoCellAnchor></xdr:wsDr>"#;
        let shapes = parse(xml, &grid(), &[], &[], &HashMap::new());
        assert_eq!(shapes.len(), 1);
        let group = shapes[0].group.as_ref().expect("in a group");
        assert_eq!((group.from_col, group.from_x), (0, 0.0));
        assert_eq!((group.to_col, group.to_x), (1, 0.0));
        assert!((group.fx - 0.25).abs() < 1e-9);
        assert!((group.fy - 0.5).abs() < 1e-9);
        assert!((group.fw - 0.25).abs() < 1e-9);
        assert!((group.fh - 0.5).abs() < 1e-9);
    }

    /// A group inside a group maps through both coordinate spaces.
    #[test]
    fn composes_a_nested_group() {
        // Outer child space 200×100; inner group at x=50 spanning half of it;
        // the child fills the inner group horizontally and half of it down.
        let xml = r#"<xdr:wsDr xmlns:xdr="x" xmlns:a="a"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:grpSp><xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="190500"/><a:chOff x="0" y="0"/><a:chExt cx="200" cy="100"/></a:xfrm></xdr:grpSpPr><xdr:grpSp><xdr:grpSpPr><a:xfrm><a:off x="50" y="0"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></xdr:grpSpPr><xdr:sp><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="50"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp></xdr:grpSp></xdr:grpSp></xdr:twoCellAnchor></xdr:wsDr>"#;
        let shapes = parse(xml, &grid(), &[], &[], &HashMap::new());
        let group = shapes[0].group.as_ref().expect("in a group");
        assert!((group.fx - 0.25).abs() < 1e-9);
        assert!((group.fy - 0.0).abs() < 1e-9);
        assert!((group.fw - 0.5).abs() < 1e-9);
        assert!((group.fh - 0.5).abs() < 1e-9);
    }

    /// `lstStyle` holds defaults; a run's own `rPr` beats them, and the
    /// attributes a run omits are inherited rather than switched off — the
    /// buttons in the design documents read 8pt bold dark grey, while their
    /// `lstStyle` says 12pt white bold.
    #[test]
    fn a_run_beats_the_list_style_it_inherits_from() {
        let xml = r#"<xdr:wsDr xmlns:xdr="x" xmlns:a="a"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr><xdr:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="1200" b="1"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:r><a:rPr sz="800"><a:solidFill><a:srgbClr val="616466"/></a:solidFill><a:latin typeface="メイリオ"/></a:rPr><a:t>x</a:t></a:r></a:p></xdr:txBody></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>"#;
        let theme = vec![Some("#FFFFFF".to_owned()), None, None];
        let shapes = parse(xml, &grid(), &theme, &[], &HashMap::new());
        let text = shapes[0].text.as_ref().expect("text");
        assert_eq!(text.size_pt, 8.0);
        assert_eq!(text.color, "#616466");
        let run = &text.paras[0][0];
        assert!(run.bold, "inherited from lstStyle");
        assert_eq!(run.family.as_deref(), Some("メイリオ"));
    }

    /// `<a:grpFill/>` draws in the group's own fill, which in the design
    /// documents is the grey behind their dropdown arrows.
    #[test]
    fn a_child_can_take_its_groups_fill() {
        let xml = r#"<xdr:wsDr xmlns:xdr="x" xmlns:a="a"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:grpSp><xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="190500"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></xdr:grpSpPr><xdr:sp><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="triangle"><a:avLst/></a:prstGeom><a:grpFill/></xdr:spPr></xdr:sp></xdr:grpSp></xdr:twoCellAnchor></xdr:wsDr>"#;
        let shapes = parse(xml, &grid(), &[], &[], &HashMap::new());
        assert_eq!(shapes[0].fill.as_ref().map(|fill| fill.color.as_str()), Some("#BFBFBF"));
    }

    /// `<a:pPr rtl="1">` makes a paragraph's base direction right-to-left, and
    /// a paragraph that names none takes the direction `lstStyle` gives.
    #[test]
    fn reads_the_direction_a_paragraph_flows_in() {
        let xml = r#"<xdr:wsDr xmlns:xdr="x" xmlns:a="a"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr><xdr:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr rtl="1"/></a:lstStyle><a:p><a:pPr rtl="1"/><a:r><a:t>a</a:t></a:r></a:p><a:p><a:pPr/><a:r><a:t>b</a:t></a:r></a:p><a:p><a:pPr rtl="0"/><a:r><a:t>c</a:t></a:r></a:p></xdr:txBody></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>"#;
        let shapes = parse(xml, &grid(), &[], &[], &HashMap::new());
        let text = shapes[0].text.as_ref().expect("text");
        // The middle paragraph names none, so it takes `lstStyle`'s direction.
        assert_eq!(text.rtl, vec![true, true, false]);
    }

    /// `algn` belongs to each paragraph; a paragraph that names none takes the
    /// alignment `lstStyle` gives, and one value covers a shape whose
    /// paragraphs agree.
    #[test]
    fn reads_each_paragraphs_alignment() {
        let xml = r#"<xdr:wsDr xmlns:xdr="x" xmlns:a="a"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr><xdr:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr algn="ctr"/></a:lstStyle><a:p><a:pPr algn="l"/><a:r><a:t>a</a:t></a:r></a:p><a:p><a:pPr/><a:r><a:t>b</a:t></a:r></a:p><a:p><a:pPr algn="r"/><a:r><a:t>c</a:t></a:r></a:p></xdr:txBody></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>"#;
        let shapes = parse(xml, &grid(), &[], &[], &HashMap::new());
        let text = shapes[0].text.as_ref().expect("text");
        assert_eq!(text.aligns, vec!["left", "center", "right"]);
        assert_eq!(text.align, "left");
    }

    /// Nothing is repeated when every paragraph agrees.
    #[test]
    fn leaves_out_alignments_that_all_agree() {
        let xml = r#"<xdr:wsDr xmlns:xdr="x" xmlns:a="a"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr><xdr:txBody><a:bodyPr/><a:p><a:pPr algn="ctr"/><a:r><a:t>a</a:t></a:r></a:p><a:p><a:pPr algn="ctr"/><a:r><a:t>b</a:t></a:r></a:p></xdr:txBody></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>"#;
        let shapes = parse(xml, &grid(), &[], &[], &HashMap::new());
        let text = shapes[0].text.as_ref().expect("text");
        assert_eq!(text.align, "center");
        assert!(text.aligns.is_empty());
    }

    /// A line's join is miter, round or bevel, and a miter's `lim` is a
    /// percentage: 200000 is twice the line width.
    #[test]
    fn reads_the_join_of_a_line() {
        let xml = r#"<xdr:wsDr xmlns:xdr="x" xmlns:a="a"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:spPr><a:prstGeom prst="triangle"><a:avLst/></a:prstGeom><a:ln w="127000"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:miter lim="200000"/></a:ln></xdr:spPr></xdr:sp></xdr:twoCellAnchor><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:spPr><a:prstGeom prst="triangle"><a:avLst/></a:prstGeom><a:ln w="127000"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:round/></a:ln></xdr:spPr></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>"#;
        let shapes = parse(xml, &grid(), &[], &[], &HashMap::new());
        let miter = shapes[0].line.as_ref().expect("line");
        assert_eq!(miter.join.as_deref(), Some("miter"));
        assert_eq!(miter.miter_limit, Some(2.0));
        let round = shapes[1].line.as_ref().expect("line");
        assert_eq!(round.join.as_deref(), Some("round"));
        assert_eq!(round.miter_limit, None);
    }

    /// `<a:br/>` is an explicit line break, kept inside its paragraph so the
    /// renderer's `pre-wrap` preserves it.
    #[test]
    fn reads_a_line_break_inside_a_paragraph() {
        let xml = r#"<xdr:wsDr xmlns:xdr="x" xmlns:a="a"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr><xdr:txBody><a:bodyPr/><a:p><a:r><a:t>a</a:t></a:r><a:br/><a:r><a:t>b</a:t></a:r></a:p></xdr:txBody></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>"#;
        let shapes = parse(xml, &grid(), &[], &[], &HashMap::new());
        let text = shapes[0].text.as_ref().expect("text");
        assert_eq!(text.paras.len(), 1);
        let joined: String = text.paras[0].iter().map(|run| run.text.as_str()).collect();
        assert_eq!(joined, "a\nb");
    }

    /// A custom geometry carries its own guides and paths, with points written
    /// as numbers or as guide names.
    #[test]
    fn reads_a_custom_geometry() {
        let xml = r#"<xdr:wsDr xmlns:xdr="x" xmlns:a="a"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:spPr><a:custGeom><a:avLst/><a:gdLst><a:gd name="g1" fmla="*/ h 2 3"/></a:gdLst><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="0" b="0"/><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="g1"/></a:moveTo><a:lnTo><a:pt x="100" y="g1"/></a:lnTo><a:lnTo><a:pt x="50" y="0"/></a:lnTo><a:close/></a:path><a:path w="100" h="100" fill="none" stroke="0"><a:moveTo><a:pt x="0" y="50"/></a:moveTo><a:arcTo wR="40" hR="40" stAng="10800000" swAng="10800000"/></a:path></a:pathLst></a:custGeom></xdr:spPr></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>"#;
        let shapes = parse(xml, &grid(), &[], &[], &HashMap::new());
        let shape = &shapes[0];
        assert_eq!(shape.kind, "custom");
        let custom = shape.custom.as_ref().expect("custom geometry");
        assert_eq!(custom.guides.len(), 1);
        assert_eq!(custom.guides[0].name, "g1");
        assert_eq!(custom.guides[0].formula, "*/ h 2 3");
        assert_eq!(custom.paths.len(), 2);

        let first = &custom.paths[0];
        assert_eq!((first.w, first.h), (100.0, 100.0));
        assert_eq!(first.fill, None);
        assert_eq!(first.stroke, None);
        assert_eq!(first.commands.len(), 4);
        assert_eq!(first.commands[0].kind, "moveTo");
        assert_eq!(first.commands[0].points[0], ["0".to_owned(), "g1".to_owned()]);
        assert_eq!(first.commands[3].kind, "close");

        let second = &custom.paths[1];
        assert_eq!(second.fill.as_deref(), Some("none"));
        assert_eq!(second.stroke, Some(false));
        let arc = second.commands[1].arc.as_ref().expect("arc");
        assert_eq!(arc, &["40".to_owned(), "40".to_owned(), "10800000".to_owned(), "10800000".to_owned()]);
    }

    /// A line without a colour takes the style refs' black — the callout's own
    /// case, where every colour comes from the theme.
    #[test]
    fn a_line_without_a_colour_takes_the_style_refs_black() {
        let theme = vec![Some("#FFFFFF".to_owned()), Some("#0D0D0D".to_owned())];
        let xml = r#"<xdr:wsDr xmlns:xdr="x" xmlns:a="a"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:spPr><a:prstGeom prst="leftBrace"><a:avLst/></a:prstGeom><a:ln w="19050"/></xdr:spPr><xdr:style><a:lnRef idx="1"><a:schemeClr val="dk1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="dk1"/></a:fillRef></xdr:style></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>"#;
        let shapes = parse(xml, &grid(), &theme, &[0.5], &HashMap::new());
        let shape = &shapes[0];
        assert_eq!(shape.kind, "leftBrace");
        // `fillRef idx="0"` is the style's way of saying "no fill".
        assert!(shape.fill.is_none());
        let line = shape.line.as_ref().expect("outline");
        assert_eq!(line.color, "#0D0D0D");
        assert_eq!(line.width, 2.0);
    }

    /// `<a:latin typeface>` names the face a shape's text is written in — the
    /// design documents' callouts ask for メイリオ and ＭＳ ゴシック, and with
    /// nothing reading it they were drawn in whatever the page used. A
    /// `+mn-ea`-style name is a slot in the theme's font scheme rather than a
    /// font, which is why the parser is handed that scheme.
    #[test]
    fn reads_the_face_a_run_is_written_in() {
        let fonts: HashMap<String, String> = [
            ("mn-lt".to_owned(), "Calibri".to_owned()),
            ("mn-ea".to_owned(), "メイリオ".to_owned()),
        ]
        .into_iter()
        .collect();
        let xml = r#"<xdr:wsDr xmlns:xdr="x" xmlns:a="a"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr><xdr:txBody><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1100"><a:latin typeface="+mn-lt"/></a:defRPr></a:pPr><a:r><a:rPr sz="1100"><a:latin typeface="ＭＳ ゴシック"/></a:rPr><a:t>a</a:t></a:r><a:r><a:rPr sz="1100"><a:latin typeface="+mn-ea"/></a:rPr><a:t>b</a:t></a:r><a:r><a:rPr sz="1100"/><a:t>c</a:t></a:r></a:p></xdr:txBody></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>"#;
        let shapes = parse(xml, &grid(), &[], &[], &fonts);
        let text = shapes[0].text.as_ref().expect("the shape has text");
        assert_eq!(text.family.as_deref(), Some("Calibri"), "the paragraph default");
        let runs = &text.paras[0];
        assert_eq!(runs[0].family.as_deref(), Some("ＭＳ ゴシック"));
        assert_eq!(runs[1].family.as_deref(), Some("メイリオ"), "a theme slot");
        assert_eq!(runs[2].family, None, "a run that names none inherits");
    }

    /// A face carries its own classification — what a viewer should fall back
    /// to when the named font is not installed.
    #[test]
    fn reads_the_classification_of_a_face() {
        let xml = r#"<xdr:wsDr xmlns:xdr="x" xmlns:a="a"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr><xdr:txBody><a:bodyPr/><a:p><a:r><a:rPr><a:latin typeface="メイリオ" panose="020B0604030504040204" pitchFamily="50" charset="-128"/><a:ea typeface="ＭＳ ゴシック"/><a:cs typeface="Arial"/></a:rPr><a:t>x</a:t></a:r></a:p></xdr:txBody></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>"#;
        let shapes = parse(xml, &grid(), &[], &[], &HashMap::new());
        let run = &shapes[0].text.as_ref().expect("text").paras[0][0];
        assert_eq!(run.panose.as_deref(), Some("020B0604030504040204"));
        assert_eq!(run.charset, Some(-128));
        assert_eq!(run.pitch, Some(50));
        assert_eq!(run.ea.as_deref(), Some("ＭＳ ゴシック"));
        assert_eq!(run.cs.as_deref(), Some("Arial"));
    }

    /// Excel writes a shape that only *looks* unrotated as a half turn plus both
    /// flips. Reading either half alone turns the shape — and its label — the
    /// wrong way round, which is what made every callout in the design documents
    /// come out upside down.
    #[test]
    fn reads_a_half_turn_and_its_cancelling_flips() {
        let xml = r#"<xdr:wsDr xmlns:xdr="x" xmlns:a="a"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:spPr><a:xfrm rot="10800000" flipH="1" flipV="1"><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm><a:prstGeom prst="borderCallout2"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>"#;
        let shapes = parse(xml, &grid(), &[], &[], &HashMap::new());
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
        let shapes = parse(xml, &grid(), &theme, &[], &HashMap::new());
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
