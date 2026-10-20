import { memo, useEffect, useId, useMemo, useRef, useState, useSyncExternalStore, type ReactNode } from "react";
import { createPortal } from "react-dom";

import { copyText } from "@/lib/clipboard";
import { shapeFontStack } from "@/lib/canvasText";
import {
  shapeDebugEnabled,
  subscribeShapeDebug,
} from "@/lib/debugPanels";
import { readMedia, type Shape, type ShapeArrow, type ShapeStroke } from "@/lib/excel";
import { SHAPE_PATTERN_TILES } from "@/lib/shapePatterns";
import {
  loadPresets,
  presetsReady,
  resolveCustom,
  resolvePreset,
  type PresetPath,
} from "@/lib/presetShapes";

/** Each gradient needs its own id; shapes are addressed by position in the list. */
function gradientId(prefix: string, index: number): string {
  return `${prefix}-g${index}`;
}

/** A cropped picture is clipped to its shape's box, so it needs an id too. */
function cropId(prefix: string, index: number): string {
  return `${prefix}-c${index}`;
}

/**
 * Presets whose box may be perfectly flat — a connector run along a single row
 * or column is exactly zero high, and dropping it would lose most of them.
 */
const LINE_PRESET = /(Connector|^line$)/;

/**
 * What a shape falls back to when the table has nothing for it — an unknown
 * preset, or the moment before the table finishes loading. Drawn as an outline
 * of its own box: the layout is right, and it cannot hide the cells under it.
 */
function boxOutline(w: number, h: number): PresetPath {
  return { d: `M0 0H${w}V${h}H0Z`, filled: false, stroked: true };
}

/**
 * Arrowhead sizes, as multiples of the line's width. Excel takes them from the
 * file's own `w` and `len` (sm, med, lg) scaled by the stroke, so a hairline
 * would otherwise get an invisible head — hence the floors, which are what the
 * reference diagrams show for their 0.5pt arrows.
 */
const ARROW_LENGTH = { sm: 4, med: 6, lg: 8 } as const;
const ARROW_WIDTH = { sm: 2, med: 3, lg: 4 } as const;

function arrowSize(arrow: ShapeArrow, lineWidth: number) {
  return {
    length: Math.max(5, (ARROW_LENGTH[arrow.length] ?? 6) * lineWidth),
    half: Math.max(2.5, (ARROW_WIDTH[arrow.width] ?? 3) * lineWidth),
  };
}

/** Where along the marker the line's end should land. */
function arrowTip(arrow: ShapeArrow, length: number): number {
  // An oval head is a dot sitting *on* the end, not a tip pointing at it.
  return arrow.kind === "oval" ? length / 2 : length;
}

/**
 * A marker per distinct arrowhead. The id has to be readable from both the
 * `defs` that defines it and the shape that points at it, so it is built from
 * what the arrow looks like rather than from a counter.
 */
function arrowId(prefix: string, arrow: ShapeArrow, line: ShapeStroke): string {
  const color = line.color.replace(/[^0-9a-f]/gi, "");
  return `${prefix}-a-${arrow.kind}-${arrow.width}-${arrow.length}-${color}-${Math.round(line.width * 100)}`;
}

/** A marker's outline in its own box, tip at the right edge. */
function arrowShape(arrow: ShapeArrow, length: number, half: number): ReactNode {
  const height = half * 2;
  switch (arrow.kind) {
    case "oval":
      return <ellipse cx={length / 2} cy={half} rx={half} ry={half} />;
    case "diamond":
      return <path d={`M 0 ${half} L ${length / 2} 0 L ${length} ${half} L ${length / 2} ${height} Z`} />;
    case "stealth":
      return (
        <path
          d={`M 0 0 L ${length} ${half} L 0 ${height} L ${length * 0.6} ${half} Z`}
        />
      );
    case "arrow":
      // An open head, drawn as a stroke rather than a filled area.
      return (
        <path
          d={`M 0 0 L ${length} ${half} L 0 ${height}`}
          fill="none"
          stroke="currentColor"
          strokeWidth={Math.max(1, half / 3)}
        />
      );
    default:
      return <path d={`M 0 0 L ${length} ${half} L 0 ${height} Z`} />;
  }
}

/** Points to pixels: the same 96/72 Excel works in. */
const PT_TO_PX = 4 / 3;

/**
 * Excel's preset dash names as SVG dash arrays. The lengths are Excel's own, in
 * multiples of the stroke width — which is why they have to be multiplied by it
 * here: a `dash` array written as plain pixels turns a 4px border into an
 * almost solid line, where Excel draws long dashes.
 */
const DASHES: Record<string, number[]> = {
  solid: [],
  dash: [4, 3],
  dashDot: [4, 3, 1, 3],
  lgDash: [8, 3],
  lgDashDot: [8, 3, 1, 3],
  lgDashDotDot: [8, 3, 1, 3, 1, 3],
  sysDash: [3, 1],
  sysDot: [1, 1],
  sysDashDot: [3, 1, 1, 1],
  sysDashDotDot: [3, 1, 1, 1, 1, 1],
};

function dashArray(dash: string | null, width: number): string | undefined {
  const pattern = dash ? DASHES[dash] : undefined;
  if (!pattern || pattern.length === 0) return undefined;
  return pattern.map((part) => Number((part * width).toFixed(2))).join(" ");
}

/** The media types a browser can draw. WMF and EMF, which older documents use
 * for clip art, cannot be shown and are left out rather than shown broken. */
const MIME: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  svg: "image/svg+xml",
  webp: "image/webp",
  tif: "image/tiff",
  tiff: "image/tiff",
};

/** Images already read, keyed by workbook and part, so a sheet switch does not
 * ask for them again. */
const mediaCache = new Map<string, string>();

function mimeOf(part: string): string | null {
  const extension = part.split(".").pop()?.toLowerCase() ?? "";
  return MIME[extension] ?? null;
}

/** Reads every image a sheet's shapes refer to, once each, in one call. */
function useMedia(path: string, parts: string[]): Map<string, string> {
  const [, bump] = useState(0);

  useEffect(() => {
    let alive = true;
    const missing = parts.filter((part) => !mediaCache.has(`${path}#${part}`));
    if (missing.length > 0) {
      readMedia(path, missing)
        .then((found) => {
          for (const [part, data] of Object.entries(found)) {
            mediaCache.set(`${path}#${part}`, data);
          }
          if (alive) bump((tick) => tick + 1);
        })
        .catch(() => {
          // A picture we cannot read is simply not drawn; the cells still are.
        });
    }
    return () => {
      alive = false;
    };
  }, [path, parts]);

  const ready = new Map<string, string>();
  for (const part of parts) {
    const data = mediaCache.get(`${path}#${part}`);
    if (data) ready.set(part, data);
  }
  return ready;
}

/**
 * Excel gives a linear gradient as a direction in degrees, clockwise from
 * pointing right. SVG wants the two ends of the line it runs along, in the
 * shape's own 0-to-1 box — which is the unit SVG already uses.
 */
function linearAttrs(angle: number) {
  const radians = (angle * Math.PI) / 180;
  const dx = Math.cos(radians) / 2;
  const dy = Math.sin(radians) / 2;
  return { x1: 0.5 - dx, y1: 0.5 - dy, x2: 0.5 + dx, y2: 0.5 + dy };
}

/** Presentation attributes an SVG element inherits down to its children. */
interface StrokeAttrs {
  stroke?: string;
  strokeWidth?: number;
  strokeOpacity?: number;
  strokeDasharray?: string;
  strokeLinecap?: "butt";
  strokeLinejoin?: "miter" | "round" | "bevel";
  strokeMiterlimit?: number;
}

function strokeProps(line: ShapeStroke | null): StrokeAttrs {
  if (!line) return { stroke: "none" };
  return {
    stroke: line.color,
    strokeWidth: line.width,
    strokeOpacity: line.alpha,
    strokeDasharray: dashArray(line.dash, line.width),
    strokeLinecap: "butt",
    strokeLinejoin: line.join ?? "miter",
    // SVG's own default limit is 4; Excel draws mitered joins out to 8.
    strokeMiterlimit: line.miterLimit ?? 8,
  };
}

function flexSide(align: string): "center" | "flex-end" | "flex-start" {
  return align === "center" ? "center" : align === "right" ? "flex-end" : "flex-start";
}

function textOf(shape: Shape, w: number, h: number): ReactNode {
  const text = shape.text;
  if (!text) return null;
  const justify =
    text.valign === "middle" ? "center" : text.valign === "bottom" ? "flex-end" : "flex-start";
  const align = flexSide(text.align);
  return (
    <foreignObject width={w} height={h}>
      <div
        style={{
          display: "flex",
          flexDirection: "column",
          justifyContent: justify,
          alignItems: align,
          width: "100%",
          height: "100%",
          boxSizing: "border-box",
          padding: `${text.insets[1]}px ${text.insets[2]}px ${text.insets[3]}px ${text.insets[0]}px`,
          color: text.color,
          fontSize: text.sizePt * PT_TO_PX,
          fontFamily: shapeFontStack(
            text.family,
            text.ea,
            text.cs,
            text.panose,
            text.charset,
            text.pitch,
          ),
          fontWeight: text.paras.some((para) => para.some((run) => run.bold)) ? 700 : undefined,
          lineHeight: 1.2,
          // Excel's `wrap="none"` lets the text run past the shape, the way an
          // unstyled cell spills into its neighbours.
          whiteSpace: text.wrap ? "pre-wrap" : "pre",
          overflow: "visible",
          textAlign: text.align,
          wordBreak: "break-word",
        }}
      >
        {text.paras.map((para, index) => {
          const paraAlign = text.aligns?.[index] ?? text.align;
          return (
            // An empty paragraph is a blank line in Excel, and empty blocks have
            // no height of their own — without this a callout's leading blank
            // lines collapse and its text rises into whatever they were clearing.
            // One line is `lineHeight` (1.2) of the text's own size.
            <div
              key={index}
              // A right-to-left paragraph's base direction decides how mixed
              // Hebrew/Arabic and Latin text orders itself; Excel keeps `algn`
              // physical, so only the direction changes here.
              dir={text.rtl?.[index] ? "rtl" : undefined}
              style={{
                minHeight: "1.2em",
                // `algn` is per paragraph in the file; the container carries
                // the shape's default for the ones that agree with it.
                textAlign: paraAlign,
                alignSelf: flexSide(paraAlign),
              }}
            >
              {para.map((run, runIndex) => (
                <span
                  key={runIndex}
                  style={{
                    fontWeight: run.bold ? 700 : undefined,
                    fontStyle: run.italic ? "italic" : undefined,
                    // A run that names no face keeps the shape's, so this stays
                    // out of the way rather than restating it.
                    fontFamily: shapeFontStack(
                      run.family,
                      run.ea,
                      run.cs,
                      run.panose,
                      run.charset,
                      run.pitch,
                    ),
                  }}
                >
                  {run.text}
                </span>
              ))}
            </div>
          );
        })}
      </div>
    </foreignObject>
  );
}

/**
 * What the hidden readout says about the sheet itself, gathered by the grid
 * because that is where the layout is known. See `debugPanels.ts`.
 */
export interface ShapeDebugFacts {
  sheet: string;
  rows: number;
  cols: number;
  fontFamily: string | null;
  fontSizePx: number;
  /** The workbook font's digit width, measured — the unit column widths use. */
  digitWidth: number;
  defaultColChars: number | null;
  defaultColPx: number;
  defaultRowPx: number;
  showGrid: boolean;
  merges: number;
  /** Columns the file sized explicitly, in pixels. */
  colWidths: Map<number, number>;
  printArea: { row: number; col: number; rows: number; cols: number } | null;
  /** The selected cell's resolved style, already laid out as printed. */
  cellReadout: string;
  /** What the last frame repainted, how long each layer took, and what it drew. */
  timings: string;
}

export interface ShapeLayerProps {
  shapes: Shape[];
  /** Pixel position of a column's left edge. */
  colX: (col: number) => number;
  /** Pixel position of a sheet row's top edge. A filtered-out row takes its place
   * from the next visible one, so a shape shrinks rather than disappearing. */
  rowY: (row: number) => number;
  /** The data area's size, which the layer is clipped to. */
  width: number;
  height: number;
  path: string;
  /**
   * A shape whose anchor is inside a frozen band is pinned: the layer is
   * translated by the scroll, so its own position has to add it back. Zero
   * unless the sheet freezes something, which is what keeps a scroll from
   * re-rendering this layer on ordinary sheets.
   */
  frozenRows?: number;
  frozenCols?: number;
  /**
   * The grid's zoom as a factor. A shape's hatch is 8 screen pixels however far
   * the sheet is zoomed — Excel's own behaviour — so the pattern is scaled back
   * by it inside the layer, which the layer's CSS zoom would otherwise carry.
   */
  zoom?: number;
  /**
   * How far each shape has been dragged from where the file puts it, in sheet
   * pixels. Session state only: a shape that covers the cells under it can be
   * moved aside, and nothing is written back.
   */
  offsets?: ReadonlyMap<number, { x: number; y: number }>;
  /** Called with a shape's new offset while it is being dragged. */
  onMoveShape?: (index: number, x: number, y: number) => void;
  /** Read out when Ctrl+Shift+D has switched the hidden panel on. */
  debug?: ShapeDebugFacts | null;
}

/**
 * The shapes and pictures a sheet floats above its cells.
 *
 * One SVG laid over the cells in the same coordinate space, rather than a
 * second scroller: two scrollers drift apart as soon as their widths differ,
 * and the grid has been down that road already.
 */
export const ShapeLayer = memo(function ShapeLayer({
  shapes,
  colX,
  rowY,
  width,
  height,
  path,
  frozenRows = 0,
  frozenCols = 0,
  zoom = 1,
  offsets,
  onMoveShape,
  debug,
}: ShapeLayerProps) {
  const parts = useMemo(
    () =>
      shapes
        .map((shape) => shape.media)
        .filter((part): part is string => part !== null && mimeOf(part) !== null),
    [shapes],
  );
  const media = useMedia(path, parts);
  // Gradients are referenced by id from the shapes that use them, so the layer
  // needs a prefix of its own. React's ids carry colons, which a `url(#…)`
  // fragment cannot.
  const ids = useId().replace(/:/g, "");

  /**
   * The preset table is fetched once per session, and the layer has to redraw
   * when it lands: until then every shape falls back to its own box, which is
   * the right layout and the wrong drawing.
   */
  const [table, setTable] = useState(0);
  void table;
  useEffect(() => {
    if (presetsReady()) return;
    void loadPresets().then(() => setTable((tick) => tick + 1));
  }, []);

  /**
   * The arrowheads this sheet needs, one marker each. Collected rather than
   * emitted per shape, so a hundred arrows share a handful of markers.
   */
  const arrows = useMemo(() => {
    const found = new Map<string, { arrow: ShapeArrow; line: ShapeStroke }>();
    for (const shape of shapes) {
      const line = shape.line;
      if (!line) continue;
      for (const arrow of [line.head, line.tail]) {
        if (arrow) found.set(arrowId(ids, arrow, line), { arrow, line });
      }
    }
    return found;
  }, [shapes, ids]);

  /**
   * One `<pattern>` per distinct hatch. `patternUnits="userSpaceOnUse"` anchors
   * it in the user space of whoever references it — the shape's own translated
   * group — so every shape starts its hatch at its own top-left, which is where
   * Excel puts it. The pattern is scaled back by the layer's zoom: Excel keeps a
   * shape's hatch at 8 screen pixels however far the sheet is zoomed.
   */
  const patterns = useMemo(() => {
    const found = new Map<string, { id: string; kind: string; fg: string; bg: string }>();
    for (const shape of shapes) {
      const pattern = shape.fill?.pattern;
      if (!pattern || !SHAPE_PATTERN_TILES[pattern.kind]) continue;
      const fg = pattern.fg ?? "#000000";
      const bg = pattern.bg ?? "#FFFFFF";
      const key = `${pattern.kind}|${fg}|${bg}`;
      if (!found.has(key)) {
        found.set(key, { id: `${ids}-p${found.size}`, kind: pattern.kind, fg, bg });
      }
    }
    return found;
  }, [shapes, ids]);

  const patternId = (shape: Shape): string | null => {
    const pattern = shape.fill?.pattern;
    if (!pattern) return null;
    return patterns.get(`${pattern.kind}|${pattern.fg ?? "#000000"}|${pattern.bg ?? "#FFFFFF"}`)?.id ?? null;
  };

  /**
   * Each shape's box, in the layer's own pixels. A radial gradient is declared
   * in the `<defs>` above the shapes, so the box has to be known there too —
   * computed once here rather than twice.
   */
  const boxes = shapes.map((shape) => {
    const group = shape.group;
    const gLeft = group ? colX(group.fromCol) + group.fromX : 0;
    const gTop = group ? rowY(group.fromRow) + group.fromY : 0;
    const gWidth = group ? Math.max(0, colX(group.toCol) + group.toX - gLeft) : 0;
    const gHeight = group ? Math.max(0, rowY(group.toRow) + group.toY - gTop) : 0;
    const x = group ? gLeft + group.fx * gWidth : colX(shape.fromCol) + shape.fromX;
    const y = group ? gTop + group.fy * gHeight : rowY(shape.fromRow) + shape.fromY;
    const w = group ? group.fw * gWidth : Math.max(0, colX(shape.toCol) + shape.toX - x);
    const h = group ? group.fh * gHeight : Math.max(0, rowY(shape.toRow) + shape.toY - y);
    return { x, y, w, h, gWidth, gHeight };
  });

  /**
   * The drag in progress. The pointer is captured on the shape's own `<g>` so
   * it can leave the shape — or the layer's clipped area — without the move
   * stopping, and the offset is reported in sheet pixels, not screen ones.
   */
  const drag = useRef<{
    index: number;
    startX: number;
    startY: number;
    from: { x: number; y: number };
  } | null>(null);
  /**
   * The offsets of a drag still in progress. Kept here rather than lifted on
   * every move: the pointer fires many times a second and this layer is small,
   * while the grid that owns the committed offsets is not. Only the release
   * reports the final position up.
   */
  const [live, setLive] = useState<Map<number, { x: number; y: number }>>(() => new Map());

  const beginDrag = (event: React.PointerEvent<SVGGElement>, index: number) => {
    if (event.button !== 0 || !onMoveShape) return;
    event.preventDefault();
    event.stopPropagation();
    // Capture keeps the drag alive when the pointer leaves the shape. A
    // synthetic pointer has no active id to capture, which is not an error here.
    try {
      event.currentTarget.setPointerCapture(event.pointerId);
    } catch {
      // no capture; the handlers on the shape still see the move
    }
    drag.current = {
      index,
      startX: event.clientX,
      startY: event.clientY,
      from: live.get(index) ?? offsets?.get(index) ?? { x: 0, y: 0 },
    };
  };

  const dragTo = (event: React.PointerEvent<SVGGElement>) => {
    const current = drag.current;
    if (!current) return;
    const scale = zoom || 1;
    const next = {
      x: current.from.x + (event.clientX - current.startX) / scale,
      y: current.from.y + (event.clientY - current.startY) / scale,
    };
    setLive((all) => new Map(all).set(current.index, next));
  };

  const endDrag = (event: React.PointerEvent<SVGGElement>) => {
    const current = drag.current;
    if (!current) return;
    drag.current = null;
    if (event.currentTarget.hasPointerCapture(event.pointerId)) {
      event.currentTarget.releasePointerCapture(event.pointerId);
    }
    const at = live.get(current.index);
    if (at && onMoveShape) onMoveShape(current.index, at.x, at.y);
    setLive((all) => {
      const next = new Map(all);
      next.delete(current.index);
      return next;
    });
  };

  /** Whether the hidden text readout is on; Ctrl+Shift+D flips it. */
  const shapeDebug = useSyncExternalStore(subscribeShapeDebug, shapeDebugEnabled);
  /**
   * Dev builds only — the packaged app ships without the readout at all.
   * `import.meta.env.DEV` is a constant, so everything below it folds away in a
   * release build rather than merely staying hidden.
   */
  const facts = import.meta.env.DEV ? debug : null;

  /**
   * Rows for that readout: the sheet's own numbers first, then every shape it
   * floats — not just the ones with text, since a wrong box is the other half
   * of the same question. The sheet's half is built here, before the early
   * return below, because most sheets have no shapes at all and their facts are
   * the whole of what there is to report. `debugPanels.ts` says how it is
   * switched on.
   */
  const debugRows: string[] = [];
  if (shapeDebug && facts) {
    const sized = [...facts.colWidths.entries()]
      .sort(([a], [b]) => a - b)
      .slice(0, 12)
      .map(([col, px]) => `${col}=${px}px`)
      .join("  ");
    const area = facts.printArea;
    debugRows.push(
      [
        `sheet ${facts.sheet}   ${facts.rows} rows × ${facts.cols} cols   gridlines=${facts.showGrid ? "on" : "off"}`,
        `font ${facts.fontFamily ?? "?"} at ${facts.fontSizePx}px   digit width ${facts.digitWidth}px`,
        `default col ${facts.defaultColChars ?? "?"} chars → ${facts.defaultColPx}px   default row ${facts.defaultRowPx}px`,
        `print area ${area ? `r${area.row} c${area.col} ${area.rows}×${area.cols}` : "none"}   merges ${facts.merges}`,
        `sized columns (${facts.colWidths.size}): ${sized || "none"}`,
        `shapes ${shapes.length}`,
        "",
        `selected cell\n${facts.cellReadout}`,
        "",
        facts.timings,
        "",
      ].join("\n"),
    );
  }

  if (shapes.length === 0) {
    return shapeDebug && facts
      ? createPortal(<ShapeDebug rows={debugRows} />, document.body)
      : null;
  }

  return (
    // Positioning and click-through belong to the caller, which is the only
    // place that knows where the data area starts.
    <svg width={width} height={height}>
      <defs>
        {shapes.map((shape, index) => {
          const fill = shape.fill;
          if (!fill || fill.stops.length === 0) return null;
          const stops = fill.stops.map((stop, stopIndex) => (
            <stop
              key={stopIndex}
              offset={stop.pos}
              stopColor={stop.color}
              stopOpacity={stop.alpha}
            />
          ));
          const radial = fill.radial;
          if (!radial) {
            return (
              <linearGradient
                key={index}
                id={gradientId(ids, index)}
                {...linearAttrs(fill.angle)}
              >
                {stops}
              </linearGradient>
            );
          }
          const [left, top, right, bottom] = radial.focus;
          const cx = left + (1 - left - right) / 2;
          const cy = top + (1 - top - bottom) / 2;
          if (radial.kind === "circle") {
            // A circle through the box's four corners: measured off Excel's own
            // render, `circle` puts the last stop on the corners, not the edges.
            const { w, h } = boxes[index];
            return (
              <radialGradient
                key={index}
                id={gradientId(ids, index)}
                gradientUnits="userSpaceOnUse"
                cx={w * cx}
                cy={h * cy}
                r={Math.max(1, Math.hypot(w / 2, h / 2))}
              >
                {stops}
              </radialGradient>
            );
          }
          // `rect` and `shape` reach the last stop at the box's edges, which is
          // the ellipse an objectBoundingBox gradient inscribes. `shape` follows
          // a non-rectangular outline in Excel; the inscribed ellipse is the
          // closest form SVG offers without a per-shape mask.
          return (
            <radialGradient key={index} id={gradientId(ids, index)} cx={cx} cy={cy} r={0.5}>
              {stops}
            </radialGradient>
          );
        })}
        {[...patterns.values()].map((pattern) => (
          <pattern
            key={pattern.id}
            id={pattern.id}
            width={8}
            height={8}
            patternUnits="userSpaceOnUse"
            patternTransform={`scale(${1 / zoom})`}
          >
            <rect width={8} height={8} fill={pattern.bg} />
            {SHAPE_PATTERN_TILES[pattern.kind].map((row, y) =>
              [...row].map((cell, x) =>
                cell === "#" ? (
                  <rect
                    key={`${x}-${y}`}
                    x={x}
                    y={y}
                    width={1}
                    height={1}
                    fill={pattern.fg}
                  />
                ) : null,
              ),
            )}
          </pattern>
        ))}
        {[...arrows].map(([id, { arrow, line }]) => {
          const { length, half } = arrowSize(arrow, line.width);
          return (
            <marker
              key={id}
              id={id}
              // User units, not stroke widths: the size is worked out here, from
              // the line width and the file's own sm/med/lg.
              markerUnits="userSpaceOnUse"
              markerWidth={length}
              markerHeight={half * 2}
              refX={arrowTip(arrow, length)}
              refY={half}
              // Turned about-face at the start, so one marker serves both ends.
              orient="auto-start-reverse"
              viewBox={`0 0 ${length} ${half * 2}`}
              fill={line.color}
              fillOpacity={line.alpha}
              color={line.color}
            >
              {arrowShape(arrow, length, half)}
            </marker>
          );
        })}
      </defs>
      {shapes.map((shape, index) => {
        const group = shape.group;
        const { x, y, w, h, gWidth, gHeight } = boxes[index];
        // A shape in a frozen band stays put while the layer under it scrolls.
        // Excel places a shape by its anchor, so that is what decides its pane.
        // The scroll is added back imperatively — see `CanvasGrid`'s transform
        // step — because this layer does not re-render on a scroll.
        const pinnedX = (group?.fromCol ?? shape.fromCol) < frozenCols;
        const pinnedY = (group?.fromRow ?? shape.fromRow) < frozenRows;

        if (shapeDebug && facts) {
          const lines = [
            `${index + 1}. ${shape.kind}  box ${w.toFixed(1)}×${h.toFixed(1)} @ (${x.toFixed(1)}, ${y.toFixed(1)})`,
            group
              ? `   in group: anchor ${group.fromCol}+${group.fromX.toFixed(1)},${group.fromRow}+${group.fromY.toFixed(1)} · ${gWidth.toFixed(1)}×${gHeight.toFixed(1)} · at ${group.fx.toFixed(3)},${group.fy.toFixed(3)} · ${group.fw.toFixed(3)}×${group.fh.toFixed(3)}`
              : `   x: col ${shape.fromCol} + ${shape.fromX.toFixed(1)} = ${x.toFixed(1)} → col ${shape.toCol} + ${shape.toX.toFixed(1)} = ${(colX(shape.toCol) + shape.toX).toFixed(1)}`,
          ];
          if (!group) {
            lines.push(
              `   y: row ${shape.fromRow} + ${shape.fromY.toFixed(1)} = ${y.toFixed(1)} → row ${shape.toRow} + ${shape.toY.toFixed(1)} = ${(rowY(shape.toRow) + shape.toY).toFixed(1)}`,
            );
          }
          if (shape.text) {
            const t = shape.text;
            const [l, top, r, b] = t.insets;
            lines.push(
              `   insets l${l.toFixed(1)} t${top.toFixed(1)} r${r.toFixed(1)} b${b.toFixed(1)} → text ${(w - l - r).toFixed(1)}×${(h - top - b).toFixed(1)}`,
              `   wrap=${t.wrap} size=${t.sizePt}pt (${(t.sizePt * PT_TO_PX).toFixed(2)}px) align=${t.align} valign=${t.valign}`,
              ...t.paras.map(
                (para, i) =>
                  `   p${i}: ${para.map((run) => run.text).join("") || "(blank line)"}`,
              ),
            );
          }
          debugRows.push(lines.join("\n"));
        }
        // A connector is anchored between two cells on the *same* row or column
        // whenever it runs perfectly straight, so it is exactly zero in one
        // direction. Dropping it like an empty box would lose most of them.
        const flat = LINE_PRESET.test(shape.kind);
        if (w <= 0 && h <= 0) return null;
        if (!flat && (w <= 0 || h <= 0)) return null;

        // Excel encodes a shape that only *looks* unrotated as `rot="180"` plus
        // both flips — two half-turns that cancel — so the flips have to be kept
        // alongside the rotation. Dropping them (to "keep the text readable")
        // turns the label upside down, which is exactly what makes these callouts
        // look reversed. SVG applies a transform list outside-in, matching the
        // order OOXML defines: flip first, then rotate.
        //
        // The frame the geometry is drawn in, which is not always the anchor's:
        // Excel writes the *unrotated* geometry into a frame that `rot` has
        // already turned, so a quarter turn means the geometry belongs in that
        // box **transposed** and is turned onto the anchor from there. Drawing
        // straight into the anchor laid every quarter-turned connector on its
        // side, ends in the wrong corners — and only where the box is far from
        // square, which is why the squarer ones looked right.
        const quarter = Math.abs(shape.rot) % 180 === 90;
        const gw = quarter ? h : w;
        const gh = quarter ? w : h;

        const placement = [
          shape.rot ? `rotate(${shape.rot} ${w / 2} ${h / 2})` : "",
          quarter ? `translate(${(w - gw) / 2} ${(h - gh) / 2})` : "",
          shape.flipH || shape.flipV
            ? `translate(${gw / 2} ${gh / 2}) scale(${shape.flipH ? -1 : 1} ${shape.flipV ? -1 : 1}) translate(${-gw / 2} ${-gh / 2})`
            : "",
        ]
          .filter(Boolean)
          .join(" ");

        const hatch = patternId(shape);
        const fillValue = shape.fill
          ? shape.fill.stops.length > 0
            ? `url(#${gradientId(ids, index)})`
            : hatch
              ? `url(#${hatch})`
              : shape.fill.color
          : "none";
        /** The preset's own paths for this box, or its outline until we have them. */
        /** The shape's own paths, else the preset's, else its outline. */
        const paths =
          (shape.custom ? resolveCustom(shape.custom, gw, gh) : null) ??
          resolvePreset(shape.kind, gw, gh, shape.adjust) ?? [
            boxOutline(gw, gh),
          ];

        const image = shape.media ? media.get(shape.media) : undefined;
        const mime = shape.media ? mimeOf(shape.media) : null;
        // Excel stretches a picture to its anchor unless it was cropped:
        // `srcRect` cuts a fraction off each side, so the drawn image is larger
        // than the box by exactly that much and is clipped to it.
        const [cropL, cropT, cropR, cropB] = shape.crop ?? [0, 0, 0, 0];
        const keepX = 1 - cropL - cropR;
        const keepY = 1 - cropT - cropB;
        const imageW = keepX > 0.001 ? gw / keepX : gw;
        const imageH = keepY > 0.001 ? gh / keepY : gh;

        const moved = live.get(index) ?? offsets?.get(index);
        const mx = x + (moved?.x ?? 0);
        const my = y + (moved?.y ?? 0);

        return (
          <g
            key={index}
            data-shape-x={mx}
            data-shape-y={my}
            data-pin-x={pinnedX ? "1" : undefined}
            data-pin-y={pinnedY ? "1" : undefined}
            transform={`translate(${mx} ${my})`}
            // A shape takes the pointer so it can be dragged off whatever it
            // covers; the layer stays click-through everywhere else.
            style={onMoveShape ? { pointerEvents: "auto", cursor: "move" } : undefined}
            onPointerDown={(event) => beginDrag(event, index)}
            onPointerMove={dragTo}
            onPointerUp={endDrag}
            onPointerCancel={endDrag}
          >
            <g transform={placement || undefined}>
              {/* A chart or table frame is anchored like a shape but has contents
                  we do not read, so only its outline is drawn. */}
              {image && mime ? (
                <g clipPath={shape.crop ? `url(#${cropId(ids, index)})` : undefined}>
                  {shape.crop ? (
                    <clipPath id={cropId(ids, index)}>
                      <rect width={gw} height={gh} />
                    </clipPath>
                  ) : null}
                  <image
                    href={`data:${mime};base64,${image}`}
                    x={keepX > 0.001 ? -cropL * imageW : 0}
                    y={keepY > 0.001 ? -cropT * imageH : 0}
                    width={imageW}
                    height={imageH}
                    preserveAspectRatio="none"
                  />
                </g>
              ) : shape.kind === "frame" ? (
                // A chart or table frame has contents we do not read, so only
                // its outline is drawn.
                <rect width={gw} height={gh} fill="none" {...strokeProps(shape.line)} />
              ) : (
                <g
                  {...strokeProps(shape.line)}
                  // An arrowhead belongs at the end of the one path a connector
                  // resolves to; a preset that draws several gets none.
                  markerStart={
                    paths.length === 1 && shape.line?.head
                      ? `url(#${arrowId(ids, shape.line.head, shape.line)})`
                      : undefined
                  }
                  markerEnd={
                    paths.length === 1 && shape.line?.tail
                      ? `url(#${arrowId(ids, shape.line.tail, shape.line)})`
                      : undefined
                  }
                >
                  {paths.map((path, part) => (
                    <path
                      key={part}
                      d={path.d}
                      fill={path.filled ? fillValue : "none"}
                      fillOpacity={path.filled ? shape.fill?.alpha : undefined}
                      stroke={path.stroked ? undefined : "none"}
                    />
                  ))}
                </g>
              )}
              {textOf(shape, gw, gh)}
            </g>
          </g>
        );
      })}
      {shapeDebug && facts &&
        createPortal(<ShapeDebug rows={debugRows} />, document.body)}
    </svg>
  );
});

/**
 * The hidden text readout, in the portal so it is not caught by the grid's
 * scroll, zoom or clipping.
 *
 * The font is read back from the DOM rather than passed down: a shape's text
 * names no family of its own in the file, so what it actually renders in is
 * whatever it inherits — and the browser's fallback is not the file's.
 */
function ShapeDebug({ rows }: { rows: string[] }) {
  const readout = rows.join("\n");
  const [font, setFont] = useState("");
  const [copied, setCopied] = useState(false);
  const box = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const inherited = box.current?.parentElement;
    const computed = inherited ? getComputedStyle(inherited).font : "";
    if (computed) setFont(computed);
    if (readout) console.log(`sheet-lens debug\n${computed}\n${readout}`);
  }, [readout]);

  if (rows.length === 0) return null;
  const body = `inherited font: ${font || "…"}\n\n${readout}`;
  return (
    // Focusable, and not click-through: the readout is meant to be read, scrolled
    // and copied out, and Ctrl+C is caught by the grid while that has focus.
    <div
      ref={box}
      tabIndex={0}
      className="fixed right-2 bottom-9 z-50 flex max-h-[60vh] w-[34rem] flex-col rounded-sm border border-border bg-background/95 font-mono text-[10px] leading-snug text-foreground shadow-lg outline-none select-text"
    >
      <div className="flex shrink-0 items-center gap-2 border-b border-border px-2 py-1 font-sans">
        <span className="font-semibold">sheet-lens debug</span>
        <span className="text-muted-foreground">Ctrl+Shift+D hides this</span>
        <button
          type="button"
          onClick={() => {
            void copyText(body).then((ok) => {
              setCopied(ok);
              window.setTimeout(() => setCopied(false), 1200);
            });
          }}
          className="ml-auto rounded-sm border border-border px-1.5 py-0.5 hover:bg-accent"
        >
          {copied ? "copied" : "copy"}
        </button>
      </div>
      <div className="min-h-0 overflow-auto p-2 whitespace-pre-wrap">{body}</div>
    </div>
  );
}
