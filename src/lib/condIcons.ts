/**
 * The icon-set marks drawn at the left of a cell.
 *
 * Excel's own icon assets are not available here, so each family is built from
 * small vector marks in the colours Excel's render uses — the red, amber and
 * green measured off a CopyPicture — inside the same geometry: a 14px box at
 * the cell's left edge, centred vertically, with the text moved right by 16px.
 */

/** The box the mark occupies, in cell pixels at 100%. */
export const ICON_SIZE = 14;
/** How far a cell's text moves right to clear the mark. */
export const ICON_INSET = 16;

const RED = "#D65532";
const RED_EDGE = "#9E381C";
const AMBER = "#EAC282";
const AMBER_EDGE = "#A98A4E";
const GREEN = "#67A38F";
const GREEN_EDGE = "#36725E";
const GREY = "#A8A8A8";
const GREY_EDGE = "#6E6E6E";
const BLUE = "#5B9BD5";
const EMPTY = "#D9D9D9";
const INK = "#FFFFFF";

function colours(level: number, count: number, grey: boolean): [string, string] {
  if (grey) return [GREY, GREY_EDGE];
  if (level <= 0) return [RED, RED_EDGE];
  if (level >= count - 1) return [GREEN, GREEN_EDGE];
  return [AMBER, AMBER_EDGE];
}

function disc(
  ctx: CanvasRenderingContext2D,
  cx: number,
  cy: number,
  radius: number,
  fill: string,
  edge: string | null,
): void {
  ctx.beginPath();
  ctx.arc(cx, cy, radius, 0, Math.PI * 2);
  ctx.fillStyle = fill;
  ctx.fill();
  if (edge) {
    ctx.lineWidth = 1;
    ctx.strokeStyle = edge;
    ctx.stroke();
  }
}

/** An arrow pointing up, rotated to the level's direction. */
function arrow(
  ctx: CanvasRenderingContext2D,
  cx: number,
  cy: number,
  size: number,
  angle: number,
  fill: string,
  edge: string,
): void {
  const shaft = size * 0.36;
  const head = size * 0.5;
  const half = size * 0.9;
  ctx.save();
  ctx.translate(cx, cy);
  ctx.rotate(angle);
  ctx.beginPath();
  ctx.moveTo(0, -half / 2);
  ctx.lineTo(head / 2, -half / 2 + head * 0.7);
  ctx.lineTo(shaft / 2, -half / 2 + head * 0.7);
  ctx.lineTo(shaft / 2, half / 2);
  ctx.lineTo(-shaft / 2, half / 2);
  ctx.lineTo(-shaft / 2, -half / 2 + head * 0.7);
  ctx.lineTo(-head / 2, -half / 2 + head * 0.7);
  ctx.closePath();
  ctx.fillStyle = fill;
  ctx.fill();
  ctx.lineWidth = 1;
  ctx.strokeStyle = edge;
  ctx.stroke();
  ctx.restore();
}

/** A pennant with a notch, which is Excel's flag shape. */
function flag(
  ctx: CanvasRenderingContext2D,
  cx: number,
  cy: number,
  size: number,
  fill: string,
  edge: string,
): void {
  const half = size * 0.42;
  ctx.beginPath();
  ctx.moveTo(cx - half, cy - half);
  ctx.lineTo(cx + half, cy - half * 0.4);
  ctx.lineTo(cx + half * 0.2, cy);
  ctx.lineTo(cx + half, cy + half * 0.4);
  ctx.lineTo(cx - half, cy + half);
  ctx.closePath();
  ctx.fillStyle = fill;
  ctx.fill();
  ctx.lineWidth = 1;
  ctx.strokeStyle = edge;
  ctx.stroke();
}

/** The glyph inside a 3Symbols disc: a cross, an exclamation or a tick. */
function symbol(
  ctx: CanvasRenderingContext2D,
  cx: number,
  cy: number,
  size: number,
  level: number,
): void {
  const r = size * 0.18;
  ctx.strokeStyle = INK;
  ctx.lineWidth = 2;
  ctx.lineCap = "round";
  ctx.beginPath();
  if (level === 0) {
    ctx.moveTo(cx - r, cy - r);
    ctx.lineTo(cx + r, cy + r);
    ctx.moveTo(cx + r, cy - r);
    ctx.lineTo(cx - r, cy + r);
  } else if (level === 1) {
    ctx.moveTo(cx, cy - r * 1.4);
    ctx.lineTo(cx, cy + r * 0.2);
    ctx.moveTo(cx, cy + r * 1.3);
    ctx.lineTo(cx, cy + r * 1.4);
  } else {
    ctx.moveTo(cx - r, cy);
    ctx.lineTo(cx - r * 0.2, cy + r * 0.8);
    ctx.lineTo(cx + r, cy - r * 0.8);
  }
  ctx.stroke();
  ctx.lineCap = "butt";
}

/** A column chart with the bars up to the level filled, which is a rating. */
function rating(
  ctx: CanvasRenderingContext2D,
  cx: number,
  cy: number,
  size: number,
  level: number,
  count: number,
): void {
  const gap = size * 0.12;
  const width = (size * 0.86 - gap * (count - 1)) / count;
  const left = cx - (size * 0.86) / 2;
  for (let index = 0; index < count; index += 1) {
    const height = size * (0.3 + (0.6 * (index + 1)) / count);
    ctx.fillStyle = index <= level ? BLUE : EMPTY;
    ctx.fillRect(left + index * (width + gap), cy + size * 0.43 - height, width, height);
  }
}

/** A disc cut into quarters, the filled ones running clockwise from the top. */
function quarters(
  ctx: CanvasRenderingContext2D,
  cx: number,
  cy: number,
  size: number,
  level: number,
): void {
  const radius = size * 0.43;
  const start = -Math.PI / 2;
  for (let index = 0; index < 4; index += 1) {
    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.arc(cx, cy, radius, start + (index * Math.PI) / 2, start + ((index + 1) * Math.PI) / 2);
    ctx.closePath();
    ctx.fillStyle = index <= level ? BLUE : EMPTY;
    ctx.fill();
  }
  disc(ctx, cx, cy, radius, "transparent", BLUE);
}

export function paintConditionalIcon(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  size: number,
  family: string,
  level: number,
): void {
  const cx = x + size / 2;
  const cy = y + size / 2;
  const arrows = /^(\d)Arrows/.exec(family);
  if (arrows) {
    const count = Number(arrows[1]);
    const grey = family.endsWith("Gray");
    // Level 0 points down, the last points up.
    const angle = Math.PI * (1 - level / (count - 1));
    const [fill, edge] = colours(level, count, grey);
    arrow(ctx, cx, cy, size, angle, fill, edge);
    return;
  }
  if (family.startsWith("3TrafficLights") || family.startsWith("4TrafficLights")) {
    const count = family.startsWith("4") ? 4 : 3;
    const [fill, edge] = colours(level, count, false);
    disc(ctx, cx, cy, size * 0.42, fill, edge);
    return;
  }
  if (family === "3Flags") {
    const [fill, edge] = colours(level, 3, false);
    flag(ctx, cx, cy, size, fill, edge);
    return;
  }
  if (family === "3Signs") {
    const [fill, edge] = colours(level, 3, false);
    disc(ctx, cx, cy, size * 0.42, fill, edge);
    return;
  }
  if (family === "3Symbols" || family === "3Symbols2") {
    const [fill, edge] = colours(level, 3, false);
    disc(ctx, cx, cy, size * 0.42, fill, edge);
    symbol(ctx, cx, cy, size, level);
    return;
  }
  if (family === "4Rating" || family === "5Rating") {
    rating(ctx, cx, cy, size, level, family === "4Rating" ? 4 : 5);
    return;
  }
  if (family === "5Quarters") {
    quarters(ctx, cx, cy, size, level);
    return;
  }
  if (family === "4RedToBlack") {
    // From red toward black, one step per level.
    const t = level / 3;
    const mix = (channel: number) => Math.round(channel * (1 - t));
    disc(ctx, cx, cy, size * 0.42, `rgb(${mix(0xd6)}, ${mix(0x55)}, ${mix(0x32)})`, null);
    return;
  }
}
