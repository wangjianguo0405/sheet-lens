/**
 * OOXML preset geometry, evaluated.
 *
 * A drawing only ever *names* a shape's form — `prstGeom prst="curvedConnector3"`
 * plus a handful of adjustment values — so whoever draws it has to carry the
 * table of formulas ECMA-376 defines for each of the 187 presets. Hand-writing
 * them one at a time is how every earlier attempt in this project went wrong:
 * the callout's tail read its adjustments the wrong way round, the connectors
 * were drawn as arcs when they are Béziers, and every offset was scaled off the
 * shape's shorter side when the table scales it off the width and height.
 *
 * `public/preset-shapes.json` is that table — Apache POI's copy of the spec's
 * definitions, trimmed to the guides and paths and stripped of namespaces — and
 * this file turns one entry plus a box into SVG.
 *
 * The formula language is a small prefix notation: an operator followed by a
 * fixed number of operands, each of which may itself be an operator.
 */

interface Preset {
  /** Adjustment values: name and default formula, in the order the file lists them. */
  av: [string, string][];
  /** Named values the path is built from, in dependency order. */
  gd: [string, string][];
  /** The paths themselves. `f`/`s` are the element's own fill and stroke hints. */
  p: { f: string | null; s: string | null; c: (string | number)[][] }[];
}

/** One path of a preset, resolved for a box. */
export interface PresetPath {
  d: string;
  /** False when the table marks the path `fill="none"` — a connector, say. */
  filled: boolean;
  /** False when the table marks the path `stroke="false"`. */
  stroked: boolean;
}

let presets: Record<string, Preset> | null = null;
let pending: Promise<void> | null = null;

/** Fetches the table, once. Safe to call from anywhere; resolves when it is in. */
export function loadPresets(): Promise<void> {
  if (!pending) {
    pending = fetch("preset-shapes.json")
      .then((response) => response.json())
      .then((data: Record<string, Preset>) => {
        presets = data;
      })
      .catch(() => {
        // Without the table every shape falls back to its box, which is a
        // worse drawing but still the right layout.
        presets = {};
      });
  }
  return pending;
}

export function presetsReady(): boolean {
  return presets !== null;
}

/**
 * Angles in these formulas are in 1/60000 of a degree, so a full turn is
 * 21600000 units and one unit is `2 * PI / 21600000` radians.
 */
const ANGLE = (2 * Math.PI) / 21600000;

/** The values every preset starts with, from the spec's predefined formulas. */
function baseGuides(w: number, h: number, vars: Map<string, number>): void {
  const ss = Math.min(w, h);
  const set = (name: string, value: number) => vars.set(name, value);
  set("w", w);
  set("h", h);
  set("l", 0);
  set("t", 0);
  set("r", w);
  set("b", h);
  set("hc", w / 2);
  set("vc", h / 2);
  set("ss", ss);
  set("ls", Math.max(w, h));
  set("wd2", w / 2);
  set("wd3", w / 3);
  set("wd4", w / 4);
  set("wd5", w / 5);
  set("wd6", w / 6);
  set("wd8", w / 8);
  set("wd10", w / 10);
  set("wd32", w / 32);
  set("hd2", h / 2);
  set("hd3", h / 3);
  set("hd4", h / 4);
  set("hd5", h / 5);
  set("hd6", h / 6);
  set("hd8", h / 8);
  set("hd10", h / 10);
  set("hd32", h / 32);
  set("ssd2", ss / 2);
  set("ssd4", ss / 4);
  set("ssd6", ss / 6);
  set("ssd8", ss / 8);
  set("ssd16", ss / 16);
  set("ssd32", ss / 32);
  // The `cd` family is angles, not lengths — `cd2` is a half turn.
  set("cd2", 10800000);
  set("cd4", 5400000);
  set("cd8", 2700000);
  set("3cd4", 16200000);
  set("3cd8", 8100000);
  set("5cd8", 13500000);
  set("7cd8", 18900000);
}

/**
 * Evaluates one formula, returning the value and where the next token starts.
 *
 * The operators, and how many operands each swallows:
 *   one    — `val`, `abs`, `sqrt`
 *   two    — `min`, `max`, `at2`, and `sin`/`cos`/`tan` (a scale, then an angle)
 *   three  — multiply-then-divide (a·b/c), add-then-divide ((a+b)/c),
 *            add-then-subtract (a+b−c), `?:` as a>0?b:c, and `pin` as b clamped
 *            into [a, c]
 *   three  — `mod`, which is the *length* of the vector (a,b,c) rather than a
 *            remainder, and `cat2`/`sat2`, which take a scale and a pair to
 *            take the angle of
 *
 * Their names are the spec's, written without the spaces it uses.
 */
function evaluate(tokens: string[], at: number, vars: Map<string, number>): [number, number] {
  const operator = tokens[at];
  const operand = (index: number): [number, number] => evaluate(tokens, index, vars);
  switch (operator) {
    case "val": {
      const token = tokens[at + 1];
      const value = vars.get(token);
      return [(value ?? Number(token)) || 0, at + 2];
    }
    case "*/": {
      const [a, i] = operand(at + 1);
      const [b, j] = operand(i);
      const [c, k] = operand(j);
      return [(a * b) / c, k];
    }
    case "+/": {
      const [a, i] = operand(at + 1);
      const [b, j] = operand(i);
      const [c, k] = operand(j);
      return [(a + b) / c, k];
    }
    case "+-": {
      const [a, i] = operand(at + 1);
      const [b, j] = operand(i);
      const [c, k] = operand(j);
      return [a + b - c, k];
    }
    case "?:": {
      const [a, i] = operand(at + 1);
      const [b, j] = operand(i);
      const [c, k] = operand(j);
      return [a > 0 ? b : c, k];
    }
    case "pin": {
      const [low, i] = operand(at + 1);
      const [value, j] = operand(i);
      const [high, k] = operand(j);
      return [Math.min(high, Math.max(low, value)), k];
    }
    case "mod": {
      const [a, i] = operand(at + 1);
      const [b, j] = operand(i);
      const [c, k] = operand(j);
      return [Math.sqrt(a * a + b * b + c * c), k];
    }
    case "abs": {
      const [a, i] = operand(at + 1);
      return [Math.abs(a), i];
    }
    case "sqrt": {
      const [a, i] = operand(at + 1);
      return [Math.sqrt(a), i];
    }
    case "min": {
      const [a, i] = operand(at + 1);
      const [b, j] = operand(i);
      return [Math.min(a, b), j];
    }
    case "max": {
      const [a, i] = operand(at + 1);
      const [b, j] = operand(i);
      return [Math.max(a, b), j];
    }
    case "sin":
    case "cos":
    case "tan": {
      const [scale, i] = operand(at + 1);
      const [angle, j] = operand(i);
      const turn = angle * ANGLE;
      const trig = operator === "sin" ? Math.sin(turn) : operator === "cos" ? Math.cos(turn) : Math.tan(turn);
      return [scale * trig, j];
    }
    // An arc tangent: `at2` is the angle of (a, b) in the same 1/60000 units.
    case "at2": {
      const [a, i] = operand(at + 1);
      const [b, j] = operand(i);
      return [Math.atan2(b, a) / ANGLE, j];
    }
    case "cat2":
    case "sat2": {
      const [scale, i] = operand(at + 1);
      const [a, j] = operand(i);
      const [b, k] = operand(j);
      const turn = Math.atan2(b, a);
      return [scale * (operator === "cat2" ? Math.cos(turn) : Math.sin(turn)), k];
    }
    default: {
      // A bare number, or a name already in the guide list.
      const known = vars.get(operator);
      return [(known ?? Number(operator)) || 0, at + 1];
    }
  }
}

function formula(text: string, vars: Map<string, number>): number {
  return evaluate(text.trim().split(/\s+/), 0, vars)[0];
}

/** A point from the table: two formulas, or two plain numbers. */
function point(tokens: string[], at: number, vars: Map<string, number>): [number, number, number] {
  const [x, next] = evaluate(tokens, at, vars);
  const [y, after] = evaluate(tokens, next, vars);
  return [x, y, after];
}

const round = (n: number) => Math.round(n * 1000) / 1000;

/**
 * A preset resolved into SVG paths for a box of `w` by `h`.
 *
 * `adjust` holds the shape's own values in the units the file stores them —
 * hundred-thousandths — falling back to each adjustment's default from the
 * table, and `null` when the preset is unknown or the path uses a command this
 * does not implement, so the caller can draw the box instead.
 */
export function resolvePreset(
  name: string,
  w: number,
  h: number,
  adjust: number[],
): PresetPath[] | null {
  const preset = presets?.[name];
  if (!preset) return null;

  const vars = new Map<string, number>();
  baseGuides(w, h, vars);
  preset.av.forEach(([key, defaultFormula], index) => {
    const given = adjust[index];
    const value = given === undefined ? formula(defaultFormula, vars) : given;
    vars.set(key, value);
  });
  for (const [key, expression] of preset.gd) {
    vars.set(key, formula(expression, vars));
  }

  const out: PresetPath[] = [];
  for (const path of preset.p) {
    let d = "";
    for (const command of path.c) {
      const [kind, ...values] = command as [string, ...(string | number)[]];
      const tokens = values as string[];
      switch (kind) {
        case "m": {
          const [x, y] = point(tokens, 0, vars);
          d += `M${round(x)} ${round(y)}`;
          break;
        }
        case "l": {
          const [x, y] = point(tokens, 0, vars);
          d += `L${round(x)} ${round(y)}`;
          break;
        }
        case "c": {
          const [x1, y1, i] = point(tokens, 0, vars);
          const [x2, y2, j] = point(tokens, i, vars);
          const [x, y] = point(tokens, j, vars);
          d += `C${round(x1)} ${round(y1)} ${round(x2)} ${round(y2)} ${round(x)} ${round(y)}`;
          break;
        }
        case "q": {
          const [x1, y1, i] = point(tokens, 0, vars);
          const [x, y] = point(tokens, i, vars);
          d += `Q${round(x1)} ${round(y1)} ${round(x)} ${round(y)}`;
          break;
        }
        case "a": {
          // `arcTo` names an ellipse's radii and the angles it runs between, and
          // starts from wherever the path already is. Converting means working
          // back to the centre it implies, then out to the far end.
          const rx = formula(tokens[0], vars);
          const ry = formula(tokens[1], vars);
          const start = formula(tokens[2], vars) * ANGLE;
          const sweep = formula(tokens[3], vars) * ANGLE;
          const [cx, cy] = currentPoint(d, vars);
          const centreX = cx - rx * Math.cos(start);
          const centreY = cy - ry * Math.sin(start);
          const endX = centreX + rx * Math.cos(start + sweep);
          const endY = centreY + ry * Math.sin(start + sweep);
          const large = Math.abs(sweep) > Math.PI ? 1 : 0;
          const direction = sweep >= 0 ? 1 : 0;
          d += `A${round(rx)} ${round(ry)} 0 ${large} ${direction} ${round(endX)} ${round(endY)}`;
          break;
        }
        default:
          d += "Z";
          break;
      }
    }
    if (d) {
      out.push({
        // `fill="none"` is how the table says a path is only stroked, `stroke=
        // "false"` how it says it is only filled. The `darken`/`lighten` fills
        // are a shaded version of the shape's own colour, which we take as
        // plain rather than shading.
        d,
        filled: path.f !== "none",
        stroked: path.s !== "false",
      });
    }
  }
  return out.length > 0 ? out : null;
}

/**
 * The end of the path so far, which is where an `arcTo` starts from. Only the
 * last command's endpoint matters, and only the end coordinates of each command
 * type are ever read back.
 */
function currentPoint(d: string, vars: Map<string, number>): [number, number] {
  const match = d.match(/([-\d.]+) ([-\d.]+)$/);
  if (!match) return [vars.get("l") ?? 0, vars.get("t") ?? 0];
  return [Number(match[1]), Number(match[2])];
}
