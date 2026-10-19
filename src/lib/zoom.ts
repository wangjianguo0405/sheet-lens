/**
 * Excel's zoom steps. Ctrl+wheel and the status bar's slider both move along
 * this ladder rather than by a fixed increment, so the stops people expect —
 * 100% above all — are always reachable exactly.
 */
export const ZOOM_STOPS = [25, 50, 75, 100, 125, 150, 200, 300, 400] as const;

export const MIN_ZOOM = ZOOM_STOPS[0];
export const MAX_ZOOM = ZOOM_STOPS[ZOOM_STOPS.length - 1];
export const DEFAULT_ZOOM = 100;

/** The next stop up (`direction > 0`) or down, staying at the ends. */
export function stepZoom(percent: number, direction: number): number {
  if (direction > 0) {
    for (const stop of ZOOM_STOPS) {
      if (stop > percent) return stop;
    }
    return MAX_ZOOM;
  }
  let below: number = MIN_ZOOM;
  for (const stop of ZOOM_STOPS) {
    if (stop < percent) below = stop;
  }
  return below;
}

/** Where a percentage sits on the ladder, which is what the slider indexes. */
export function zoomIndex(percent: number): number {
  let best = 0;
  ZOOM_STOPS.forEach((stop, index) => {
    if (Math.abs(stop - percent) < Math.abs(ZOOM_STOPS[best] - percent)) {
      best = index;
    }
  });
  return best;
}
