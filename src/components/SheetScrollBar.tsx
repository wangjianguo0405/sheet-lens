import { useCallback, useEffect, useRef, useState } from "react";

interface SheetScrollBarProps {
  /** The grid's scroller, or null until the grid has mounted. */
  viewport: HTMLDivElement | null;
}

/**
 * The sheet's horizontal scrollbar, at the end of the tab row where Excel keeps
 * it.
 *
 * It is a scrollbar rather than a drawing of one: an element with nothing in it
 * that scrolls, kept where the sheet is — whichever of the two moves, the other
 * follows. So it is the engine's own bar, with the engine's own drag, wheel and
 * keyboard behaviour.
 *
 * Its empty content is sized to give it the *grid's own proportions*, so its
 * thumb is exactly as long as the sheet's own bar draws it — the same bar, in a
 * shorter strip. Sizing it to the grid's range instead, with a track this much
 * shorter, drew a thumb half the length of the one the sheet shows when the tabs
 * are hidden, which reads as broken next to it.
 */
export function SheetScrollBar({ viewport }: SheetScrollBarProps) {
  const mirror = useRef<HTMLDivElement>(null);
  /** How wide the empty content must be for the proportions to match. */
  const [reach, setReach] = useState(0);
  /**
   * How thick the engine's own bar is. A scrollbar sits at the bottom of its
   * box, so the box is made exactly that tall and the row centres it.
   */
  const [thickness, setThickness] = useState(0);
  /** The position this bar was last put at from the grid: its event is an echo. */
  const echoed = useRef(-1);

  /**
   * The bar's reach, worked out from the grid's own metrics rather than read
   * back off the element.
   *
   * The two boxes change at different times — the grid's content grows in the
   * commit that grows the blank area, this one's a render later — so reading
   * both in one go can pair a new range with an old one, which maps a position
   * to somewhere it never was. One expression, used by both directions, cannot
   * disagree with itself.
   */
  const ranges = useCallback(() => {
    const el = viewport;
    const bar = mirror.current;
    if (!el || !bar) return null;
    const gridRange = el.scrollWidth - el.clientWidth;
    const barClient = bar.clientWidth;
    const content = Math.max(
      barClient,
      Math.round((barClient * el.scrollWidth) / el.clientWidth),
    );
    return { gridRange, barRange: content - barClient };
  }, [viewport]);

  const sync = useCallback(() => {
    const el = viewport;
    const bar = mirror.current;
    if (!el || !bar) return;
    const r = ranges();
    if (!r) return;
    setReach(bar.clientWidth + r.barRange);
    setThickness(bar.offsetHeight - bar.clientHeight);

    // The same share of the way along, not the same pixel: the two bars have
    // different reaches and the same proportions. A drag keeps the thumb where
    // a drag keeps it — the browser is holding it — so this only ever nudges it
    // back onto the proportion, by the pixel the two mappings disagree by.
    const want = r.gridRange > 0 ? (el.scrollLeft / r.gridRange) * r.barRange : 0;
    if (Math.abs(bar.scrollLeft - want) > 1) {
      echoed.current = Math.round(want);
      bar.scrollLeft = want;
    }
  }, [ranges, viewport]);

  useEffect(() => {
    const el = viewport;
    if (!el) return;
    sync();
    el.addEventListener("scroll", sync);
    const observer = new ResizeObserver(sync);
    observer.observe(el);
    // Another sheet, another zoom, a filter or the blank area growing all
    // change how far there is to go without either box changing.
    if (mirror.current) observer.observe(mirror.current);
    const content = el.firstElementChild;
    if (content) observer.observe(content);
    return () => {
      el.removeEventListener("scroll", sync);
      observer.disconnect();
    };
  }, [viewport, sync]);

  /**
   * Dragged, rolled or keyed, this is the bar in front of the user: it leads,
   * and the grid follows it to the same share of the way along.
   *
   * Unless the event is our own write coming back. Answering that would move the
   * sheet to wherever a rounded proportion of the bar's position lands — which
   * is a few pixels out at best, and hundreds of them in the frame the blank
   * area grows and the proportions change underneath it. Mid-scroll, that is the
   * scroll being yanked backwards.
   */
  const onScroll = () => {
    const bar = mirror.current;
    if (!bar || !viewport) return;
    if (Math.abs(bar.scrollLeft - echoed.current) <= 1) return;
    const r = ranges();
    if (!r || r.barRange <= 0 || r.gridRange <= 0) return;
    const want = (bar.scrollLeft / r.barRange) * r.gridRange;
    if (Math.abs(viewport.scrollLeft - want) > 1) viewport.scrollLeft = want;
  };

  return (
    <div
      ref={mirror}
      onScroll={onScroll}
      // The bar is the only thing in here, so the box is as tall as the bar and
      // the empty content is a single pixel.
      style={thickness ? { height: thickness } : undefined}
      className="w-full overflow-x-auto overflow-y-hidden"
    >
      <div style={{ width: reach, height: 1 }} />
    </div>
  );
}
