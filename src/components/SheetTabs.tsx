import { useCallback, useEffect, useRef, useState, type ReactNode } from "react";
import { ChevronLeft, ChevronRight, Copy, Ellipsis, Eye, EyeOff } from "lucide-react";
import { useTranslation } from "react-i18next";

import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuRadioGroup,
  DropdownMenuRadioItem,
  DropdownMenuSeparator,
  DropdownMenuSub,
  DropdownMenuSubContent,
  DropdownMenuSubTrigger,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import { Accelerated, acceleratorOf, pickByAccelerator } from "@/components/Accelerated";
import { ColorPalette } from "@/components/ColorPalette";
import { type UnhideAnchor } from "@/components/UnhideDialog";
import { copyText } from "@/lib/clipboard";
import type { SheetInfo } from "@/lib/excel";
import { cn } from "@/lib/utils";

interface SheetTabsProps {
  /** The sheets the strip draws: the workbook's own order, less the hidden. */
  sheets: SheetInfo[];
  active: string;
  onSelect: (name: string) => void;
  /**
   * A click on a tab, with the modifiers that make it a group: Ctrl adds or
   * removes one sheet, Shift takes the run from the anchor to this one.
   */
  onTabClick: (name: string, mods: { ctrl: boolean; shift: boolean }) => void;
  /**
   * The sheets grouped with the active one. The active sheet is always in it;
   * what the group is for is acting on several tabs at once.
   */
  selected: Set<string>;
  /**
   * The colours the reader has given the tabs, by sheet name, over the ones the
   * workbook sets. They are the app's own, like the review marks: nothing is
   * written back, and a colour lasts as long as the tab holding it.
   */
  colors: Map<string, string>;
  /** Sets one sheet's tab colour, or takes the reader's off it when null. */
  onColor: (sheet: string, color: string | null) => void;
  /** Takes a sheet off the strip, for as long as the workbook is open. */
  onHide: (sheet: string) => void;
  /**
   * Opens the dialog that offers the hidden sheets back, told where this menu
   * stood so the dialog can take its place.
   */
  onUnhide: (at: UnhideAnchor | null) => void;
  /** How many sheets are off the strip, which is what Unhide has to offer. */
  hiddenCount: number;
  /**
   * A tab dragged to another place in the strip: both are indices into
   * `sheets`, which is what is on the strip and not the workbook's own list —
   * the hidden sheets are not there to be dragged past.
   */
  onReorder: (from: number, to: number) => void;
  /** Drawn at the right end of the row: the sheet's own scrollbar. */
  trailing?: ReactNode;
}

/** How far a press has to travel before it is a drag rather than a click. */
const DRAG_SLOP = 4;

/**
 * Where a menu stands, as the point a dialog replacing it hangs from: its left
 * edge, and its foot measured up from the window's own foot — the two the
 * dialog's fixed positioning wants.
 */
function anchorOf(node: HTMLElement | null): UnhideAnchor | null {
  if (!node) return null;
  const box = node.getBoundingClientRect();
  return { left: box.left, bottom: window.innerHeight - box.bottom };
}

/**
 * Excel's sheet name, in whichever of black and white survives the tab's colour —
 * which is what Excel does, and what decides whether a light green tab takes
 * black type or a blue one takes white. The weights are the usual
 * luma ones, so a colour lands on the side its brightness says.
 */
function readableOn(color: string): string {
  const channel = (at: number) => parseInt(color.slice(at, at + 2), 16);
  const luma =
    0.299 * channel(1) + 0.587 * channel(3) + 0.114 * channel(5);
  return luma > 128 ? "#000000" : "#ffffff";
}

/**
 * The active tab is drawn the way Windows draws a folder tab rather than filled
 * like the rest: a pale wash of the sheet's colour, brightest at the top, with
 * Excel's green rule under the name. A sheet with no colour of its own gets the
 * same shape in plain white, which is what makes an uncoloured strip readable.
 */
function activeWash(color: string): string {
  return `linear-gradient(180deg, color-mix(in srgb, ${color} 2%, #fff), color-mix(in srgb, ${color} 45%, #fff))`;
}

/**
 * Excel's sheet strip along the bottom of the sheet area: two arrows to walk
 * the tabs, and the tabs themselves. What sits at the right end is the sheet's
 * scrollbar, not this row's — see `SheetScrollBar`.
 *
 * The arrows move the tabs, not the selection, which is what they do in Excel:
 * stepping the active sheet is Ctrl+PageUp/PageDown, and that scrolls the
 * strip by itself when it has to. The strip keeps its own scrollbar hidden —
 * a press on the tabs is a press on a sheet, and a bar there would be read as
 * the sheet's.
 */
export function SheetTabs({
  sheets,
  active,
  onSelect,
  onTabClick,
  selected,
  colors,
  onColor,
  onHide,
  onUnhide,
  hiddenCount,
  onReorder,
  trailing,
}: SheetTabsProps) {
  const { t } = useTranslation();
  const strip = useRef<HTMLDivElement>(null);
  /** The "…" at each end, measured so they are not counted as tabs that fit. */
  const moreLeft = useRef<HTMLButtonElement>(null);
  const moreRight = useRef<HTMLButtonElement>(null);
  /** Where the tabs are scrolled to, which decides what the arrows can do. */
  const [bar, setBar] = useState({ left: 0, view: 0, total: 0, more: 0 });
  /** The tab a press is holding, where it started, and the gap it is over now. */
  const drag = useRef<{
    name: string;
    from: number;
    x: number;
    moved: boolean;
    gap: number | null;
  } | null>(null);
  const [dragged, setDragged] = useState<string | null>(null);
  /** The gap a drop would land in, and where to draw the line that says so. */
  const [drop, setDrop] = useState<{ gap: number; x: number } | null>(null);

  /**
   * The tab whose name is open over it. A sheet's name belongs to the file, so
   * this is the same field a double-clicked cell opens — one to read and copy
   * from, not one to change.
   */
  const [editing, setEditing] = useState<string | null>(null);
  const editorRef = useRef<HTMLDivElement>(null);
  /**
   * The tab a right-click was made on, and where on it — the menu that opens
   * from there is drawn once for the whole strip rather than once per tab, so
   * this is what says which tab it belongs to and where to put it.
   */
  const [menu, setMenu] = useState<{ sheet: string; x: number; y: number } | null>(
    null,
  );
  /**
   * Where a right-click on the tab arrows was made, since the sheet list opens
   * from the press rather than from the arrow's own box — the same footing the
   * tab's colour menu takes, and the reason both are dropdowns: a context menu
   * may not choose the side it opens on.
   */
  const [listMenu, setListMenu] = useState<{ x: number; y: number } | null>(null);
  /** The tab's own menu, measured to place the dialog Unhide opens. */
  const tabMenu = useRef<HTMLDivElement>(null);

  const sync = useCallback(() => {
    const node = strip.current;
    if (!node) return;
    setBar({
      left: node.scrollLeft,
      view: node.clientWidth,
      total: node.scrollWidth,
      // The dots are laid out beside the strip, so they are width the tabs
      // cannot use. Counting them back here is what keeps them from promising
      // themselves room to exist: without it, tabs that fit perfectly well on
      // their own would count as overflowing for as long as the dots were up.
      more:
        (moreLeft.current?.offsetWidth ?? 0) +
        (moreRight.current?.offsetWidth ?? 0),
    });
  }, []);

  useEffect(() => {
    const el = strip.current;
    if (!el) return;
    const observer = new ResizeObserver(sync);
    observer.observe(el);
    sync();
    return () => observer.disconnect();
  }, [sync]);

  // Another workbook brings tabs of other widths, and the strip's own box does
  // not change when it does — so the observer above never hears about it.
  useEffect(sync, [sheets, sync]);

  /**
   * A sheet the user never saw is a sheet they did not pick, so a switch from
   * elsewhere — Ctrl+PageDown, a search hit, Go To — scrolls its tab into view.
   */
  useEffect(() => {
    const tab = strip.current?.querySelector<HTMLElement>(
      `[data-sheet="${CSS.escape(active)}"]`,
    );
    tab?.scrollIntoView({ block: "nearest", inline: "nearest" });
  }, [active, sheets]);

  /** It belongs to a tab on a strip, and goes when that strip's tabs do. */
  useEffect(() => {
    setEditing(null);
  }, [sheets, active]);

  /**
   * Focused as it opens, with the caret at the end of the name — the footing a
   * double-clicked cell's field starts on, so Ctrl+A reads the whole name from
   * there.
   */
  useEffect(() => {
    const el = editorRef.current;
    if (!editing || !el) return;
    el.focus({ preventScroll: true });
    const range = document.createRange();
    range.selectNodeContents(el);
    range.collapse(false);
    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
  }, [editing]);

  /**
   * And nothing may change it, so that what is read is what the workbook holds.
   * The two are a pair because neither is enough alone: a native `beforeinput`
   * listener refuses every edit the browser proposes, and the observer puts
   * back whatever got in without proposing one — a context menu's Paste makes
   * its edit through `execCommand`, which fires no such event.
   */
  useEffect(() => {
    const el = editorRef.current;
    if (!editing || !el) return;
    const refuse = (event: Event) => event.preventDefault();
    el.addEventListener("beforeinput", refuse);
    const asRead = el.innerHTML;
    const observer = new MutationObserver(() => {
      if (el.innerHTML !== asRead) el.innerHTML = asRead;
    });
    observer.observe(el, { childList: true, characterData: true, subtree: true });
    return () => {
      el.removeEventListener("beforeinput", refuse);
      observer.disconnect();
    };
  }, [editing]);

  /**
   * The name owns the keyboard while it is open, and takes it in the capture
   * phase so that nothing behind it is answering the same press — Ctrl+A would
   * otherwise reach the sheet and select every cell in it while the tab's own
   * name sat there unselected. Selecting the text here rather than leaving it
   * to the browser is what makes that hold whether or not the field happens to
   * be the thing holding the focus.
   *
   * Escape and Enter put it away, as they do on a cell's field.
   */
  useEffect(() => {
    if (!editing) return;
    const onKeyDown = (event: KeyboardEvent) => {
      const field = editorRef.current;
      if (!field) return;
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "a") {
        event.preventDefault();
        event.stopPropagation();
        field.focus({ preventScroll: true });
        const range = document.createRange();
        range.selectNodeContents(field);
        const selection = window.getSelection();
        selection?.removeAllRanges();
        selection?.addRange(range);
        return;
      }
      if (event.key === "Escape" || event.key === "Enter") {
        event.preventDefault();
        event.stopPropagation();
        setEditing(null);
      }
    };
    window.addEventListener("keydown", onKeyDown, true);
    return () => window.removeEventListener("keydown", onKeyDown, true);
  }, [editing]);

  /** A press anywhere else puts it away, the way leaving a cell does in Excel. */
  useEffect(() => {
    if (!editing) return;
    const onDown = (event: PointerEvent) => {
      if (!editorRef.current?.contains(event.target as Node)) setEditing(null);
    };
    document.addEventListener("pointerdown", onDown, true);
    return () => document.removeEventListener("pointerdown", onDown, true);
  }, [editing]);

  /**
   * Where a drop would land: the gap between two tabs, and where that gap sits
   * in the strip's own scroll coordinates so the marker can be drawn on it.
   *
   * A gap is *before* the first tab whose middle is still to the right of the
   * pointer, which is the rule the strip uses anywhere a pointer picks a place.
   * The last gap is past the final tab.
   */
  const gapAt = useCallback((clientX: number) => {
    const el = strip.current;
    if (!el) return null;
    const tabs = [...el.querySelectorAll<HTMLElement>("[data-sheet]")];
    let gap = tabs.length;
    for (let i = 0; i < tabs.length; i += 1) {
      const rect = tabs[i].getBoundingClientRect();
      if (clientX < rect.left + rect.width / 2) {
        gap = i;
        break;
      }
    }
    const box = el.getBoundingClientRect();
    const edge = tabs[gap]
      ? tabs[gap].getBoundingClientRect().left
      : tabs.length > 0
        ? tabs[tabs.length - 1].getBoundingClientRect().right
        : box.left;
    return { gap, x: edge - box.left + el.scrollLeft };
  }, []);

  /**
   * A press on a tab, held until it moves, then carried to a gap and dropped
   * there — Excel's own gesture for a sheet tab, which is the one place in a
   * workbook where the strip's order is the reader's to set.
   *
   * The strip does not reorder under the pointer: the tabs stay where they are
   * and a line marks the gap the sheet would land in, so the drop can be called
   * off by letting go somewhere else. The order it leaves is the reader's while
   * the workbook is open, and nothing is written back.
   *
   * The listeners live on the window, not on the tab: a drag leaves the tab
   * immediately, and the pointer has to keep being read over its neighbours —
   * which are the gaps it is choosing between.
   */
  useEffect(() => {
    const onMove = (event: PointerEvent) => {
      const held = drag.current;
      if (!held) return;
      // A press is not a drag until it has travelled. A tab that took the
      // dragged look for every click would flicker on the way to somewhere.
      if (!held.moved) {
        if (Math.abs(event.clientX - held.x) < DRAG_SLOP) return;
        held.moved = true;
        setDragged(held.name);
      }
      const gap = gapAt(event.clientX);
      if (!gap) return;
      held.gap = gap.gap;
      setDrop(gap);
    };
    const onUp = () => {
      const held = drag.current;
      drag.current = null;
      setDragged(null);
      setDrop(null);
      if (!held || !held.moved || held.gap === null) return;
      // A gap counts the tabs as they stand, the dragged one included, so a
      // drop to its own right lands one place short of it.
      const to = held.gap > held.from ? held.gap - 1 : held.gap;
      if (to !== held.from) onReorder(held.from, to);
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
    window.addEventListener("pointercancel", onUp);
    return () => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      window.removeEventListener("pointercancel", onUp);
    };
  }, [onReorder, gapAt]);

  /** The tabs' left edges in scroll coordinates, which the arrows step between. */
  const edges = useCallback(() => {
    const el = strip.current;
    if (!el) return [];
    const origin = el.getBoundingClientRect().left - el.scrollLeft;
    return [...el.querySelectorAll<HTMLElement>("[data-sheet]")].map((tab) =>
      Math.round(tab.getBoundingClientRect().left - origin),
    );
  }, []);

  // Measured against the strip *plus* the room the dots are taking, so that
  // what this answers is "do the tabs fit without them".
  const overflowing = bar.total > bar.view + bar.more + 1;
  const atStart = bar.left <= 0;
  const atEnd = bar.left + bar.view >= bar.total - 1;

  const step = useCallback(
    (direction: 1 | -1) => {
      const node = strip.current;
      if (!node) return;
      const stops = edges();
      // The nearest tab edge strictly off to that side, so a press always
      // moves the strip by at least one tab rather than a nudge.
      const target =
        direction > 0
          ? (stops.find((edge) => edge > node.scrollLeft + 1) ?? node.scrollWidth)
          : ([...stops].reverse().find((edge) => edge < node.scrollLeft - 1) ?? 0);
      node.scrollTo({ left: target, behavior: "smooth" });
    },
    [edges],
  );

  /** Ctrl+click on an arrow: the whole way to that end of the tabs, which is
   *  what Excel's own help gives as the way to reach the first or last sheet. */
  const jump = useCallback((direction: 1 | -1) => {
    const node = strip.current;
    if (!node) return;
    node.scrollTo({
      left: direction < 0 ? 0 : node.scrollWidth,
      behavior: "smooth",
    });
  }, []);

  /**
   * The sheet the "…" at that end belongs to: the one just past the tabs that
   * are wholly on screen, which is what the dots there are the clue about.
   *
   * Pressing it makes that sheet current and brings it into view — it does not
   * merely scroll — so pressing it again walks the selection one sheet further
   * in the same direction, and the dots on that side go out once the first (or
   * last) sheet is wholly visible.
   */
  const reveal = useCallback(
    (direction: 1 | -1) => {
      const node = strip.current;
      if (!node) return;
      const box = node.getBoundingClientRect();
      const tabs = [...node.querySelectorAll<HTMLElement>("[data-sheet]")];
      const whole = tabs.filter((tab) => {
        const rect = tab.getBoundingClientRect();
        return rect.left >= box.left - 1 && rect.right <= box.right + 1;
      });
      const edge = direction < 0 ? whole[0] : whole[whole.length - 1];
      if (!edge) return;
      const name = tabs[tabs.indexOf(edge) + direction]?.dataset.sheet;
      if (name) onSelect(name);
    },
    [onSelect],
  );

  const arrowClass = (enabled: boolean) =>
    cn(
      // A disabled arrow still belongs to the right-click list, so it must not
      // hold the pointer: the press has to reach the box around it.
      "flex items-center border-border px-1.5 disabled:pointer-events-none",
      enabled
        ? "text-muted-foreground hover:bg-accent/40 hover:text-foreground"
        : "text-muted-foreground/30",
    );

  return (
    <div
      data-sheet-strip
      className="flex shrink-0 items-stretch border-t border-border bg-muted/40"
    >
      {/*
        Excel's sheet list hangs off a right-click on the tab arrows: with more
        sheets than fit, it is how you reach one without walking the arrows. The
        press opens it, and the menu is drawn further down with the tab's colour
        menu — both hang off the pointer, which a context menu cannot do.
      */}
      <div
        className="flex shrink-0 items-stretch"
        onContextMenu={(event) => {
          event.preventDefault();
          setListMenu({ x: event.clientX, y: event.clientY });
        }}
      >
        <button
          type="button"
          aria-label={t("sheets.scrollLeft")}
          title={t("sheets.scrollLeft")}
          disabled={!overflowing || atStart}
          onClick={(event) =>
            event.ctrlKey || event.metaKey ? jump(-1) : step(-1)
          }
          className={cn(arrowClass(overflowing && !atStart), "border-r")}
        >
          <ChevronLeft className="size-3.5" />
        </button>
        <button
          type="button"
          aria-label={t("sheets.scrollRight")}
          title={t("sheets.scrollRight")}
          disabled={!overflowing || atEnd}
          onClick={(event) =>
            event.ctrlKey || event.metaKey ? jump(1) : step(1)
          }
          className={cn(arrowClass(overflowing && !atEnd), "border-r")}
        >
          <ChevronRight className="size-3.5" />
        </button>
      </div>

      {/*
        Excel's "…": with tabs hidden past the end of the strip, it puts three
        dots at that end as the clue that there are more, and pressing them
        makes the sheet just beyond what is on screen current and scrolls it
        into view — see `reveal`. Both ends do that, so pressing one again walks
        the selection one sheet further that way.

        Each end keeps its slot even when nothing is hidden that way, and then
        draws *nothing*: the gap stays but the dots go. Excel's own screenshots
        of a tab bar at its first sheet show exactly that, and it is the right
        shape for the row besides — a slot that came and went would shift every
        tab along by its width the moment the strip reached that end.

        It stands beside the strip rather than over it, so the dots take the
        room the tabs then cannot: that is how Excel lays it out too, and it
        keeps the tabs scrolling under nothing.
      */}
      {overflowing && (
        <button
          ref={moreLeft}
          type="button"
          aria-label={t("sheets.moreLeft")}
          title={t("sheets.moreLeft")}
          disabled={atStart}
          onClick={() => reveal(-1)}
          className={cn(arrowClass(!atStart), "border-r px-1")}
        >
          <Ellipsis className={cn("size-3.5", atStart && "invisible")} />
        </button>
      )}

      <div
        ref={strip}
        onScroll={sync}
        className="relative flex min-w-24 flex-1 items-stretch overflow-x-auto [scrollbar-width:none] [&::-webkit-scrollbar]:hidden"
      >
        {/*
          Where the tab being carried would land. The strip itself does not
          move — Excel keeps the tabs where they are and draws a line at the gap
          the sheet would drop into, which is what makes letting go somewhere
          else the way to call the drag off.
        */}
        {drop && (
          <span
            aria-hidden
            className="pointer-events-none absolute inset-y-0 z-10 w-0.5 bg-[#107C41]"
            style={{ left: drop.x - 1 }}
          />
        )}
        {sheets.map((sheet, index) => {
          const isActive = sheet.name === active;
          const isSelected = selected.has(sheet.name);
          const title = t("sheets.tabTitle", {
            name: sheet.name,
            rows: sheet.rows,
            cols: sheet.cols,
          });
          // The reader's colour over the workbook's, which is what the tab is
          // drawn in and what the palette below rings.
          const color = colors.get(sheet.name) ?? sheet.tabColor ?? null;
          const className = cn(
            "shrink-0 border-r border-b-2 border-border px-3 py-1 text-xs whitespace-nowrap transition select-none",
            isActive
              ? "border-b-[#107C41] bg-background font-medium text-foreground"
              : isSelected
                // Grouped with the active sheet: Excel keeps them looking
                // selected, without the raised face the current one has.
                ? "border-b-[#107C41]/40 bg-background text-foreground"
                : "border-b-transparent text-muted-foreground hover:bg-accent/40 hover:text-foreground",
            // A tab filled with its own colour has the grey the others take
            // on hover sitting behind an inline fill that wins, so it
            // lightens instead — the same "this one is under the pointer".
            !isActive && color && "hover:brightness-110",
            dragged === sheet.name && "cursor-grabbing opacity-70",
          );
          // A tab that has a colour is filled with it, name and all — Excel
          // paints the whole tab, not a strip along its foot. The current sheet
          // keeps the raised tab instead, washed with that same colour, so it
          // still says which sheet it is without going solid like the others.
          const style = color
            ? isActive
              ? { backgroundImage: activeWash(color) }
              : {
                  backgroundColor: color,
                  color: readableOn(color),
                }
            : undefined;

          /**
           * The name, opened over the tab by a double-click, as a cell's text
           * is opened over its cell. A ring rather than a border: the strip's
           * tabs are laid out edge to edge, and a border would widen this one
           * and shift every tab after it the moment the field opened.
           */
          if (sheet.name === editing) {
            return (
              <div
                key={sheet.name}
                ref={editorRef}
                data-sheet={sheet.name}
                title={title}
                contentEditable
                suppressContentEditableWarning
                spellCheck={false}
                onPaste={(event) => event.preventDefault()}
                onDrop={(event) => event.preventDefault()}
                className={cn(
                  className,
                  "cursor-text outline-none select-text ring-1 ring-ring ring-inset",
                )}
                style={style}
              >
                {sheet.name}
              </div>
            );
          }

          return (
            <button
              key={sheet.name}
              type="button"
              data-sheet={sheet.name}
              onPointerDown={(event) => {
                if (event.button !== 0) return;
                drag.current = {
                  name: sheet.name,
                  from: index,
                  x: event.clientX,
                  moved: false,
                  gap: null,
                };
              }}
              onClick={(event) =>
                onTabClick(sheet.name, {
                  ctrl: event.ctrlKey || event.metaKey,
                  shift: event.shiftKey,
                })
              }
              // The second press is the one that opens the name, and the first
              // has selected the sheet by then — which is what Excel does with a
              // tab that is not the current one.
              onDoubleClick={() => setEditing(sheet.name)}
              // The colour menu, opened from where the press was rather than
              // from the tab's own box: see the strip's menu below.
              onContextMenu={(event) => {
                event.preventDefault();
                setMenu({ sheet: sheet.name, x: event.clientX, y: event.clientY });
              }}
              title={title}
              className={className}
              style={style}
            >
              {sheet.name}
            </button>
          );
        })}
      </div>

      {overflowing && (
        <button
          ref={moreRight}
          type="button"
          aria-label={t("sheets.moreRight")}
          title={t("sheets.moreRight")}
          disabled={atEnd}
          onClick={() => reveal(1)}
          className={cn(arrowClass(!atEnd), "border-l px-1")}
        >
          <Ellipsis className={cn("size-3.5", atEnd && "invisible")} />
        </button>
      )}

      {/*
        The sheet list a right-click on the arrows opens, hung off the pointer
        exactly as the tab's own menu is — and for the same reason: a context
        menu always opens to the right of whatever it is attached to, and this
        belongs at the press. It opens *downwards* from the point, so the
        pointer sits at its top-left corner; at the foot of a window there is
        nowhere below to go, so it flips up over the sheet.
      */}
      <DropdownMenu
        open={listMenu !== null}
        onOpenChange={(open) => {
          if (!open) setListMenu(null);
        }}
      >
        <DropdownMenuTrigger asChild>
          <span
            aria-hidden
            className="pointer-events-none fixed size-px"
            style={{ left: listMenu?.x ?? 0, top: listMenu?.y ?? 0 }}
          />
        </DropdownMenuTrigger>
        <DropdownMenuContent
          side="bottom"
          align="start"
          sideOffset={3}
          className="max-h-80 min-w-40"
        >
          <DropdownMenuLabel className="text-[10px]">
            {t("sheets.list")}
          </DropdownMenuLabel>
          <DropdownMenuRadioGroup value={active} onValueChange={onSelect}>
            {sheets.map((sheet) => (
              <DropdownMenuRadioItem
                key={sheet.name}
                value={sheet.name}
                className="text-xs"
              >
                <span className="truncate">{sheet.name}</span>
              </DropdownMenuRadioItem>
            ))}
          </DropdownMenuRadioGroup>
        </DropdownMenuContent>
      </DropdownMenu>

      {/*
        Excel's right-click on a tab, which for this app holds the one thing it
        can do: colour the tab. Nothing is written back, so the colour is the
        reader's own — and it lands on the tab that was pointed at rather than
        moving the selection, which a right-click on a tab does not do in Excel
        either.

        A menu of its own rather than the strip's shared context menu, because a
        context menu may not choose the side it opens on — Radix anchors one to
        the pointer and puts it to the right, and this one belongs *above* the
        pointer, the way Excel opens a tab's menu out over the sheet. So the
        anchor is a hair's-breadth box placed where the press was, and the menu
        hangs off that with the corners set here: its bottom-left corner sits
        just above and to the right of the pointer.
      */}
      <DropdownMenu
        open={menu !== null}
        onOpenChange={(open) => {
          if (!open) setMenu(null);
        }}
      >
        <DropdownMenuTrigger asChild>
          <span
            aria-hidden
            className="pointer-events-none fixed size-px"
            // The few pixels to the right are the anchor's own: the menu's left
            // edge lands on it, and an align offset is not something this menu
            // answers to.
            style={{ left: (menu?.x ?? 0) + 4, top: menu?.y ?? 0 }}
          />
        </DropdownMenuTrigger>
        <DropdownMenuContent
          ref={tabMenu}
          side="top"
          align="start"
          sideOffset={3}
          className="min-w-44"
          onKeyDown={pickByAccelerator}
        >
          {/*
            Excel's two sheet commands, above the colour, and in its order:
            Unhide is greyed until there is something to bring back, and Hide
            until there is a sheet it could take away — hiding the only one
            still on the strip would leave the workbook with nothing to show,
            which Excel refuses outright.
          */}
          <DropdownMenuItem
            data-accelerator={acceleratorOf(t("sheets.unhide"))}
            disabled={hiddenCount === 0}
            icon={<Eye />}
            onSelect={() => onUnhide(anchorOf(tabMenu.current))}
            className="text-xs"
          >
            <Accelerated label={t("sheets.unhide")} />
          </DropdownMenuItem>
          <DropdownMenuItem
            data-accelerator={acceleratorOf(t("sheets.hide"))}
            // Hiding everything on the strip would leave nothing to show, so
            // that is exactly when the command greys out.
            disabled={
              sheets.length <= (menu && selected.has(menu.sheet) ? selected.size : 1)
            }
            icon={<EyeOff />}
            onSelect={() => {
              if (menu !== null) onHide(menu.sheet);
            }}
            className="text-xs"
          >
            <Accelerated label={t("sheets.hide")} />
          </DropdownMenuItem>
          <DropdownMenuSeparator />
          {/*
            The name belongs to the file and is the one thing on a tab worth
            having somewhere else — a bug report, a chat message — so it goes
            to the clipboard as it stands, the right-clicked tab's wherever
            that tab is in the strip.
          */}
          <DropdownMenuItem
            data-accelerator={acceleratorOf(t("sheets.copyName"))}
            icon={<Copy />}
            onSelect={() => {
              if (menu !== null) void copyText(menu.sheet);
            }}
            className="text-xs"
          >
            <Accelerated label={t("sheets.copyName")} />
          </DropdownMenuItem>
          <DropdownMenuSub>
            <DropdownMenuSubTrigger className="text-xs">
              {t("sheets.tabColor")}
            </DropdownMenuSubTrigger>
            <DropdownMenuSubContent>
              <ColorPalette
                variant="dropdown"
                current={
                  menu === null
                    ? null
                    : (colors.get(menu.sheet) ??
                      sheets.find((one) => one.name === menu.sheet)?.tabColor ??
                      null)
                }
                clearLabel={t("sheets.tabColorNone")}
                onPick={(fill) => {
                  if (menu !== null) onColor(menu.sheet, fill);
                }}
              />
            </DropdownMenuSubContent>
          </DropdownMenuSub>
        </DropdownMenuContent>
      </DropdownMenu>

      {trailing}
    </div>
  );
}
