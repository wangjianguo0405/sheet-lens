import { Check, ChevronDown, Plus, X } from "lucide-react";
import { useCallback, useEffect, useLayoutEffect, useRef, useState, type MouseEvent as ReactMouseEvent } from "react";
import { useTranslation } from "react-i18next";

import {
  ContextMenu,
  ContextMenuContent,
  ContextMenuItem,
  ContextMenuSeparator,
  ContextMenuTrigger,
} from "@/components/ui/context-menu";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import { copyText } from "@/lib/clipboard";
import { cn } from "@/lib/utils";

/** How far a press has to travel before it is a drag rather than a click. */
const DRAG_SLOP = 4;

/**
 * Middle-click closes a tab, the gesture every browser puts on one. It listens
 * for `auxclick` rather than `mousedown` so that a left press still selects the
 * tab on release, and it takes the whole tab rather than just its × — the name
 * is the part being aimed at.
 */
function middleClose(close: () => void) {
  return (event: ReactMouseEvent) => {
    if (event.button !== 1) return;
    close();
  };
}

export interface FileTab {
  path: string;
  fileName: string;
}

interface FileTabsProps {
  tabs: FileTab[];
  active: string | null;
  onSelect: (path: string) => void;
  onClose: (path: string) => void;
  /** Reads a workbook again from disk, for a file that has changed under us. */
  onReload: (path: string) => void;
  /** A tab dragged to another place in the strip: both are indices into `tabs`. */
  onReorder: (from: number, to: number) => void;
  /** Whether the start page has a tab of its own at the moment. */
  homeOpen: boolean;
  /** Whether that tab is the one showing. */
  atHome: boolean;
  /** Show the start page. */
  onHome: () => void;
  /** Close its tab, which goes back to the workbook that was showing. */
  onCloseHome: () => void;
  /**
   * The "+" at the end of the strip. It opens the picker when the start page is
   * already showing, and opens the start page otherwise.
   */
  onNew: () => void;
}

export function FileTabs({
  tabs,
  active,
  onSelect,
  onClose,
  onReload,
  onReorder,
  homeOpen,
  atHome,
  onHome,
  onCloseHome,
  onNew,
}: FileTabsProps) {
  const { t } = useTranslation();
  const strip = useRef<HTMLDivElement>(null);
  /** The ▼ itself, measured so it is not counted as room the tabs have. */
  const more = useRef<HTMLButtonElement>(null);
  /**
   * What the strip is holding against what it can show, which is what decides
   * whether the ▼ is needed. The ▼'s own width is added back on the right —
   * counting it would keep the button up for good once a file was closed, since
   * the room it takes can be the room that would have made the tabs fit.
   */
  const [bar, setBar] = useState({ view: 0, total: 0, more: 0 });
  /** The press that may yet become a drag, and which tab it is holding. */
  const drag = useRef<{ path: string; index: number; x: number; moved: boolean } | null>(
    null,
  );
  const [dragged, setDragged] = useState<string | null>(null);

  const sync = useCallback(() => {
    const node = strip.current;
    if (!node) return;
    setBar({
      view: node.clientWidth,
      total: node.scrollWidth,
      more: more.current?.offsetWidth ?? 0,
    });
  }, []);

  useEffect(() => {
    const node = strip.current;
    if (!node) return;
    const observer = new ResizeObserver(sync);
    observer.observe(node);
    sync();
    return () => observer.disconnect();
  }, [sync]);

  // Another workbook brings a tab of another width, and the strip's own box does
  // not change when it does — so the observer above never hears about it.
  useEffect(sync, [tabs, homeOpen, sync]);

  const overflowing = bar.total > bar.view + bar.more + 1;

  /**
   * A tab that has just been added is the one being looked for, and it arrives
   * at the end of the strip — past the edge whenever the strip is full. So the
   * strip scrolls it in: pressing "+", or opening a workbook any other way,
   * shows the new tab rather than leaving it beyond the last one on screen.
   *
   * Laid out rather than painted, because scrolling after the frame is drawn is
   * a frame with the tab still hidden — which is the whole thing being fixed.
   */
  const had = useRef({ files: 0, home: false });
  useLayoutEffect(() => {
    const grew = tabs.length > had.current.files || (homeOpen && !had.current.home);
    had.current = { files: tabs.length, home: homeOpen };
    if (!grew) return;
    // Whichever kind of tab it is, the new one is the last in the strip.
    strip.current?.lastElementChild?.scrollIntoView({ block: "nearest", inline: "end" });
  }, [tabs.length, homeOpen]);

  /**
   * The tab the pointer is over, as a place in the strip. Each tab is passed
   * once the pointer is past its middle, which is what makes the one being
   * dragged change places with the one under it rather than only at its edge.
   */
  const placeAt = (clientX: number) => {
    const nodes = [
      ...(strip.current?.querySelectorAll<HTMLElement>("[data-file-tab]") ?? []),
    ];
    let place = 0;
    nodes.forEach((node, index) => {
      const rect = node.getBoundingClientRect();
      if (clientX >= rect.left + rect.width / 2) place = index;
    });
    return place;
  };

  /**
   * The listeners live on the window, not on the tab: a drag leaves the tab
   * immediately, and the pointer has to keep being read over its neighbours —
   * which are the tabs it is changing places with.
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
        setDragged(held.path);
      }
      const place = placeAt(event.clientX);
      if (place === held.index) return;
      onReorder(held.index, place);
      // The strip has been reordered under the pointer, so the tab being
      // dragged is at `place` now and the next move is measured from there.
      held.index = place;
    };
    const onUp = () => {
      if (!drag.current) return;
      drag.current = null;
      setDragged(null);
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
    window.addEventListener("pointercancel", onUp);
    return () => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      window.removeEventListener("pointercancel", onUp);
    };
  }, [onReorder]);

  /**
   * A press on a tab, held until it moves. The × is the one part of a tab that
   * is not a handle: it closes, so a press on it is left to do that.
   */
  const holdTab = (event: React.PointerEvent, path: string, index: number) => {
    if (event.button !== 0) return;
    if ((event.target as HTMLElement).closest("[data-tab-close]")) return;
    drag.current = { path, index, x: event.clientX, moved: false };
  };

  if (tabs.length === 0 && !homeOpen) return null;

  return (
    <div className="flex shrink-0 items-stretch border-b border-border bg-muted/40">
      {/*
        The tabs alone scroll; what stands at either end of them does not.

        A middle press is held back from the browser, which would otherwise read
        it as the start of its own auto-scroll — and that only bites once the
        strip is full, which is exactly when a middle press is worth having: it
        is what closes a tab here. Taking the press leaves the `auxclick` that
        follows to do the closing.
      */}
      <div
        ref={strip}
        onScroll={sync}
        onMouseDown={(event) => {
          if (event.button === 1) event.preventDefault();
        }}
        className="flex min-w-0 flex-1 items-stretch overflow-x-auto"
      >
        {tabs.map((tab, index) => {
          const isActive = tab.path === active;
          return (
            <ContextMenu key={tab.path}>
              <ContextMenuTrigger asChild>
                <div
                  data-file-tab={tab.path}
                  onPointerDown={(event) => holdTab(event, tab.path, index)}
                  className={cn(
                    "group flex max-w-64 shrink-0 items-center gap-1 border-r border-border px-2 py-1 text-xs select-none",
                    isActive
                      ? "bg-background text-foreground"
                      : "text-muted-foreground hover:bg-accent/40",
                    dragged === tab.path && "cursor-grabbing opacity-70",
                  )}
                  onAuxClick={middleClose(() => onClose(tab.path))}
                >
                  <button
                    type="button"
                    className="truncate"
                    title={tab.path}
                    onClick={() => onSelect(tab.path)}
                  >
                    {tab.fileName}
                  </button>
                  <button
                    type="button"
                    data-tab-close
                    aria-label={t("tabs.close", { name: tab.fileName })}
                    className={cn(
                      "rounded-sm p-0.5 transition-opacity hover:bg-accent",
                      isActive ? "opacity-60" : "opacity-0 group-hover:opacity-60",
                    )}
                    onClick={() => onClose(tab.path)}
                  >
                    <X className="size-3" />
                  </button>
                </div>
              </ContextMenuTrigger>
              {/* The name and the path are the two things worth having somewhere
                  else — a bug report, a chat message — and neither is easy to get
                  at otherwise, since the tab shows only the name and truncates
                  even that. Above them sits the one thing a tab can *do*: read its
                  file again, for the workbook that changed underneath us — which
                  Excel being open on it is the usual cause of. */}
              <ContextMenuContent className="min-w-44">
                <ContextMenuItem
                  className="text-xs"
                  onSelect={() => onReload(tab.path)}
                >
                  {t("tabs.reload")}
                </ContextMenuItem>
                <ContextMenuSeparator />
                <ContextMenuItem
                  className="text-xs"
                  onSelect={() => void copyText(tab.fileName)}
                >
                  {t("tabs.copyName")}
                </ContextMenuItem>
                <ContextMenuItem
                  className="text-xs"
                  onSelect={() => void copyText(tab.path)}
                >
                  {t("tabs.copyPath")}
                </ContextMenuItem>
              </ContextMenuContent>
            </ContextMenu>
          );
        })}

        {/*
          The start page, as a tab of its own. It exists only once the "+" has
          been pressed, and its own × puts the workbook that was showing back.
        */}
        {homeOpen && (
          <div
            className={cn(
              "group flex shrink-0 items-center gap-1 border-r border-border px-2 py-1 text-xs",
              atHome ? "bg-background text-foreground" : "text-muted-foreground hover:bg-accent/40",
            )}
            onAuxClick={middleClose(onCloseHome)}
          >
            <button type="button" className="truncate" onClick={onHome}>
              {t("tabs.new")}
            </button>
            <button
              type="button"
              aria-label={t("tabs.close", { name: t("tabs.new") })}
              className={cn(
                "rounded-sm p-0.5 transition-opacity hover:bg-accent",
                atHome ? "opacity-60" : "opacity-0 group-hover:opacity-60",
              )}
              onClick={onCloseHome}
            >
              <X className="size-3" />
            </button>
          </div>
        )}

      </div>

      {/*
        What is open, listed. The tabs that no longer fit are still open files,
        and this is how one is reached without walking the strip to the end —
        Chrome and Firefox keep the same chevron at the end of their tab strips,
        and it appears on the same condition: the moment the tabs stop fitting.
        It stands beside the strip rather than over it, so nothing scrolls under
        it and no tab is ever half behind it.
      */}
      {overflowing && (
        <DropdownMenu>
          <DropdownMenuTrigger asChild>
            <button
              ref={more}
              type="button"
              aria-label={t("tabs.list")}
              title={t("tabs.list")}
              className="flex shrink-0 items-center border-l border-border px-2 text-muted-foreground hover:bg-accent/40 hover:text-foreground"
            >
              <ChevronDown className="size-3.5" />
            </button>
          </DropdownMenuTrigger>
          <DropdownMenuContent align="end" className="min-w-56">
            <DropdownMenuLabel className="text-[10px]">
              {t("tabs.list")}
            </DropdownMenuLabel>
            {tabs.map((tab) => (
              <Row
                key={tab.path}
                label={tab.fileName}
                title={tab.path}
                current={tab.path === active}
                onSelect={() => onSelect(tab.path)}
                onClose={() => onClose(tab.path)}
                closeLabel={t("tabs.close", { name: tab.fileName })}
              />
            ))}
            {homeOpen && (
              <Row
                label={t("tabs.new")}
                current={atHome}
                onSelect={onHome}
                onClose={onCloseHome}
                closeLabel={t("tabs.close", { name: t("tabs.new") })}
              />
            )}
          </DropdownMenuContent>
        </DropdownMenu>
      )}

      <button
        type="button"
        aria-label={t("tabs.openAnother")}
        title={t("tabs.openAnother")}
        onClick={onNew}
        className="flex shrink-0 items-center border-l border-border px-2 text-muted-foreground hover:bg-accent/40 hover:text-foreground"
      >
        <Plus className="size-3.5" />
      </button>
    </div>
  );
}

/**
 * One line of the open-files list: what it is, whether it is the one showing,
 * and a way to close it.
 *
 * The × is a button inside the item, so its press has to be kept from reaching
 * the item — otherwise closing a file would also switch to it and shut the list,
 * which is the opposite of what a list you are tidying up should do.
 */
function Row({
  label,
  title,
  current,
  onSelect,
  onClose,
  closeLabel,
}: {
  label: string;
  title?: string;
  current: boolean;
  onSelect: () => void;
  onClose: () => void;
  closeLabel: string;
}) {
  const keep = (event: ReactMouseEvent) => event.stopPropagation();
  return (
    <DropdownMenuItem className="text-xs" onSelect={onSelect}>
      <Check className={cn("size-3.5 shrink-0", !current && "invisible")} />
      <span className="min-w-0 flex-1 truncate" title={title ?? label}>
        {label}
      </span>
      <button
        type="button"
        aria-label={closeLabel}
        title={closeLabel}
        onPointerDown={keep}
        onPointerUp={keep}
        onClick={(event) => {
          keep(event);
          onClose();
        }}
        className="shrink-0 rounded-sm p-0.5 text-muted-foreground hover:bg-accent hover:text-foreground"
      >
        <X className="size-3" />
      </button>
    </DropdownMenuItem>
  );
}
