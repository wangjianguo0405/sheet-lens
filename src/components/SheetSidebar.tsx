import { useState } from "react";
import { useTranslation } from "react-i18next";

import { Hint } from "@/components/Hint";
import { NoteList } from "@/components/NoteList";
import { SheetList } from "@/components/SheetList";
import type { UnhideAnchor } from "@/components/UnhideDialog";
import type { SheetInfo } from "@/lib/excel";

export interface SheetSidebarProps {
  path: string;
  sheets: SheetInfo[];
  /** The sheets the reader has hidden, by name. */
  hidden: string[];
  active: string;
  colors: Map<string, string>;
  width: number;
  onSelect: (name: string) => void;
  onColor: (sheet: string, color: string | null) => void;
  onHide: (name: string) => void;
  /** Opens the Unhide dialog where the menu that asked for it stood. */
  onUnhide: (at: UnhideAnchor | null) => void;
  onGo: (row: number, col: number) => void;
  /** Points at a note from its row, so the sheet can show it in place. */
  onPreview: (note: { row: number; col: number } | null) => void;
  onDrag: (event: React.PointerEvent<HTMLDivElement>) => void;
}

/**
 * The left-hand sidebar: the workbook's sheets, or the notes on the sheet on
 * screen, behind a two-tab header. Both read the same grid, so the switch is
 * only which list is under it.
 */
export function SheetSidebar({
  path,
  sheets,
  hidden,
  active,
  colors,
  width,
  onSelect,
  onColor,
  onHide,
  onUnhide,
  onGo,
  onPreview,
  onDrag,
}: SheetSidebarProps) {
  const { t } = useTranslation();
  /** Which panel the left sidebar shows: the sheet list, or the sheet's notes. */
  const [tab, setTab] = useState<"sheets" | "notes">("sheets");

  return (
    <>
      <div className="flex min-h-0 shrink-0 flex-col" style={{ width }}>
        {/* Two panels share the sidebar; the grid is the same either way, so the
            switch costs nothing. */}
        <div className="flex shrink-0 border-b border-border bg-sidebar">
          {(["sheets", "notes"] as const).map((which) => (
            <button
              key={which}
              type="button"
              onClick={() => setTab(which)}
              className={
                tab === which
                  ? "flex-1 border-b-2 border-[#107C41] px-2 py-1 text-[11px] font-medium text-foreground"
                  : "flex-1 border-b-2 border-transparent px-2 py-1 text-[11px] text-muted-foreground hover:text-foreground"
              }
            >
              {which === "sheets" ? t("sheets.list") : t("notes.title")}
            </button>
          ))}
        </div>
        <div className="min-h-0 flex-1">
          {tab === "sheets" ? (
            <SheetList
              path={path}
              sheets={sheets}
              hidden={new Set(hidden)}
              active={active}
              onSelect={onSelect}
              colors={colors}
              onColor={onColor}
              onHide={onHide}
              onUnhide={onUnhide}
            />
          ) : (
            <NoteList path={path} sheet={active} onGo={onGo} onPreview={onPreview} />
          )}
        </div>
      </div>
      <Hint label={t("app.dragToResize")}>
        <div
          role="separator"
          aria-orientation="vertical"
          aria-label={t("app.resizeSheetList")}
          onPointerDown={onDrag}
          className="w-1 shrink-0 cursor-col-resize transition-colors hover:bg-[#107C41]/40"
        />
      </Hint>
    </>
  );
}
