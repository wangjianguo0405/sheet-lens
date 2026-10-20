import { useEffect, useState } from "react";
import { useTranslation } from "react-i18next";

import { ScrollArea } from "@/components/ui/scroll-area";
import { cellRef } from "@/lib/column";
import { sheetLayout, type NoteCell } from "@/lib/excel";

interface NoteListProps {
  path: string;
  sheet: string;
  /** Takes the grid to the note's cell, as Go To does. */
  onGo: (row: number, col: number) => void;
  /**
   * The note the pointer is on, so the grid can show it where its author left
   * it — the same box a hover over the cell draws. Null takes it away.
   */
  onPreview: (note: { row: number; col: number } | null) => void;
}

/**
 * The sheet's notes, in reading order. Read-only: the wording comes from the
 * file's comments part, and this viewer never writes one back.
 *
 * The layout is asked for here rather than passed down from the grid — the
 * backend keeps the parsed sheet, so the second ask costs a lookup and the
 * panel then works whether or not the grid is on this sheet.
 */
export function NoteList({ path, sheet, onGo, onPreview }: NoteListProps) {
  const { t } = useTranslation();
  const [notes, setNotes] = useState<NoteCell[] | null>(null);

  // Leaving the panel, or the sheet, takes any preview down with it.
  useEffect(() => () => onPreview(null), [path, sheet, onPreview]);

  useEffect(() => {
    let cancelled = false;
    setNotes(null);
    sheetLayout(path, sheet)
      .then((layout) => {
        if (cancelled) return;
        const sorted = [...(layout.notes ?? [])].sort(
          (a, b) => a.row - b.row || a.col - b.col,
        );
        setNotes(sorted);
      })
      .catch(() => {
        if (!cancelled) setNotes([]);
      });
    return () => {
      cancelled = true;
    };
  }, [path, sheet]);

  return (
    <div className="flex h-full min-h-0 flex-col border-r border-border bg-sidebar">
      <div className="shrink-0 border-b border-border px-2 py-1.5 text-[11px] text-muted-foreground">
        {t("notes.count", { count: notes?.length ?? 0 })}
      </div>
      <ScrollArea className="min-h-0 flex-1">
        <div className="flex flex-col py-1">
          {notes?.map((note) => {
            const text = (note.text ?? "").trim();
            return (
              <button
                type="button"
                key={`${note.row}:${note.col}`}
                onClick={() => onGo(note.row, note.col)}
                onPointerEnter={() => onPreview({ row: note.row, col: note.col })}
                onPointerLeave={() => onPreview(null)}
                onFocus={() => onPreview({ row: note.row, col: note.col })}
                onBlur={() => onPreview(null)}
                title={text || t("notes.blank")}
                className="flex w-full items-baseline gap-2 px-2 py-1 text-left text-xs hover:bg-accent"
              >
                <span className="w-10 shrink-0 font-mono text-[11px] text-muted-foreground">
                  {cellRef(note.row, note.col)}
                </span>
                <span
                  className={
                    text
                      ? "min-w-0 flex-1 truncate"
                      : "min-w-0 flex-1 truncate text-muted-foreground"
                  }
                >
                  {text || t("notes.blank")}
                </span>
              </button>
            );
          })}
          {notes?.length === 0 && (
            <p className="px-2 py-3 text-[11px] text-muted-foreground">
              {t("notes.empty")}
            </p>
          )}
        </div>
      </ScrollArea>
    </div>
  );
}
