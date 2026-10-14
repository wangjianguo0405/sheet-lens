import { FileSpreadsheet } from "lucide-react";
import { useState } from "react";
import { useTranslation } from "react-i18next";

import { Button } from "@/components/ui/button";
import { fileName } from "@/lib/recent";
import { cn } from "@/lib/utils";

/** How much of the history shows before the rest goes behind "More". */
const HISTORY_SHOWN = 6;

interface WelcomeProps {
  onOpen: () => void;
  /** Newest first, the same list the File menu shows. */
  recentFiles: string[];
  /** Those of them the last attempt to open failed on, shown greyed. */
  missingFiles: string[];
  onOpenRecent: (path: string) => void;
  error: string | null;
}

/**
 * One key, as a browser's own help would draw it: the parts of a chord each in
 * their own box, so that Ctrl+Home reads as two keys pressed together rather
 * than as one long word.
 */
function Keys({ combo }: { combo: string }) {
  return (
    <span className="flex items-center gap-0.5">
      {combo.split("+").map((part) => (
        <kbd key={part} className="rounded-sm bg-muted px-1 text-[10px]">
          {part}
        </kbd>
      ))}
    </span>
  );
}

/**
 * The start page: what to open on the left, and what the keyboard does on the
 * right.
 *
 * The history is here rather than only behind the File menu because an empty
 * window is the one place it is worth the space: the workbook being come back
 * to is usually one of the last few, and picking it out of a list beats walking
 * the file picker again. It is the File menu's list, so anything taken out
 * there goes from here too.
 */
export function Welcome({
  onOpen,
  recentFiles,
  missingFiles,
  onOpenRecent,
  error,
}: WelcomeProps) {
  const { t } = useTranslation();
  const missing = new Set(missingFiles);
  /** Whether the history is showing all of itself or only the newest few. */
  const [fullHistory, setFullHistory] = useState(false);

  /**
   * Every key the app answers to, in the order a reader meets them: the file,
   * then the sheet, then moving around inside it. The sheet's are not usable
   * until something is open, which is what the page is for.
   */
  const shortcuts: { keys: string[]; label: string }[] = [
    { keys: ["Ctrl+O"], label: t("welcome.sc.open") },
    { keys: ["Ctrl+F"], label: t("welcome.sc.search") },
    { keys: ["Ctrl+G"], label: t("welcome.sc.goTo") },
    { keys: ["Ctrl+W"], label: t("welcome.sc.close") },
    { keys: ["Ctrl+A"], label: t("welcome.sc.selectAll") },
    { keys: ["Ctrl+C"], label: t("welcome.sc.copy") },
    { keys: ["←↑↓→"], label: t("welcome.sc.move") },
    { keys: ["Shift+←↑↓→"], label: t("welcome.sc.extend") },
    { keys: ["Ctrl+←↑↓→"], label: t("welcome.sc.edge") },
    { keys: ["PageUp", "PageDown"], label: t("welcome.sc.page") },
    { keys: ["Home", "End"], label: t("welcome.sc.rowEnds") },
    { keys: ["Ctrl+Home", "Ctrl+End"], label: t("welcome.sc.sheetEnds") },
    { keys: ["Ctrl+PageUp", "Ctrl+PageDown"], label: t("welcome.sc.sheetStep") },
    { keys: ["Esc"], label: t("welcome.sc.escape") },
  ];

  return (
    <div className="flex h-full w-full items-center justify-center overflow-auto p-6">
      <div className="flex w-full max-w-3xl flex-col gap-6 md:flex-row md:gap-0">
        {/* What to open */}
        <div className="min-w-0 flex-1 md:pr-6">
          <div className="flex items-center gap-2">
            <FileSpreadsheet className="size-5 text-muted-foreground/60" />
            <h1 className="text-sm font-semibold">SheetLens</h1>
          </div>
          {/* Two lines rather than one paragraph: what the app is, then what
              to do with it — which is the part worth reading twice. */}
          <p className="mt-1 text-xs text-muted-foreground">
            {t("welcome.tagline")}
          </p>
          <p className="text-xs text-muted-foreground">{t("welcome.hint")}</p>

          <Button size="sm" className="mt-4 h-8 text-xs" onClick={onOpen}>
            {t("welcome.open")}
          </Button>

          {error && (
            <p className="mt-3 max-w-md rounded-sm bg-destructive/10 px-2 py-1 text-[11px] text-destructive">
              {error}
            </p>
          )}

          <section className="mt-6">
            <h2 className="text-[10px] font-medium tracking-wide text-muted-foreground uppercase">
              {t("welcome.history")}
            </h2>
            {recentFiles.length === 0 ? (
              <p className="mt-1.5 text-[11px] text-muted-foreground">
                {t("welcome.noHistory")}
              </p>
            ) : (
              <ul className="mt-1 max-h-64 overflow-y-auto">
                {(fullHistory
                  ? recentFiles
                  : recentFiles.slice(0, HISTORY_SHOWN)
                ).map((path) => (
                  <li key={path}>
                    <button
                      type="button"
                      onClick={() => onOpenRecent(path)}
                      // A workbook that would not open is greyed and left in
                      // place, exactly as the File menu leaves it: a click is
                      // what gives it another try.
                      title={
                        missing.has(path)
                          ? t("toolbar.missing", { path })
                          : path
                      }
                      className={cn(
                        "flex w-full items-center rounded-sm px-1.5 py-1 text-left text-xs transition-colors hover:bg-accent/50",
                        missing.has(path)
                          ? "text-muted-foreground"
                          : "text-foreground/80",
                      )}
                    >
                      <span className="truncate">{fileName(path)}</span>
                    </button>
                  </li>
                ))}
                {/* The rest of the list, for a history with more in it than
                    the page has room for — the File menu's own "More", in the
                    one place a reader is already looking at the list. */}
                {!fullHistory && recentFiles.length > HISTORY_SHOWN && (
                  <li>
                    <button
                      type="button"
                      onClick={() => setFullHistory(true)}
                      className="flex w-full items-center rounded-sm px-1.5 py-1 text-left text-xs text-muted-foreground transition-colors hover:bg-accent/50"
                    >
                      {t("welcome.more")}
                    </button>
                  </li>
                )}
              </ul>
            )}
          </section>
        </div>

        {/* What the keyboard does. A rule down its left edge, the way the
            app's own panels are held apart: without it the two columns are
            told apart by the space between them and nothing else. */}
        <section className="min-w-0 shrink-0 md:w-84 md:border-l md:border-border md:pl-6">
          <h2 className="text-[10px] font-medium tracking-wide text-muted-foreground uppercase">
            {t("welcome.shortcuts")}
          </h2>
          <ul className="mt-1.5 space-y-1">
            {shortcuts.map(({ keys, label }) => (
              <li key={label} className="flex items-baseline gap-3 text-[11px]">
                <span className="flex w-36 shrink-0 flex-wrap items-center gap-x-1 gap-y-0.5">
                  {keys.map((combo, index) => (
                    <span key={combo} className="flex items-center gap-1">
                      {/* Two keys that do the same thing read as one entry:
                          PageUp or PageDown, Ctrl+Home or Ctrl+End. */}
                      {index > 0 && (
                        <span className="text-muted-foreground/60">/</span>
                      )}
                      <Keys combo={combo} />
                    </span>
                  ))}
                </span>
                <span className="text-muted-foreground">{label}</span>
              </li>
            ))}
          </ul>
        </section>
      </div>
    </div>
  );
}
