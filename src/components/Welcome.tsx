import { FileSpreadsheet, FolderOpen, Pin, PinOff } from "lucide-react";
import { useState } from "react";
import { useTranslation } from "react-i18next";

import {
  fileName,
  folderName,
  pinFile,
  readPinnedFiles,
  unpinFile,
} from "@/lib/recent";
import { cn } from "@/lib/utils";

/** How much of the history shows before the rest goes behind "More". */
const HISTORY_SHOWN = 6;

/**
 * What a thing you can click is coloured: the Excel green the app already marks
 * its live parts with — the current sheet's rule, the panel splitters. A row
 * that is the same grey as the prose around it does not read as a link, and
 * everything in this column that is not a folder path is one.
 */
const LINK = "text-[#107C41]";

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
 * One workbook, in either list: where it is, what it is called, and the pin
 * that moves it between the two. The name and the pin are two buttons side by
 * side rather than one button with a button in it, which HTML has no such
 * thing as — and the pin has to be its own target, or clicking it would open
 * the workbook it was meant to fix.
 */
function Row({
  path,
  missing,
  fixed,
  onOpen,
  onToggle,
}: {
  path: string;
  missing: boolean;
  /** Whether this workbook is fixed, which decides what the pin does. */
  fixed: boolean;
  onOpen: (path: string) => void;
  onToggle: (path: string) => void;
}) {
  const { t } = useTranslation();
  const label = t(fixed ? "welcome.unpin" : "welcome.pin", {
    name: fileName(path),
  });

  return (
    <li className="group flex items-center">
      <button
        type="button"
        onClick={() => onOpen(path)}
        // A workbook that would not open is greyed and left in place, exactly
        // as the File menu leaves it: a click is what gives it another try.
        title={missing ? t("toolbar.missing", { path }) : path}
        className={cn(
          "flex min-w-0 flex-1 items-center gap-2 rounded-sm px-1.5 py-1 text-left text-xs transition-colors hover:bg-accent/50",
          missing ? "text-muted-foreground" : LINK,
        )}
      >
        {/* Where it is, as well as what it is called: two workbooks of the
            same name in two folders are told apart by nothing else. The folder
            is the part that gives way when the row runs out. */}
        <span className="min-w-0 truncate group-hover:underline">
          {fileName(path)}
        </span>
        <span className="min-w-0 flex-1 truncate text-muted-foreground">
          {folderName(path)}
        </span>
      </button>
      <button
        type="button"
        onClick={() => onToggle(path)}
        aria-label={label}
        title={label}
        className={cn(
          "shrink-0 rounded-sm p-1 text-muted-foreground transition-opacity hover:bg-accent hover:text-foreground hover:opacity-100",
          // A fixed workbook shows the way back out of being fixed without
          // being asked for; the history offers the way in only under a
          // pointer already on the row.
          fixed ? "opacity-60" : "opacity-0 group-hover:opacity-60",
        )}
      >
        {fixed ? <PinOff className="size-3" /> : <Pin className="size-3" />}
      </button>
    </li>
  );
}

/**
 * The start page: what to open on the left, and what the keyboard does on the
 * right.
 *
 * The two lists are here rather than only behind the File menu because an empty
 * window is the one place they are worth the space: the workbook being come
 * back to is usually one of the last few, and picking it out of a list beats
 * walking the file picker again. The history is the File menu's own list, so
 * anything taken out there goes from here too — and the fixed ones are the
 * handful that should not have to be found again at all.
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
  /** Fixed by hand, and kept here rather than in the app: the page is its only
      reader, and mounting reads the list back out of storage. */
  const [pinnedFiles, setPinnedFiles] = useState<string[]>(readPinnedFiles);
  const pinned = new Set(pinnedFiles);
  /** A fixed workbook is shown above, so the history leaves it out rather
      than offering the same workbook twice. */
  const history = recentFiles.filter((path) => !pinned.has(path));

  const togglePinned = (path: string) =>
    setPinnedFiles(pinned.has(path) ? unpinFile(path) : pinFile(path));

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
      <div className="flex w-full max-w-4xl flex-col gap-6 md:flex-row md:gap-0">
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

          {/* The one thing the page is for, as a row with an icon rather than
              as a filled button: it heads the same column as the list under
              it, and a block of solid colour on an empty window is a shout.
              A size up from the list, since it is the page's whole purpose and
              the names under it are only where you were. */}
          <button
            type="button"
            onClick={onOpen}
            className={cn(
              LINK,
              "mt-5 flex w-full items-center gap-2 rounded-sm px-1.5 py-1.5 text-left text-sm transition-colors hover:bg-accent/50 hover:underline",
            )}
          >
            <FolderOpen className="size-4 shrink-0" />
            {t("welcome.open")}
          </button>

          {error && (
            <p className="mt-3 max-w-md rounded-sm bg-destructive/10 px-2 py-1 text-[11px] text-destructive">
              {error}
            </p>
          )}

          <section className="mt-6">
            {pinnedFiles.length === 0 && history.length === 0 ? (
              <>
                <h2 className="text-[10px] font-medium tracking-wide text-muted-foreground uppercase">
                  {t("welcome.history")}
                </h2>
                <p className="mt-1.5 text-[11px] text-muted-foreground">
                  {t("welcome.noHistory")}
                </p>
              </>
            ) : (
              <>
                {pinnedFiles.length > 0 && (
                  <>
                    <h2 className="text-[10px] font-medium tracking-wide text-muted-foreground uppercase">
                      {t("welcome.pinned")}
                    </h2>
                    <ul className="mt-1 max-h-64 overflow-y-auto">
                      {pinnedFiles.map((path) => (
                        <Row
                          key={path}
                          path={path}
                          missing={missing.has(path)}
                          fixed
                          onOpen={onOpenRecent}
                          onToggle={togglePinned}
                        />
                      ))}
                    </ul>
                  </>
                )}

                {history.length > 0 && (
                  <>
                    <h2
                      className={cn(
                        "text-[10px] font-medium tracking-wide text-muted-foreground uppercase",
                        pinnedFiles.length > 0 && "mt-5",
                      )}
                    >
                      {t("welcome.history")}
                    </h2>
                    <ul className="mt-1 max-h-64 overflow-y-auto">
                      {(fullHistory ? history : history.slice(0, HISTORY_SHOWN)).map(
                        (path) => (
                          <Row
                            key={path}
                            path={path}
                            missing={missing.has(path)}
                            fixed={false}
                            onOpen={onOpenRecent}
                            onToggle={togglePinned}
                          />
                        ),
                      )}
                      {/* The rest of the list, for a history with more in it
                          than the page has room for — the File menu's own
                          "More", in the one place a reader is already looking
                          at the list. */}
                      {!fullHistory && history.length > HISTORY_SHOWN && (
                        <li>
                          <button
                            type="button"
                            onClick={() => setFullHistory(true)}
                            className={cn(
                              LINK,
                              "flex w-full items-center rounded-sm px-1.5 py-1 text-left text-xs transition-colors hover:bg-accent/50 hover:underline",
                            )}
                          >
                            {t("welcome.more")}
                          </button>
                        </li>
                      )}
                    </ul>
                  </>
                )}
              </>
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
