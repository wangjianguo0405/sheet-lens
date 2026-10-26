import { FileSpreadsheet, Folder, FolderOpen, Pin, PinOff, Trash2 } from "lucide-react";
import { useState, type ReactNode } from "react";
import { useTranslation } from "react-i18next";

import {
  fileName,
  folderKey,
  folderName,
  foldersOf,
  pinDir,
  pinFile,
  readPinnedDirs,
  readPinnedFiles,
  unpinDir,
  unpinFile,
} from "@/lib/recent";
import { cn } from "@/lib/utils";
import { Hint } from "@/components/Hint";

/**
 * One column's height, and how much of it the fixed run may take.
 *
 * A column is the scroller, not each list under it: one fixed folder beside
 * ten fixed workbooks has to come out the same height, and a heading that
 * scrolls away is no heading. The fixed run sticks at the top of its column,
 * so what was fixed stays in reach while the rest goes by — but only for a
 * share of the column, or a long enough fixed list would leave nothing below.
 */
const COLUMN_HEIGHT = "md:h-72";
const FIXED_HEIGHT = "max-h-40";

/** A group's heading, as every column of the page draws it. */
const HEAD = "text-[10px] font-medium tracking-wide text-muted-foreground uppercase";

/**
 * What a thing you can click is coloured: the Excel green the app already marks
 * its live parts with — the current sheet's rule, the panel splitters. A row
 * that is the same grey as the prose around it does not read as a link, and
 * everything in this column that is not a folder path is one.
 */
const LINK = "text-[#107C41]";

/**
 * A folder row's two little buttons — the pin and the bin. They wait under the
 * pointer, as the file rows' pin does: a column of folders reads as folders,
 * with what can be done to one arriving only once it is pointed at.
 */
const ACTION =
  "shrink-0 rounded-sm p-1 text-muted-foreground opacity-0 transition-opacity group-hover:opacity-60 hover:bg-accent hover:text-foreground hover:opacity-100";

interface WelcomeProps {
  onOpen: () => void;
  /** Opens the picker already inside this folder. */
  onOpenDirectory: (path: string) => void;
  /**
   * Takes a folder off the page — its pin, and every workbook it holds out of
   * the file lists, which is where the folder would otherwise come back from.
   */
  onForgetFolder: (folder: string) => void;
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
      <Hint label={missing ? t("toolbar.missing", { path }) : path}>
        <button
          type="button"
          onClick={() => onOpen(path)}
          // A workbook that would not open is greyed and left in place, exactly
          // as the File menu leaves it: a click is what gives it another try.
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
      </Hint>
      <Hint label={label}>
        <button
          type="button"
          onClick={() => onToggle(path)}
          aria-label={label}
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
      </Hint>
    </li>
  );
}

/**
 * One folder in the page's first column: what it is called, where it is, and
 * the pin that keeps it there. A press opens the picker inside it rather than
 * listing what is in it — a folder is a place to start from, and the picker is
 * where a workbook is chosen anyway.
 */
function DirRow({
  path,
  fixed,
  onOpen,
  onToggle,
  onForget,
}: {
  path: string;
  /** Whether this folder is fixed, which decides what the pin does. */
  fixed: boolean;
  onOpen: (path: string) => void;
  onToggle: (path: string) => void;
  /** Drops the folder and the workbooks in it from the page's lists. */
  onForget: (path: string) => void;
}) {
  const { t } = useTranslation();
  const name = fileName(path);
  const label = t(fixed ? "welcome.unpinDir" : "welcome.pinDir", { name });
  const removeLabel = t("welcome.removeDir", { name });

  return (
    <li className="group flex items-center">
      <Hint label={t("welcome.openFolder", { path })}>
        <button
          type="button"
          onClick={() => onOpen(path)}
          className={cn(
            LINK,
            "flex min-w-0 flex-1 items-center gap-2 rounded-sm px-1.5 py-1 text-left text-xs transition-colors hover:bg-accent/50",
          )}
        >
          <Folder className="size-3.5 shrink-0" />
          <span className="min-w-0 truncate group-hover:underline">{name}</span>
          <span className="min-w-0 flex-1 truncate text-muted-foreground">
            {folderName(path)}
          </span>
        </button>
      </Hint>
      <Hint label={label}>
        <button
          type="button"
          onClick={() => onToggle(path)}
          aria-label={label}
          className={ACTION}
        >
          {fixed ? <PinOff className="size-3" /> : <Pin className="size-3" />}
        </button>
      </Hint>
      <Hint label={removeLabel}>
        <button
          type="button"
          onClick={() => onForget(path)}
          aria-label={removeLabel}
          className={ACTION}
        >
          <Trash2 className="size-3" />
        </button>
      </Hint>
    </li>
  );
}

/**
 * One column of the page: the fixed run at its top, then the rest.
 *
 * The column scrolls rather than each of its lists, so the two columns come
 * out one height whatever they hold, and the fixed run stays under its heading
 * at the top while the rest of the column goes by. The rows are handed in
 * made, since a folder and a workbook are drawn differently.
 */
function Column({
  fixed,
  rest,
  fixedLabel,
  restLabel,
  emptyLabel,
}: {
  /** The rows of the fixed run, in the order they are to be drawn. */
  fixed: ReactNode[];
  /** Everything else, under its own heading. */
  rest: ReactNode[];
  fixedLabel: string;
  restLabel: string;
  /** What the column says when it has nothing to list at all. */
  emptyLabel: string;
}) {
  return (
    <div className={cn("overflow-y-auto", COLUMN_HEIGHT)}>
      {fixed.length === 0 && rest.length === 0 ? (
        <>
          <h2 className={HEAD}>{restLabel}</h2>
          <p className="mt-1.5 text-[11px] text-muted-foreground">{emptyLabel}</p>
        </>
      ) : (
        <>
          {fixed.length > 0 && (
            <div
              className={cn(
                "sticky top-0 z-10 bg-background",
                rest.length > 0 && "border-b border-border pb-1",
              )}
            >
              <h2 className={HEAD}>{fixedLabel}</h2>
              <ul className={cn("mt-1 overflow-y-auto", FIXED_HEIGHT)}>{fixed}</ul>
            </div>
          )}
          {rest.length > 0 && (
            <>
              <h2 className={cn(HEAD, fixed.length > 0 && "mt-3")}>{restLabel}</h2>
              <ul className="mt-1">{rest}</ul>
            </>
          )}
        </>
      )}
    </div>
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
  onOpenDirectory,
  onForgetFolder,
  recentFiles,
  missingFiles,
  onOpenRecent,
  error,
}: WelcomeProps) {
  const { t } = useTranslation();
  const missing = new Set(missingFiles);
  /** Fixed by hand, and kept here rather than in the app: the page is its only
      reader, and mounting reads the list back out of storage. */
  const [pinnedFiles, setPinnedFiles] = useState<string[]>(readPinnedFiles);
  const pinned = new Set(pinnedFiles);
  /** The folders kept beside them, and the ones the history has passed through. */
  const [pinnedDirs, setPinnedDirs] = useState<string[]>(readPinnedDirs);
  const pinnedDirKeys = new Set(pinnedDirs.map(folderKey));
  const folderHistory = foldersOf([...recentFiles, ...pinnedFiles]).filter(
    (dir) => !pinnedDirKeys.has(folderKey(dir)),
  );
  /** A fixed workbook is shown above, so the history leaves it out rather
      than offering the same workbook twice. */
  const history = recentFiles.filter((path) => !pinned.has(path));

  const togglePinned = (path: string) =>
    setPinnedFiles(pinned.has(path) ? unpinFile(path) : pinFile(path));

  const togglePinnedDir = (path: string) =>
    setPinnedDirs(
      pinnedDirKeys.has(folderKey(path)) ? unpinDir(path) : pinDir(path),
    );

  /**
   * Drops a folder and the workbooks in it. The lists themselves are written
   * by the app — the history is its own state — so this reads the pinned ones
   * back once it has, rather than guessing at what it removed.
   */
  const forgetFolder = (path: string) => {
    onForgetFolder(path);
    setPinnedFiles(readPinnedFiles());
    setPinnedDirs(readPinnedDirs());
  };

  /**
   * Every key the app answers to, in the order a reader meets them: the file,
   * then the sheet, then moving around inside it. The sheet's are not usable
   * until something is open, which is what the page is for.
   */
  const shortcuts: { keys: string[]; label: string }[] = [
    { keys: ["Ctrl+O"], label: t("welcome.sc.open") },
    { keys: ["Ctrl+W"], label: t("welcome.sc.close") },
    { keys: ["Ctrl+Tab", "Ctrl+Shift+Tab"], label: t("welcome.sc.workbook") },
    { keys: ["Ctrl+F"], label: t("welcome.sc.search") },
    { keys: ["Ctrl+G", "F5"], label: t("welcome.sc.goTo") },
    { keys: ["Ctrl+A"], label: t("welcome.sc.selectAll") },
    { keys: ["Shift+Space", "Ctrl+Space"], label: t("welcome.sc.rowColumn") },
    { keys: ["Ctrl+C"], label: t("welcome.sc.copy") },
    { keys: ["←↑↓→"], label: t("welcome.sc.move") },
    { keys: ["Shift+←↑↓→"], label: t("welcome.sc.extend") },
    { keys: ["Ctrl+←↑↓→"], label: t("welcome.sc.edge") },
    { keys: ["Ctrl+Shift+Home", "Ctrl+Shift+End"], label: t("welcome.sc.extendEnds") },
    { keys: ["PageUp", "PageDown"], label: t("welcome.sc.page") },
    { keys: ["Alt+PageUp", "Alt+PageDown"], label: t("welcome.sc.pageAcross") },
    { keys: ["Enter", "Tab"], label: t("welcome.sc.enterTab") },
    { keys: ["Home", "End"], label: t("welcome.sc.rowEnds") },
    { keys: ["Ctrl+Home", "Ctrl+End"], label: t("welcome.sc.sheetEnds") },
    { keys: ["Ctrl+PageUp", "Ctrl+PageDown"], label: t("welcome.sc.sheetStep") },
    { keys: ["Ctrl+B", "Ctrl+I", "Ctrl+U"], label: t("welcome.sc.format") },
    { keys: ["Ctrl+Shift+L"], label: t("welcome.sc.filter") },
    { keys: ["Esc"], label: t("welcome.sc.escape") },
  ];
  // Read down each column and on into the next, as a printed list would be.
  const columns = 3;
  const perColumn = Math.ceil(shortcuts.length / columns);
  const groups = Array.from({ length: columns }, (_, index) =>
    shortcuts.slice(index * perColumn, (index + 1) * perColumn),
  );

  return (
    <div className="flex h-full w-full overflow-auto p-6">
      {/* `m-auto` centres the page in a window tall enough for it and lets it
          start at the top in one that is not — centring it with the flexbox
          properties alone put the title above the scroller's reach. */}
      <div className="m-auto flex w-full max-w-6xl flex-col gap-6">
        {/* Above the three columns rather than inside one of them: the name,
            what the app is and the one thing the page is for belong to the
            page, not to any one of the lists under it. */}
        <header className="flex flex-col items-center gap-1 text-center">
          <div className="flex items-center gap-2">
            <FileSpreadsheet className="size-5 text-muted-foreground/60" />
            <h1 className="text-sm font-semibold">SheetLens</h1>
          </div>
          {/* Two lines rather than one paragraph: what the app is, then what
              to do with it — which is the part worth reading twice. */}
          <p className="text-xs text-muted-foreground">{t("welcome.tagline")}</p>
          <p className="text-xs text-muted-foreground">{t("welcome.hint")}</p>

          {/* The one thing the page is for, as a row with an icon rather than
              as a filled button: a block of solid colour on an empty window is
              a shout. A size up from the lists below it, since it is the
              page's whole purpose and the names under it are only where you
              were. */}
          <button
            type="button"
            onClick={onOpen}
            className={cn(
              LINK,
              "mt-4 flex items-center gap-2 rounded-sm px-3 py-1.5 text-sm transition-colors hover:bg-accent/50 hover:underline",
            )}
          >
            <FolderOpen className="size-4 shrink-0" />
            {t("welcome.open")}
          </button>

          {error && (
            <p className="mt-2 max-w-md rounded-sm bg-destructive/10 px-2 py-1 text-[11px] text-destructive">
              {error}
            </p>
          )}
        </header>

        <div className="flex flex-col gap-6 md:flex-row md:gap-0">
          {/* Where to look. A workbook is often one of a folder's several, and
              finding the folder is the part that takes the walking. */}
          <section className="min-w-0 shrink-0 md:w-80 md:border-r md:border-border md:pr-6">
            <Column
              fixed={pinnedDirs.map((path) => (
                <DirRow
                  key={`fixed:${path}`}
                  path={path}
                  fixed
                  onOpen={onOpenDirectory}
                  onToggle={togglePinnedDir}
                  onForget={forgetFolder}
                />
              ))}
              rest={folderHistory.map((path) => (
                <DirRow
                  key={`recent:${path}`}
                  path={path}
                  fixed={false}
                  onOpen={onOpenDirectory}
                  onToggle={togglePinnedDir}
                  onForget={forgetFolder}
                />
              ))}
              fixedLabel={t("welcome.pinnedDirs")}
              restLabel={t("welcome.folders")}
              emptyLabel={t("welcome.noFolders")}
            />
          </section>

          {/* What to open */}
          <div className="min-w-0 flex-1 md:px-6">
            <Column
              fixed={pinnedFiles.map((path) => (
                <Row
                  key={path}
                  path={path}
                  missing={missing.has(path)}
                  fixed
                  onOpen={onOpenRecent}
                  onToggle={togglePinned}
                />
              ))}
              rest={history.map((path) => (
                <Row
                  key={path}
                  path={path}
                  missing={missing.has(path)}
                  fixed={false}
                  onOpen={onOpenRecent}
                  onToggle={togglePinned}
                />
              ))}
              fixedLabel={t("welcome.pinned")}
              restLabel={t("welcome.history")}
              emptyLabel={t("welcome.noHistory")}
            />
          </div>
        </div>

        {/* What the keyboard does, across the foot of the page: a rule above it
            holds it apart from the two lists, and its own three columns keep a
            list this long from running down the window. */}
        <section className="min-w-0 shrink-0 border-t border-border pt-4">
          <h2 className={HEAD}>
            {t("welcome.shortcuts")}
          </h2>
          <div className="mt-1.5 grid gap-x-8 md:grid-cols-3">
            {groups.map((group, index) => (
              <ul key={index} className="space-y-1">
                {group.map(({ keys, label }) => (
                  <li key={label} className="flex items-baseline gap-3 text-[11px]">
                    <span className="flex w-36 shrink-0 flex-wrap items-center gap-x-1 gap-y-0.5">
                      {keys.map((combo, position) => (
                        <span key={combo} className="flex items-center gap-1">
                          {/* Two keys that do the same thing read as one entry:
                              PageUp or PageDown, Ctrl+Home or Ctrl+End. */}
                          {position > 0 && (
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
            ))}
          </div>
        </section>
      </div>
    </div>
  );
}
