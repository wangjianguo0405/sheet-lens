import { useCallback, useEffect, useState } from "react";
import { useTranslation } from "react-i18next";
import { open as openFolderDialog } from "@tauri-apps/plugin-dialog";
import {
  ChevronDown,
  ChevronRight,
  FileSpreadsheet,
  Folder,
  FolderOpen,
  Plus,
  RotateCw,
  X,
} from "lucide-react";

import { Hint } from "@/components/Hint";
import {
  listDirectory,
  readTreeOpen,
  readTreeRoots,
  storeTreeOpen,
  storeTreeRoots,
  type DirEntry,
} from "@/lib/fileTree";
import { fileName, folderKey, folderName } from "@/lib/recent";
import { cn } from "@/lib/utils";

/** A group's heading, the same type the start page and the sheet list use. */
const HEAD = "text-[10px] font-medium tracking-wide text-muted-foreground uppercase";

/** The tree's own little buttons, which wait under the pointer like the page's. */
const ACTION =
  "shrink-0 rounded-sm p-1 text-muted-foreground opacity-0 transition-opacity group-hover:opacity-60 hover:bg-accent hover:text-foreground hover:opacity-100";

/** What a workbook's row is coloured — the Excel green the app links in. */
const LINK = "text-[#107C41]";

/** One folder's level: the listing, or the read that failed. */
type Node =
  | { state: "loading" }
  | { state: "error"; message: string }
  | { state: "ready"; entries: DirEntry[] };

interface FileTreeProps {
  /** The workbook on screen, marked when the tree reaches it. */
  currentPath: string | null;
  onOpenFile: (path: string) => void;
}

/**
 * The reader's folder tree: roots they add by hand, opened one level at a
 * time.
 *
 * It runs on the real filesystem, through the app's own `list_directory` — a
 * folder is read when it is opened and not before, so a root as wide as a
 * checkout costs nothing until it is walked. Clicking a workbook opens it the
 * same way the File menu's history does.
 */
export function FileTree({ currentPath, onOpenFile }: FileTreeProps) {
  const { t } = useTranslation();
  const [roots, setRoots] = useState<string[]>(readTreeRoots);
  const [open, setOpen] = useState<Set<string>>(() => new Set(readTreeOpen()));
  const [nodes, setNodes] = useState<Map<string, Node>>(new Map());
  /** The reads on their way, so a second expand of the same row asks once. */
  const [inflight] = useState(() => new Set<string>());

  const load = useCallback(
    (path: string) => {
      if (inflight.has(path)) return;
      inflight.add(path);
      setNodes((prev) => new Map(prev).set(path, { state: "loading" }));
      listDirectory(path)
        .then((entries) => setNodes((prev) => new Map(prev).set(path, { state: "ready", entries })))
        .catch((reason: unknown) =>
          setNodes((prev) => new Map(prev).set(path, { state: "error", message: String(reason) })),
        )
        .finally(() => inflight.delete(path));
    },
    [inflight],
  );

  // The folders left open last time come back open, and their contents have to
  // be asked for again: only the opening is remembered, not the listing.
  useEffect(() => {
    for (const path of readTreeOpen()) load(path);
  }, [load]);

  const toggle = (path: string) => {
    const next = new Set(open);
    if (next.has(path)) next.delete(path);
    else {
      next.add(path);
      const node = nodes.get(path);
      if (!node || node.state === "error") load(path);
    }
    setOpen(next);
    storeTreeOpen([...next]);
  };

  /** Drops every listing and reads the open folders again — what SVN update changed. */
  const refresh = () => {
    setNodes(new Map());
    for (const path of open) load(path);
  };

  /**
   * A folder picked rather than typed: the same dialog the open button uses,
   * asking for a folder instead of workbooks.
   */
  const addRoot = async () => {
    try {
      const picked = await openFolderDialog({ directory: true, multiple: false });
      if (typeof picked !== "string") return;
      if (roots.some((root) => folderKey(root) === folderKey(picked))) return;
      setRoots(storeTreeRoots([...roots, picked]));
      // Adding a root is asking to look inside it, so it opens at once.
      const next = new Set(open);
      next.add(picked);
      setOpen(next);
      storeTreeOpen([...next]);
      load(picked);
    } catch {
      // A dialog that cannot open leaves the tree with what it had.
    }
  };

  const removeRoot = (path: string) =>
    setRoots(storeTreeRoots(roots.filter((root) => root !== path)));

  const folderRow = (path: string, depth: number, root: boolean) => {
    const node = nodes.get(path);
    const expanded = open.has(path);
    const Chevron = expanded ? ChevronDown : ChevronRight;
    const Face = expanded ? FolderOpen : Folder;
    return (
      <li key={path}>
        <div
          className="group flex items-center rounded-sm pr-0.5 hover:bg-accent/50"
          style={{ paddingLeft: 4 + depth * 12 }}
        >
          <button
            type="button"
            onClick={() => toggle(path)}
            aria-expanded={expanded}
            className="flex min-w-0 flex-1 items-center gap-1 px-1 py-0.5 text-left text-xs"
          >
            <Chevron className="size-3 shrink-0 text-muted-foreground" />
            <Face className="size-3.5 shrink-0 text-muted-foreground" />
            <span className="whitespace-nowrap">{fileName(path)}</span>
            {/* A root shows where it comes from: its own name is the leaf every
                checkout repeats, and the folder above it is the one that tells
                one checkout from another. */}
            {root && (
              <span className="min-w-0 flex-1 whitespace-nowrap text-[10px] text-muted-foreground">
                {folderName(path)}
              </span>
            )}
          </button>
          {root && (
            <Hint label={t("tree.removeRoot", { name: fileName(path) })}>
              <button
                type="button"
                aria-label={t("tree.removeRoot", { name: fileName(path) })}
                onClick={() => removeRoot(path)}
                className={ACTION}
              >
                <X className="size-3" />
              </button>
            </Hint>
          )}
        </div>
        {expanded && (
          <ul>
            {node === undefined || node.state === "loading" ? (
              <li
                className="px-1 py-0.5 text-[11px] text-muted-foreground"
                style={{ paddingLeft: 22 + depth * 12 }}
              >
                {t("tree.loading")}
              </li>
            ) : node.state === "error" ? (
              <li style={{ paddingLeft: 22 + depth * 12 }}>
                <Hint label={node.message}>
                  <span className="block truncate px-1 py-0.5 text-[11px] text-destructive">
                    {t("tree.error")}
                  </span>
                </Hint>
              </li>
            ) : (
              node.entries.map((entry) =>
                entry.dir ? folderRow(entry.path, depth + 1, false) : fileRow(entry, depth + 1),
              )
            )}
          </ul>
        )}
      </li>
    );
  };

  const fileRow = (entry: DirEntry, depth: number) => {
    const current = currentPath !== null && folderKey(currentPath) === folderKey(entry.path);
    return (
      <li key={entry.path}>
        <button
          type="button"
          onClick={() => onOpenFile(entry.path)}
          title={entry.path}
          className={cn(
            "flex w-full min-w-0 items-center gap-1 rounded-sm px-1 py-0.5 text-left text-xs transition-colors hover:bg-accent/50 hover:underline",
            LINK,
            current && "bg-accent/60 font-medium",
          )}
          style={{ paddingLeft: 16 + depth * 12 }}
        >
          <FileSpreadsheet className="size-3.5 shrink-0" />
          <span className="whitespace-nowrap">{entry.name}</span>
        </button>
      </li>
    );
  };

  return (
    <div className="flex min-h-0 flex-1 flex-col">
      <header className="group flex shrink-0 items-center gap-1 px-2 py-1">
        <span className={cn(HEAD, "min-w-0 flex-1 truncate")}>{t("tree.title")}</span>
        <div className="flex items-center gap-0.5">
          <Hint label={t("tree.refresh")}>
            <button
              type="button"
              aria-label={t("tree.refresh")}
              onClick={refresh}
              className={ACTION}
            >
              <RotateCw className="size-3" />
            </button>
          </Hint>
          <Hint label={t("tree.addRoot")}>
            <button
              type="button"
              aria-label={t("tree.addRoot")}
              onClick={() => void addRoot()}
              className={ACTION}
            >
              <Plus className="size-3" />
            </button>
          </Hint>
        </div>
      </header>
      {/* The list is as wide as its widest row: a deep folder's indent and a
          long name both run past the panel's edge, and the list scrolls
          sideways to reach them rather than cutting the row off where the
          panel does. */}
      <div className="min-h-0 flex-1 overflow-auto pb-1">
        {roots.length === 0 ? (
          <p className="px-2 py-1 text-[11px] text-muted-foreground">{t("tree.empty")}</p>
        ) : (
          <ul className="w-max min-w-full">{roots.map((root) => folderRow(root, 0, true))}</ul>
        )}
      </div>
    </div>
  );
}
