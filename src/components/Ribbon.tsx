import {
  AlignCenter,
  AlignCenterVertical,
  AlignEndVertical,
  AlignLeft,
  AlignRight,
  AlignStartVertical,
  Bold,
  ChevronDown,
  Columns3,
  Italic,
  ListFilter,
  ListTree,
  PaintBucket,
  PanelBottom,
  PanelLeft,
  Rows3,
  Snowflake,
  SplitSquareHorizontal,
  Underline,
  Ungroup,
  WrapText,
} from "lucide-react";
import { useEffect, useState } from "react";
import { useTranslation } from "react-i18next";

import { Button } from "@/components/ui/button";
import { ColorPalette } from "@/components/ColorPalette";
import { Hint } from "@/components/Hint";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuRadioGroup,
  DropdownMenuRadioItem,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import type { ActiveCellStyle, CellFormat } from "@/lib/formats";
import { listFonts } from "@/lib/fonts";
import { cn } from "@/lib/utils";

/**
 * The faces to offer when the machine cannot be asked — a build without the
 * backend command, or a page running outside Tauri. The list is the fonts the
 * documents this viewer is built for are actually set in, plus the Latin faces
 * Excel itself starts from.
 */
const FONT_FAMILIES = [
  "Calibri",
  "Arial",
  "Times New Roman",
  "游ゴシック",
  "Yu Gothic",
  "ＭＳ Ｐゴシック",
  "ＭＳ ゴシック",
  "Meiryo",
  "等线",
  "宋体",
  "微软雅黑",
];

interface RibbonProps {
  sidebarOpen: boolean;
  onToggleSidebar: () => void;
  sheetTabsOpen: boolean;
  onToggleSheetTabs: () => void;
  /** Shades the selection, or lifts the shading off it when `fill` is null. */
  onMark: (fill: string | null) => void;
  /** What the bucket fills with when it is pressed rather than its caret. */
  markColor: string;
  /** The active cell's style, as the font and alignment groups show it. */
  activeStyle: ActiveCellStyle | null;
  /** Lays a font or alignment choice over the selection; null takes it off. */
  onFormat: (patch: CellFormat | null) => void;
  /** Whether any frozen band is in force, the file's included. */
  frozen: boolean;
  onToggleFreeze: () => void;
  splitOn: boolean;
  onToggleSplit: () => void;
  filterOn: boolean;
  /** Adds a filter over the current region, or takes the one in force away. */
  onToggleFilter: () => void;
  onGroup: (action: "rows" | "cols" | "ungroup") => void;
}

/**
 * One labelled block of the ribbon. The label sits under its controls the way
 * Excel's group names do, and a rule between groups is what tells one from the
 * next — without it a row of small icons reads as one undifferentiated strip.
 */
function Group({
  label,
  children,
}: {
  label: string;
  children: React.ReactNode;
}) {
  return (
    <div className="flex shrink-0 flex-col justify-between gap-1 border-l border-border px-2 first:border-l-0 first:pl-0">
      <div className="flex items-center gap-0.5">{children}</div>
      <span className="text-center text-[9px] leading-none tracking-wide text-muted-foreground/80 uppercase">
        {label}
      </span>
    </div>
  );
}

/** A ribbon button that reads as pressed, for the toggles among them. */
function Toggle({  pressed,
  onClick,
  title,
  children,
}: {
  pressed: boolean;
  onClick: () => void;
  title: string;
  children: React.ReactNode;
}) {
  return (
    <Hint label={title}>
      <Button
        size="sm"
        variant={pressed ? "secondary" : "ghost"}
        className="h-6 min-w-6 gap-1 px-1 text-[11px]"
        onClick={onClick}
        aria-pressed={pressed}
        aria-label={title}
      >
        {children}
      </Button>
    </Hint>
  );
}

/**
 * The font size box. It keeps a draft of what is being typed so the cell's own
 * size does not overwrite a half-typed number; the choice lands on Enter or on
 * leaving the box, and the box then takes the size in force again.
 */
function SizeBox({
  sizePx,
  onPick,
}: {
  sizePx: number;
  onPick: (sizePx: number) => void;
}) {
  const points = Math.round(((sizePx * 3) / 4) * 10) / 10;
  const [draft, setDraft] = useState(String(points));
  useEffect(() => setDraft(String(points)), [points]);

  const commit = () => {
    const asked = Number(draft.replace(",", "."));
    if (Number.isFinite(asked) && asked > 0) onPick((asked * 4) / 3);
    else setDraft(String(points));
  };

  return (
    <input
      type="text"
      inputMode="decimal"
      value={draft}
      onChange={(event) => setDraft(event.target.value)}
      onBlur={commit}
      onKeyDown={(event) => {
        if (event.key === "Enter") {
          event.preventDefault();
          commit();
          (event.target as HTMLInputElement).blur();
        }
      }}
      className="h-6 w-11 rounded-sm border border-input bg-transparent px-1 text-[11px] tabular-nums outline-none focus-visible:border-ring"
    />
  );
}

/**
 * The ribbon under the file tabs: what can be done to the sheet on screen.
 *
 * Excel's ribbon is a strip of labelled groups, and this is a short version of
 * one. It only exists while a workbook is open — a start page has nothing to
 * format — and everything it changes is the app's own, never the file's.
 */
export function Ribbon({
  sidebarOpen,
  onToggleSidebar,
  sheetTabsOpen,
  onToggleSheetTabs,
  onMark,
  markColor,
  activeStyle,
  onFormat,
  frozen,
  onToggleFreeze,
  splitOn,
  onToggleSplit,
  filterOn,
  onToggleFilter,
  onGroup,
}: RibbonProps) {
  const { t } = useTranslation();

  const style = activeStyle?.style ?? null;
  const family = activeStyle?.family ?? null;
  const horizontal = style?.horizontal ?? null;
  const vertical = style?.vertical ?? "bottom";
  /**
   * What the machine actually has, asked for once. The built-in list stands
   * until the answer arrives, and for good if there is none.
   */
  const [installedFonts, setInstalledFonts] = useState<string[]>([]);
  useEffect(() => {
    let cancelled = false;
    void listFonts()
      .then((names) => {
        if (!cancelled && names.length > 0) setInstalledFonts(names);
      })
      .catch(() => {
        // Nothing to say: the list below is still a usable one.
      });
    return () => {
      cancelled = true;
    };
  }, []);
  const families = installedFonts.length > 0 ? installedFonts : FONT_FAMILIES;
  const fonts = family && !families.includes(family) ? [family, ...families] : families;

  return (
    <div className="flex shrink-0 items-stretch gap-1 overflow-x-auto border-b border-border bg-muted/20 px-2 py-1">
      <Group label={t("ribbon.view")}>
        <Toggle
          pressed={sidebarOpen}
          onClick={onToggleSidebar}
          title={
            sidebarOpen ? t("toolbar.hideSheetList") : t("toolbar.showSheetList")
          }
        >
          <PanelLeft className="size-3.5" />
        </Toggle>
        <Toggle
          pressed={sheetTabsOpen}
          onClick={onToggleSheetTabs}
          title={
            sheetTabsOpen ? t("toolbar.hideSheetTabs") : t("toolbar.showSheetTabs")
          }
        >
          <PanelBottom className="size-3.5" />
        </Toggle>
        <Toggle
          pressed={frozen}
          onClick={onToggleFreeze}
          title={frozen ? t("ribbon.unfreeze") : t("ribbon.freeze")}
        >
          <Snowflake className="size-3.5" />
        </Toggle>
        <Toggle
          pressed={splitOn}
          onClick={onToggleSplit}
          title={splitOn ? t("ribbon.unsplit") : t("ribbon.split")}
        >
          <SplitSquareHorizontal className="size-3.5" />
        </Toggle>
      </Group>

      {/* Excel's Font group: the face, the size, and the three switches. */}
      <Group label={t("ribbon.font")}>
        <DropdownMenu>
          <Hint label={t("ribbon.fontFamily")}>
            <DropdownMenuTrigger asChild>
              <Button
                size="sm"
                variant="ghost"
                className="h-6 max-w-32 justify-between gap-1 px-1.5 text-[11px] font-normal"
              >
                <span className="truncate">{family ?? t("ribbon.defaultFont")}</span>
                <ChevronDown className="size-3 shrink-0 opacity-60" />
              </Button>
            </DropdownMenuTrigger>
          </Hint>
          <DropdownMenuContent align="start" className="max-h-80 min-w-44 overflow-y-auto">
            <DropdownMenuRadioGroup
              value={family ?? ""}
              onValueChange={(value) => onFormat({ family: value })}
            >
              {fonts.map((name) => (
                <DropdownMenuRadioItem
                  key={name}
                  value={name}
                  className="text-xs"
                  style={{ fontFamily: name }}
                >
                  {name}
                </DropdownMenuRadioItem>
              ))}
            </DropdownMenuRadioGroup>
          </DropdownMenuContent>
        </DropdownMenu>

        <SizeBox
          sizePx={activeStyle?.sizePx ?? 12}
          onPick={(sizePx) => onFormat({ sizePx })}
        />

        <Toggle
          pressed={style?.bold ?? false}
          onClick={() => onFormat({ bold: !(style?.bold ?? false) })}
          title={t("ribbon.bold")}
        >
          <Bold className="size-3.5" />
        </Toggle>
        <Toggle
          pressed={style?.italic ?? false}
          onClick={() => onFormat({ italic: !(style?.italic ?? false) })}
          title={t("ribbon.italic")}
        >
          <Italic className="size-3.5" />
        </Toggle>
        <Toggle
          pressed={style?.underline === "single" || style?.underline === "double"}
          onClick={() => onFormat({ underline: style?.underline ? null : "single" })}
          title={t("ribbon.underline")}
        >
          <Underline className="size-3.5" />
        </Toggle>
      </Group>

      {/* Alignment, split the way Excel splits it: across, then down. */}
      <Group label={t("ribbon.align")}>
        <Toggle
          pressed={horizontal === "left"}
          onClick={() => onFormat({ horizontal: "left" })}
          title={t("ribbon.alignLeft")}
        >
          <AlignLeft className="size-3.5" />
        </Toggle>
        <Toggle
          pressed={horizontal === "center"}
          onClick={() => onFormat({ horizontal: "center" })}
          title={t("ribbon.alignCenter")}
        >
          <AlignCenter className="size-3.5" />
        </Toggle>
        <Toggle
          pressed={horizontal === "right"}
          onClick={() => onFormat({ horizontal: "right" })}
          title={t("ribbon.alignRight")}
        >
          <AlignRight className="size-3.5" />
        </Toggle>
        <Toggle
          pressed={vertical === "top"}
          onClick={() => onFormat({ vertical: "top" })}
          title={t("ribbon.alignTop")}
        >
          <AlignStartVertical className="size-3.5" />
        </Toggle>
        <Toggle
          pressed={vertical === "middle"}
          onClick={() => onFormat({ vertical: "middle" })}
          title={t("ribbon.alignMiddle")}
        >
          <AlignCenterVertical className="size-3.5" />
        </Toggle>
        <Toggle
          pressed={vertical === "bottom"}
          onClick={() => onFormat({ vertical: "bottom" })}
          title={t("ribbon.alignBottom")}
        >
          <AlignEndVertical className="size-3.5" />
        </Toggle>
        <Toggle
          pressed={style?.wrap ?? false}
          onClick={() => onFormat({ wrap: !(style?.wrap ?? false) })}
          title={t("ribbon.wrap")}
        >
          <WrapText className="size-3.5" />
        </Toggle>
      </Group>

      {/* The review fill, as Excel's paint bucket: the bucket itself fills the
          selection with the colour it is holding, and the caret beside it opens
          the palette. Its colours are the app's own — the file is never
          written to. */}
      <Group label={t("ribbon.color")}>
        <div className="flex items-stretch">
          <Hint label={t("toolbar.fill")}>
            <Button
              size="sm"
              variant="ghost"
              className="h-6 rounded-r-none px-1.5"
              onClick={() => onMark(markColor)}
            >
              <span className="flex flex-col items-center justify-center">
                <PaintBucket className="size-3.5" />
                <span
                  aria-hidden
                  className="mt-px h-[3px] w-4 rounded-[1px] ring-1 ring-foreground/20 ring-inset"
                  style={{ background: markColor }}
                />
              </span>
            </Button>
          </Hint>
          <DropdownMenu>
            <Hint label={t("toolbar.fill")}>
              <DropdownMenuTrigger asChild>
                <Button
                  size="sm"
                  variant="ghost"
                  className="h-6 rounded-l-none px-0.5"
                >
                  <ChevronDown className="size-3" />
                </Button>
              </DropdownMenuTrigger>
            </Hint>
            <DropdownMenuContent align="start" className="min-w-[13.5rem]">
              <ColorPalette variant="dropdown" current={markColor} onPick={onMark} />
            </DropdownMenuContent>
          </DropdownMenu>
        </div>
      </Group>

      {/* Data: the filter switch and the outline groups, where Excel keeps
          them. Both are session-only, like everything else on this strip. */}
      <Group label={t("ribbon.data")}>
        <Toggle
          pressed={filterOn}
          onClick={onToggleFilter}
          title={filterOn ? t("ribbon.clearFilter") : t("ribbon.filter")}
        >
          <ListFilter className="size-3.5" />
        </Toggle>

        <DropdownMenu>
          <Hint label={t("ribbon.group")}>
            <DropdownMenuTrigger asChild>
              <Button
                size="sm"
                variant="ghost"
                className={cn("h-6 gap-1 px-1.5 text-[11px] font-normal")}
              >
                <ListTree className="size-3.5" />
                <ChevronDown className="size-3" />
              </Button>
            </DropdownMenuTrigger>
          </Hint>
          <DropdownMenuContent align="start" className="min-w-40">
            <DropdownMenuItem
              className="text-xs"
              icon={<Rows3 />}
              onSelect={() => onGroup("rows")}
            >
              {t("ribbon.groupRows")}
            </DropdownMenuItem>
            <DropdownMenuItem
              className="text-xs"
              icon={<Columns3 />}
              onSelect={() => onGroup("cols")}
            >
              {t("ribbon.groupCols")}
            </DropdownMenuItem>
            <DropdownMenuSeparator />
            <DropdownMenuItem
              className="text-xs"
              icon={<Ungroup />}
              onSelect={() => onGroup("ungroup")}
            >
              {t("ribbon.ungroup")}
            </DropdownMenuItem>
          </DropdownMenuContent>
        </DropdownMenu>
      </Group>
    </div>
  );
}
