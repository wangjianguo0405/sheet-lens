import type { KeyboardEvent } from "react";

/**
 * Excel's menu accelerators, written into the label the way a Windows menu
 * string carries one: `非表示(&H)` draws the `H` underlined, and pressing it
 * picks the item whether or not the pointer is anywhere near.
 *
 * Shared by the two menus that carry them — a row or column heading's, and a
 * sheet tab's — so the two spell their letters the same way.
 */
export function acceleratorOf(label: string): string {
  const at = label.indexOf("&");
  return at >= 0 && at + 1 < label.length ? label[at + 1].toUpperCase() : "";
}

export function Accelerated({ label }: { label: string }) {
  const at = label.indexOf("&");
  if (at < 0 || at + 1 >= label.length) return <>{label}</>;
  return (
    <>
      {label.slice(0, at)}
      <span className="underline">{label[at + 1]}</span>
      {label.slice(at + 2)}
    </>
  );
}

/** Pressing an item's underlined letter picks it, as Excel's menus do. */
export function pickByAccelerator(event: KeyboardEvent<HTMLElement>): void {
  if (event.altKey || event.ctrlKey || event.metaKey || event.key.length !== 1) return;
  const item = event.currentTarget.querySelector(
    `[data-accelerator="${event.key.toUpperCase()}"]`,
  );
  if (!item) return;
  event.preventDefault();
  (item as HTMLElement).click();
}
