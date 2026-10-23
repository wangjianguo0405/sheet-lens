import type { ReactElement, ReactNode } from "react";

import { Tooltip, TooltipContent, TooltipTrigger } from "@/components/ui/tooltip";

/**
 * The app's hover hint, in place of the browser's `title`.
 *
 * A `title` waits about a second before it shows, cannot be styled, and reads
 * as the operating system rather than as this app. This one appears at once and
 * is drawn by the app; the wrapped element keeps its own `aria-label`, and the
 * tooltip supplies the description for the ones that have none.
 *
 * Only the *action* of a control goes in here. A `title` that repeats text the
 * element already shows — a truncated file name, a cell's note — is not a hint
 * and stays on the element.
 */
export function Hint({
  label,
  children,
}: {
  label: ReactNode;
  children: ReactElement;
}) {
  if (label === null || label === undefined || label === "") return children;
  return (
    <Tooltip>
      <TooltipTrigger asChild>{children}</TooltipTrigger>
      <TooltipContent>{label}</TooltipContent>
    </Tooltip>
  );
}
