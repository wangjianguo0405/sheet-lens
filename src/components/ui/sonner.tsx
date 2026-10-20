import type { CSSProperties } from "react";
import { Toaster as Sonner, type ToasterProps } from "sonner";

/**
 * Sonner's toasts, painted in the app's own tokens rather than its default
 * greys, so a notice reads as part of the window. Sits just above the status
 * bar at the right, where the eye already is after a menu action.
 *
 * The theme is fixed to light: the app never puts `.dark` on the document, so
 * following the system here would hand a dark toast to a light window.
 */
export function Toaster(props: ToasterProps) {
  return (
    <Sonner
      theme="light"
      position="bottom-right"
      offset={{ bottom: 36, right: 12 }}
      className="toaster group"
      style={
        {
          "--normal-bg": "var(--popover)",
          "--normal-text": "var(--popover-foreground)",
          "--normal-border": "var(--border)",
          "--border-radius": "var(--radius-md)",
        } as CSSProperties
      }
      toastOptions={{
        classNames: {
          toast: "text-xs shadow-lg",
          title: "text-xs font-medium",
          description: "text-xs text-muted-foreground",
          actionButton: "text-xs",
          cancelButton: "text-xs",
        },
      }}
      {...props}
    />
  );
}
