/**
 * The hidden readouts, off by default and toggled with Ctrl+Shift+D.
 *
 * They are deliberately kept out of the way rather than deleted: the layout
 * questions they answer — which column offsets a shape was resolved from, what
 * its box left for text, which font it inherited — come up again while the grid
 * is being worked on, and the numbers are easier to read off a panel than to
 * reconstruct.
 *
 * Dev builds only. The shortcut that flips them sits behind
 * `import.meta.env.DEV` in `App.tsx`, so a packaged build cannot bring them up
 * at all.
 */
let enabled = false;
const listeners = new Set<() => void>();

export function shapeDebugEnabled(): boolean {
  return enabled;
}

export function toggleShapeDebug(): void {
  enabled = !enabled;
  for (const listener of listeners) listener();
}

export function subscribeShapeDebug(listener: () => void): () => void {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}
