/**
 * Writes text to the system clipboard.
 *
 * The async clipboard API needs a permission the Tauri webview does not always
 * grant, so this falls back to the old selection-based command, which works
 * without one.
 */
export async function copyText(text: string): Promise<boolean> {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    // The fallback has to move the focus onto its scratch textarea — that is what
    // makes `execCommand` copy at all — so removing the textarea afterwards
    // leaves the focus on the body. Everything the app binds to the sheet then
    // stops answering until something is clicked, which reads as the shortcut
    // that just ran being the last one that works. Focus goes back where it was.
    const previous = document.activeElement as HTMLElement | null;
    try {
      const area = document.createElement("textarea");
      area.value = text;
      area.setAttribute("readonly", "");
      area.style.position = "fixed";
      area.style.top = "-1000px";
      document.body.appendChild(area);
      area.select();
      const copied = document.execCommand("copy");
      document.body.removeChild(area);
      return copied;
    } catch {
      return false;
    } finally {
      if (previous?.isConnected) previous.focus({ preventScroll: true });
    }
  }
}
