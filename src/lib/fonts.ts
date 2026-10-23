import { invoke } from "@tauri-apps/api/core";

/**
 * Every font family this machine has, as the system reports them.
 *
 * The webview cannot enumerate installed fonts — `document.fonts` only knows
 * the faces a page loaded — so the backend asks GDI. An empty answer (a build
 * without that half, or a run outside Tauri) leaves the caller to its own list.
 */
export async function listFonts(): Promise<string[]> {
  return invoke<string[]>("list_fonts");
}
