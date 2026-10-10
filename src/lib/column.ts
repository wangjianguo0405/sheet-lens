export function columnLabel(index: number): string {
  let n = index + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

export function columnIndex(label: string): number {
  let n = 0;
  for (const ch of label.toUpperCase()) {
    const v = ch.charCodeAt(0) - 64;
    if (v < 1 || v > 26) return -1;
    n = n * 26 + v;
  }
  return n - 1;
}

export function cellRef(row: number, col: number): string {
  return `${columnLabel(col)}${row + 1}`;
}

/** Accepts `G408`, `$G$408` and `Sheet1!G408` (the sheet part is ignored). */
export function parseRef(input: string): { row: number; col: number } | null {
  const target = input.includes("!") ? input.slice(input.lastIndexOf("!") + 1) : input;
  const match = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(target.trim());
  if (!match) return null;
  const col = columnIndex(match[1]);
  if (col < 0) return null;
  const row = Number(match[2]) - 1;
  if (row < 0) return null;
  return { row, col };
}
