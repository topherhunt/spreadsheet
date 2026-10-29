// A1-style addressing. Rows and columns are 0-based everywhere in the app; only
// the A1 strings are 1-based.

export interface CellAddr { row: number; col: number }
export interface Range { r1: number; c1: number; r2: number; c2: number }

export const MAX_ROWS = 1048576;
export const MAX_COLS = 16384;

export function colToName(col: number): string {
  let n = col + 1;
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function nameToCol(name: string): number {
  let n = 0;
  for (const ch of name.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export function toA1(row: number, col: number): string {
  return colToName(col) + (row + 1);
}

export function parseA1(ref: string): CellAddr {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref);
  if (!m) throw new Error(`Not a cell reference: ${ref}`);
  return { row: parseInt(m[2], 10) - 1, col: nameToCol(m[1]) };
}

/** Parses "A1:C3", "A1", "A:C" or "2:5" into a normalized range. */
export function parseRange(ref: string): Range {
  const [a, b] = ref.split(':');
  if (b === undefined) {
    const p = parseA1(a);
    return { r1: p.row, c1: p.col, r2: p.row, c2: p.col };
  }
  const whole = (x: string) => x.replace(/\$/g, '');
  const wa = whole(a), wb = whole(b);
  if (/^[A-Za-z]+$/.test(wa) && /^[A-Za-z]+$/.test(wb)) {
    return normRange({ r1: 0, c1: nameToCol(wa), r2: MAX_ROWS - 1, c2: nameToCol(wb) });
  }
  if (/^\d+$/.test(wa) && /^\d+$/.test(wb)) {
    return normRange({ r1: parseInt(wa, 10) - 1, c1: 0, r2: parseInt(wb, 10) - 1, c2: MAX_COLS - 1 });
  }
  const p = parseA1(a), q = parseA1(b);
  return normRange({ r1: p.row, c1: p.col, r2: q.row, c2: q.col });
}

export function rangeToA1(r: Range): string {
  if (r.r1 === r.r2 && r.c1 === r.c2) return toA1(r.r1, r.c1);
  return `${toA1(r.r1, r.c1)}:${toA1(r.r2, r.c2)}`;
}

export function normRange(r: Range): Range {
  return {
    r1: Math.min(r.r1, r.r2), c1: Math.min(r.c1, r.c2),
    r2: Math.max(r.r1, r.r2), c2: Math.max(r.c1, r.c2),
  };
}

export function inRange(r: Range, row: number, col: number): boolean {
  return row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2;
}

export function rangesIntersect(a: Range, b: Range): boolean {
  return a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
}

/** Space-separated list of ranges, as in sqref attributes. */
export function parseSqref(sqref: string): Range[] {
  return sqref.trim().split(/\s+/).filter(Boolean).map(parseRange);
}

export function sqrefToString(ranges: Range[]): string {
  return ranges.map(rangeToA1).join(' ');
}

/** Sheet names need quoting in formulas unless they are plain identifiers. */
export function quoteSheetName(name: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) && !/^(TRUE|FALSE)$/i.test(name)) return name;
  return `'${name.replace(/'/g, "''")}'`;
}
