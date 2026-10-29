import { MAX_COLS, MAX_ROWS, normRange, Range } from '../../core/addr';
import { Sheet } from '../../core/model';

export interface Pos { row: number; col: number }

/** Grows a range until no merged block is cut by its edge. */
export function expandMerges(sheet: Sheet, r: Range): Range {
  let out = { ...r };
  for (let changed = true; changed;) {
    changed = false;
    for (const m of sheet.merges) {
      if (m.r2 < out.r1 || m.r1 > out.r2 || m.c2 < out.c1 || m.c1 > out.c2) continue;
      const n = { r1: Math.min(out.r1, m.r1), c1: Math.min(out.c1, m.c1), r2: Math.max(out.r2, m.r2), c2: Math.max(out.c2, m.c2) };
      if (n.r1 !== out.r1 || n.c1 !== out.c1 || n.r2 !== out.r2 || n.c2 !== out.c2) { out = n; changed = true; }
    }
  }
  return out;
}

export function mergeAt(sheet: Sheet, row: number, col: number): Range | undefined {
  return sheet.merges.find((m) => row >= m.r1 && row <= m.r2 && col >= m.c1 && col <= m.c2);
}

/**
 * The selected cells: the range from `anchor` (the active cell) to `head`,
 * plus earlier ranges kept with Cmd/Ctrl-click.
 */
export class Selection {
  anchor: Pos = { row: 0, col: 0 };
  head: Pos = { row: 0, col: 0 };
  extra: Range[] = [];

  constructor(private sheet: () => Sheet) {}

  primary(): Range {
    return expandMerges(this.sheet(), normRange({ r1: this.anchor.row, c1: this.anchor.col, r2: this.head.row, c2: this.head.col }));
  }

  ranges(): Range[] { return [...this.extra, this.primary()]; }

  set(anchor: Pos, head: Pos = anchor, keep = false): void {
    if (keep) this.extra.push(this.primary()); else this.extra = [];
    this.anchor = { ...anchor };
    this.head = { ...head };
  }

  extendTo(head: Pos): void { this.head = { ...head }; }

  isWholeRows(r = this.primary()): boolean { return r.c1 === 0 && r.c2 === MAX_COLS - 1; }
  isWholeCols(r = this.primary()): boolean { return r.r1 === 0 && r.r2 === MAX_ROWS - 1; }
  isSingleCell(): boolean {
    const r = this.primary();
    if (this.extra.length) return false;
    if (r.r1 === r.r2 && r.c1 === r.c2) return true;
    const m = mergeAt(this.sheet(), r.r1, r.c1);
    return !!m && m.r1 === r.r1 && m.c1 === r.c1 && m.r2 === r.r2 && m.c2 === r.c2;
  }

  contains(row: number, col: number): boolean {
    return this.ranges().some((r) => row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2);
  }
}
