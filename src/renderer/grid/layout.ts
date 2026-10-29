import { Sheet } from '../../core/model';

/** Excel column width (characters) to pixels at 100%, as Excel computes it for Calibri 11. */
export function colPx(width: number): number { return Math.round(width * 7 + 5); }
export function pxToColWidth(px: number): number { return Math.max(0, Math.round(((px - 5) / 7) * 100) / 100); }
export function rowPx(points: number): number { return Math.round((points * 96) / 72); }
export function pxToRowHeight(px: number): number { return Math.round(((px * 72) / 96) * 100) / 100; }

/**
 * Pixel offsets along one axis (rows or columns): a default size, and a sparse
 * set of indexes with their own size (0 when hidden). Lookups are binary
 * searches, so a million rows cost nothing until they're customized.
 */
export class Axis {
  private keys: number[];
  private prefix: number[];

  constructor(public count: number, readonly def: number, private custom: Map<number, number>) {
    this.keys = [...custom.keys()].sort((a, b) => a - b);
    this.prefix = [0];
    for (const k of this.keys) this.prefix.push(this.prefix[this.prefix.length - 1] + custom.get(k)! - def);
  }

  size(i: number): number {
    const s = this.custom.get(i);
    return s === undefined ? this.def : s;
  }

  /** Pixel position where index `i` starts. */
  offset(i: number): number {
    let lo = 0, hi = this.keys.length;
    while (lo < hi) { const m = (lo + hi) >> 1; if (this.keys[m] < i) lo = m + 1; else hi = m; }
    return i * this.def + this.prefix[lo];
  }

  total(): number { return this.offset(this.count); }

  /** Index whose span contains `px` (clamped to the axis). Zero-size indexes are never returned. */
  indexAt(px: number): number {
    if (px <= 0) return this.nextVisible(0);
    let lo = 0, hi = this.count - 1;
    while (lo < hi) {
      const m = (lo + hi + 1) >> 1;
      if (this.offset(m) <= px) lo = m; else hi = m - 1;
    }
    return this.size(lo) ? lo : this.nextVisible(lo);
  }

  nextVisible(i: number): number {
    while (i < this.count - 1 && !this.size(i)) i++;
    return i;
  }

  prevVisible(i: number): number {
    while (i > 0 && !this.size(i)) i--;
    return i;
  }
}

export function rowAxis(sheet: Sheet, count: number, zoom: number): Axis {
  const custom = new Map<number, number>();
  const def = Math.round(rowPx(sheet.defaultRowHeight) * zoom);
  sheet.rows.forEach((r, i) => {
    if (!r) return;
    if (r.hidden || r.filtered) custom.set(i, 0);
    else if (r.height !== undefined) custom.set(i, Math.max(1, Math.round(rowPx(r.height) * zoom)));
  });
  return new Axis(count, def, custom);
}

export function colAxis(sheet: Sheet, count: number, zoom: number): Axis {
  const custom = new Map<number, number>();
  const def = Math.round(colPx(sheet.defaultColWidth) * zoom);
  sheet.cols.forEach((c, i) => {
    if (!c) return;
    if (c.hidden) custom.set(i, 0);
    else if (c.width !== undefined) custom.set(i, Math.round(colPx(c.width) * zoom));
  });
  return new Axis(count, def, custom);
}
