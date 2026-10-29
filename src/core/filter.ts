import { Range } from './addr';
import { parseInput } from './csv';
import { Cell, ColumnFilter, ensureRow, FilterCondition, getCell, isError, Scalar, Sheet, sheetExtent, SortSpec, Workbook } from './model';
import { displayCell } from './numfmt';
import { translateFormula } from './formula/refs';

export function cellShown(wb: Workbook, c: Cell | undefined): string {
  return displayCell(wb, c).text;
}

function asNumber(c: Cell | undefined): number | null {
  return c && typeof c.v === 'number' ? c.v : null;
}

/** Whether a cell passes a condition. Text tests are case-insensitive; comparisons are numeric when both sides are numbers. */
export function meetsCondition(wb: Workbook, cond: FilterCondition, c: Cell | undefined): boolean {
  const text = cellShown(wb, c).toLowerCase();
  const a = (cond.a ?? '').toLowerCase();
  const num = asNumber(c);
  const an = parseInput(cond.a ?? ''), bn = parseInput(cond.b ?? '');
  const cmp = (x: Scalar): number | null => {
    if (num !== null && typeof x === 'number') return num - x;
    if (num !== null || typeof x === 'number') return null;
    return text.localeCompare(String(x ?? '').toLowerCase());
  };
  switch (cond.op) {
    case 'empty': return text === '';
    case 'notEmpty': return text !== '';
    case 'contains': return text.includes(a);
    case 'notContains': return !text.includes(a);
    case 'startsWith': return text.startsWith(a);
    case 'endsWith': return text.endsWith(a);
    case 'equals': return num !== null && typeof an === 'number' ? num === an : text === a;
    case 'notEquals': return num !== null && typeof an === 'number' ? num !== an : text !== a;
    case 'gt': { const d = cmp(an); return d !== null && d > 0; }
    case 'gte': { const d = cmp(an); return d !== null && d >= 0; }
    case 'lt': { const d = cmp(an); return d !== null && d < 0; }
    case 'lte': { const d = cmp(an); return d !== null && d <= 0; }
    case 'between': { const lo = cmp(an), hi = cmp(bn); return lo !== null && hi !== null && lo >= 0 && hi <= 0; }
    case 'notBetween': { const lo = cmp(an), hi = cmp(bn); return lo !== null && hi !== null && (lo < 0 || hi > 0); }
  }
}

export function passes(wb: Workbook, cf: ColumnFilter, c: Cell | undefined): boolean {
  if (cf.values && !cf.values.includes(cellShown(wb, c))) return false;
  if (cf.condition && !meetsCondition(wb, cf.condition, c)) return false;
  return true;
}

/** Grows the filter range down to the sheet's last used row, as Sheets does when data is added below it. */
export function extendFilterRange(sheet: Sheet): void {
  if (!sheet.filter) return;
  const ext = sheetExtent(sheet);
  if (ext.rows - 1 > sheet.filter.range.r2) sheet.filter.range = { ...sheet.filter.range, r2: ext.rows - 1 };
}

/** Recomputes which rows the sheet's filter hides. */
export function applyFilter(wb: Workbook, sheet: Sheet): void {
  for (const r of sheet.rows) if (r?.filtered) delete r.filtered;
  const f = sheet.filter;
  if (!f) return;
  const cols = Object.entries(f.columns).map(([k, v]) => [Number(k), v] as const);
  if (!cols.length) return;
  for (let r = f.range.r1 + 1; r <= f.range.r2; r++) {
    const ok = cols.every(([col, cf]) => passes(wb, cf, getCell(sheet, r, col)));
    if (!ok) ensureRow(sheet, r).filtered = true;
  }
  sheet.pristine = false;
}

/** Distinct displayed values of a filter column, for the filter menu. Rows hidden by other columns' filters are left out. */
export function columnValues(wb: Workbook, sheet: Sheet, col: number): { text: string; count: number }[] {
  const f = sheet.filter;
  if (!f) throw new Error('Sheet has no filter');
  const others = Object.entries(f.columns).filter(([k]) => Number(k) !== col).map(([k, v]) => [Number(k), v] as const);
  const counts = new Map<string, { count: number; sort: Scalar }>();
  for (let r = f.range.r1 + 1; r <= f.range.r2; r++) {
    if (!others.every(([c, cf]) => passes(wb, cf, getCell(sheet, r, c)))) continue;
    const cell = getCell(sheet, r, col);
    const text = cellShown(wb, cell);
    const e = counts.get(text);
    if (e) e.count++; else counts.set(text, { count: 1, sort: cell?.v ?? null });
  }
  return [...counts.entries()].sort((a, b) => compareValues(a[1].sort, b[1].sort)).map(([text, { count }]) => ({ text, count }));
}

function rank(v: Scalar): number {
  if (typeof v === 'number') return 0;
  if (typeof v === 'string') return v === '' ? 4 : 1;
  if (typeof v === 'boolean') return 2;
  if (isError(v)) return 3;
  return 4;
}

const collator = new Intl.Collator(undefined, { sensitivity: 'base', numeric: true });

/** Spreadsheet sort order: numbers, text, booleans, errors, then blanks. */
export function compareValues(a: Scalar, b: Scalar): number {
  const ra = rank(a), rb = rank(b);
  if (ra !== rb) return ra - rb;
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  if (typeof a === 'string' && typeof b === 'string') return collator.compare(a, b);
  if (typeof a === 'boolean' && typeof b === 'boolean') return Number(a) - Number(b);
  return 0;
}

/**
 * Sorts the rows of `range` (excluding a header row when `header`). Only cells
 * inside the range's columns move; formulas are translated so references to
 * their own row keep pointing at it. Blanks stay last in both directions.
 */
export function sortRange(sheet: Sheet, range: Range, specs: SortSpec[], header: boolean): void {
  const top = range.r1 + (header ? 1 : 0);
  if (top >= range.r2) return;
  const rows: { from: number; cells: (Cell | undefined)[] }[] = [];
  for (let r = top; r <= range.r2; r++) {
    const cells: (Cell | undefined)[] = [];
    for (let c = range.c1; c <= range.c2; c++) cells.push(getCell(sheet, r, c));
    rows.push({ from: r, cells });
  }
  rows.sort((x, y) => {
    for (const s of specs) {
      const a = x.cells[s.col - range.c1]?.v ?? null, b = y.cells[s.col - range.c1]?.v ?? null;
      const blankA = rank(a) === 4, blankB = rank(b) === 4;
      if (blankA !== blankB) return blankA ? 1 : -1;
      const d = compareValues(a, b);
      if (d) return s.desc ? -d : d;
    }
    return x.from - y.from;
  });
  rows.forEach((row, i) => {
    const to = top + i;
    row.cells.forEach((c, j) => {
      const col = range.c1 + j;
      if (c && c.f && to !== row.from) c.f = translateFormula(c.f, to - row.from, 0);
      if (c) ensureRow(sheet, to).cells[col] = c;
      else { const r = sheet.rows[to]; if (r) delete r.cells[col]; }
    });
  });
  sheet.pristine = false;
}

/** The block of data around the used area: what "filter"/"auto-table" apply to when one cell is selected. */
export function dataRegion(sheet: Sheet): Range | null {
  const ext = sheetExtent(sheet);
  if (!ext.rows || !ext.cols) return null;
  let r1 = 0;
  while (r1 < ext.rows && !sheet.rows[r1]?.cells.some((c) => c && c.v !== null)) r1++;
  let c1 = 0;
  while (c1 < ext.cols && !sheet.rows.some((r) => { const c = r?.cells[c1]; return c && c.v !== null; })) c1++;
  if (r1 >= ext.rows || c1 >= ext.cols) return null;
  return { r1, c1, r2: ext.rows - 1, c2: ext.cols - 1 };
}
