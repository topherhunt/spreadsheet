import { MAX_COLS, MAX_ROWS, parseRange, Range, rangeToA1 } from './addr';
import { renameSheetInFormula, invalidateSheetInFormula, shiftFormula, StructuralChange } from './formula/refs';
import { Cell, newSheet, Sheet, Workbook } from './model';

// Row/column insertion and deletion, and sheet add/rename/delete/move, applied
// to the model. Formulas everywhere in the workbook follow the change. The
// engine is rebuilt by the caller afterwards.

/** Shifts one axis of a range; null when the range is entirely deleted. */
export function shiftSpan(lo: number, hi: number, at: number, count: number): [number, number] | null {
  if (count > 0) return [lo >= at ? lo + count : lo, hi >= at ? hi + count : hi];
  const delEnd = at - count - 1;
  if (lo >= at && hi <= delEnd) return null;
  const shift = (x: number) => (x > delEnd ? x + count : x >= at ? at : x);
  const nlo = lo >= at && lo <= delEnd ? at : shift(lo);
  const nhi = hi >= at && hi <= delEnd ? at - 1 : shift(hi);
  return [nlo, Math.max(nlo, nhi)];
}

export function shiftRange(r: Range, axis: 'row' | 'col', at: number, count: number): Range | null {
  if (axis === 'row') {
    const s = shiftSpan(r.r1, r.r2, at, count);
    return s && { ...r, r1: s[0], r2: s[1] };
  }
  const s = shiftSpan(r.c1, r.c2, at, count);
  return s && { ...r, c1: s[0], c2: s[1] };
}

function eachFormulaCell(wb: Workbook, fn: (c: Cell, sheet: Sheet) => void): void {
  for (const s of wb.sheets) for (const r of s.rows) if (r) for (const c of r.cells) if (c?.f) fn(c, s);
}

function shiftArray<T>(arr: (T | undefined)[], at: number, count: number): (T | undefined)[] {
  if (arr.length <= at) return arr;
  const head = arr.slice(0, at);
  const tail = count > 0 ? arr.slice(at) : arr.slice(at - count);
  const out: (T | undefined)[] = head;
  out.length = Math.max(0, at + (count > 0 ? count : 0));
  // Keep holes as holes: copying index by index preserves sparseness.
  tail.forEach((v, i) => { out[at + (count > 0 ? count : 0) + i] = v; });
  return out;
}

/** Inserts (count > 0) or deletes (count < 0) rows or columns starting at `at`. */
export function shiftCells(wb: Workbook, si: number, axis: 'row' | 'col', at: number, count: number): void {
  const sheet = wb.sheets[si];
  if (count > 0) {
    const pushedOff = axis === 'row' ? sheet.rows.length > MAX_ROWS - count : sheet.cols.length > MAX_COLS - count || sheet.rows.some((r) => r && r.cells.length > MAX_COLS - count);
    if (pushedOff) throw new Error(`Can’t insert ${axis === 'row' ? 'rows' : 'columns'}: that would push cells off the end of the sheet.`);
  }
  const change: StructuralChange = { sheet: sheet.name, axis, at, count };
  eachFormulaCell(wb, (c, s) => { c.f = shiftFormula(c.f!, s.name, change); });
  for (const s of wb.sheets) for (const r of s.rows) if (r) for (const c of r.cells) {
    if (c?.link?.startsWith('#')) c.link = '#' + shiftFormula(c.link.slice(1), s.name, change);
  }

  if (axis === 'row') sheet.rows = shiftArray(sheet.rows, at, count);
  else {
    for (const r of sheet.rows) if (r) r.cells = shiftArray(r.cells, at, count);
    sheet.cols = shiftArray(sheet.cols, at, count);
  }

  for (const r of sheet.rows) if (r) for (const c of r.cells) {
    if (!c?.arrayRef) continue;
    const moved = shiftRange(parseRange(c.arrayRef), axis, at, count);
    if (moved) c.arrayRef = rangeToA1(moved); else delete c.arrayRef;
  }
  sheet.merges = sheet.merges.map((m) => shiftRange(m, axis, at, count)).filter((m): m is Range => !!m && (m.r1 !== m.r2 || m.c1 !== m.c2));

  if (sheet.filter) {
    const moved = shiftRange(sheet.filter.range, axis, at, count);
    if (!moved) sheet.filter = null;
    else {
      const columns: typeof sheet.filter.columns = {};
      for (const [k, v] of Object.entries(sheet.filter.columns)) {
        const col = Number(k);
        if (axis === 'row') { columns[col] = v; continue; }
        const s = shiftSpan(col, col, at, count);
        if (s) columns[s[0]] = v;
      }
      sheet.filter = { range: moved, columns };
    }
  }

  const fz = axis === 'row' ? 'rows' : 'cols';
  if (sheet.freeze[fz] > at) {
    const s = shiftSpan(0, sheet.freeze[fz] - 1, at, count);
    sheet.freeze = { ...sheet.freeze, [fz]: s ? s[1] + 1 : at };
  }

  for (const n of wb.definedNames) {
    const scope = n.localSheet !== undefined ? wb.sheets[n.localSheet].name : sheet.name;
    n.ref = shiftFormula(n.ref, scope, change);
  }
  sheet.pristine = false;
  markFormulaSheetsChanged(wb);
}

function markFormulaSheetsChanged(wb: Workbook): void {
  // A shifted formula elsewhere makes that sheet's XML stale too.
  for (const s of wb.sheets) if (s.kind === 'worksheet' && s.rows.some((r) => r?.cells.some((c) => c?.f))) s.pristine = false;
}

export function uniqueSheetName(wb: Workbook, base: string): string {
  const taken = new Set(wb.sheets.map((s) => s.name.toLowerCase()));
  if (!taken.has(base.toLowerCase())) return base;
  for (let i = 2; ; i++) if (!taken.has(`${base} ${i}`.toLowerCase())) return `${base} ${i}`;
}

export function validSheetName(wb: Workbook, name: string, except?: number): string | null {
  if (!name.trim()) return 'A sheet name can’t be empty.';
  if (name.length > 31) return 'Sheet names can be at most 31 characters.';
  if (/[\\/?*[\]:]/.test(name)) return 'Sheet names can’t contain \\ / ? * [ ] or :';
  if (/^'|'$/.test(name)) return 'Sheet names can’t start or end with an apostrophe.';
  if (wb.sheets.some((s, i) => i !== except && s.name.toLowerCase() === name.toLowerCase())) return 'Another sheet already has that name.';
  return null;
}

export function addSheet(wb: Workbook, at: number, name = uniqueSheetName(wb, `Sheet${wb.sheets.length + 1}`)): void {
  wb.sheets.splice(at, 0, newSheet(name));
  for (const n of wb.definedNames) if (n.localSheet !== undefined && n.localSheet >= at) n.localSheet++;
  if (wb.activeSheet >= at) wb.activeSheet++;
}

export function renameSheet(wb: Workbook, si: number, name: string): void {
  const err = validSheetName(wb, name, si);
  if (err) throw new Error(err);
  const old = wb.sheets[si].name;
  eachFormulaCell(wb, (c) => { c.f = renameSheetInFormula(c.f!, old, name); });
  for (const n of wb.definedNames) n.ref = renameSheetInFormula(n.ref, old, name);
  for (const s of wb.sheets) for (const r of s.rows) if (r) for (const c of r.cells) {
    if (c?.link?.startsWith('#')) c.link = '#' + renameSheetInFormula(c.link.slice(1), old, name);
  }
  wb.sheets[si].name = name;
  wb.sheets[si].pristine = false;
  markFormulaSheetsChanged(wb);
}

export function deleteSheet(wb: Workbook, si: number): void {
  if (wb.sheets.filter((s) => s.state === 'visible').length <= 1 && wb.sheets[si].state === 'visible') throw new Error('A workbook needs at least one visible sheet.');
  const name = wb.sheets[si].name;
  wb.sheets.splice(si, 1);
  eachFormulaCell(wb, (c) => { c.f = invalidateSheetInFormula(c.f!, name); });
  wb.definedNames = wb.definedNames.filter((n) => n.localSheet !== si);
  for (const n of wb.definedNames) {
    if (n.localSheet !== undefined && n.localSheet > si) n.localSheet--;
    n.ref = invalidateSheetInFormula(n.ref, name);
  }
  if (wb.activeSheet >= wb.sheets.length || wb.activeSheet > si) wb.activeSheet = Math.max(0, wb.activeSheet - 1);
  while (wb.sheets[wb.activeSheet].state !== 'visible') wb.activeSheet = (wb.activeSheet + 1) % wb.sheets.length;
  markFormulaSheetsChanged(wb);
}

export function moveSheet(wb: Workbook, from: number, to: number): void {
  const order = wb.sheets.map((_, i) => i);
  const [moved] = order.splice(from, 1);
  order.splice(to, 0, moved);
  const newIndex = new Map(order.map((old, i) => [old, i]));
  wb.sheets = order.map((i) => wb.sheets[i]);
  for (const n of wb.definedNames) if (n.localSheet !== undefined) n.localSheet = newIndex.get(n.localSheet)!;
  wb.activeSheet = newIndex.get(wb.activeSheet)!;
}
