import { normRange, Range } from './addr';
import { applyFilter, dataRegion, extendFilterRange, sortRange } from './filter';
import { Engine } from './formula/engine';
import { moveRefsInFormula, tokenize, translateFormula } from './formula/refs';
import { fileFormula } from './formula/xlfn';
import { parseTyped } from './input';
import { Cell, cloneCell, ColumnFilter, ensureRow, getCell, Run, Sheet, SortSpec, Workbook } from './model';
import { cellFormat, isDateFormat } from './numfmt';
import { deriveXf, StylePatch } from './styles';
import * as structure from './structure';

export type ChangeKind = 'cells' | 'layout' | 'sheets';
type Listener = (kind: ChangeKind, sheet: number) => void;

interface Step { label: string; sheet: number; undo(): void; redo(): void }

/** Deep copy of a sheet's editable state (everything but the source package it came from). */
function cloneSheetState(s: Sheet): Sheet {
  return {
    ...s,
    rows: s.rows.map((r) => r && { ...r, cells: r.cells.map((c) => c && cloneCell(c)), keep: r.keep && { ...r.keep } }),
    cols: s.cols.map((c) => c && { ...c, keep: c.keep && { ...c.keep } }),
    merges: s.merges.map((m) => ({ ...m })),
    freeze: { ...s.freeze },
    filter: s.filter && structuredClone(s.filter),
    features: new Set(s.features),
  };
}

const DATE_TYPES: Record<string, string> = { NUMBER_DATE: 'm/d/yyyy', NUMBER_DATETIME: 'm/d/yyyy h:mm:ss', NUMBER_TIME: 'h:mm:ss' };

/**
 * An open workbook: the model, its formula engine, and undo history. Every
 * edit the UI makes goes through here.
 */
export class Doc {
  readonly engine: Engine;
  dirty = false;
  private undoStack: Step[] = [];
  private redoStack: Step[] = [];
  private listeners: Listener[] = [];

  constructor(public wb: Workbook) {
    this.engine = new Engine(wb);
  }

  onChange(fn: Listener): void { this.listeners.push(fn); }

  private emit(kind: ChangeKind, sheet: number): void {
    for (const fn of this.listeners) fn(kind, sheet);
  }

  private record(step: Step, kind: ChangeKind): void {
    this.undoStack.push(step);
    if (this.undoStack.length > 500) this.undoStack.shift();
    this.redoStack = [];
    this.dirty = true;
    this.emit(kind, step.sheet);
  }

  canUndo(): boolean { return this.undoStack.length > 0; }
  canRedo(): boolean { return this.redoStack.length > 0; }

  undo(): Step | undefined {
    const s = this.undoStack.pop();
    if (!s) return;
    s.undo();
    this.redoStack.push(s);
    this.dirty = true;
    this.emit('sheets', s.sheet);
    return s;
  }

  redo(): Step | undefined {
    const s = this.redoStack.pop();
    if (!s) return;
    s.redo();
    this.undoStack.push(s);
    this.dirty = true;
    this.emit('sheets', s.sheet);
    return s;
  }

  // -------------------------------------------------------------------------
  // Cells

  /** Replaces cells, recalculating and recording one undo step. Returns the step's redo cells. */
  putCells(sheet: number, cells: { row: number; col: number; cell: Cell | undefined }[], label: string): { cell: Cell | undefined }[] {
    return this.putEntries(sheet, cells.map((e) => ({ ...e, sheet })), label);
  }

  /** putCells across sheets; `sheet` is the one the undo step returns to. */
  private putEntries(sheet: number, after: { sheet: number; row: number; col: number; cell: Cell | undefined }[], label: string): { cell: Cell | undefined }[] {
    if (!after.length) return [];
    const before = after.map(({ sheet: si, row, col }) => {
      const c = getCell(this.wb.sheets[si], row, col);
      return { sheet: si, row, col, cell: c && cloneCell(c) };
    });
    this.engine.apply(after.map((e) => ({ ...e, cell: e.cell && cloneCell(e.cell) })));
    this.record({
      label, sheet,
      undo: () => this.engine.apply(before.map((e) => ({ ...e, cell: e.cell && cloneCell(e.cell) }))),
      redo: () => this.engine.apply(after.map((e) => ({ ...e, cell: e.cell && cloneCell(e.cell) }))),
    }, 'cells');
    return after;
  }

  /** The cell that typing `text` into an existing cell produces. */
  typedCell(sheet: number, row: number, col: number, text: string): Cell | undefined {
    const ex = getCell(this.wb.sheets[sheet], row, col);
    if (text === '') return ex?.s || ex?.link ? { v: null, s: ex.s, link: ex.link } : undefined;
    const t = parseTyped(text, this.wb.date1904);
    const cell: Cell = { v: t.v };
    if (ex?.s) cell.s = ex.s;
    if (ex?.link) cell.link = ex.link;
    if (t.formula !== undefined) cell.f = fileFormula(t.formula);
    if (t.fmt && cellFormat(this.wb, ex) === 'General') cell.s = deriveXf(this.wb.styles, cell.s, { numFmt: t.fmt });
    return cell;
  }

  setInput(sheet: number, row: number, col: number, text: string): void {
    const [after] = this.putCells(sheet, [{ row, col, cell: this.typedCell(sheet, row, col, text) }], 'Edit cell');
    const s = this.autoFormatFormula(sheet, row, col);
    if (s !== undefined) after.cell!.s = s;
  }

  /**
   * Gives a new formula returning a date a date format, as Sheets does: from
   * the function's result type, or else from the first referenced cell.
   */
  private autoFormatFormula(sheet: number, row: number, col: number): number | undefined {
    const c = getCell(this.wb.sheets[sheet], row, col);
    if (!c?.f || typeof c.v !== 'number' || cellFormat(this.wb, c) !== 'General') return;
    let fmt: string | undefined = DATE_TYPES[this.engine.valueType(sheet, row, col)];
    if (!fmt) {
      const ref = tokenize(c.f).find((t) => t.type === 'ref');
      if (ref?.type === 'ref' && !ref.error && ref.a.row !== undefined && ref.a.col !== undefined) {
        const si = ref.sheet === undefined ? sheet : this.wb.sheets.findIndex((s) => s.name.toLowerCase() === ref.sheet!.toLowerCase());
        const code = si >= 0 ? cellFormat(this.wb, getCell(this.wb.sheets[si], ref.a.row, ref.a.col)) : 'General';
        if (isDateFormat(code)) fmt = code;
      }
    }
    if (!fmt) return;
    c.s = deriveXf(this.wb.styles, c.s, { numFmt: fmt });
    return c.s;
  }

  /** Writes a block of typed text (a paste from another app) starting at row/col. */
  pasteText(sheet: number, row: number, col: number, grid: string[][]): void {
    const out: { row: number; col: number; cell: Cell | undefined }[] = [];
    grid.forEach((line, y) => line.forEach((text, x) => out.push({ row: row + y, col: col + x, cell: this.typedCell(sheet, row + y, col + x, text) })));
    this.putCells(sheet, out, 'Paste');
  }

  /**
   * Pastes cells copied inside the app; formulas move relative to where they
   * came from. `size` larger than the grid repeats it (one copied cell pasted
   * over a selection).
   */
  pasteCells(sheet: number, row: number, col: number, grid: (Cell | undefined)[][], from: { row: number; col: number }, valuesOnly = false, size = { rows: grid.length, cols: grid[0].length }): void {
    const out: { row: number; col: number; cell: Cell | undefined }[] = [];
    const h = grid.length, w = grid[0].length;
    for (let y = 0; y < size.rows; y++) for (let x = 0; x < size.cols; x++) {
      const src = grid[y % h][x % w];
      let cell = src && cloneCell(src);
      if (cell) {
        delete cell.spill; delete cell.sst; delete cell.frozen; delete cell.keep; delete cell.arrayRef;
        if (valuesOnly) { delete cell.f; delete cell.rich; delete cell.s; delete cell.link; delete cell.raw; }
        else if (cell.f) { cell.f = translateFormula(cell.f, row + y - (from.row + (y % h)), col + x - (from.col + (x % w))); cell.v = null; }
      }
      if (valuesOnly && cell && cell.v === null) cell = undefined;
      out.push({ row: row + y, col: col + x, cell });
    }
    this.putCells(sheet, out, 'Paste');
  }

  /** Copies the top row/left column of a range across the rest of it (Ctrl+D / Ctrl+R). */
  fill(sheet: number, range: Range, dir: 'down' | 'right'): void {
    const s = this.wb.sheets[sheet];
    const out: { row: number; col: number; cell: Cell | undefined }[] = [];
    for (let r = range.r1; r <= range.r2; r++) for (let c = range.c1; c <= range.c2; c++) {
      const sr = dir === 'down' ? range.r1 : r, sc = dir === 'right' ? range.c1 : c;
      if (sr === r && sc === c) continue;
      const src = getCell(s, sr, sc);
      let cell = src && cloneCell(src);
      if (cell) {
        delete cell.spill; delete cell.sst; delete cell.frozen; delete cell.keep; delete cell.arrayRef; delete cell.raw;
        if (cell.f) { cell.f = translateFormula(cell.f, r - sr, c - sc); cell.v = null; }
      }
      if (cell && cell.v === null && !cell.f && !cell.s) cell = undefined;
      out.push({ row: r, col: c, cell });
    }
    this.putCells(sheet, out, dir === 'down' ? 'Fill down' : 'Fill right');
  }

  /**
   * Drag-fill from `src` across `target` (which extends src in one direction).
   * Numeric lines of two or more values with a constant step continue as a
   * series; anything else repeats, with formulas moved relative to their copy.
   */
  fillSeries(sheet: number, src: Range, target: Range): void {
    const s = this.wb.sheets[sheet];
    const vertical = target.r1 !== src.r1 || target.r2 !== src.r2;
    const n = vertical ? src.r2 - src.r1 + 1 : src.c2 - src.c1 + 1;
    const lines = vertical ? [src.c1, src.c2] : [src.r1, src.r2];
    const out: { row: number; col: number; cell: Cell | undefined }[] = [];
    for (let line = lines[0]; line <= lines[1]; line++) {
      const at = (i: number) => (vertical ? { row: src.r1 + i, col: line } : { row: line, col: src.c1 + i });
      const seq = Array.from({ length: n }, (_, i) => { const p = at(i); return getCell(s, p.row, p.col); });
      const nums = seq.map((c) => (c && !c.f && typeof c.v === 'number' ? c.v : null));
      let step: number | null = null;
      if (n >= 2 && nums.every((v) => v !== null)) {
        const d = nums[1]! - nums[0]!;
        step = nums.every((v, i) => i === 0 || Math.abs(v! - nums[i - 1]! - d) <= 1e-9) ? d : null;
      }
      const from = vertical ? target.r1 - src.r1 : target.c1 - src.c1;
      const to = vertical ? target.r2 - src.r1 : target.c2 - src.c1;
      for (let i = from; i <= to; i++) {
        if (i >= 0 && i < n) continue;
        const k = ((i % n) + n) % n;
        const p = at(i), sp = at(k);
        const base = seq[k];
        let cell = base && cloneCell(base);
        if (cell) {
          delete cell.spill; delete cell.sst; delete cell.frozen; delete cell.keep; delete cell.arrayRef; delete cell.raw;
          if (step !== null) cell.v = nums[0]! + step * i;
          else if (cell.f) { cell.f = translateFormula(cell.f, p.row - sp.row, p.col - sp.col); cell.v = null; }
        }
        out.push({ ...p, cell });
      }
    }
    this.putCells(sheet, out, 'Fill');
  }

  /**
   * Moves cells (cut and paste) and clears the source. As in Sheets, references
   * anywhere in the workbook that point wholly inside the moved block follow it,
   * including the moved formulas' own references into the block.
   */
  moveCells(sheet: number, from: Range, row: number, col: number): void {
    const s = this.wb.sheets[sheet];
    const dRow = row - from.r1, dCol = col - from.c1;
    const follow = (f: string, si: number) => moveRefsInFormula(f, this.wb.sheets[si].name, s.name, from, dRow, dCol);
    const entries = new Map<string, { sheet: number; row: number; col: number; cell: Cell | undefined }>();
    const put = (si: number, r: number, c: number, cell: Cell | undefined) => entries.set(`${si}:${r}:${c}`, { sheet: si, row: r, col: c, cell });
    this.wb.sheets.forEach((sh, si) => sh.rows.forEach((rw, r) => rw?.cells.forEach((c, ci) => {
      if (!c?.f) return;
      const f = follow(c.f, si);
      if (f !== c.f) put(si, r, ci, { ...cloneCell(c), f });
    })));
    for (let r = from.r1; r <= from.r2; r++) for (let c = from.c1; c <= from.c2; c++) put(sheet, r, c, undefined);
    for (let r = from.r1; r <= from.r2; r++) for (let c = from.c1; c <= from.c2; c++) {
      const ex = getCell(s, r, c);
      const cell = ex && !ex.spill ? cloneCell(ex) : undefined;
      if (cell) { delete cell.sst; delete cell.arrayRef; }
      if (cell?.f) cell.f = follow(cell.f, sheet);
      put(sheet, r + dRow, c + dCol, cell);
    }
    this.putEntries(sheet, [...entries.values()], 'Move');
  }

  clear(sheet: number, ranges: Range[]): void {
    const s = this.wb.sheets[sheet];
    const out: { row: number; col: number; cell: Cell | undefined }[] = [];
    for (const rg of ranges) for (let r = rg.r1; r <= rg.r2; r++) {
      const row = s.rows[r];
      if (!row) continue;
      for (let c = rg.c1; c <= Math.min(rg.c2, row.cells.length - 1); c++) {
        const ex = row.cells[c];
        if (!ex || ex.spill) continue;
        out.push({ row: r, col: c, cell: ex.s ? { v: null, s: ex.s } : undefined });
      }
    }
    this.putCells(sheet, out, 'Clear');
  }

  /** Applies formatting to every cell in the ranges, including empty ones. */
  setStyle(sheet: number, ranges: Range[], patch: StylePatch, label = 'Format'): void {
    this.setStyles(sheet, [{ ranges, patch }], label);
  }

  /** Several patches as one undo step; where parts overlap, later patches apply on top (border presets). */
  setStyles(sheet: number, parts: { ranges: Range[]; patch: StylePatch }[], label = 'Format'): void {
    const s = this.wb.sheets[sheet];
    const work = new Map<string, { row: number; col: number; cell: Cell }>();
    for (const { ranges, patch } of parts) {
      const cache = new Map<number | undefined, number>();
      for (const rg of ranges) for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
        const key = `${r}:${c}`;
        let w = work.get(key);
        if (!w) { const ex = getCell(s, r, c); w = { row: r, col: c, cell: ex ? cloneCell(ex) : { v: null } }; work.set(key, w); }
        const cell = w.cell;
        let s2 = cache.get(cell.s);
        if (s2 === undefined) { s2 = deriveXf(this.wb.styles, cell.s, patch); cache.set(cell.s, s2); }
        cell.s = s2 || undefined;
        if (cell.rich && fontPatch(patch)) cell.rich = cell.rich.map((run) => ({ ...run, font: stripFont(run.font, patch) }));
      }
    }
    this.putCells(sheet, [...work.values()], label);
  }

  /** Removes all formatting (cell styles and rich-text runs), keeping values and links. */
  clearFormat(sheet: number, ranges: Range[]): void {
    const s = this.wb.sheets[sheet];
    const out: { row: number; col: number; cell: Cell | undefined }[] = [];
    for (const rg of ranges) for (let r = rg.r1; r <= rg.r2; r++) {
      const row = s.rows[r];
      if (!row) continue;
      for (let c = rg.c1; c <= Math.min(rg.c2, row.cells.length - 1); c++) {
        const ex = row.cells[c];
        if (!ex || (!ex.s && !ex.rich)) continue;
        const cell = cloneCell(ex);
        delete cell.s; delete cell.rich; delete cell.sst;
        out.push({ row: r, col: c, cell: cell.v === null && !cell.f && !cell.link ? undefined : cell });
      }
    }
    this.putCells(sheet, out, 'Clear formatting');
  }

  /** Sets a cell's text as formatted runs. A single unformatted run becomes plain text. */
  setRich(sheet: number, row: number, col: number, runs: Run[]): void {
    const ex = getCell(this.wb.sheets[sheet], row, col);
    const text = runs.map((r) => r.text).join('');
    const plain = runs.every((r) => !r.link && (!r.font || !Object.keys(r.font).length));
    if (plain) { this.setInput(sheet, row, col, text); return; }
    const cell: Cell = { v: text, rich: runs.filter((r) => r.text) };
    if (ex?.s) cell.s = ex.s;
    // XLSX has one hyperlink per cell: the first linked run's target.
    const link = runs.find((r) => r.link)?.link ?? ex?.link;
    if (link) cell.link = link;
    this.putCells(sheet, [{ row, col, cell }], 'Edit cell');
  }

  setLink(sheet: number, row: number, col: number, url: string | undefined): void {
    const ex = getCell(this.wb.sheets[sheet], row, col);
    const cell: Cell = ex ? cloneCell(ex) : { v: url ?? null };
    if (url) cell.link = url; else delete cell.link;
    if (cell.rich) cell.rich = cell.rich.map((r) => ({ ...r, link: undefined }));
    if (url && (cell.v === null || cell.v === '')) cell.v = url;
    this.putCells(sheet, [{ row, col, cell }], url ? 'Insert link' : 'Remove link');
  }

  // -------------------------------------------------------------------------
  // Whole-sheet changes, undone by restoring snapshots

  /** Runs `fn` on one sheet (or the whole workbook with sheet -1), recording a snapshot-based undo step. */
  private snapshotStep(label: string, sheet: number, kind: ChangeKind, recalc: boolean, fn: () => void): void {
    const take = (): { sheets: Sheet[]; names: Workbook['definedNames']; active: number; edits: Workbook['edits'] } => ({
      sheets: sheet < 0 ? this.wb.sheets.map(cloneSheetState) : [cloneSheetState(this.wb.sheets[sheet])],
      names: structuredClone(this.wb.definedNames),
      active: this.wb.activeSheet,
      edits: [...this.wb.edits],
    });
    const restore = (snap: ReturnType<typeof take>) => {
      if (sheet < 0) this.wb.sheets = snap.sheets.map(cloneSheetState);
      else this.wb.sheets[sheet] = cloneSheetState(snap.sheets[0]);
      for (const s of this.wb.sheets) s.pristine = false;
      this.wb.definedNames = structuredClone(snap.names);
      this.wb.activeSheet = snap.active;
      this.wb.edits = [...snap.edits];
      if (recalc) this.engine.rebuild();
    };
    const before = take();
    fn();
    if (recalc) this.engine.rebuild();
    const after = take();
    this.record({ label, sheet: Math.max(0, sheet), undo: () => restore(before), redo: () => restore(after) }, kind);
  }

  insertRows(sheet: number, at: number, count: number): void {
    this.snapshotStep(count > 0 ? 'Insert rows' : 'Delete rows', -1, 'layout', true, () => structure.shiftCells(this.wb, sheet, 'row', at, count));
  }

  insertCols(sheet: number, at: number, count: number): void {
    this.snapshotStep(count > 0 ? 'Insert columns' : 'Delete columns', -1, 'layout', true, () => structure.shiftCells(this.wb, sheet, 'col', at, count));
  }

  addSheet(at = this.wb.activeSheet + 1): void {
    this.snapshotStep('Add sheet', -1, 'sheets', true, () => { structure.addSheet(this.wb, at); this.wb.activeSheet = at; });
  }

  renameSheet(sheet: number, name: string): void {
    const err = structure.validSheetName(this.wb, name, sheet);
    if (err) throw new Error(err);
    this.snapshotStep('Rename sheet', -1, 'sheets', true, () => structure.renameSheet(this.wb, sheet, name));
  }

  deleteSheet(sheet: number): void {
    this.snapshotStep('Delete sheet', -1, 'sheets', true, () => structure.deleteSheet(this.wb, sheet));
  }

  moveSheet(from: number, to: number): void {
    this.snapshotStep('Move sheet', -1, 'sheets', true, () => structure.moveSheet(this.wb, from, to));
  }

  /** Switching sheets is view state: it marks the document changed but isn't an undo step. */
  setActiveSheet(sheet: number): void {
    if (this.wb.activeSheet === sheet) return;
    this.wb.activeSheet = sheet;
    this.emit('sheets', sheet);
  }

  /** Sets column widths (character units): one for all, or one per column. Undefined restores the default. */
  setColWidth(sheet: number, cols: number[], width: number | undefined | number[]): void {
    this.snapshotStep('Resize columns', sheet, 'layout', false, () => {
      const s = this.wb.sheets[sheet];
      cols.forEach((c, i) => {
        const w = Array.isArray(width) ? width[i] : width;
        const info = { ...(s.cols[c] ?? {}) };
        if (w === undefined) delete info.width; else info.width = w;
        s.cols[c] = info;
      });
      s.pristine = false;
    });
  }

  /** Sets row heights (points): one for all, or one per row. Undefined restores the default. */
  setRowHeight(sheet: number, rows: number[], height: number | undefined | number[]): void {
    this.snapshotStep('Resize rows', sheet, 'layout', false, () => {
      const s = this.wb.sheets[sheet];
      rows.forEach((r, i) => {
        const hgt = Array.isArray(height) ? height[i] : height;
        const row = ensureRow(s, r);
        if (hgt === undefined) delete row.height; else row.height = hgt;
      });
      s.pristine = false;
    });
  }

  setHidden(sheet: number, axis: 'row' | 'col', indexes: number[], hidden: boolean): void {
    this.snapshotStep(hidden ? 'Hide' : 'Unhide', sheet, 'layout', false, () => {
      const s = this.wb.sheets[sheet];
      for (const i of indexes) {
        if (axis === 'row') { const r = ensureRow(s, i); if (hidden) r.hidden = true; else delete r.hidden; }
        else { const c = { ...(s.cols[i] ?? {}) }; if (hidden) c.hidden = true; else delete c.hidden; s.cols[i] = c; }
      }
      s.pristine = false;
    });
  }

  setFreeze(sheet: number, rows: number, cols: number): void {
    this.snapshotStep('Freeze', sheet, 'layout', false, () => {
      const s = this.wb.sheets[sheet];
      s.freeze = { rows, cols };
      s.pristine = false;
    });
  }

  setGridLines(sheet: number, show: boolean): void {
    this.snapshotStep(show ? 'Show gridlines' : 'Hide gridlines', sheet, 'layout', false, () => {
      const s = this.wb.sheets[sheet];
      s.showGridLines = show;
      s.pristine = false;
    });
  }

  setMerge(sheet: number, range: Range, merge: boolean): void {
    this.snapshotStep(merge ? 'Merge cells' : 'Unmerge cells', sheet, 'layout', false, () => {
      const s = this.wb.sheets[sheet];
      const r = normRange(range);
      s.merges = s.merges.filter((m) => m.r2 < r.r1 || m.r1 > r.r2 || m.c2 < r.c1 || m.c1 > r.c2);
      if (merge && (r.r1 !== r.r2 || r.c1 !== r.c2)) s.merges.push(r);
      s.pristine = false;
    });
  }

  // -------------------------------------------------------------------------
  // Filtering and sorting

  /** Turns the filter on over `range` (or the data region), or off when one exists. */
  toggleFilter(sheet: number, range?: Range): void {
    const s = this.wb.sheets[sheet];
    this.snapshotStep(s.filter ? 'Remove filter' : 'Create filter', sheet, 'layout', false, () => {
      if (s.filter) { s.filter = null; applyFilter(this.wb, s); s.pristine = false; return; }
      const rg = range && (range.r1 !== range.r2 || range.c1 !== range.c2) ? normRange(range) : dataRegion(s);
      if (!rg) throw new Error('There’s no data to filter.');
      s.filter = { range: rg, columns: {} };
      s.pristine = false;
    });
  }

  /** Sets (or with null, clears) one column's filter and re-applies the filter. */
  setColumnFilter(sheet: number, col: number, cf: ColumnFilter | null): void {
    const s = this.wb.sheets[sheet];
    if (!s.filter) throw new Error('Sheet has no filter');
    this.snapshotStep('Filter', sheet, 'layout', false, () => {
      const f = s.filter!;
      extendFilterRange(s);
      const columns = { ...f.columns };
      if (cf && (cf.values || cf.condition)) columns[col] = cf; else delete columns[col];
      s.filter = { ...f, columns };
      applyFilter(this.wb, s);
    });
  }

  clearFilters(sheet: number): void {
    const s = this.wb.sheets[sheet];
    if (!s.filter) return;
    this.snapshotStep('Clear filters', sheet, 'layout', false, () => {
      s.filter = { ...s.filter!, columns: {} };
      applyFilter(this.wb, s);
    });
  }

  sort(sheet: number, range: Range, specs: SortSpec[], header: boolean): void {
    const s = this.wb.sheets[sheet];
    this.snapshotStep('Sort', -1, 'cells', true, () => {
      if (s.filter) extendFilterRange(s);
      sortRange(s, normRange(range), specs, header);
      if (s.filter) applyFilter(this.wb, s);
    });
  }

  /** Sorts the filtered table by one column, from its filter menu. */
  sortFilterColumn(sheet: number, col: number, desc: boolean): void {
    const s = this.wb.sheets[sheet];
    if (!s.filter) throw new Error('Sheet has no filter');
    extendFilterRange(s);
    this.sort(sheet, s.filter.range, [{ col, desc }], true);
  }

  /**
   * One click for a downloaded report: bold header row, frozen header, and a
   * filter over the data. The bold header doesn't count as formatting a CSV
   * needs XLSX for.
   */
  autoTable(sheet: number): void {
    const s = this.wb.sheets[sheet];
    const rg = dataRegion(s);
    if (!rg) throw new Error('There’s no data to make a table from.');
    this.snapshotStep('Auto table', sheet, 'layout', false, () => {
      const cache = new Map<number | undefined, number>();
      for (let c = rg.c1; c <= rg.c2; c++) {
        const ex = getCell(s, rg.r1, c);
        let xf = cache.get(ex?.s);
        if (xf === undefined) { xf = deriveXf(this.wb.styles, ex?.s, { bold: true }); cache.set(ex?.s, xf); }
        if (!ex?.s) this.wb.viewStyles.add(xf);
        ensureRow(s, rg.r1).cells[c] = { ...(ex ?? { v: null }), s: xf };
      }
      s.freeze = { rows: rg.r1 + 1, cols: s.freeze.cols };
      if (!s.filter) s.filter = { range: rg, columns: {} };
      s.pristine = false;
    });
  }

  markSaved(): void { this.dirty = false; }
}

function fontPatch(p: StylePatch): boolean {
  return ['bold', 'italic', 'underline', 'strike', 'fontName', 'fontSize', 'color'].some((k) => k in p);
}

/** Formatting a whole cell overrides the same property on its runs. */
function stripFont(font: Run['font'], p: StylePatch): Run['font'] {
  if (!font) return font;
  const f = { ...font };
  if ('bold' in p) delete f.bold;
  if ('italic' in p) delete f.italic;
  if ('underline' in p) delete f.underline;
  if ('strike' in p) delete f.strike;
  if ('fontName' in p) delete f.name;
  if ('fontSize' in p) delete f.size;
  if ('color' in p) delete f.color;
  return Object.keys(f).length ? f : undefined;
}
