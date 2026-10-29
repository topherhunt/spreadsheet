import { DetailedCellError, EmptyValue, ExportedCellChange, HyperFormula, type CellValue, type RawCellContent } from 'hyperformula';
import { inRange, MAX_COLS, MAX_ROWS, parseRange, rangeToA1 } from '../addr';
import { Cell, getCell, isError, Scalar, setCell, Sheet, Workbook } from '../model';
import { displayFormula } from './xlfn';
import { functionNames, has3dRef } from './refs';
import './functions';

export interface CellEdit { sheet: number; row: number; col: number; cell: Cell | undefined }
export interface Changed { sheet: number; row: number; col: number }

const KNOWN = new Set(HyperFormula.getRegisteredFunctionNames('enGB'));

function valueInput(v: Scalar): RawCellContent {
  if (v === null) return null;
  if (typeof v === 'string') return "'" + v; // never let the engine reinterpret text
  if (isError(v)) return v.error;
  return v;
}

/** What the engine holds for a cell: its formula, or its value when it has no live formula. */
function cellInput(c: Cell | undefined): RawCellContent {
  if (!c || c.spill) return null;
  if (c.f && !c.frozen) return '=' + displayFormula(c.f);
  return valueInput(c.v);
}

function fromEngine(v: CellValue): Scalar {
  if (v instanceof DetailedCellError) return { error: v.value };
  if (v === undefined) return null;
  return v;
}

/**
 * Keeps a HyperFormula instance in step with the workbook model. The model is
 * the source of truth for formulas; the engine supplies their values. Formulas
 * it can't evaluate (unknown functions, unparsable syntax, names it can't
 * resolve) are "frozen": the file's cached value stands in for them.
 */
export class Engine {
  hf!: HyperFormula;
  private ids = new Map<number, number>(); // workbook sheet index -> engine sheet id
  private indexOf = new Map<number, number>();
  /** Formulas that produce arrays, as "sheet:row:col". */
  private anchors = new Set<string>();

  constructor(private wb: Workbook) {
    this.build();
    this.freezeLegacySpills();
  }

  /** Recreates the engine from the model, for structural edits (sheets, rows, columns), and refreshes every formula's value. */
  rebuild(): void {
    this.hf.destroy();
    this.build();
    this.wb.sheets.forEach((s, si) => {
      if (s.kind !== 'worksheet') return;
      s.rows.forEach((r, ri) => r?.cells.forEach((c, ci) => {
        if (c?.f && !c.frozen) this.syncValue(si, ri, ci, this.value(si, ri, ci));
      }));
    });
    for (const key of this.anchors) {
      const [si, row, col] = key.split(':').map(Number);
      const c = getCell(this.wb.sheets[si], row, col)!;
      if (c.frozen) continue;
      this.syncArrayRef(si, row, col, c);
      const area = parseRange(c.arrayRef!);
      for (let y = area.r1; y <= area.r2; y++) for (let x = area.c1; x <= area.c2; x++) {
        if (y !== row || x !== col) this.syncValue(si, y, x, this.value(si, y, x));
      }
    }
  }

  destroy(): void {
    this.hf.destroy();
  }

  private build(): void {
    const wb = this.wb;
    this.ids.clear();
    this.indexOf.clear();
    this.anchors.clear();
    const sheets: Record<string, RawCellContent[][]> = {};
    wb.sheets.forEach((s, si) => {
      if (s.kind !== 'worksheet') return;
      sheets[s.name] = sheetInput(s);
      s.rows.forEach((r, ri) => r?.cells.forEach((c, ci) => { if (c?.f && c.arrayRef) this.anchors.add(`${si}:${ri}:${ci}`); }));
    });
    this.hf = HyperFormula.buildFromSheets(sheets, {
      licenseKey: 'gpl-v3',
      maxRows: MAX_ROWS,
      maxColumns: MAX_COLS,
      useArrayArithmetic: true,
      evaluateNullToZero: true,
      useColumnIndex: true,
      ignoreWhiteSpace: 'any',
      // Excel's 1900 system counts a nonexistent Feb 29, 1900; serials from March 1900 on are the same either way.
      nullDate: wb.date1904 ? { year: 1904, month: 1, day: 1 } : { year: 1899, month: 12, day: 31 },
      leapYear1900: !wb.date1904,
      dateFormats: ['MM/DD/YYYY', 'MM/DD/YY', 'YYYY/MM/DD'],
    });
    wb.sheets.forEach((s, i) => {
      if (s.kind !== 'worksheet') return;
      const id = this.hf.getSheetId(s.name);
      if (id === undefined) throw new Error(`Engine lost sheet ${s.name}`);
      this.ids.set(i, id);
      this.indexOf.set(id, i);
    });
    // Excel treats TRUE/FALSE as constants; HyperFormula only knows the functions.
    this.hf.addNamedExpression('TRUE', '=TRUE()');
    this.hf.addNamedExpression('FALSE', '=FALSE()');
    for (const n of wb.definedNames) {
      if (n.name.startsWith('_xlnm.')) continue;
      const scope = n.localSheet === undefined ? undefined : this.ids.get(n.localSheet);
      if (n.localSheet !== undefined && scope === undefined) continue;
      const expr = n.ref.startsWith('=') ? n.ref : '=' + displayFormula(n.ref);
      if (!this.hf.isItPossibleToAddNamedExpression(n.name, expr, scope)) continue; // formulas using it get frozen below
      this.hf.addNamedExpression(n.name, expr, scope);
    }
    this.freezeUnsupported();
  }

  private freezeUnsupported(): void {
    const frozen: { sheet: number; row: number; col: number; c: Cell }[] = [];
    this.wb.sheets.forEach((s, si) => {
      const id = this.ids.get(si);
      if (id === undefined) return;
      s.rows.forEach((r, ri) => r?.cells.forEach((c, ci) => {
        if (!c?.f || c.frozen) return;
        const formula = displayFormula(c.f);
        if (has3dRef(formula)) {
          this.wb.unsupportedFunctions.add('3-D sheet references');
        } else {
          const v = this.hf.getCellValue({ sheet: id, row: ri, col: ci });
          if (!(v instanceof DetailedCellError)) return;
          const unknownFn = /^Function name (.+) not recognized/.exec(v.message);
          const unresolved = v.type === 'ERROR' || unknownFn || /^Named expression .+ not recognized/.test(v.message);
          if (!unresolved) return;
          for (const fn of functionNames(formula)) if (!KNOWN.has(fn)) this.wb.unsupportedFunctions.add(fn);
        }
        if (c.v === null) return; // nothing cached to stand in for it
        frozen.push({ sheet: id, row: ri, col: ci, c });
      }));
    });
    this.freeze(frozen);
  }

  /**
   * A file formula saved without an array range is one Excel evaluated to a
   * single value (implicit intersection, as in =A1:A3*2 beside row 2). The
   * engine spills it instead, so it keeps the file's value.
   */
  private freezeLegacySpills(): void {
    const frozen: { sheet: number; row: number; col: number; c: Cell }[] = [];
    this.wb.sheets.forEach((s, si) => {
      const id = this.ids.get(si);
      if (id === undefined) return;
      s.rows.forEach((r, ri) => r?.cells.forEach((c, ci) => {
        if (!c?.f || c.frozen || c.arrayRef || c.v === null) return;
        const v = this.hf.getCellValue({ sheet: id, row: ri, col: ci });
        const size = this.arraySize(id, ri, ci);
        const spills = (v instanceof DetailedCellError && v.type === 'SPILL') || (size && (size.h > 1 || size.w > 1));
        if (!spills) return;
        this.wb.unsupportedFunctions.add('implicit intersection');
        frozen.push({ sheet: id, row: ri, col: ci, c });
      }));
    });
    this.freeze(frozen);
  }

  private freeze(frozen: { sheet: number; row: number; col: number; c: Cell }[]): void {
    if (!frozen.length) return;
    this.hf.batch(() => {
      for (const { sheet, row, col, c } of frozen) {
        c.frozen = true;
        this.hf.setCellContents({ sheet, row, col }, [[valueInput(c.v)]]);
      }
    });
  }

  /** Writes cells into the model and the engine, then copies recalculated values back. */
  apply(edits: CellEdit[]): Changed[] {
    const changes = this.hf.batch(() => {
      for (const e of edits) {
        const sheet = this.wb.sheets[e.sheet];
        setCell(sheet, e.row, e.col, e.cell);
        sheet.pristine = false;
        const id = this.ids.get(e.sheet);
        if (id === undefined) throw new Error(`Sheet ${sheet.name} isn't a worksheet`);
        this.hf.setCellContents({ sheet: id, row: e.row, col: e.col }, [[cellInput(e.cell)]]);
      }
    });
    const out: Changed[] = edits.map(({ sheet, row, col }) => ({ sheet, row, col }));
    const reported: Changed[] = [...out];
    for (const ch of changes) {
      if (!(ch instanceof ExportedCellChange)) continue;
      const si = this.indexOf.get(ch.sheet);
      if (si === undefined) throw new Error(`Engine changed unknown sheet ${ch.sheet}`);
      reported.push({ sheet: si, row: ch.row, col: ch.col });
      if (this.syncValue(si, ch.row, ch.col, fromEngine(ch.newValue))) out.push({ sheet: si, row: ch.row, col: ch.col });
    }
    for (const e of edits) {
      if (!e.cell?.f || e.cell.frozen) continue;
      this.syncValue(e.sheet, e.row, e.col, this.value(e.sheet, e.row, e.col));
      this.anchors.add(`${e.sheet}:${e.row}:${e.col}`);
    }
    for (const key of [...this.anchors]) {
      const [si, row, col] = key.split(':').map(Number);
      const c = getCell(this.wb.sheets[si], row, col);
      if (!c?.f || c.frozen) { this.anchors.delete(key); continue; }
      const area = c.arrayRef ? parseRange(c.arrayRef) : { r1: row, c1: col, r2: row, c2: col };
      if (!reported.some((o) => o.sheet === si && inRange(area, o.row, o.col))) continue;
      out.push(...this.syncArrayRef(si, row, col, c));
    }
    return out;
  }

  value(sheet: number, row: number, col: number): Scalar {
    const id = this.ids.get(sheet);
    if (id === undefined) throw new Error(`Sheet ${sheet} isn't a worksheet`);
    return fromEngine(this.hf.getCellValue({ sheet: id, row, col }));
  }

  /** Evaluates a formula (with leading '=') against the workbook without storing it. */
  evaluate(sheet: number, formula: string): Scalar {
    const id = this.ids.get(sheet);
    if (id === undefined) throw new Error(`Sheet ${sheet} isn't a worksheet`);
    const v = this.hf.calculateFormula(formula, id);
    if (Array.isArray(v)) return fromEngine(v[0]?.[0] as CellValue);
    return fromEngine(v as CellValue);
  }

  /** How the engine sees a computed number ('NUMBER_DATE', 'NUMBER_PERCENT', ...), for auto-formatting new formulas. */
  valueType(sheet: number, row: number, col: number): string {
    const id = this.ids.get(sheet);
    if (id === undefined) throw new Error(`Sheet ${sheet} isn't a worksheet`);
    return this.hf.getCellValueDetailedType({ sheet: id, row, col });
  }

  /** Puts an engine value into the model. Returns whether anything changed. */
  private syncValue(si: number, row: number, col: number, v: Scalar): boolean {
    const sheet = this.wb.sheets[si];
    const c = getCell(sheet, row, col);
    if (c && !c.spill) {
      if (c.f === undefined || c.frozen) return false; // plain values are the engine's input, not its output
      if (sameValue(c.v, v)) return false;
      c.v = v;
      delete c.sst;
      sheet.pristine = false;
      return true;
    }
    if (v === null) {
      if (!c) return false;
      // An emptied spill cell keeps its style; otherwise it goes.
      if (c.s) { const { s } = c; setCell(sheet, row, col, { v: null, s }); } else setCell(sheet, row, col, undefined);
    } else if (c) {
      if (sameValue(c.v, v)) return false;
      c.v = v;
      delete c.sst;
      delete c.rich;
    } else {
      setCell(sheet, row, col, { v, spill: true });
    }
    sheet.pristine = false;
    return true;
  }

  /**
   * Records the extent of a formula's spilled result so the file can store it
   * as an array formula, and clears cells a shrunken result no longer covers
   * (the engine doesn't report those). Returns the cleared cells.
   */
  private syncArrayRef(si: number, row: number, col: number, c: Cell): Changed[] {
    const size = this.arraySize(this.ids.get(si)!, row, col);
    if (!size) return [];
    const { h, w } = size;
    if (h === 1 && w === 1 && !c.arrayRef) return [];
    const next = { r1: row, c1: col, r2: row + h - 1, c2: col + w - 1 };
    const cleared: Changed[] = [];
    if (c.arrayRef) {
      const sheet = this.wb.sheets[si];
      const prev = parseRange(c.arrayRef);
      for (let y = prev.r1; y <= prev.r2; y++) for (let x = prev.c1; x <= prev.c2; x++) {
        if (inRange(next, y, x) || !getCell(sheet, y, x)?.spill) continue;
        if (this.syncValue(si, y, x, null)) cleared.push({ sheet: si, row: y, col: x });
      }
    }
    c.arrayRef = rangeToA1(next);
    return cleared;
  }

  /** The extent of a formula's array result, or undefined when it isn't an array. */
  private arraySize(id: number, row: number, col: number): { h: number; w: number } | undefined {
    // The array size isn't public API; fail loudly if HyperFormula's internals move.
    const graph = (this.hf as unknown as { dependencyGraph: { getCell(a: object): unknown } }).dependencyGraph;
    if (!graph || typeof graph.getCell !== 'function') throw new Error('HyperFormula dependency graph is unavailable');
    const vertex = graph.getCell({ sheet: id, row, col }) as { array?: { raw?(): unknown[][]; error?: unknown } } | undefined;
    if (!vertex?.array || vertex.array.error) return undefined; // an errored array's value is the error
    if (typeof vertex.array.raw !== 'function') throw new Error('HyperFormula array values changed shape');
    // The engine sizes arrays by their inputs (FILTER of 4 rows is 4 tall); the result is what's non-empty.
    const raw = vertex.array.raw();
    let h = 0, w = 0;
    raw.forEach((r, y) => r.forEach((v, x) => {
      if (v === EmptyValue || v === null || v === undefined) return;
      h = Math.max(h, y + 1);
      w = Math.max(w, x + 1);
    }));
    return { h: Math.max(h, 1), w: Math.max(w, 1) };
  }
}

function sameValue(a: Scalar, b: Scalar): boolean {
  if (isError(a) || isError(b)) return isError(a) && isError(b) && a.error === b.error;
  return a === b;
}

function sheetInput(s: Sheet): RawCellContent[][] {
  const out: RawCellContent[][] = [];
  s.rows.forEach((r, ri) => {
    if (!r) return;
    const row: RawCellContent[] = [];
    r.cells.forEach((c, ci) => { const v = cellInput(c); if (v !== null) row[ci] = v; });
    if (row.length) out[ri] = row;
  });
  // The engine rejects holey arrays.
  return Array.from(out, (row) => (row ? Array.from(row, (v) => (v === undefined ? null : v)) : []));
}
