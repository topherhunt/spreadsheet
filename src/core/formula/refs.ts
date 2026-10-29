import { colToName, MAX_COLS, MAX_ROWS, nameToCol, quoteSheetName } from '../addr';

// A lightweight formula tokenizer that understands exactly enough to move
// references around: copy/paste/fill translation, shared-formula expansion,
// and shifting for inserted/deleted rows and columns. Everything else in the
// formula is carried through as opaque text, so a formula we only shift comes
// back byte-identical apart from the references that actually moved.

export interface RefPart { col?: number; colAbs: boolean; row?: number; rowAbs: boolean }

export type Token =
  | { type: 'text'; text: string }
  | { type: 'func'; text: string; name: string }
  | { type: 'ref'; sheet?: string; a: RefPart; b?: RefPart; error?: boolean; src?: string };

const CELL_RE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/;
const COL_RE = /^(\$?)([A-Za-z]{1,3})$/;
const ROW_RE = /^(\$?)(\d+)$/;

function parseCellPart(s: string): RefPart | null {
  const m = CELL_RE.exec(s);
  if (!m) return null;
  const col = nameToCol(m[2]);
  const row = parseInt(m[4], 10) - 1;
  if (col >= MAX_COLS || row < 0 || row >= MAX_ROWS) return null;
  return { col, colAbs: m[1] === '$', row, rowAbs: m[3] === '$' };
}

function parseColPart(s: string): RefPart | null {
  const m = COL_RE.exec(s);
  if (!m) return null;
  const col = nameToCol(m[2]);
  return col < MAX_COLS ? { col, colAbs: m[1] === '$', rowAbs: false } : null;
}

function parseRowPart(s: string): RefPart | null {
  const m = ROW_RE.exec(s);
  if (!m) return null;
  const row = parseInt(m[2], 10) - 1;
  return row >= 0 && row < MAX_ROWS ? { row, rowAbs: m[1] === '$', colAbs: false } : null;
}

const IDENT_CHAR = /[A-Za-z0-9_.$\\À-￿]/;

export function tokenize(formula: string): Token[] {
  const out: Token[] = [];
  let text = '';
  const flush = () => { if (text) { out.push({ type: 'text', text }); text = ''; } };
  let i = 0;
  const n = formula.length;

  // Reads an identifier-ish chunk at i (letters, digits, $, _, ., backslash).
  const readIdent = (at: number): string => {
    let j = at;
    while (j < n && IDENT_CHAR.test(formula[j])) j++;
    return formula.slice(at, j);
  };

  // Tries to read a reference (no sheet prefix) at position `at`.
  const readRef = (at: number): { a: RefPart; b?: RefPart; end: number } | null => {
    const first = readIdent(at);
    if (!first) return null;
    let end = at + first.length;
    if (formula[end] === '(') return null; // function call
    if (formula[end] === ':') {
      const second = readIdent(end + 1);
      if (second && formula[end + 1 + second.length] !== '(') {
        const ca = parseCellPart(first), cb = parseCellPart(second);
        if (ca && cb) return { a: ca, b: cb, end: end + 1 + second.length };
        const cola = parseColPart(first), colb = parseColPart(second);
        if (cola && colb) return { a: cola, b: colb, end: end + 1 + second.length };
        const rowa = parseRowPart(first), rowb = parseRowPart(second);
        if (rowa && rowb) return { a: rowa, b: rowb, end: end + 1 + second.length };
      }
    }
    const c = parseCellPart(first);
    if (c && formula[end] !== '[') return { a: c, end };
    return null;
  };

  while (i < n) {
    const ch = formula[i];
    if (ch === '"') {
      let j = i + 1;
      while (j < n) {
        if (formula[j] === '"') { if (formula[j + 1] === '"') { j += 2; continue; } break; }
        j++;
      }
      text += formula.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (ch === '[') {
      // External-workbook prefix or structured reference: opaque, balanced.
      let depth = 0, j = i;
      for (; j < n; j++) {
        if (formula[j] === '[') depth++;
        else if (formula[j] === ']') { depth--; if (depth === 0) break; }
      }
      text += formula.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (ch === '#') {
      const m = /^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|GETTING_DATA)/i.exec(formula.slice(i));
      if (m) { text += m[0]; i += m[0].length; continue; }
    }
    if (ch === "'") {
      // Quoted sheet name followed by '!'
      let j = i + 1;
      while (j < n) {
        if (formula[j] === "'") { if (formula[j + 1] === "'") { j += 2; continue; } break; }
        j++;
      }
      if (formula[j + 1] === '!') {
        const sheet = formula.slice(i + 1, j).replace(/''/g, "'");
        const ref = readRef(j + 2);
        if (ref) {
          flush();
          out.push({ type: 'ref', sheet, a: ref.a, b: ref.b, src: formula.slice(i, ref.end) });
          i = ref.end;
          continue;
        }
        if (formula.startsWith('#REF!', j + 2)) {
          flush();
          out.push({ type: 'ref', sheet, a: { colAbs: false, rowAbs: false }, error: true, src: formula.slice(i, j + 7) });
          i = j + 7;
          continue;
        }
      }
      text += formula.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    const prev = i > 0 ? formula[i - 1] : '';
    if (IDENT_CHAR.test(ch) && !(prev && IDENT_CHAR.test(prev))) {
      const ident = readIdent(i);
      const after = i + ident.length;
      if (formula[after] === '(' && !/^\d/.test(ident)) {
        flush();
        out.push({ type: 'func', text: ident, name: ident.toUpperCase() });
        i = after;
        continue;
      }
      if (formula[after] === '!' && ident && !/^[\d.]/.test(ident)) {
        const ref = readRef(after + 1);
        if (ref) {
          flush();
          out.push({ type: 'ref', sheet: ident, a: ref.a, b: ref.b, src: formula.slice(i, ref.end) });
          i = ref.end;
          continue;
        }
        if (formula.startsWith('#REF!', after + 1)) {
          flush();
          out.push({ type: 'ref', sheet: ident, a: { colAbs: false, rowAbs: false }, error: true, src: formula.slice(i, after + 6) });
          i = after + 6;
          continue;
        }
      }
      // Numbers like 1.5E+3 must not be mistaken for refs; a ref never starts with a digit
      // except row ranges (1:3), which readRef handles.
      if (/^\d+(\.\d*)?([Ee])?$/.test(ident) && /[Ee]$/.test(ident) && /[+-]/.test(formula[after] ?? '')) {
        const m = /^[+-]\d+/.exec(formula.slice(after));
        text += ident + (m ? m[0] : '');
        i = after + (m ? m[0].length : 0);
        continue;
      }
      const ref = readRef(i);
      if (ref) {
        flush();
        out.push({ type: 'ref', a: ref.a, b: ref.b, src: formula.slice(i, ref.end) });
        i = ref.end;
        continue;
      }
      text += ident;
      i = after;
      continue;
    }
    text += ch;
    i++;
  }
  flush();
  return out;
}

function partToString(p: RefPart): string {
  let s = '';
  if (p.col !== undefined) s += (p.colAbs ? '$' : '') + colToName(p.col);
  if (p.row !== undefined) s += (p.rowAbs ? '$' : '') + (p.row + 1);
  return s;
}

export function refToString(t: Extract<Token, { type: 'ref' }>): string {
  const prefix = t.sheet !== undefined ? quoteSheetName(t.sheet) + '!' : '';
  if (t.error) return prefix + '#REF!';
  return prefix + partToString(t.a) + (t.b ? ':' + partToString(t.b) : '');
}

/** Joins tokens back into a formula; refs that kept their source text are written as found. */
export function untokenize(tokens: Token[]): string {
  return tokens.map((t) => (t.type === 'ref' ? t.src ?? refToString(t) : t.text)).join('');
}

/**
 * Rewrites only the reference tokens that `fn` changes, leaving every other
 * byte of the formula (including how untouched references were written) as is.
 */
function rewriteRefs(formula: string, fn: (t: Extract<Token, { type: 'ref' }>) => Extract<Token, { type: 'ref' }> | null): string {
  const toks = tokenize(formula);
  let changed = false;
  const out = toks.map((t) => {
    if (t.type !== 'ref' || t.error) return t;
    const r = fn(t);
    if (r === t) return t;
    changed = true;
    return r ? { ...r, src: undefined } : { ...t, a: { colAbs: false, rowAbs: false }, b: undefined, error: true, src: undefined };
  });
  if (!changed) return formula;
  return untokenize(out);
}

/** Moves relative references by (dRow, dCol), as copying a formula does. Out-of-sheet refs become #REF!. */
export function translateFormula(formula: string, dRow: number, dCol: number): string {
  if (dRow === 0 && dCol === 0) return formula;
  const move = (p: RefPart): RefPart | null => {
    const q = { ...p };
    if (q.row !== undefined && !q.rowAbs) { q.row += dRow; if (q.row < 0 || q.row >= MAX_ROWS) return null; }
    if (q.col !== undefined && !q.colAbs) { q.col += dCol; if (q.col < 0 || q.col >= MAX_COLS) return null; }
    return q;
  };
  return rewriteRefs(formula, (t) => {
    // Whole-column refs don't move vertically and whole-row refs don't move sideways.
    const a = move(t.a);
    const b = t.b ? move(t.b) : undefined;
    if (!a || b === null) return null;
    const same = (p: RefPart, q: RefPart | undefined) => !!q && p.row === q.row && p.col === q.col;
    if (same(t.a, a) && (t.b === undefined || same(t.b, b))) return t;
    return { ...t, a, b };
  });
}

export interface StructuralChange {
  sheet: string;          // sheet whose rows/cols change
  axis: 'row' | 'col';
  at: number;             // first index affected
  count: number;          // >0 inserted, <0 deleted
}

/**
 * Adjusts a formula living on `formulaSheet` for rows/cols inserted or deleted
 * on `change.sheet`. A reference wholly inside a deleted block becomes #REF!;
 * a range that partially overlaps shrinks.
 */
export function shiftFormula(formula: string, formulaSheet: string, change: StructuralChange): string {
  const key = change.axis;
  const max = key === 'row' ? MAX_ROWS : MAX_COLS;
  const { at, count } = change;
  return rewriteRefs(formula, (t) => {
    const target = t.sheet ?? formulaSheet;
    if (target.toLowerCase() !== change.sheet.toLowerCase()) return t;
    const get = (p: RefPart) => (key === 'row' ? p.row : p.col);
    const set = (p: RefPart, v: number): RefPart => (key === 'row' ? { ...p, row: v } : { ...p, col: v });
    const a0 = get(t.a);
    if (a0 === undefined) return t; // whole-row ref under a column change, or vice versa
    const b0 = t.b ? get(t.b) : a0;
    if (b0 === undefined) return t;
    let lo = Math.min(a0, b0), hi = Math.max(a0, b0);
    if (count > 0) {
      if (lo >= at) lo += count;
      if (hi >= at) hi += count;
      // A range pushed past the sheet's end keeps its end there, as in Excel.
      if (lo >= max) return null;
      hi = Math.min(hi, max - 1);
    } else {
      const delEnd = at - count - 1; // inclusive
      if (lo >= at && hi <= delEnd) return null;
      const shift = (x: number) => (x > delEnd ? x + count : x >= at ? at : x);
      const nlo = lo >= at && lo <= delEnd ? at : shift(lo);
      const nhi = hi >= at && hi <= delEnd ? at - 1 : shift(hi);
      lo = nlo; hi = Math.max(nlo, nhi);
    }
    if (lo === Math.min(a0, b0) && hi === Math.max(a0, b0)) return t;
    const a = set(t.a, a0 <= b0 ? lo : hi);
    const b = t.b ? set(t.b, a0 <= b0 ? hi : lo) : undefined;
    return { ...t, a, b };
  });
}

/**
 * Follows a cut and paste: refs to `movedSheet` lying wholly inside `from` (0-based,
 * inclusive) move by (dRow, dCol), absolute or not. Refs that only overlap it stay put.
 */
export function moveRefsInFormula(formula: string, formulaSheet: string, movedSheet: string, from: { r1: number; c1: number; r2: number; c2: number }, dRow: number, dCol: number): string {
  const inside = (p: RefPart) => p.row !== undefined && p.col !== undefined && p.row >= from.r1 && p.row <= from.r2 && p.col >= from.c1 && p.col <= from.c2;
  const move = (p: RefPart): RefPart => ({ ...p, row: p.row! + dRow, col: p.col! + dCol });
  return rewriteRefs(formula, (t) => {
    if ((t.sheet ?? formulaSheet).toLowerCase() !== movedSheet.toLowerCase()) return t;
    if (!inside(t.a) || (t.b && !inside(t.b))) return t;
    return { ...t, a: move(t.a), b: t.b && move(t.b) };
  });
}

export function renameSheetInFormula(formula: string, oldName: string, newName: string): string {
  return rewriteRefs(formula, (t) => (t.sheet !== undefined && t.sheet.toLowerCase() === oldName.toLowerCase() ? { ...t, sheet: newName } : t));
}

/** Refs to `sheet` become #REF! (the sheet was deleted). */
export function invalidateSheetInFormula(formula: string, sheet: string): string {
  return rewriteRefs(formula, (t) => (t.sheet !== undefined && t.sheet.toLowerCase() === sheet.toLowerCase() ? null : t));
}

export function functionNames(formula: string): string[] {
  return tokenize(formula).filter((t): t is Extract<Token, { type: 'func' }> => t.type === 'func').map((t) => t.name);
}

/** Whether the formula spans sheets (Sheet1:Sheet3!A1), which the engine misreads as a ref to the last sheet. */
export function has3dRef(formula: string): boolean {
  const tokens = tokenize(formula);
  return tokens.some((t, i) => {
    if (t.type !== 'ref' || t.sheet === undefined) return false;
    if (t.sheet.includes(':')) return true; // quoted: 'Jan:Mar'!A1
    const prev = tokens[i - 1];
    return prev?.type === 'text' && /[A-Za-z0-9_.À-￿]:$/.test(prev.text);
  });
}
