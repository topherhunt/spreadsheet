import { strFromU8 } from 'fflate';
import { parseA1, parseRange, parseSqref, Range, rangeToA1, toA1 } from '../addr';
import { invalidateSheetInFormula, renameSheetInFormula, shiftFormula, StructuralChange } from '../formula/refs';
import { Sheet, SheetEdit, Workbook } from '../model';
import { shiftRange } from '../structure';
import { relIdOf } from './read';
import { childElements, createLike, firstChild, localName, parseXml, serializeXml, XEl } from './xml';

// File parts the model doesn't hold (tables, pivot tables, comments, drawings, charts,
// conditional formatting, data validation) follow row/column and sheet edits the way
// Excel moves them. Each save rebuilds them from the package as read by replaying
// `wb.edits`, so undoing past a save restores them exactly.

/** Matches no sheet name, so unqualified refs in a part that isn't on the edited sheet stay put. */
const NO_SHEET = '\u0000';

function followFormula(f: string, e: SheetEdit, hostPath: string | undefined): string {
  if (e.kind === 'rename') return renameSheetInFormula(f, e.from, e.to);
  if (e.kind === 'drop') return invalidateSheetInFormula(f, e.name);
  return shiftFormula(f, e.path !== undefined && e.path === hostPath ? e.change.sheet : NO_SHEET, e.change);
}

/** The row/column changes made to the sheet stored at `path`. */
function shiftsOn(wb: Workbook, path: string): StructuralChange[] {
  return wb.edits.flatMap((e) => (e.kind === 'shift' && e.path === path ? [e.change] : []));
}

/** Where index `p` ends up; null when it was deleted. */
function point(p: number, c: StructuralChange): number | null {
  if (c.count > 0) return p >= c.at ? p + c.count : p;
  const end = c.at - c.count - 1;
  return p > end ? p + c.count : p >= c.at ? null : p;
}

/** Where a cell ends up; a deleted cell is null, or with `collapse` moves to the first row/column after the deletion. */
function cellRef(ref: string, c: StructuralChange, collapse = false): string | null {
  const a = parseA1(ref);
  const p = point(c.axis === 'row' ? a.row : a.col, c) ?? (collapse ? c.at : null);
  if (p === null) return null;
  return c.axis === 'row' ? toA1(p, a.col) : toA1(a.row, p);
}

function sqref(s: string, c: StructuralChange): string {
  return parseSqref(s).map((r) => shiftRange(r, c.axis, c.at, c.count)).filter((r): r is Range => !!r).map(rangeToA1).join(' ');
}

function setText(el: XEl, s: string): void {
  while (el.firstChild) el.removeChild(el.firstChild);
  el.appendChild(el.ownerDocument!.createTextNode(s));
}

function descendants(el: XEl, names: Set<string>, out: XEl[] = []): XEl[] {
  for (const c of childElements(el)) {
    if (names.has(localName(c))) out.push(c);
    descendants(c, names, out);
  }
  return out;
}

function followFormulas(root: XEl, names: Set<string>, edits: SheetEdit[], hostPath: string | undefined): void {
  for (const el of descendants(root, names)) {
    const before = el.textContent ?? '';
    let f = before;
    for (const e of edits) f = followFormula(f, e, hostPath);
    if (f !== before) setText(el, f);
  }
}

// ---------------------------------------------------------------------------
// Sheet XML: conditional formatting, data validation, sparklines, selections

const SHEET_FORMULAS = new Set(['formula', 'formula1', 'formula2', 'f']);
/** Containers that Excel rejects (or that mean nothing) once their last child is gone. */
const DROP_WHEN_EMPTY = new Set(['conditionalFormattings', 'dataValidations', 'sparklines', 'sparklineGroup', 'sparklineGroups', 'ignoredErrors', 'protectedRanges', 'ext', 'extLst']);

function removeUp(el: XEl): void {
  let cur: XEl | null = el;
  while (cur) {
    const parent = cur.parentNode as XEl | null;
    if (!parent) return;
    parent.removeChild(cur);
    cur = DROP_WHEN_EMPTY.has(localName(parent)) && !childElements(parent).length ? parent : null;
  }
}

function moveSheetRefs(el: XEl, c: StructuralChange): void {
  for (const child of childElements(el)) {
    const name = localName(child);
    if (name === 'sheetData') continue;
    if (name === 'selection') {
      const active = child.getAttribute('activeCell');
      const moved = active ? cellRef(active, c, true) : null;
      if (moved) child.setAttribute('activeCell', moved);
      const sq = child.getAttribute('sqref');
      if (sq) child.setAttribute('sqref', sqref(sq, c) || moved || 'A1');
      continue;
    }
    const sq = child.getAttribute('sqref');
    if (sq) {
      const moved = sqref(sq, c);
      if (!moved) { removeUp(child); continue; }
      child.setAttribute('sqref', moved);
    }
    if (name === 'sqref') {
      const moved = sqref(child.textContent ?? '', c);
      if (!moved) { removeUp(child.parentNode as XEl); continue; }
      setText(child, moved);
      continue;
    }
    moveSheetRefs(child, c);
  }
}

/** Brings a sheet's own unmodeled elements (its XML as read, without sheetData) in line with `wb.edits`. */
export function followSheetXml(root: XEl, wb: Workbook, path: string): void {
  if (!wb.edits.length) return;
  for (const e of wb.edits) {
    followFormulas(root, SHEET_FORMULAS, [e], path);
    for (const v of descendants(root, new Set(['cfvo']))) {
      if (v.getAttribute('type') === 'formula' && v.getAttribute('val')) v.setAttribute('val', followFormula(v.getAttribute('val')!, e, path));
    }
    if (e.kind === 'shift' && e.path === path) moveSheetRefs(root, e.change);
  }
  for (const dv of descendants(root, new Set(['dataValidations']))) {
    if (dv.getAttribute('count') !== null) dv.setAttribute('count', String(childElements(dv, 'dataValidation').length));
  }
}

// ---------------------------------------------------------------------------
// Tables and pivot tables

export interface TableState {
  path: string;
  ref: Range;
  header: boolean;
  totals: boolean;
  /** Each column's index among the part's tableColumns, or none for a column inserted since. */
  columns: { src?: number; name: string }[];
}

function readTable(path: string, root: XEl): TableState {
  const cols = firstChild(root, 'tableColumns');
  return {
    path,
    ref: parseRange(root.getAttribute('ref')!),
    header: root.getAttribute('headerRowCount') !== '0',
    totals: (root.getAttribute('totalsRowCount') ?? '0') !== '0',
    columns: (cols ? childElements(cols, 'tableColumn') : []).map((c, i) => ({ src: i, name: c.getAttribute('name') ?? '' })),
  };
}

function stepTable(t: TableState, c: StructuralChange): TableState | null {
  const r = t.ref;
  const end = c.at - c.count - 1;
  const gone = (i: number) => c.count < 0 && i >= c.at && i <= end;
  if (c.axis === 'row') {
    if (gone(r.r1) && gone(r.r2)) return null;
    if (t.header && gone(r.r1)) throw new Error('Can’t delete a table’s header row on its own. Delete all of the table’s rows, or leave its header row.');
    const totals = t.totals && !gone(r.r2);
    const ref = shiftRange(r, 'row', c.at, c.count)!;
    // A table keeps at least one data row, as in Excel.
    const min = (t.header ? 1 : 0) + 1 + (totals ? 1 : 0);
    if (ref.r2 - ref.r1 + 1 < min) ref.r2 = ref.r1 + min - 1;
    return { ...t, ref, totals };
  }
  let columns = t.columns;
  if (c.count > 0 && c.at > r.c1 && c.at <= r.c2) {
    const taken = new Set(columns.map((col) => col.name.toLowerCase()));
    const added: TableState['columns'] = [];
    for (let n = 1; added.length < c.count; n++) {
      if (!taken.has(`column${n}`)) added.push({ name: `Column${n}` });
    }
    columns = [...columns.slice(0, c.at - r.c1), ...added, ...columns.slice(c.at - r.c1)];
  } else if (c.count < 0) {
    columns = columns.filter((_, i) => !gone(r.c1 + i));
    if (!columns.length) return null;
  }
  return { ...t, ref: shiftRange(r, 'col', c.at, c.count)!, columns };
}

function originalXml(wb: Workbook, path: string): XEl | null {
  const bytes = wb.src?.original.get(path);
  return bytes ? (parseXml(strFromU8(bytes)).documentElement as unknown as XEl) : null;
}

function relsOf(sheet: Sheet, type: string): { id: string; target: string }[] {
  return (sheet.src?.doc ? sheet.src.rels : []).filter((r) => !r.external && r.type.endsWith(`/${type}`));
}

/** The sheet's tables as they stand after `wb.edits` (and `extra`), by relationship; null for a deleted table. Throws if an edit isn't allowed. */
export function tablesOf(wb: Workbook, sheet: Sheet, extra?: StructuralChange): { relId: string; table: TableState | null }[] {
  const shifts = sheet.src ? shiftsOn(wb, sheet.src.path) : [];
  if (extra) shifts.push(extra);
  return relsOf(sheet, 'table').flatMap((rel) => {
    const root = originalXml(wb, rel.target);
    if (!root) return [];
    let t: TableState | null = readTable(rel.target, root);
    for (const c of shifts) if (t) t = stepTable(t, c);
    return [{ relId: rel.id, table: t }];
  });
}

function pivotRefs(wb: Workbook, sheet: Sheet): { path: string; ref: Range }[] {
  const shifts = sheet.src ? shiftsOn(wb, sheet.src.path) : [];
  return relsOf(sheet, 'pivotTable').flatMap((rel) => {
    const root = originalXml(wb, rel.target);
    const loc = root ? firstChild(root, 'location') : null;
    if (!loc?.getAttribute('ref')) return [];
    let ref = parseRange(loc.getAttribute('ref')!);
    for (const c of shifts) ref = shiftRange(ref, c.axis, c.at, c.count) ?? ref;
    return [{ path: rel.target, ref }];
  });
}

/** Refuses a row/column change Excel refuses: one that cuts through a pivot table, or deletes a table's header row on its own. */
export function checkShift(wb: Workbook, sheet: Sheet, c: StructuralChange): void {
  tablesOf(wb, sheet, c);
  const what = c.axis === 'row' ? 'rows' : 'columns';
  for (const { ref } of pivotRefs(wb, sheet)) {
    const lo = c.axis === 'row' ? ref.r1 : ref.c1, hi = c.axis === 'row' ? ref.r2 : ref.c2;
    const cuts = c.count > 0 ? c.at > lo && c.at <= hi : c.at <= hi && c.at - c.count - 1 >= lo;
    if (cuts) throw new Error(`Can’t ${c.count > 0 ? 'insert' : 'delete'} ${what} there: that would change part of a pivot table.`);
  }
}

function writeTable(root: XEl, t: TableState, edits: SheetEdit[], hostPath: string): void {
  followFormulas(root, new Set(['calculatedColumnFormula', 'totalsRowFormula']), edits, hostPath);
  const was = readTable(t.path, root);
  const same = rangeToA1(was.ref) === rangeToA1(t.ref) && was.totals === t.totals && was.columns.length === t.columns.length && t.columns.every((c, i) => c.src === i);
  if (same) return;
  root.setAttribute('ref', rangeToA1(t.ref));
  const cols = firstChild(root, 'tableColumns');
  const orig = cols ? childElements(cols, 'tableColumn') : [];
  if (!t.totals && root.getAttribute('totalsRowCount') !== null) {
    root.removeAttribute('totalsRowCount');
    for (const c of orig) {
      c.removeAttribute('totalsRowFunction');
      c.removeAttribute('totalsRowLabel');
      for (const f of childElements(c, 'totalsRowFormula')) c.removeChild(f);
    }
  }
  const af = firstChild(root, 'autoFilter');
  if (af) {
    af.setAttribute('ref', rangeToA1({ ...t.ref, r2: t.ref.r2 - (t.totals ? 1 : 0) }));
    const newIndex = new Map(t.columns.flatMap((c, i) => (c.src === undefined ? [] : [[c.src, i] as const])));
    for (const fc of childElements(af, 'filterColumn')) {
      const i = newIndex.get(parseInt(fc.getAttribute('colId') ?? '-1', 10));
      if (i === undefined) af.removeChild(fc); else fc.setAttribute('colId', String(i));
    }
  }
  // A remembered sort order is dropped once the table's shape changes, rather than pointing at the wrong cells.
  for (const s of [...(af ? childElements(af, 'sortState') : []), ...childElements(root, 'sortState')]) s.parentNode!.removeChild(s);
  if (cols) {
    let id = Math.max(0, ...orig.map((c) => parseInt(c.getAttribute('id') ?? '0', 10)));
    for (const c of orig) cols.removeChild(c);
    for (const c of t.columns) {
      if (c.src !== undefined) { cols.appendChild(orig[c.src]); continue; }
      const el = createLike(cols, 'tableColumn');
      el.setAttribute('id', String(++id));
      el.setAttribute('name', c.name);
      cols.appendChild(el);
    }
    cols.setAttribute('count', String(t.columns.length));
  }
}

// ---------------------------------------------------------------------------
// Comments, notes and drawings

function followCellRefs(root: XEl, name: string, shifts: StructuralChange[]): void {
  for (const el of descendants(root, new Set([name]))) {
    let ref: string | null = el.getAttribute('ref');
    for (const c of shifts) if (ref) ref = cellRef(ref, c);
    if (ref) el.setAttribute('ref', ref); else el.parentNode!.removeChild(el);
  }
}

/** A position (cell index plus offset) moved by one change; a deleted cell collapses onto the edge of the deletion. */
function movePos(p: number, off: number, c: StructuralChange): [number, number] {
  const np = point(p, c);
  return np === null ? [c.at, 0] : [np, off];
}

function followDrawing(root: XEl, shifts: StructuralChange[]): void {
  for (const a of descendants(root, new Set(['twoCellAnchor', 'oneCellAnchor']))) {
    const mode = localName(a) === 'oneCellAnchor' ? 'oneCell' : a.getAttribute('editAs') ?? 'twoCell';
    if (mode === 'absolute') continue;
    const markers = [firstChild(a, 'from'), firstChild(a, 'to')].filter((m): m is XEl => !!m);
    for (const c of shifts) {
      const idx = (m: XEl) => firstChild(m, c.axis)!, off = (m: XEl) => firstChild(m, `${c.axis}Off`)!;
      const [from, to] = markers;
      const p0 = parseInt(idx(from).textContent!, 10);
      const [p, o] = movePos(p0, parseInt(off(from).textContent!, 10), c);
      if (to && mode === 'oneCell') setText(idx(to), String(parseInt(idx(to).textContent!, 10) + p - p0));
      else if (to) {
        const [q, qo] = movePos(parseInt(idx(to).textContent!, 10), parseInt(off(to).textContent!, 10), c);
        setText(idx(to), String(q));
        setText(off(to), String(qo));
      }
      setText(idx(from), String(p));
      setText(off(from), String(o));
    }
  }
}

/** Legacy VML (note boxes, form controls) isn't reliably well-formed XML, so its few numbers are edited in place. */
function followVml(xml: string, shifts: StructuralChange[]): string {
  return xml.replace(/<v:shape\b[\s\S]*?<\/v:shape>/g, (shape) => {
    const num = (tag: string) => { const m = new RegExp(`<(\\w+:)?${tag}>\\s*(\\d+)\\s*</(\\w+:)?${tag}>`).exec(shape); return m ? parseInt(m[2], 10) : null; };
    let row = num('Row'), col = num('Column');
    const am = /<((?:\w+:)?)Anchor>([^<]*)<\/(?:\w+:)?Anchor>/.exec(shape);
    let anchor = am ? am[2].split(',').map((s) => parseInt(s.trim(), 10)) : null;
    for (const c of shifts) {
      const cur = c.axis === 'row' ? row : col;
      if (cur !== null) {
        const moved = point(cur, c);
        if (moved === null && /ObjectType="Note"/.test(shape)) return '';
        if (c.axis === 'row') row = moved ?? c.at; else col = moved ?? c.at;
      }
      if (anchor && anchor.length === 8) {
        const pairs = c.axis === 'row' ? [2, 6] : [0, 4];
        for (const i of pairs) [anchor[i], anchor[i + 1]] = movePos(anchor[i], anchor[i + 1], c);
      }
    }
    let out = shape;
    if (row !== null) out = out.replace(/(<(\w+:)?Row>)\s*\d+\s*(<\/(\w+:)?Row>)/, `$1${row}$3`);
    if (col !== null) out = out.replace(/(<(\w+:)?Column>)\s*\d+\s*(<\/(\w+:)?Column>)/, `$1${col}$3`);
    if (am && anchor) out = out.replace(am[0], `<${am[1]}Anchor>${anchor.join(', ')}</${am[1]}Anchor>`);
    return out;
  });
}

// ---------------------------------------------------------------------------
// Saving

export interface FollowTarget {
  files: Map<string, Uint8Array>;
  setText(path: string, xml: string): void;
  remove(path: string): void;
  setOverride(part: string, contentType: string | null): void;
}

const CT_TABLE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';

/** Rebuilds every followed part outside the sheet XML from the package as read and `wb.edits`. */
export function followParts(pkg: FollowTarget, wb: Workbook): void {
  const original = wb.src?.original;
  if (!original) return;
  // A part the edits leave unchanged keeps its original bytes, so untouched files still round-trip exactly.
  const restore = (path: string) => { const b = original.get(path)!; if (pkg.files.get(path) !== b) pkg.files.set(path, b); };
  const rebuild = (path: string, fn: (root: XEl) => void) => {
    const bytes = original.get(path);
    if (!bytes || !pkg.files.has(path)) return;
    const doc = parseXml(strFromU8(bytes));
    const before = serializeXml(doc);
    fn(doc.documentElement as unknown as XEl);
    const after = serializeXml(doc);
    if (after === before) restore(path); else pkg.setText(path, after);
  };
  const edits = wb.edits;

  for (const sheet of wb.sheets) {
    const path = sheet.src?.doc ? sheet.src.path : null;
    if (!path) continue;
    const shifts = shiftsOn(wb, path);
    const parts = (type: string) => relsOf(sheet, type).map((r) => r.target);
    const tables = new Map(tablesOf(wb, sheet).map((x) => [x.relId, x.table]));
    for (const rel of relsOf(sheet, 'table')) {
      if (!tables.has(rel.id)) continue;
      const t = tables.get(rel.id);
      if (!t) { if (pkg.files.has(rel.target)) { pkg.remove(rel.target); pkg.setOverride(rel.target, null); } continue; }
      if (!pkg.files.has(rel.target)) { pkg.files.set(rel.target, original.get(rel.target)!); pkg.setOverride(rel.target, CT_TABLE); }
      rebuild(rel.target, (root) => writeTable(root, t, edits, path));
    }
    for (const p of [...parts('comments'), ...parts('threadedComment'), ...parts('drawing'), ...parts('vmlDrawing'), ...parts('pivotTable')]) {
      if (!original.has(p) || !pkg.files.has(p)) continue;
      if (!shifts.length) { restore(p); continue; }
      if (p.endsWith('.vml')) {
        const xml = strFromU8(original.get(p)!), moved = followVml(xml, shifts);
        if (moved === xml) restore(p); else pkg.setText(p, moved);
        continue;
      }
      rebuild(p, (root) => {
        const name = localName(root);
        if (name === 'comments') followCellRefs(root, 'comment', shifts);
        else if (name === 'ThreadedComments') followCellRefs(root, 'threadedComment', shifts);
        else if (name === 'wsDr') followDrawing(root, shifts);
        else if (name === 'pivotTableDefinition') {
          const loc = firstChild(root, 'location');
          if (loc?.getAttribute('ref')) loc.setAttribute('ref', rangeToA1(pivotRefs(wb, sheet).find((x) => x.path === p)!.ref));
        }
      });
    }
  }

  for (const path of original.keys()) {
    const chart = /^xl\/charts\/[^/]+\.xml$/.test(path);
    const cache = /^xl\/pivotCache\/pivotCacheDefinition[^/]*\.xml$/.test(path);
    if (!chart && !cache) continue;
    if (!edits.length) { if (pkg.files.has(path)) restore(path); continue; }
    rebuild(path, (root) => {
      if (chart) { followFormulas(root, new Set(['f']), edits, undefined); return; }
      const ws = descendants(root, new Set(['worksheetSource']))[0];
      if (!ws || ws.getAttributeNode('r:id') || !ws.getAttribute('sheet') || !ws.getAttribute('ref')) return;
      let name = ws.getAttribute('sheet')!, ref: Range = parseRange(ws.getAttribute('ref')!);
      for (const e of edits) {
        if (e.kind === 'rename' && e.from.toLowerCase() === name.toLowerCase()) name = e.to;
        if (e.kind === 'shift' && e.change.sheet.toLowerCase() === name.toLowerCase()) ref = shiftRange(ref, e.change.axis, e.change.at, e.change.count) ?? ref;
      }
      ws.setAttribute('sheet', name);
      ws.setAttribute('ref', rangeToA1(ref));
    });
  }
}

/** Removes the `<tablePart>` of each table that row/column edits deleted, with its relationship. */
export function dropDeletedTables(root: XEl, wb: Workbook, sheet: Sheet, removeRel: (id: string) => void): void {
  const parts = firstChild(root, 'tableParts');
  if (!parts) return;
  for (const { relId, table } of tablesOf(wb, sheet)) {
    if (table) continue;
    for (const tp of childElements(parts, 'tablePart')) if (relIdOf(tp) === relId) parts.removeChild(tp);
    removeRel(relId);
  }
  const left = childElements(parts, 'tablePart').length;
  if (left) parts.setAttribute('count', String(left)); else root.removeChild(parts);
}
