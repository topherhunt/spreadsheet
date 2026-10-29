import { unzipSync, strFromU8 } from 'fflate';
import { parseRange, parseSqref, parseA1, rangeToA1 } from '../addr';
import {
  Cell, ColInfo, ColumnFilter, DefinedName, FilterCondition, newSheet, Row, Run, RunFont, Scalar, Sheet, Workbook,
} from '../model';
import { emptyStyleSheet, parseStyles, resolveColor, StyleSheet, ColorSpec } from '../styles';
import { translateFormula } from '../formula/refs';
import { serialFromIso } from '../dates';
import {
  assertNoDtd, attrsToXml, childElements, decodeXml, findClose, firstChild, localName, parseAttrs, parseXml, scanTags, XDoc, XEl,
} from './xml';

export interface Rel { id: string; type: string; target: string; external: boolean }

export interface SstEntry { text: string; runs?: Run[] }

export interface XlsxSource {
  /** Every part of the original package, byte-for-byte. */
  files: Map<string, Uint8Array>;
  workbookPath: string;
  workbookDoc: XDoc;
  workbookRels: Rel[];
  sst: SstEntry[];
  sstPath?: string;
  stylesPath?: string;
}

export interface SheetSource {
  path: string;
  relId: string;
  sheetId: string;
  /** Sheet XML with an empty <sheetData/>, as a DOM to patch on save. */
  doc?: XDoc;
  rels: Rel[];
  /** JSON snapshots of the modeled sheet-level elements, to tell what changed on save. */
  loaded: Record<string, string>;
  /** Raw hyperlink elements by A1 ref, to write back unchanged ones exactly. */
  links: Map<string, { target: string; attrs: Record<string, string> }>;
}

const REL_NS_SUFFIX = 'relationships';

export function partDir(path: string): string {
  const i = path.lastIndexOf('/');
  return i < 0 ? '' : path.slice(0, i + 1);
}

export function resolveTarget(base: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (partDir(base) + target).split('/');
  const out: string[] = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p !== '.') out.push(p);
  }
  return out.join('/');
}

export function relsPathFor(part: string): string {
  return `${partDir(part)}_rels/${part.slice(partDir(part).length)}.rels`;
}

export function parseRels(files: Map<string, Uint8Array>, part: string): Rel[] {
  const bytes = files.get(relsPathFor(part));
  if (!bytes) return [];
  const doc = parseXml(strFromU8(bytes));
  return childElements(doc.documentElement as unknown as XEl, 'Relationship').map((r) => ({
    id: r.getAttribute('Id')!,
    type: r.getAttribute('Type') ?? '',
    external: r.getAttribute('TargetMode') === 'External',
    target: r.getAttribute('TargetMode') === 'External' ? r.getAttribute('Target')! : resolveTarget(part, r.getAttribute('Target')!),
  }));
}

export function relIdOf(el: XEl): string | null {
  for (let i = 0; i < el.attributes.length; i++) {
    const a = el.attributes[i];
    if ((a.localName ?? a.name.replace(/^.*:/, '')) === 'id' && (a.namespaceURI ?? '').endsWith(REL_NS_SUFFIX)) return a.value;
  }
  return null;
}

function relType(r: Rel): string {
  return r.type.replace(/^.*\//, '');
}

function text(files: Map<string, Uint8Array>, path: string): string | undefined {
  const b = files.get(path);
  return b ? strFromU8(b) : undefined;
}

// ---------------------------------------------------------------------------
// Shared strings and rich text

const RPR_TAGS = /<(?:\w+:)?(b|i|strike|u|sz|color|rFont|vertAlign)((?:\s+[^\s=>/]+\s*=\s*"[^"]*")*)\s*\/?>/g;

export function parseRunFont(rPr: string, ss: StyleSheet): RunFont {
  const f: RunFont = {};
  const re = new RegExp(RPR_TAGS.source, 'g');
  for (let m = re.exec(rPr); m; m = re.exec(rPr)) {
    const a = parseAttrs(m[2]);
    const on = a.val === undefined || a.val === '1' || a.val === 'true';
    switch (m[1]) {
      case 'b': f.bold = on; break;
      case 'i': f.italic = on; break;
      case 'strike': f.strike = on; break;
      case 'u': f.underline = a.val !== 'none'; break;
      case 'sz': f.size = parseFloat(a.val); break;
      case 'rFont': f.name = a.val; break;
      case 'color': {
        const c: ColorSpec = {};
        if (a.rgb) c.rgb = a.rgb.slice(-6);
        if (a.theme) c.theme = parseInt(a.theme, 10);
        if (a.indexed) c.indexed = parseInt(a.indexed, 10);
        if (a.tint) c.tint = parseFloat(a.tint);
        const hex = resolveColor(ss, c);
        if (hex) f.color = hex;
        break;
      }
    }
  }
  return f;
}

/** Parses the content of an <si> or <is> element into text and (if rich) runs. */
export function parseStringItem(xml: string, from: number, to: number, ss: StyleSheet): SstEntry {
  let textOut = '';
  let runs: Run[] | undefined;
  let inRun = false, runText = '', rPr: string | undefined;
  let skipDepth = 0; // inside <rPh> (phonetic) whose text isn't part of the value
  let tStart = -1;
  for (const t of scanTags(xml, from, to)) {
    if (t.name === 'rPh' || t.name === 'phoneticPr') {
      if (t.name === 'rPh' && !t.selfClose) skipDepth += t.close ? -1 : 1;
      continue;
    }
    if (skipDepth) continue;
    if (t.name === 'r' && !t.selfClose) {
      if (!t.close) { inRun = true; runText = ''; rPr = undefined; runs ??= []; }
      else {
        inRun = false;
        runs!.push(rPr !== undefined ? { text: runText, rPr, rPrFont: parseRunFont(rPr, ss), font: parseRunFont(rPr, ss) } : { text: runText });
      }
    } else if (t.name === 'rPr' && !t.close) {
      if (t.selfClose) rPr = '';
      else { const c = findClose(xml, 'rPr', t.end); rPr = xml.slice(t.start, c.end); }
    } else if (t.name === 't') {
      if (t.selfClose) continue;
      if (!t.close) tStart = t.end;
      else if (tStart >= 0) {
        const s = decodeXml(xml.slice(tStart, t.start));
        textOut += s;
        if (inRun) runText += s;
        tStart = -1;
      }
    }
  }
  return runs ? { text: textOut, runs } : { text: textOut };
}

function parseSst(xml: string, ss: StyleSheet): SstEntry[] {
  assertNoDtd(xml);
  const out: SstEntry[] = [];
  let siStart = -1;
  for (const t of scanTags(xml)) {
    if (t.name !== 'si') continue;
    if (t.selfClose) { out.push({ text: '' }); continue; }
    if (!t.close) siStart = t.end;
    else { out.push(parseStringItem(xml, siStart, t.start, ss)); siStart = -1; }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Worksheet

const FEATURE_ELEMENTS: Record<string, string> = {
  conditionalFormatting: 'Conditional formatting',
  dataValidations: 'Data validation',
  drawing: 'Charts, images or shapes',
  legacyDrawing: 'Notes/comments',
  legacyDrawingHF: 'Header/footer images',
  tableParts: 'Tables',
  sheetProtection: 'Sheet protection',
  protectedRanges: 'Protected ranges',
  scenarios: 'Scenarios',
  dataConsolidate: 'Data consolidation',
  customSheetViews: 'Custom views',
  oleObjects: 'Embedded objects',
  controls: 'Form controls',
  webPublishItems: 'Web publishing',
  picture: 'Background picture',
};

const REL_FEATURES: Record<string, string> = {
  pivotTable: 'Pivot tables',
  comments: 'Notes/comments',
  threadedComment: 'Threaded comments',
  table: 'Tables',
  drawing: 'Charts, images or shapes',
};

function bool(v: string | undefined): boolean {
  return v === '1' || v === 'true';
}

function cellValueOf(t: string | undefined, raw: string | undefined, sst: SstEntry[], date1904: boolean): { v: Scalar; entry?: SstEntry; sst?: number } {
  if (raw === undefined) return { v: null };
  switch (t) {
    case 's': {
      const i = parseInt(raw, 10);
      const e = sst[i];
      if (!e) throw new Error(`Shared string index ${i} is out of range`);
      return { v: e.text, entry: e, sst: i };
    }
    case 'str': return { v: raw };
    case 'b': return { v: raw === '1' || raw === 'true' };
    case 'e': return { v: { error: raw } };
    case 'd': {
      // ISO 8601 date cell (mostly Strict files); Excel itself converts these to serials.
      const n = serialFromIso(raw, date1904);
      return { v: n ?? raw };
    }
    default: {
      const n = Number(raw);
      return { v: Number.isFinite(n) ? n : raw };
    }
  }
}

function parseSheetData(xml: string, from: number, to: number, sheet: Sheet, sst: SstEntry[], ss: StyleSheet, date1904: boolean): void {
  let row: Row | undefined;
  let rowIdx = -1;
  let colCursor = -1;
  let cell: { attrs: Record<string, string>; col: number; f?: string; fAttrs?: Record<string, string>; v?: string; is?: SstEntry; extra?: string } | undefined;
  let textStart = -1;
  let inF = false, inV = false;
  const shared = new Map<string, { f: string; row: number; col: number }>();
  const pendingShared: { row: number; col: number; si: string }[] = [];

  const tags = scanTags(xml, from, to);
  for (const t of tags) {
    switch (t.name) {
      case 'row': {
        if (t.close) { row = undefined; break; }
        const a = parseAttrs(t.attrs);
        rowIdx = a.r ? parseInt(a.r, 10) - 1 : rowIdx + 1;
        colCursor = -1;
        const r: Row = sheet.rows[rowIdx] ?? { cells: [] };
        const keep: Record<string, string> = {};
        for (const [k, v] of Object.entries(a)) {
          if (k === 'r' || k === 'spans') continue;
          if (k === 'ht') r.height = parseFloat(v);
          else if (k === 'hidden') { if (bool(v)) r.hidden = true; }
          else if (k === 'customHeight') continue;
          else keep[k] = v;
        }
        if (Object.keys(keep).length) r.keep = keep;
        sheet.rows[rowIdx] = r;
        row = t.selfClose ? undefined : r;
        break;
      }
      case 'c': {
        if (t.close) { finishCell(); break; }
        const attrs = parseAttrs(t.attrs);
        const col = attrs.r ? parseA1(attrs.r).col : colCursor + 1;
        if (attrs.r) {
          const rr = parseA1(attrs.r).row;
          if (rr !== rowIdx) throw new Error(`Cell ${attrs.r} is inside row ${rowIdx + 1}`);
        }
        colCursor = col;
        cell = { attrs, col };
        if (t.selfClose) finishCell();
        break;
      }
      case 'f': {
        if (!cell) break;
        if (t.close) { cell.f = decodeXml(xml.slice(textStart, t.start)); inF = false; break; }
        cell.fAttrs = parseAttrs(t.attrs);
        if (t.selfClose) { cell.f = cell.f ?? ''; break; }
        inF = true; textStart = t.end;
        break;
      }
      case 'v': {
        if (!cell) break;
        if (t.close) { cell.v = decodeXml(xml.slice(textStart, t.start)); inV = false; break; }
        if (t.selfClose) { cell.v = ''; break; }
        inV = true; textStart = t.end;
        break;
      }
      case 'is': {
        if (!cell || t.close) break;
        if (t.selfClose) { cell.is = { text: '' }; break; }
        const c = findClose(xml, 'is', t.end);
        cell.is = parseStringItem(xml, t.end, c.innerEnd, ss);
        skipTo(c.end);
        break;
      }
      case 'extLst': {
        if (!cell || t.close || t.selfClose) break;
        const c = findClose(xml, 'extLst', t.end);
        cell.extra = (cell.extra ?? '') + xml.slice(t.start, c.end);
        skipTo(c.end);
        break;
      }
      default:
        break;
    }
  }
  void inF; void inV;

  for (const p of pendingShared) {
    const m = shared.get(p.si);
    if (!m) throw new Error(`Shared formula ${p.si} has no master cell`);
    const c = sheet.rows[p.row]!.cells[p.col]!;
    c.f = translateFormula(m.f, p.row - m.row, p.col - m.col);
  }

  // The other cells of an array formula's range hold its results, which the engine recomputes.
  sheet.rows.forEach((r, ri) => r?.cells.forEach((c, ci) => {
    if (!c?.arrayRef) return;
    const rg = parseRange(c.arrayRef);
    for (let y = rg.r1; y <= rg.r2; y++) for (let x = rg.c1; x <= rg.c2; x++) {
      if (y === ri && x === ci) continue;
      const o = sheet.rows[y]?.cells[x];
      if (o && o.f === undefined) o.spill = true;
    }
  }));

  function skipTo(pos: number) {
    // Advance the generator past nested content we consumed ourselves.
    for (let nx = tags.next(); !nx.done; nx = tags.next()) if (nx.value.end >= pos) return;
  }

  function finishCell() {
    if (!cell || !row) { cell = undefined; return; }
    const { attrs, col } = cell;
    const t = attrs.t;
    let out: Cell;
    if (t === 'inlineStr' || cell.is) {
      const e = cell.is ?? { text: '' };
      out = { v: e.text };
      if (e.runs) out.rich = e.runs;
    } else {
      const val = cellValueOf(t, cell.v, sst, date1904);
      out = { v: val.v };
      if (val.sst !== undefined) out.sst = val.sst;
      // A copy: the shared-string table must keep the original runs to detect edits.
      if (val.entry?.runs) out.rich = structuredClone(val.entry.runs);
    }
    if (attrs.s !== undefined && attrs.s !== '0') out.s = parseInt(attrs.s, 10);
    const cKeep: Record<string, string> = {};
    for (const [k, v] of Object.entries(attrs)) if (k !== 'r' && k !== 's' && k !== 't') cKeep[k] = v;
    if (cell.f !== undefined && cell.fAttrs) {
      const fa = cell.fAttrs;
      const fKeep: Record<string, string> = {};
      for (const [k, v] of Object.entries(fa)) if (k !== 't' && k !== 'ref' && k !== 'si') fKeep[k] = v;
      if (fa.t === 'shared') {
        if (cell.f) shared.set(fa.si, { f: cell.f, row: rowIdx, col });
        else pendingShared.push({ row: rowIdx, col, si: fa.si });
        out.f = cell.f;
      } else if (fa.t === 'array') {
        out.f = cell.f;
        out.arrayRef = fa.ref ?? rangeToA1({ r1: rowIdx, c1: col, r2: rowIdx, c2: col });
      } else if (fa.t === 'dataTable') {
        // What-if data tables have no formula text of their own; keep the whole <f> as is.
        fKeep.t = 'dataTable';
        if (fa.ref) fKeep.ref = fa.ref;
        out.f = cell.f;
        out.frozen = true;
      } else {
        out.f = cell.f;
      }
      if (Object.keys(fKeep).length) (out.keep ??= {}).fAttrs = fKeep;
    } else if (cell.fAttrs && cell.f === undefined) {
      out.f = '';
    }
    if (Object.keys(cKeep).length) (out.keep ??= {}).cAttrs = cKeep;
    if (cell.extra) (out.keep ??= {}).extra = cell.extra;
    row.cells[col] = out;
    cell = undefined;
  }
}

function parseColorAttr(el: XEl | null, ss: StyleSheet): string | undefined {
  if (!el) return undefined;
  const c: ColorSpec = {};
  if (el.getAttribute('rgb')) c.rgb = el.getAttribute('rgb')!.slice(-6);
  if (el.getAttribute('theme')) c.theme = parseInt(el.getAttribute('theme')!, 10);
  if (el.getAttribute('indexed')) c.indexed = parseInt(el.getAttribute('indexed')!, 10);
  if (el.getAttribute('tint')) c.tint = parseFloat(el.getAttribute('tint')!);
  return resolveColor(ss, c);
}

function parseFilterColumn(fc: XEl, sheet: Sheet): ColumnFilter | null {
  const filters = firstChild(fc, 'filters');
  if (filters) {
    const values = childElements(filters, 'filter').map((f) => f.getAttribute('val') ?? '');
    if (bool(filters.getAttribute('blank') ?? undefined)) values.push('');
    if (childElements(filters, 'dateGroupItem').length) sheet.features.add('Date-grouped filters');
    return { values };
  }
  const custom = firstChild(fc, 'customFilters');
  if (custom) {
    const items = childElements(custom, 'customFilter').map((c) => ({ op: c.getAttribute('operator') ?? 'equal', val: c.getAttribute('val') ?? '' }));
    const and = bool(custom.getAttribute('and') ?? undefined);
    const cond = customToCondition(items, and);
    if (cond) return { condition: cond };
  }
  sheet.features.add('Advanced filter criteria');
  return null;
}

function customToCondition(items: { op: string; val: string }[], and: boolean): FilterCondition | null {
  if (items.length === 1) {
    const { op, val } = items[0];
    const wild = /^\*(.*)\*$/.exec(val);
    if (op === 'equal' && wild) return { op: 'contains', a: wild[1] };
    if (op === 'notEqual' && wild) return { op: 'notContains', a: wild[1] };
    if (op === 'notEqual' && val === ' ') return { op: 'notEmpty' };
    if (op === 'equal' && val.endsWith('*') && !val.startsWith('*')) return { op: 'startsWith', a: val.slice(0, -1) };
    if (op === 'equal' && val.startsWith('*') && !val.endsWith('*')) return { op: 'endsWith', a: val.slice(1) };
    const map: Record<string, FilterCondition['op']> = {
      equal: 'equals', notEqual: 'notEquals', greaterThan: 'gt', greaterThanOrEqual: 'gte', lessThan: 'lt', lessThanOrEqual: 'lte',
    };
    if (map[op]) return { op: map[op], a: val };
  }
  if (items.length === 2) {
    const [x, y] = items;
    if (and && x.op === 'greaterThanOrEqual' && y.op === 'lessThanOrEqual') return { op: 'between', a: x.val, b: y.val };
    if (!and && x.op === 'lessThan' && y.op === 'greaterThan') return { op: 'notBetween', a: x.val, b: y.val };
  }
  return null;
}

function parseWorksheet(xml: string, sheet: Sheet, src: SheetSource, sst: SstEntry[], ss: StyleSheet, files: Map<string, Uint8Array>, date1904: boolean): void {
  assertNoDtd(xml);
  // Cut <sheetData> out so the rest can be a small DOM; the rows are scanned directly.
  let dataOpen: { start: number; end: number; selfClose: boolean } | undefined;
  for (const t of scanTags(xml)) {
    if (t.name === 'sheetData' && !t.close) { dataOpen = { start: t.start, end: t.end, selfClose: t.selfClose }; break; }
  }
  if (!dataOpen) throw new Error(`${src.path} has no <sheetData>`);
  let rest: string;
  if (dataOpen.selfClose) rest = xml;
  else {
    const close = findClose(xml, 'sheetData', dataOpen.end);
    parseSheetData(xml, dataOpen.end, close.innerEnd, sheet, sst, ss, date1904);
    const tagName = xml.slice(dataOpen.start + 1, dataOpen.end - 1).trim().split(/\s/)[0];
    rest = xml.slice(0, dataOpen.start) + `<${tagName}/>` + xml.slice(close.end);
  }
  const doc = parseXml(rest);
  src.doc = doc;
  const root = doc.documentElement as unknown as XEl;

  const sheetPr = firstChild(root, 'sheetPr');
  if (sheetPr) sheet.tabColor = parseColorAttr(firstChild(sheetPr, 'tabColor'), ss);

  const views = firstChild(root, 'sheetViews');
  const view = views ? firstChild(views, 'sheetView') : null;
  if (view) {
    sheet.rightToLeft = bool(view.getAttribute('rightToLeft') ?? undefined);
    sheet.showGridLines = view.getAttribute('showGridLines') !== '0' && view.getAttribute('showGridLines') !== 'false';
    const pane = firstChild(view, 'pane');
    if (pane && (pane.getAttribute('state') === 'frozen' || pane.getAttribute('state') === 'frozenSplit')) {
      sheet.freeze = {
        rows: Math.round(parseFloat(pane.getAttribute('ySplit') || '0')),
        cols: Math.round(parseFloat(pane.getAttribute('xSplit') || '0')),
      };
    } else if (pane) {
      sheet.features.add('Split panes');
    }
  }

  const fmt = firstChild(root, 'sheetFormatPr');
  if (fmt) {
    if (fmt.getAttribute('defaultRowHeight')) sheet.defaultRowHeight = parseFloat(fmt.getAttribute('defaultRowHeight')!);
    if (fmt.getAttribute('defaultColWidth')) sheet.defaultColWidth = parseFloat(fmt.getAttribute('defaultColWidth')!);
    else sheet.defaultColWidth = (parseFloat(fmt.getAttribute('baseColWidth') || '8') * 7 + 5) / 7;
  } else {
    sheet.defaultColWidth = 9.140625;
    sheet.defaultRowHeight = 15;
  }

  const cols = firstChild(root, 'cols');
  if (cols) for (const c of childElements(cols, 'col')) {
    const min = parseInt(c.getAttribute('min')!, 10) - 1;
    const max = Math.min(parseInt(c.getAttribute('max')!, 10) - 1, 16383);
    const info: ColInfo = {};
    const keep: Record<string, string> = {};
    for (let i = 0; i < c.attributes.length; i++) {
      const a = c.attributes[i];
      switch (a.name) {
        case 'min': case 'max': case 'customWidth': break;
        case 'width': info.width = parseFloat(a.value); break;
        case 'hidden': if (bool(a.value)) info.hidden = true; break;
        case 'style': info.style = parseInt(a.value, 10); break;
        default: keep[a.name] = a.value;
      }
    }
    if (Object.keys(keep).length) info.keep = keep;
    for (let i = min; i <= max; i++) sheet.cols[i] = { ...info };
  }

  const merges = firstChild(root, 'mergeCells');
  if (merges) sheet.merges = childElements(merges, 'mergeCell').map((m) => parseRange(m.getAttribute('ref')!));

  const af = firstChild(root, 'autoFilter');
  if (af && af.getAttribute('ref')) {
    const range = parseRange(af.getAttribute('ref')!);
    sheet.filter = { range, columns: {} };
    for (const fc of childElements(af, 'filterColumn')) {
      const colId = parseInt(fc.getAttribute('colId')!, 10);
      const f = parseFilterColumn(fc, sheet);
      if (f) sheet.filter.columns[range.c1 + colId] = f;
    }
    if (firstChild(af, 'sortState')) sheet.features.add('Saved sort order');
    // Excel stores filter-hidden rows as plain hidden rows; attribute them to the filter.
    if (Object.keys(sheet.filter.columns).length) {
      for (let y = range.r1 + 1; y <= range.r2; y++) {
        const r = sheet.rows[y];
        if (r?.hidden) { r.hidden = false; r.filtered = true; }
      }
    }
  }

  const links = firstChild(root, 'hyperlinks');
  if (links) for (const h of childElements(links, 'hyperlink')) {
    const ref = h.getAttribute('ref');
    if (!ref) continue;
    const rid = relIdOf(h);
    const rel = rid ? src.rels.find((r) => r.id === rid) : undefined;
    let target = rel ? rel.target : '';
    const loc = h.getAttribute('location');
    if (loc) target = rel ? `${target}#${loc}` : `#${loc}`;
    if (!target) continue;
    const attrs: Record<string, string> = {};
    for (let i = 0; i < h.attributes.length; i++) {
      const a = h.attributes[i];
      if (a.name !== 'ref' && relIdOf(h) !== a.value) attrs[a.name] = a.value;
    }
    for (const r of parseSqref(ref)) {
      for (let y = r.r1; y <= r.r2; y++) for (let x = r.c1; x <= r.c2; x++) {
        const row = sheet.rows[y] ?? (sheet.rows[y] = { cells: [] });
        const c = row.cells[x] ?? (row.cells[x] = { v: null });
        c.link = target;
        src.links.set(rangeToA1({ r1: y, c1: x, r2: y, c2: x }), { target, attrs });
      }
    }
  }

  for (const el of childElements(root)) {
    const f = FEATURE_ELEMENTS[localName(el)];
    if (f) sheet.features.add(f);
    if (localName(el) === 'extLst') {
      const s = el.toString();
      if (/sparklineGroups/.test(s)) sheet.features.add('Sparklines');
      if (/conditionalFormattings/.test(s)) sheet.features.add('Conditional formatting');
      if (/dataValidations/.test(s)) sheet.features.add('Data validation');
      if (/slicerList/.test(s)) sheet.features.add('Slicers');
    }
  }
  for (const r of src.rels) {
    const f = REL_FEATURES[relType(r)];
    if (f) sheet.features.add(f);
    if (relType(r) === 'drawing' && files.has(r.target)) {
      const d = text(files, r.target)!;
      if (/<(\w+:)?chart\b/.test(d) || /relationships\/chart"/.test(text(files, relsPathFor(r.target)) ?? '')) sheet.features.add('Charts');
      if (/<(\w+:)?pic\b/.test(d)) sheet.features.add('Images');
    }
  }
  if (sheet.features.has('Charts') || sheet.features.has('Images')) sheet.features.delete('Charts, images or shapes');

  src.loaded = snapshotSheet(sheet);
}

export function snapshotSheet(sheet: Sheet): Record<string, string> {
  return {
    cols: JSON.stringify(sheet.cols),
    merges: JSON.stringify(sheet.merges),
    freeze: JSON.stringify(sheet.freeze),
    filter: JSON.stringify(sheet.filter),
    tab: JSON.stringify([sheet.tabColor, sheet.showGridLines, sheet.rightToLeft]),
    defaults: JSON.stringify([sheet.defaultColWidth, sheet.defaultRowHeight]),
  };
}

// ---------------------------------------------------------------------------
// Workbook

export function readXlsx(bytes: Uint8Array, format: 'xlsx' | 'xlsm' = 'xlsx'): Workbook {
  let entries: Record<string, Uint8Array>;
  try {
    entries = unzipSync(bytes);
  } catch (e) {
    throw new Error(`This file is not a valid XLSX package (${(e as Error).message}).`);
  }
  const files = new Map(Object.entries(entries));
  if (!files.has('[Content_Types].xml')) throw new Error('This file is not a valid XLSX package (no [Content_Types].xml).');

  const rootRels = parseRels(files, '');
  const officeDoc = rootRels.find((r) => relType(r) === 'officeDocument');
  if (!officeDoc) throw new Error('This package has no workbook part.');
  const workbookPath = officeDoc.target;
  const wbXml = text(files, workbookPath);
  if (!wbXml) throw new Error(`Workbook part ${workbookPath} is missing.`);
  const workbookDoc = parseXml(wbXml);
  const wbRoot = workbookDoc.documentElement as unknown as XEl;
  const workbookRels = parseRels(files, workbookPath);

  const stylesRel = workbookRels.find((r) => relType(r) === 'styles');
  const themeRel = workbookRels.find((r) => relType(r) === 'theme');
  const styles = stylesRel && files.has(stylesRel.target)
    ? parseStyles(text(files, stylesRel.target)!, themeRel ? text(files, themeRel.target) : undefined)
    : emptyStyleSheet();

  const sstRel = workbookRels.find((r) => relType(r) === 'sharedStrings');
  const sst = sstRel && files.has(sstRel.target) ? parseSst(text(files, sstRel.target)!, styles) : [];

  const wbPr = firstChild(wbRoot, 'workbookPr');
  const date1904 = bool(wbPr?.getAttribute('date1904') ?? undefined);

  const features = new Set<string>();
  if (files.has('xl/vbaProject.bin') || workbookRels.some((r) => relType(r) === 'vbaProject')) features.add('Macros (VBA)');
  if (workbookRels.some((r) => relType(r) === 'pivotCacheDefinition')) features.add('Pivot tables');
  if (workbookRels.some((r) => relType(r) === 'externalLink')) features.add('Links to other workbooks');
  if (workbookRels.some((r) => relType(r) === 'connections')) features.add('Data connections');
  if ([...files.keys()].some((p) => p.startsWith('customXml/') || p.includes('/customXml/'))) { /* passive, preserved */ }
  if (firstChild(wbRoot, 'workbookProtection')) features.add('Workbook protection');

  const wb: Workbook = {
    sheets: [], styles, definedNames: [], date1904, activeSheet: 0, format,
    features, unsupportedFunctions: new Set(), viewStyles: new Set(),
    src: { files, workbookPath, workbookDoc, workbookRels, sst, sstPath: sstRel?.target, stylesPath: stylesRel?.target },
  };

  const sheetsEl = firstChild(wbRoot, 'sheets');
  if (!sheetsEl) throw new Error('Workbook has no <sheets>.');
  for (const s of childElements(sheetsEl, 'sheet')) {
    const name = s.getAttribute('name')!;
    const rid = relIdOf(s);
    const rel = workbookRels.find((r) => r.id === rid);
    if (!rel) throw new Error(`Sheet "${name}" has no relationship ${rid}`);
    const sheet = newSheet(name);
    sheet.pristine = true;
    const st = s.getAttribute('state');
    if (st === 'hidden' || st === 'veryHidden') sheet.state = st;
    const src: SheetSource = { path: rel.target, relId: rid!, sheetId: s.getAttribute('sheetId')!, rels: parseRels(files, rel.target), loaded: {}, links: new Map() };
    sheet.src = src;
    const xml = text(files, rel.target);
    if (relType(rel) !== 'worksheet' || xml === undefined) {
      sheet.kind = 'other';
      sheet.features.add(relType(rel) === 'chartsheet' ? 'Chart sheet' : relType(rel) === 'dialogsheet' ? 'Dialog sheet' : 'Macro sheet');
    } else {
      try {
        parseWorksheet(xml, sheet, src, sst, styles, files, date1904);
      } catch (e) {
        throw new Error(`Couldn't read sheet "${name}": ${(e as Error).message}`);
      }
    }
    wb.sheets.push(sheet);
  }

  const names = firstChild(wbRoot, 'definedNames');
  if (names) for (const d of childElements(names, 'definedName')) {
    const dn: DefinedName = { name: d.getAttribute('name')!, ref: d.textContent ?? '' };
    if (d.getAttribute('localSheetId') !== null && d.getAttribute('localSheetId') !== '') dn.localSheet = parseInt(d.getAttribute('localSheetId')!, 10);
    if (bool(d.getAttribute('hidden') ?? undefined)) dn.hidden = true;
    const keep: Record<string, string> = {};
    for (let i = 0; i < d.attributes.length; i++) {
      const a = d.attributes[i];
      if (!['name', 'localSheetId', 'hidden'].includes(a.name)) keep[a.name] = a.value;
    }
    if (Object.keys(keep).length) dn.keep = keep;
    wb.definedNames.push(dn);
  }

  const views = firstChild(wbRoot, 'bookViews');
  const v = views ? firstChild(views, 'workbookView') : null;
  const active = v ? parseInt(v.getAttribute('activeTab') || '0', 10) : 0;
  wb.activeSheet = Math.min(Math.max(active, 0), wb.sheets.length - 1);
  if (wb.sheets[wb.activeSheet].state !== 'visible') wb.activeSheet = Math.max(0, wb.sheets.findIndex((s) => s.state === 'visible'));
  return wb;
}

export { attrsToXml };
