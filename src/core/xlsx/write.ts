import { strFromU8, strToU8, zipSync, Zippable } from 'fflate';
import { colToName, quoteSheetName, Range, rangeToA1, toA1 } from '../addr';
import { Cell, ColInfo, ColumnFilter, DefinedName, isError, Run, RunFont, Sheet, sheetExtent, Workbook } from '../model';
import { serializeStyles, stylesChanged } from '../styles';
import { parseRels, partDir, relsPathFor, snapshotSheet, SstEntry, XlsxSource } from './read';
import {
  attrsToXml, childElements, createLike, escapeXml, firstChild, insertInOrder, localName, parseXml, serializeXml, XDoc, XEl,
} from './xml';

// Saving patches the original package: every part we don't model is copied
// through untouched, untouched sheets keep their exact bytes, and modified
// sheets keep every element except the ones regenerated from the model.

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_REL_DOC = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';

const CT = {
  worksheet: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
  sharedStrings: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
  styles: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
  workbook: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  workbookMacro: 'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
};

const WORKSHEET_ORDER = [
  'sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges',
  'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting',
  'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks',
  'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF',
  'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst',
];

const WORKBOOK_ORDER = [
  'fileVersion', 'fileSharing', 'workbookPr', 'workbookProtection', 'bookViews', 'sheets', 'functionGroups', 'externalReferences',
  'definedNames', 'calcPr', 'oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing',
  'fileRecoveryPr', 'webPublishObjects', 'extLst',
];

const SHEET_PR_ORDER = ['tabColor', 'outlinePr', 'pageSetUpPr'];

const NEW_SHEET_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_REL_DOC}"><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;

// ---------------------------------------------------------------------------
// Package plumbing

class Pkg {
  files: Map<string, Uint8Array>;
  private docs = new Map<string, XDoc>();
  /** Parts whose DOM was edited; only these are re-serialized, the rest keep their bytes. */
  private dirty = new Set<string>();
  /** Base URI for relationship types, which differs between Transitional and Strict files. */
  relBase: string;

  constructor(files: Map<string, Uint8Array>) {
    this.files = new Map(files);
    const root = this.doc('_rels/.rels');
    const office = childElements(root.documentElement as unknown as XEl, 'Relationship')
      .find((r) => (r.getAttribute('Type') ?? '').endsWith('/officeDocument'));
    if (!office) throw new Error('Package has no officeDocument relationship');
    this.relBase = office.getAttribute('Type')!.replace(/\/officeDocument$/, '');
  }

  touch(path: string): void { this.dirty.add(path); }

  /** A part as a DOM; call touch() after editing it so zip() writes it back. */
  doc(path: string): XDoc {
    let d = this.docs.get(path);
    if (!d) {
      const b = this.files.get(path);
      if (!b) throw new Error(`Package part ${path} is missing`);
      d = parseXml(strFromU8(b));
      this.docs.set(path, d);
    }
    return d;
  }

  hasDoc(path: string): boolean { return this.docs.has(path) || this.files.has(path); }

  setText(path: string, xml: string): void {
    this.docs.delete(path);
    this.dirty.delete(path);
    this.files.set(path, strToU8(xml));
  }

  remove(path: string): void {
    this.docs.delete(path);
    this.dirty.delete(path);
    this.files.delete(path);
  }

  private rels(part: string): XEl {
    const p = relsPathFor(part);
    if (!this.hasDoc(p)) this.setText(p, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS_PKG_REL}"/>`);
    this.touch(p);
    return this.doc(p).documentElement as unknown as XEl;
  }

  addRel(part: string, type: string, target: string, external = false): string {
    const root = this.rels(part);
    const ids = new Set(childElements(root, 'Relationship').map((r) => r.getAttribute('Id')));
    let n = ids.size + 1;
    while (ids.has(`rId${n}`)) n++;
    const el = createLike(root, 'Relationship');
    el.setAttribute('Id', `rId${n}`);
    el.setAttribute('Type', `${this.relBase}/${type}`);
    el.setAttribute('Target', external ? target : relativeTarget(part, target));
    if (external) el.setAttribute('TargetMode', 'External');
    root.appendChild(el);
    return `rId${n}`;
  }

  removeRels(part: string, pred: (type: string, id: string) => boolean): void {
    const p = relsPathFor(part);
    if (!this.hasDoc(p)) return;
    const root = this.doc(p).documentElement as unknown as XEl;
    for (const r of childElements(root, 'Relationship')) {
      if (pred((r.getAttribute('Type') ?? '').replace(/^.*\//, ''), r.getAttribute('Id')!)) { root.removeChild(r); this.touch(p); }
    }
  }

  setOverride(part: string, contentType: string | null): void {
    this.touch('[Content_Types].xml');
    const root = this.doc('[Content_Types].xml').documentElement as unknown as XEl;
    for (const o of childElements(root, 'Override')) {
      if (o.getAttribute('PartName') === `/${part}`) root.removeChild(o);
    }
    if (contentType) {
      const o = createLike(root, 'Override');
      o.setAttribute('PartName', `/${part}`);
      o.setAttribute('ContentType', contentType);
      root.appendChild(o);
    }
  }

  zip(): Uint8Array {
    for (const path of this.dirty) this.files.set(path, strToU8(serializeXml(this.docs.get(path)!)));
    this.dirty.clear();
    const z: Zippable = {};
    // [Content_Types].xml first, as Office writes it; the rest in original order.
    const ct = this.files.get('[Content_Types].xml');
    if (!ct) throw new Error('Package has no [Content_Types].xml');
    z['[Content_Types].xml'] = ct;
    for (const [path, bytes] of this.files) if (path !== '[Content_Types].xml') z[path] = bytes;
    return zipSync(z, { level: 6 });
  }
}

function relativeTarget(fromPart: string, target: string): string {
  const from = partDir(fromPart).split('/').filter(Boolean);
  const to = target.split('/');
  let i = 0;
  while (i < from.length && i < to.length - 1 && from[i] === to[i]) i++;
  return [...Array(from.length - i).fill('..'), ...to.slice(i)].join('/');
}

function blankPackage(macro: boolean): XlsxSource {
  const files = new Map<string, Uint8Array>();
  const x = (s: string) => strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${s}`);
  files.set('[Content_Types].xml', x(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="${macro ? CT.workbookMacro : CT.workbook}"/></Types>`));
  files.set('_rels/.rels', x(`<Relationships xmlns="${NS_PKG_REL}"><Relationship Id="rId1" Type="${NS_REL_DOC}/officeDocument" Target="xl/workbook.xml"/></Relationships>`));
  const wbXml = `<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL_DOC}"><bookViews><workbookView/></bookViews><sheets/><calcPr calcId="191029"/></workbook>`;
  files.set('xl/workbook.xml', x(wbXml));
  return { files, workbookPath: 'xl/workbook.xml', workbookDoc: parseXml(wbXml), workbookRels: [], sst: [] };
}

// ---------------------------------------------------------------------------
// Shared strings

class SstWriter {
  private byText = new Map<string, number>();
  added: SstEntry[] = [];
  constructor(private existing: SstEntry[]) {
    existing.forEach((e, i) => { if (!e.runs && !this.byText.has(e.text)) this.byText.set(e.text, i); });
  }

  /** The string's index, reusing the cell's original entry while it still matches exactly. Records it on the cell. */
  index(c: Cell): number {
    const prev = c.sst === undefined ? undefined : this.existing[c.sst];
    if (prev && prev.text === c.v && JSON.stringify(prev.runs) === JSON.stringify(c.rich)) return c.sst!;
    let i: number;
    if (c.rich) {
      this.added.push({ text: c.v as string, runs: structuredClone(c.rich) });
      i = this.existing.length + this.added.length - 1;
    } else {
      const t = c.v as string;
      const hit = this.byText.get(t);
      if (hit !== undefined) i = hit;
      else {
        this.added.push({ text: t });
        i = this.existing.length + this.added.length - 1;
        this.byText.set(t, i);
      }
    }
    c.sst = i;
    return i;
  }

  get total(): number { return this.existing.length + this.added.length; }
}

function tXml(s: string): string {
  const sp = /^\s|\s$|\n/.test(s) ? ' xml:space="preserve"' : '';
  return `<t${sp}>${escapeXml(s)}</t>`;
}

function rPrXml(f: RunFont): string {
  let s = '';
  if (f.name) s += `<rFont val="${escapeXml(f.name)}"/>`;
  if (f.bold) s += '<b/>';
  if (f.italic) s += '<i/>';
  if (f.strike) s += '<strike/>';
  if (f.color) s += `<color rgb="FF${f.color.replace('#', '').toUpperCase()}"/>`;
  if (f.size) s += `<sz val="${f.size}"/>`;
  if (f.underline) s += '<u/>';
  return `<rPr>${s}</rPr>`;
}

function sameFont(a: RunFont | undefined, b: RunFont | undefined): boolean {
  return JSON.stringify(a ?? {}) === JSON.stringify(b ?? {});
}

export function runsXml(runs: Run[]): string {
  return runs.map((r) => {
    let rPr = '';
    if (r.rPr !== undefined && sameFont(r.font, r.rPrFont)) rPr = r.rPr;
    else if (r.font && Object.values(r.font).some((v) => v !== undefined)) rPr = rPrXml(r.font);
    return `<r>${rPr}${tXml(r.text)}</r>`;
  }).join('');
}

function siXml(e: SstEntry): string {
  return `<si>${e.runs ? runsXml(e.runs) : tXml(e.text)}</si>`;
}

function writeSst(pkg: Pkg, src: XlsxSource, sst: SstWriter): void {
  if (!sst.added.length) return;
  const added = sst.added.map(siXml).join('');
  if (src.sstPath && pkg.files.has(src.sstPath)) {
    // Append in place so the original entries (and their exact markup) stay put.
    let xml = strFromU8(pkg.files.get(src.sstPath)!);
    xml = xml.replace(/<(\w+:)?sst\b([^>]*?)\s*\/>/, (_m, p: string | undefined, a: string) => `<${p ?? ''}sst${a}></${p ?? ''}sst>`);
    const close = xml.search(/<\/(\w+:)?sst>\s*$/);
    if (close < 0) throw new Error('sharedStrings.xml has no closing </sst>');
    const prefix = /^<\/(\w+:)?/.exec(xml.slice(close))![1] ?? '';
    const addedXml = prefix ? added.replace(/<(\/?)(si|r|rPr|t|rFont|b|i|strike|color|sz|u)\b/g, `<$1${prefix}$2`) : added;
    let head = xml.slice(0, close);
    head = head.replace(/<(\w+:)?sst\b[^>]*>/, (open) => open
      .replace(/\s+count="\d*"/, '')
      .replace(/\s+uniqueCount="\d*"/, '')
      .replace(/(\s*\/?)>$/, ` uniqueCount="${sst.total}"$1>`));
    pkg.setText(src.sstPath, head + addedXml + xml.slice(close));
  } else {
    const path = 'xl/sharedStrings.xml';
    pkg.setText(path, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="${NS_MAIN}" uniqueCount="${sst.total}">${added}</sst>`);
    pkg.addRel(src.workbookPath, 'sharedStrings', path);
    pkg.setOverride(path, CT.sharedStrings);
    src.sstPath = path;
  }
}

// ---------------------------------------------------------------------------
// Sheet data

function numStr(n: number): string {
  if (!Number.isFinite(n)) throw new Error(`Can't save a non-finite number (${n})`);
  return String(n);
}

const EXCEL_ERRORS = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!']);

function cellXml(c: Cell, ref: string, p: string, sst: SstWriter): string {
  const keep = c.keep;
  let t: string | undefined;
  let v: string | undefined;
  const val = c.v;
  if (val === null) v = undefined;
  else if (typeof val === 'number') v = numStr(val);
  else if (typeof val === 'boolean') { t = 'b'; v = val ? '1' : '0'; }
  // The engine has errors Excel doesn't (#CYCLE!, #ERROR!); Excel rejects those as a cached value.
  else if (isError(val)) { t = 'e'; v = EXCEL_ERRORS.has(val.error) ? val.error : '#VALUE!'; }
  else if (c.f !== undefined) { t = 'str'; v = val; }
  else { t = 's'; v = String(sst.index(c)); }

  const attrs: Record<string, string | number | undefined> = { r: ref, s: c.s || undefined, t, ...(keep?.cAttrs ?? {}) };
  let inner = '';
  if (c.f !== undefined) {
    const fa: Record<string, string | undefined> = { ...(keep?.fAttrs ?? {}) };
    if (c.arrayRef) { fa.t = 'array'; fa.ref = c.arrayRef; }
    inner += c.f === '' ? `<${p}f${attrsToXml(fa)}/>` : `<${p}f${attrsToXml(fa)}>${escapeXml(c.f)}</${p}f>`;
  }
  if (v !== undefined) inner += `<${p}v>${escapeXml(v)}</${p}v>`;
  if (keep?.extra) inner += keep.extra;
  return inner ? `<${p}c${attrsToXml(attrs)}>${inner}</${p}c>` : `<${p}c${attrsToXml(attrs)}/>`;
}

function sheetDataXml(sheet: Sheet, p: string, sst: SstWriter): string {
  const out: string[] = [];
  sheet.rows.forEach((row, ri) => {
    if (!row) return;
    let cells = '';
    let min = -1, max = -1;
    row.cells.forEach((c, ci) => {
      if (!c) return;
      if (c.v === null && c.f === undefined && !c.s && !c.keep) return;
      if (min < 0) min = ci;
      max = ci;
      cells += cellXml(c, toA1(ri, ci), p, sst);
    });
    const hidden = row.hidden || row.filtered;
    if (!cells && row.height === undefined && !hidden && !row.keep) return;
    const attrs: Record<string, string | number | undefined> = {
      r: ri + 1,
      spans: min >= 0 ? `${min + 1}:${max + 1}` : undefined,
      ...(row.keep ?? {}),
      ht: row.height,
      customHeight: row.height !== undefined ? 1 : undefined,
      hidden: hidden ? 1 : undefined,
    };
    out.push(cells ? `<${p}row${attrsToXml(attrs)}>${cells}</${p}row>` : `<${p}row${attrsToXml(attrs)}/>`);
  });
  return out.join('');
}

// ---------------------------------------------------------------------------
// Sheet-level elements

function replaceChild(root: XEl, name: string, build: ((el: XEl) => void) | null, order: string[]): void {
  const old = firstChild(root, name);
  if (old) root.removeChild(old);
  if (!build) return;
  const el = createLike(root, name);
  build(el);
  insertInOrder(root, el, order);
}

function child(parent: XEl, name: string, attrs: Record<string, string | number | undefined> = {}): XEl {
  const el = createLike(parent, name);
  for (const [k, v] of Object.entries(attrs)) if (v !== undefined) el.setAttribute(k, String(v));
  parent.appendChild(el);
  return el;
}

function colsEqual(a: ColInfo | undefined, b: ColInfo | undefined): boolean {
  return JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
}

function writeCols(root: XEl, sheet: Sheet): void {
  const spans: { min: number; max: number; info: ColInfo }[] = [];
  sheet.cols.forEach((info, i) => {
    if (!info || (info.width === undefined && !info.hidden && info.style === undefined && !info.keep)) return;
    const last = spans[spans.length - 1];
    if (last && last.max === i - 1 && colsEqual(last.info, info)) last.max = i;
    else spans.push({ min: i, max: i, info });
  });
  replaceChild(root, 'cols', spans.length ? (el) => {
    for (const s of spans) {
      child(el, 'col', {
        min: s.min + 1, max: s.max + 1,
        width: s.info.width ?? (s.info.hidden ? 0 : sheet.defaultColWidth),
        style: s.info.style,
        hidden: s.info.hidden ? 1 : undefined,
        customWidth: s.info.width !== undefined ? 1 : undefined,
        ...(s.info.keep ?? {}),
      });
    }
  } : null, WORKSHEET_ORDER);
}

function writeFreeze(root: XEl, sheet: Sheet): void {
  let views = firstChild(root, 'sheetViews');
  if (!views) {
    views = createLike(root, 'sheetViews');
    insertInOrder(root, views, WORKSHEET_ORDER);
  }
  let view = firstChild(views, 'sheetView');
  if (!view) view = child(views, 'sheetView', { workbookViewId: 0 });
  for (const el of childElements(view)) {
    const n = localName(el);
    if (n === 'pane' || n === 'selection') view.removeChild(el);
  }
  const { rows, cols } = sheet.freeze;
  if (!rows && !cols) return;
  const activePane = rows && cols ? 'bottomRight' : rows ? 'bottomLeft' : 'topRight';
  const pane = createLike(view, 'pane');
  if (cols) pane.setAttribute('xSplit', String(cols));
  if (rows) pane.setAttribute('ySplit', String(rows));
  pane.setAttribute('topLeftCell', toA1(rows, cols));
  pane.setAttribute('activePane', activePane);
  pane.setAttribute('state', 'frozen');
  view.insertBefore(pane, view.firstChild);
  const sel = createLike(view, 'selection');
  sel.setAttribute('pane', activePane);
  view.insertBefore(sel, pane.nextSibling);
}

function writeViewFlags(root: XEl, sheet: Sheet): void {
  const views = firstChild(root, 'sheetViews');
  const view = views ? firstChild(views, 'sheetView') : null;
  if (view) {
    if (sheet.rightToLeft) view.setAttribute('rightToLeft', '1'); else view.removeAttribute('rightToLeft');
    if (!sheet.showGridLines) view.setAttribute('showGridLines', '0'); else view.removeAttribute('showGridLines');
  }
  let pr = firstChild(root, 'sheetPr');
  const tab = pr ? firstChild(pr, 'tabColor') : null;
  if (tab) pr!.removeChild(tab);
  if (sheet.tabColor) {
    if (!pr) { pr = createLike(root, 'sheetPr'); insertInOrder(root, pr, WORKSHEET_ORDER); }
    const el = createLike(pr, 'tabColor');
    el.setAttribute('rgb', `FF${sheet.tabColor.replace('#', '').toUpperCase()}`);
    insertInOrder(pr, el, SHEET_PR_ORDER);
  }
}

function writeDefaults(root: XEl, sheet: Sheet): void {
  let fmt = firstChild(root, 'sheetFormatPr');
  if (!fmt) { fmt = createLike(root, 'sheetFormatPr'); insertInOrder(root, fmt, WORKSHEET_ORDER); }
  fmt.setAttribute('defaultRowHeight', String(sheet.defaultRowHeight));
  fmt.setAttribute('defaultColWidth', String(sheet.defaultColWidth));
}

export function filterColumnXml(el: XEl, cf: ColumnFilter): void {
  if (cf.values) {
    const blank = cf.values.includes('');
    const filters = child(el, 'filters', { blank: blank ? 1 : undefined });
    for (const v of cf.values) if (v !== '') child(filters, 'filter', { val: v });
    return;
  }
  const c = cf.condition;
  if (!c) throw new Error('Filter column has neither values nor a condition');
  if (c.op === 'empty') { child(el, 'filters', { blank: 1 }); return; }
  const single: Record<string, [string, (a: string) => string]> = {
    notEmpty: ['notEqual', () => ' '],
    contains: ['equal', (a) => `*${a}*`], notContains: ['notEqual', (a) => `*${a}*`],
    startsWith: ['equal', (a) => `${a}*`], endsWith: ['equal', (a) => `*${a}`],
    equals: ['equal', (a) => a], notEquals: ['notEqual', (a) => a],
    gt: ['greaterThan', (a) => a], gte: ['greaterThanOrEqual', (a) => a], lt: ['lessThan', (a) => a], lte: ['lessThanOrEqual', (a) => a],
  };
  if (c.op === 'between' || c.op === 'notBetween') {
    const between = c.op === 'between';
    const custom = child(el, 'customFilters', { and: between ? 1 : undefined });
    child(custom, 'customFilter', { operator: between ? 'greaterThanOrEqual' : 'lessThan', val: c.a ?? '' });
    child(custom, 'customFilter', { operator: between ? 'lessThanOrEqual' : 'greaterThan', val: c.b ?? '' });
    return;
  }
  const [operator, val] = single[c.op];
  const custom = child(el, 'customFilters');
  child(custom, 'customFilter', { operator, val: val(c.a ?? '') });
}

function writeAutoFilter(root: XEl, sheet: Sheet): void {
  const f = sheet.filter;
  replaceChild(root, 'autoFilter', f ? (el) => {
    el.setAttribute('ref', rangeToA1(f.range));
    for (const [col, cf] of Object.entries(f.columns).sort((a, b) => Number(a[0]) - Number(b[0]))) {
      const fc = child(el, 'filterColumn', { colId: Number(col) - f.range.c1 });
      filterColumnXml(fc, cf);
    }
  } : null, WORKSHEET_ORDER);
}

function writeMerges(root: XEl, sheet: Sheet): void {
  replaceChild(root, 'mergeCells', sheet.merges.length ? (el) => {
    el.setAttribute('count', String(sheet.merges.length));
    for (const m of sheet.merges) child(el, 'mergeCell', { ref: rangeToA1(m) });
  } : null, WORKSHEET_ORDER);
}

function relPrefix(root: XEl): string {
  for (let i = 0; i < root.attributes.length; i++) {
    const a = root.attributes[i];
    if (a.name.startsWith('xmlns:') && a.value.endsWith('/relationships')) return a.name.slice(6);
  }
  root.setAttribute('xmlns:r', NS_REL_DOC);
  return 'r';
}

function cellLinks(sheet: Sheet): Map<string, string> {
  const out = new Map<string, string>();
  sheet.rows.forEach((row, ri) => row?.cells.forEach((c, ci) => {
    if (!c) return;
    const link = c.link ?? c.rich?.find((r) => r.link)?.link;
    if (link) out.set(toA1(ri, ci), link);
  }));
  return out;
}

function writeHyperlinks(pkg: Pkg, root: XEl, sheet: Sheet, path: string): void {
  const links = cellLinks(sheet);
  const loaded = sheet.src?.links;
  const same = loaded && loaded.size === links.size && [...links].every(([ref, t]) => loaded.get(ref)?.target === t);
  if (same) return;
  // Hyperlink relationships belong to <hyperlinks> alone, so they are rebuilt with it.
  pkg.removeRels(path, (type) => type === 'hyperlink');
  const prefix = links.size ? relPrefix(root) : '';
  replaceChild(root, 'hyperlinks', links.size ? (el) => {
    for (const [ref, target] of links) {
      const h = child(el, 'hyperlink', { ref });
      const prev = loaded?.get(ref);
      if (prev && prev.target === target) {
        for (const [k, v] of Object.entries(prev.attrs)) if (k !== 'location') h.setAttribute(k, v);
      }
      const hash = target.indexOf('#');
      if (hash === 0) h.setAttribute('location', target.slice(1));
      else {
        const url = hash > 0 && !/^[a-z]+:/i.test(target) ? target.slice(0, hash) : target;
        if (hash > 0 && url !== target) h.setAttribute('location', target.slice(hash + 1));
        h.setAttributeNS(root.lookupNamespaceURI(prefix) ?? NS_REL_DOC, `${prefix}:id`, pkg.addRel(path, 'hyperlink', url, true));
      }
    }
  } : null, WORKSHEET_ORDER);
}

function writeDimension(root: XEl, sheet: Sheet): void {
  const ext = sheetExtent(sheet);
  const ref = ext.rows && ext.cols ? rangeToA1({ r1: 0, c1: 0, r2: ext.rows - 1, c2: ext.cols - 1 }) : 'A1';
  replaceChild(root, 'dimension', (el) => el.setAttribute('ref', ref), WORKSHEET_ORDER);
}

function writeSheet(pkg: Pkg, sheet: Sheet, path: string, sst: SstWriter): void {
  const doc = sheet.src?.doc ? (sheet.src.doc.cloneNode(true) as XDoc) : parseXml(NEW_SHEET_XML);
  const root = doc.documentElement as unknown as XEl;
  const loaded = sheet.src?.loaded ?? {};
  const now = snapshotSheet(sheet);
  writeDimension(root, sheet);
  if (now.tab !== loaded.tab) writeViewFlags(root, sheet);
  if (now.freeze !== loaded.freeze) writeFreeze(root, sheet);
  if (now.defaults !== loaded.defaults) writeDefaults(root, sheet);
  if (now.cols !== loaded.cols) writeCols(root, sheet);
  if (now.filter !== loaded.filter) writeAutoFilter(root, sheet);
  if (now.merges !== loaded.merges) writeMerges(root, sheet);
  writeHyperlinks(pkg, root, sheet, path);

  const sd = firstChild(root, 'sheetData');
  if (!sd) throw new Error(`Sheet ${sheet.name} lost its <sheetData>`);
  const p = sd.prefix ? `${sd.prefix}:` : '';
  const xml = serializeXml(doc);
  const placeholder = new RegExp(`<${p}sheetData\\s*/>`);
  if (!placeholder.test(xml)) throw new Error('Could not find the <sheetData/> placeholder');
  const data = sheetDataXml(sheet, p, sst);
  pkg.setText(path, xml.replace(placeholder, () => (data ? `<${p}sheetData>${data}</${p}sheetData>` : `<${p}sheetData/>`)));
}

// ---------------------------------------------------------------------------
// Workbook

function filterDbRef(sheet: Sheet, r: Range): string {
  const abs = (row: number, col: number) => `$${colToName(col)}$${row + 1}`;
  return `${quoteSheetName(sheet.name)}!${abs(r.r1, r.c1)}:${abs(r.r2, r.c2)}`;
}

/** Defined names to write: the model's, with each sheet's hidden _FilterDatabase name kept in step with its filter. */
export function namesToWrite(wb: Workbook): DefinedName[] {
  const FD = '_xlnm._FilterDatabase';
  const fresh = (i: number): DefinedName => ({ name: FD, localSheet: i, hidden: true, ref: filterDbRef(wb.sheets[i], wb.sheets[i].filter!.range) });
  const unchanged = (s: Sheet) => s.src?.loaded.filter === JSON.stringify(s.filter);
  const out: DefinedName[] = [];
  const seen = new Set<number>();
  for (const n of wb.definedNames) {
    if (n.name !== FD) { out.push(n); continue; }
    const s = n.localSheet === undefined ? undefined : wb.sheets[n.localSheet];
    if (!s?.filter) continue;
    seen.add(n.localSheet!);
    out.push(unchanged(s) ? n : { ...n, ...fresh(n.localSheet!) });
  }
  wb.sheets.forEach((s, i) => { if (s.filter && !seen.has(i)) out.push(fresh(i)); });
  return out;
}

function writeDefinedNames(root: XEl, names: DefinedName[]): void {
  replaceChild(root, 'definedNames', names.length ? (el) => {
    for (const n of names) {
      const d = child(el, 'definedName', { name: n.name, localSheetId: n.localSheet, hidden: n.hidden ? 1 : undefined, ...(n.keep ?? {}) });
      d.appendChild(el.ownerDocument!.createTextNode(n.ref));
    }
  } : null, WORKBOOK_ORDER);
}

function originalNames(root: XEl): string {
  const el = firstChild(root, 'definedNames');
  if (!el) return '[]';
  return JSON.stringify(childElements(el, 'definedName').map((d) => [d.getAttribute('name'), d.getAttribute('localSheetId') || null, d.textContent]));
}

function nameKey(names: DefinedName[]): string {
  return JSON.stringify(names.map((n) => [n.name, n.localSheet === undefined ? null : String(n.localSheet), n.ref]));
}

function nextSheetPath(pkg: Pkg): string {
  for (let n = 1; ; n++) {
    const p = `xl/worksheets/sheet${n}.xml`;
    if (!pkg.files.has(p)) return p;
  }
}

/** Serializes the workbook to XLSX bytes. The workbook's source package is updated to match, so repeated saves stay patches. */
export function writeXlsx(wb: Workbook): Uint8Array {
  const macro = wb.format === 'xlsm';
  const src = wb.src ?? blankPackage(macro);
  const pkg = new Pkg(src.files);
  const sst = new SstWriter(src.sst);
  const wbDoc = pkg.doc(src.workbookPath);
  const wbRoot = wbDoc.documentElement as unknown as XEl;
  let changed = false;

  // Sheets: map model sheets onto workbook <sheet> entries, adding and removing parts as needed.
  const sheetsEl = firstChild(wbRoot, 'sheets');
  if (!sheetsEl) throw new Error('Workbook has no <sheets>');
  const original = childElements(sheetsEl);
  const relAttr = (el: XEl) => {
    for (let i = 0; i < el.attributes.length; i++) {
      const a = el.attributes[i];
      if (a.name.endsWith(':id') && (a.namespaceURI ?? '').endsWith('/relationships')) return a;
    }
    return null;
  };
  const byRel = new Map(original.map((el) => [relAttr(el)?.value, el]));
  const kept = new Set<string>();
  let maxSheetId = Math.max(0, ...original.map((el) => parseInt(el.getAttribute('sheetId') ?? '0', 10)));
  const newSheetEls: XEl[] = [];
  for (const sheet of wb.sheets) {
    if (sheet.src && byRel.has(sheet.src.relId)) {
      const el = byRel.get(sheet.src.relId)!;
      kept.add(sheet.src.relId);
      if (el.getAttribute('name') !== sheet.name) { el.setAttribute('name', sheet.name); changed = true; }
      if ((el.getAttribute('state') || 'visible') !== sheet.state) {
        if (sheet.state === 'visible') el.removeAttribute('state'); else el.setAttribute('state', sheet.state);
        changed = true;
      }
      newSheetEls.push(el);
      if (sheet.kind === 'worksheet' && !sheet.pristine) { writeSheet(pkg, sheet, sheet.src.path, sst); changed = true; }
    } else {
      const path = nextSheetPath(pkg);
      const relId = pkg.addRel(src.workbookPath, 'worksheet', path);
      pkg.setOverride(path, CT.worksheet);
      const el = createLike(sheetsEl, 'sheet');
      el.setAttribute('name', sheet.name);
      el.setAttribute('sheetId', String(++maxSheetId));
      if (sheet.state !== 'visible') el.setAttribute('state', sheet.state);
      const rp = relPrefix(wbRoot);
      el.setAttributeNS(wbRoot.lookupNamespaceURI(rp) ?? NS_REL_DOC, `${rp}:id`, relId);
      newSheetEls.push(el);
      writeSheet(pkg, sheet, path, sst);
      sheet.src = { path, relId, sheetId: String(maxSheetId), rels: [], loaded: {}, links: new Map() };
      changed = true;
    }
  }
  for (const el of original) {
    const id = relAttr(el)?.value;
    if (id && kept.has(id)) continue;
    const rel = src.workbookRels.find((r) => r.id === id);
    if (rel) {
      pkg.remove(rel.target);
      pkg.remove(relsPathFor(rel.target));
      pkg.setOverride(rel.target, null);
    }
    pkg.removeRels(src.workbookPath, (_t, rid) => rid === id);
    changed = true;
  }
  const orderChanged = original.length !== newSheetEls.length || original.some((el, i) => el !== newSheetEls[i]);
  if (orderChanged) {
    for (const el of original) sheetsEl.removeChild(el);
    for (const el of newSheetEls) sheetsEl.appendChild(el);
    changed = true;
  }

  const names = namesToWrite(wb);
  if (nameKey(names) !== originalNames(wbRoot)) { writeDefinedNames(wbRoot, names); changed = true; }

  if (!wb.src || stylesChanged(wb.styles)) {
    const xml = serializeStyles(wb.styles);
    if (src.stylesPath) pkg.setText(src.stylesPath, xml);
    else {
      const path = 'xl/styles.xml';
      pkg.setText(path, xml);
      pkg.addRel(src.workbookPath, 'styles', path);
      pkg.setOverride(path, CT.styles);
      src.stylesPath = path;
    }
  }

  writeSst(pkg, src, sst);

  if (changed) {
    pkg.touch(src.workbookPath);
    // Cell values may have changed under formulas: drop the stale calc chain and ask Excel to recalculate.
    const chain = src.workbookRels.find((r) => r.type.endsWith('/calcChain'));
    if (chain) {
      pkg.remove(chain.target);
      pkg.setOverride(chain.target, null);
      pkg.removeRels(src.workbookPath, (t) => t === 'calcChain');
    }
    let calc = firstChild(wbRoot, 'calcPr');
    if (!calc) { calc = createLike(wbRoot, 'calcPr'); insertInOrder(wbRoot, calc, WORKBOOK_ORDER); }
    calc.setAttribute('fullCalcOnLoad', '1');

    const views = firstChild(wbRoot, 'bookViews');
    const view = views ? firstChild(views, 'workbookView') : null;
    if (view) {
      const n = wb.sheets.length;
      const active = parseInt(view.getAttribute('activeTab') ?? '0', 10);
      if (!wb.src) view.setAttribute('activeTab', String(wb.activeSheet));
      else if (active >= n) view.setAttribute('activeTab', String(n - 1));
      const first = parseInt(view.getAttribute('firstSheet') ?? '0', 10);
      if (first >= n) view.removeAttribute('firstSheet');
    }
  }

  const bytes = pkg.zip();
  // Saving makes this output the new baseline for the next save.
  wb.src = { ...src, files: pkg.files, workbookDoc: parseXml(strFromU8(pkg.files.get(src.workbookPath)!)), sst: [...src.sst, ...sst.added] };
  wb.src.workbookRels = parseRels(pkg.files, src.workbookPath);
  for (const s of wb.sheets) {
    if (s.kind === 'worksheet' && !s.pristine) rebaseSheet(pkg, s);
    s.pristine = true;
  }
  wb.styles.loaded = {
    fonts: wb.styles.fonts.length, fills: wb.styles.fills.length, borders: wb.styles.borders.length, xfs: wb.styles.xfs.length,
    numFmtIds: new Set(wb.styles.numFmts.keys()),
  };
  if (src.stylesPath) wb.styles.doc = parseXml(strFromU8(pkg.files.get(src.stylesPath)!));
  return bytes;
}

/** After a save, the written sheet becomes the source the next save patches. */
function rebaseSheet(pkg: Pkg, sheet: Sheet): void {
  const src = sheet.src!;
  const xml = strFromU8(pkg.files.get(src.path)!);
  const m = /<(\w+:)?sheetData\b[^>]*?(\/>|>[\s\S]*?<\/(\w+:)?sheetData>)/.exec(xml);
  if (!m) throw new Error(`Saved sheet ${sheet.name} has no <sheetData>`);
  const p = m[1] ?? '';
  src.doc = parseXml(xml.slice(0, m.index) + `<${p}sheetData/>` + xml.slice(m.index + m[0].length));
  src.rels = parseRels(pkg.files, src.path);
  src.links = new Map([...cellLinks(sheet)].map(([ref, target]) => [ref, { target, attrs: {} }]));
  src.loaded = snapshotSheet(sheet);
}
