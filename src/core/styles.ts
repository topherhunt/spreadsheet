import { childElements, firstChild, parseXml, serializeXml, XDoc, XEl, escapeXml, localName } from './xlsx/xml';
import type { RunFont } from './model';

// The workbook's style table. Existing entries are never modified or removed:
// a formatting change derives a new cellXfs entry (appending fonts/fills/etc.
// as needed), so every style a file came with is written back exactly.

export interface ColorSpec { rgb?: string; theme?: number; tint?: number; indexed?: number; auto?: boolean }

export interface FontDef {
  bold: boolean; italic: boolean; underline: string | null; strike: boolean;
  size: number; name: string; color?: ColorSpec; vertAlign?: string; scheme?: string; family?: string;
}

export interface FillDef { pattern: string; fg?: ColorSpec; bg?: ColorSpec; gradient?: boolean }

export interface BorderSide { style: string; color?: ColorSpec }
export interface BorderDef { left?: BorderSide; right?: BorderSide; top?: BorderSide; bottom?: BorderSide }

export interface AlignDef {
  horizontal?: string; vertical?: string; wrapText?: boolean; indent?: number; textRotation?: number; shrinkToFit?: boolean;
}

export interface XfDef { numFmtId: number; fontId: number; fillId: number; borderId: number; align: AlignDef }

export interface StyleSheet {
  numFmts: Map<number, string>;
  fonts: FontDef[];
  fills: FillDef[];
  borders: BorderDef[];
  xfs: XfDef[];
  /** Raw XML for each entry that came from the file; appended entries have none. */
  fontXml: (string | undefined)[];
  fillXml: (string | undefined)[];
  borderXml: (string | undefined)[];
  xfXml: (string | undefined)[];
  themeColors: string[];   // by Excel theme index (lt1, dk1, lt2, dk2, accent1..6, hlink, folHlink)
  themeFonts: { major: string; minor: string };
  indexedColors: string[];
  /** Counts of entries loaded from the file; anything past these was added this session. */
  loaded: { fonts: number; fills: number; borders: number; xfs: number; numFmtIds: Set<number> };
  doc?: XDoc;
}

export const BUILTIN_NUMFMTS: Record<number, string> = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
  5: '"$"#,##0_);\\("$"#,##0\\)', 6: '"$"#,##0_);[Red]\\("$"#,##0\\)', 7: '"$"#,##0.00_);\\("$"#,##0.00\\)', 8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
  9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'm/d/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM',
  20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yyyy h:mm',
  37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
  41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)', 42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)', 44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
  45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@',
};

const DEFAULT_INDEXED = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];

const DEFAULT_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];

// ---------------------------------------------------------------------------
// Parsing

function parseColor(el: XEl | null): ColorSpec | undefined {
  if (!el) return undefined;
  const c: ColorSpec = {};
  const rgb = el.getAttribute('rgb');
  if (rgb) c.rgb = rgb.length === 8 ? rgb.slice(2) : rgb;
  if (el.hasAttribute('theme')) c.theme = parseInt(el.getAttribute('theme')!, 10);
  if (el.hasAttribute('indexed')) c.indexed = parseInt(el.getAttribute('indexed')!, 10);
  if (el.hasAttribute('tint')) c.tint = parseFloat(el.getAttribute('tint')!);
  if (el.getAttribute('auto') === '1' || el.getAttribute('auto') === 'true') c.auto = true;
  return c;
}

function boolEl(el: XEl | null): boolean {
  if (!el) return false;
  const v = el.getAttribute('val');
  return v === null || v === '' || v === '1' || v === 'true';
}

export function parseFontEl(el: XEl): FontDef {
  const u = firstChild(el, 'u');
  const valOf = (n: string) => firstChild(el, n)?.getAttribute('val') ?? undefined;
  return {
    bold: boolEl(firstChild(el, 'b')),
    italic: boolEl(firstChild(el, 'i')),
    strike: boolEl(firstChild(el, 'strike')),
    underline: u ? (u.getAttribute('val') || 'single') : null,
    size: parseFloat(valOf('sz') ?? '11'),
    name: valOf('name') ?? valOf('rFont') ?? 'Calibri',
    color: parseColor(firstChild(el, 'color')),
    vertAlign: valOf('vertAlign'),
    scheme: valOf('scheme'),
    family: valOf('family'),
  };
}

function parseFill(el: XEl): FillDef {
  const p = firstChild(el, 'patternFill');
  if (!p) return { pattern: 'none', gradient: !!firstChild(el, 'gradientFill') };
  return { pattern: p.getAttribute('patternType') || 'none', fg: parseColor(firstChild(p, 'fgColor')), bg: parseColor(firstChild(p, 'bgColor')) };
}

function parseBorder(el: XEl): BorderDef {
  const side = (n: string): BorderSide | undefined => {
    const s = firstChild(el, n);
    if (!s || !s.getAttribute('style') || s.getAttribute('style') === 'none') return undefined;
    return { style: s.getAttribute('style')!, color: parseColor(firstChild(s, 'color')) };
  };
  return { left: side('left') ?? side('start'), right: side('right') ?? side('end'), top: side('top'), bottom: side('bottom') };
}

function parseXf(el: XEl): XfDef {
  const a = firstChild(el, 'alignment');
  const align: AlignDef = {};
  if (a) {
    if (a.getAttribute('horizontal')) align.horizontal = a.getAttribute('horizontal')!;
    if (a.getAttribute('vertical')) align.vertical = a.getAttribute('vertical')!;
    if (a.getAttribute('wrapText') === '1' || a.getAttribute('wrapText') === 'true') align.wrapText = true;
    if (a.getAttribute('shrinkToFit') === '1') align.shrinkToFit = true;
    if (a.getAttribute('indent')) align.indent = parseInt(a.getAttribute('indent')!, 10);
    if (a.getAttribute('textRotation')) align.textRotation = parseInt(a.getAttribute('textRotation')!, 10);
  }
  const n = (k: string) => parseInt(el.getAttribute(k) || '0', 10);
  return { numFmtId: n('numFmtId'), fontId: n('fontId'), fillId: n('fillId'), borderId: n('borderId'), align };
}

function outer(el: XEl): string {
  return el.toString();
}

export function parseTheme(xml: string | undefined): { colors: string[]; fonts: { major: string; minor: string } } {
  if (!xml) return { colors: DEFAULT_THEME.slice(), fonts: { major: 'Calibri Light', minor: 'Calibri' } };
  const doc = parseXml(xml);
  const colors = DEFAULT_THEME.slice();
  const scheme = doc.getElementsByTagName('a:clrScheme')[0];
  if (scheme) {
    const byName: Record<string, string> = {};
    for (const c of childElements(scheme)) {
      const inner = childElements(c)[0];
      if (!inner) continue;
      byName[localName(c)] = inner.getAttribute('lastClr') || inner.getAttribute('val') || '000000';
    }
    const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
    order.forEach((n, i) => { if (byName[n]) colors[i] = byName[n]; });
  }
  const latin = (tag: string) => {
    const f = doc.getElementsByTagName(tag)[0];
    const l = f ? childElements(f, 'latin')[0] : undefined;
    return l?.getAttribute('typeface') || undefined;
  };
  return { colors, fonts: { major: latin('a:majorFont') ?? 'Calibri Light', minor: latin('a:minorFont') ?? 'Calibri' } };
}

export function parseStyles(xml: string, themeXml?: string): StyleSheet {
  const doc = parseXml(xml);
  const root = doc.documentElement as unknown as XEl;
  const theme = parseTheme(themeXml);
  const ss = emptyStyleSheet();
  ss.doc = doc;
  ss.themeColors = theme.colors;
  ss.themeFonts = theme.fonts;
  ss.fonts = []; ss.fills = []; ss.borders = []; ss.xfs = [];

  const numFmts = firstChild(root, 'numFmts');
  if (numFmts) for (const n of childElements(numFmts, 'numFmt')) {
    const id = parseInt(n.getAttribute('numFmtId')!, 10);
    ss.numFmts.set(id, n.getAttribute('formatCode') ?? 'General');
    ss.loaded.numFmtIds.add(id);
  }
  const fonts = firstChild(root, 'fonts');
  if (fonts) for (const f of childElements(fonts, 'font')) { ss.fonts.push(parseFontEl(f)); ss.fontXml.push(outer(f)); }
  const fills = firstChild(root, 'fills');
  if (fills) for (const f of childElements(fills, 'fill')) { ss.fills.push(parseFill(f)); ss.fillXml.push(outer(f)); }
  const borders = firstChild(root, 'borders');
  if (borders) for (const b of childElements(borders, 'border')) { ss.borders.push(parseBorder(b)); ss.borderXml.push(outer(b)); }
  const xfs = firstChild(root, 'cellXfs');
  if (xfs) for (const x of childElements(xfs, 'xf')) { ss.xfs.push(parseXf(x)); ss.xfXml.push(outer(x)); }

  const colors = firstChild(root, 'colors');
  const idx = colors ? firstChild(colors, 'indexedColors') : null;
  if (idx) ss.indexedColors = childElements(idx, 'rgbColor').map((c) => (c.getAttribute('rgb') || 'FF000000').slice(-6));

  if (ss.fonts.length === 0) { ss.fonts.push(defaultFont()); ss.fontXml.push(undefined); }
  if (ss.fills.length === 0) { ss.fills.push({ pattern: 'none' }, { pattern: 'gray125' }); ss.fillXml.push(undefined, undefined); }
  if (ss.borders.length === 0) { ss.borders.push({}); ss.borderXml.push(undefined); }
  if (ss.xfs.length === 0) { ss.xfs.push({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, align: {} }); ss.xfXml.push(undefined); }
  ss.loaded = { fonts: ss.fonts.length, fills: ss.fills.length, borders: ss.borders.length, xfs: ss.xfs.length, numFmtIds: ss.loaded.numFmtIds };
  return ss;
}

function defaultFont(): FontDef {
  return { bold: false, italic: false, underline: null, strike: false, size: 11, name: 'Calibri', color: { theme: 1 }, family: '2', scheme: 'minor' };
}

/** A fresh style table for workbooks that didn't come from an XLSX (CSV, new). */
export function emptyStyleSheet(): StyleSheet {
  return {
    numFmts: new Map(), fonts: [{ ...defaultFont(), name: 'Arial', size: 10, scheme: undefined, color: undefined }],
    fills: [{ pattern: 'none' }, { pattern: 'gray125' }], borders: [{}],
    xfs: [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, align: {} }],
    fontXml: [undefined], fillXml: [undefined, undefined], borderXml: [undefined], xfXml: [undefined],
    themeColors: DEFAULT_THEME.slice(), themeFonts: { major: 'Calibri Light', minor: 'Calibri' },
    indexedColors: DEFAULT_INDEXED.slice(),
    loaded: { fonts: 0, fills: 0, borders: 0, xfs: 0, numFmtIds: new Set() },
  };
}

// ---------------------------------------------------------------------------
// Color resolution

function hexToRgb(h: string): [number, number, number] {
  return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
}

function applyTint(hex: string, tint: number): string {
  // Excel applies tint in HSL space to luminance.
  const [r, g, b] = hexToRgb(hex).map((x) => x / 255);
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0, s = 0, l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const hue2rgb = (p: number, q: number, t: number) => {
    if (t < 0) t += 1; if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  let rr = l, gg = l, bb = l;
  if (s !== 0) {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    rr = hue2rgb(p, q, h + 1 / 3); gg = hue2rgb(p, q, h); bb = hue2rgb(p, q, h - 1 / 3);
  }
  return [rr, gg, bb].map((x) => Math.round(x * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
}

/** Resolves to "#RRGGBB", or undefined for automatic/system colors. */
export function resolveColor(ss: StyleSheet, c: ColorSpec | undefined): string | undefined {
  if (!c || c.auto) return undefined;
  let hex: string | undefined;
  if (c.rgb) hex = c.rgb.toUpperCase();
  else if (c.theme !== undefined) hex = ss.themeColors[c.theme];
  else if (c.indexed !== undefined) {
    if (c.indexed >= 64) return undefined; // system foreground/background
    hex = (ss.indexedColors[c.indexed] ?? DEFAULT_INDEXED[c.indexed]);
  }
  if (!hex) return undefined;
  if (c.tint) hex = applyTint(hex, c.tint);
  return '#' + hex;
}

export function numFmtCode(ss: StyleSheet, id: number): string {
  return ss.numFmts.get(id) ?? BUILTIN_NUMFMTS[id] ?? 'General';
}

export function fontName(ss: StyleSheet, f: FontDef): string {
  if (f.scheme === 'minor') return ss.themeFonts.minor;
  if (f.scheme === 'major') return ss.themeFonts.major;
  return f.name;
}

// ---------------------------------------------------------------------------
// Deriving styles

export interface StylePatch {
  bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean;
  fontName?: string; fontSize?: number;
  /** "#RRGGBB" or null for automatic. */
  color?: string | null;
  /** "#RRGGBB" or null for no fill. */
  fill?: string | null;
  numFmt?: string;
  horizontal?: string | null; vertical?: string | null; wrapText?: boolean;
  /** Border sides to set; null removes the side. */
  border?: Partial<Record<'left' | 'right' | 'top' | 'bottom', BorderSide | null>>;
}

function key(o: unknown): string { return JSON.stringify(o); }

function findOrAdd<T>(list: T[], xml: (string | undefined)[], item: T): number {
  const k = key(item);
  const i = list.findIndex((x) => key(x) === k);
  if (i >= 0) return i;
  list.push(item); xml.push(undefined);
  return list.length - 1;
}

function colorFromHex(hex: string | null): ColorSpec | undefined {
  return hex ? { rgb: hex.replace('#', '').toUpperCase() } : undefined;
}

/** Returns the index of a cellXfs entry equal to `base` with `patch` applied, adding one if needed. */
export function deriveXf(ss: StyleSheet, base: number | undefined, patch: StylePatch): number {
  const src = ss.xfs[base ?? 0];
  if (!src) throw new Error(`Style index ${base} out of range`);
  const xf: XfDef = { ...src, align: { ...src.align } };

  const fontKeys: (keyof StylePatch)[] = ['bold', 'italic', 'underline', 'strike', 'fontName', 'fontSize', 'color'];
  if (fontKeys.some((k) => k in patch)) {
    const f: FontDef = { ...ss.fonts[src.fontId] };
    if (patch.bold !== undefined) f.bold = patch.bold;
    if (patch.italic !== undefined) f.italic = patch.italic;
    if (patch.underline !== undefined) f.underline = patch.underline ? 'single' : null;
    if (patch.strike !== undefined) f.strike = patch.strike;
    if (patch.fontName !== undefined) { f.name = patch.fontName; f.scheme = undefined; }
    if (patch.fontSize !== undefined) f.size = patch.fontSize;
    if ('color' in patch) f.color = colorFromHex(patch.color ?? null);
    xf.fontId = findOrAdd(ss.fonts, ss.fontXml, f);
  }
  if ('fill' in patch) {
    const fill: FillDef = patch.fill ? { pattern: 'solid', fg: colorFromHex(patch.fill), bg: { indexed: 64 } } : { pattern: 'none' };
    xf.fillId = patch.fill ? findOrAdd(ss.fills, ss.fillXml, fill) : 0;
  }
  if (patch.numFmt !== undefined) xf.numFmtId = numFmtId(ss, patch.numFmt);
  if ('horizontal' in patch) { if (patch.horizontal) xf.align.horizontal = patch.horizontal; else delete xf.align.horizontal; }
  if ('vertical' in patch) { if (patch.vertical) xf.align.vertical = patch.vertical; else delete xf.align.vertical; }
  if (patch.wrapText !== undefined) { if (patch.wrapText) xf.align.wrapText = true; else delete xf.align.wrapText; }
  if (patch.border) {
    const b: BorderDef = { ...ss.borders[src.borderId] };
    for (const side of ['left', 'right', 'top', 'bottom'] as const) {
      if (!(side in patch.border)) continue;
      const v = patch.border[side];
      if (v) b[side] = v; else delete b[side];
    }
    xf.borderId = findOrAdd(ss.borders, ss.borderXml, b);
  }
  return findOrAdd(ss.xfs, ss.xfXml, xf);
}

export function numFmtId(ss: StyleSheet, code: string): number {
  for (const [id, c] of Object.entries(BUILTIN_NUMFMTS)) if (c === code) return Number(id);
  for (const [id, c] of ss.numFmts) if (c === code) return id;
  let id = 164;
  while (ss.numFmts.has(id)) id++;
  ss.numFmts.set(id, code);
  return id;
}

/** Effective font of a rich-text run: the cell's font overlaid with the run's own. */
export function runFontFromDef(ss: StyleSheet, f: FontDef): RunFont {
  return {
    bold: f.bold, italic: f.italic, underline: !!f.underline, strike: f.strike,
    size: f.size, name: fontName(ss, f), color: resolveColor(ss, f.color),
  };
}

// ---------------------------------------------------------------------------
// Serialization of appended entries

function colorXml(tag: string, c: ColorSpec | undefined): string {
  if (!c) return '';
  let a = '';
  if (c.auto) a += ' auto="1"';
  if (c.rgb) a += ` rgb="FF${c.rgb}"`;
  if (c.theme !== undefined) a += ` theme="${c.theme}"`;
  if (c.indexed !== undefined) a += ` indexed="${c.indexed}"`;
  if (c.tint) a += ` tint="${c.tint}"`;
  return `<${tag}${a}/>`;
}

export function fontXml(f: FontDef, rich = false): string {
  let s = '';
  if (f.bold) s += '<b/>';
  if (f.italic) s += '<i/>';
  if (f.strike) s += '<strike/>';
  if (f.underline) s += f.underline === 'single' ? '<u/>' : `<u val="${f.underline}"/>`;
  if (f.vertAlign) s += `<vertAlign val="${f.vertAlign}"/>`;
  s += `<sz val="${f.size}"/>`;
  s += colorXml('color', f.color);
  s += `<${rich ? 'rFont' : 'name'} val="${escapeXml(f.name)}"/>`;
  if (f.family) s += `<family val="${f.family}"/>`;
  if (f.scheme) s += `<scheme val="${f.scheme}"/>`;
  return rich ? `<rPr>${s}</rPr>` : `<font>${s}</font>`;
}

function fillXml(f: FillDef): string {
  if (f.pattern === 'none' || !f.fg) return `<fill><patternFill patternType="${f.pattern}"/></fill>`;
  return `<fill><patternFill patternType="${f.pattern}">${colorXml('fgColor', f.fg)}${colorXml('bgColor', f.bg)}</patternFill></fill>`;
}

function borderXml(b: BorderDef): string {
  const side = (n: 'left' | 'right' | 'top' | 'bottom') => {
    const s = b[n];
    return s ? `<${n} style="${s.style}">${colorXml('color', s.color)}</${n}>` : `<${n}/>`;
  };
  return `<border>${side('left')}${side('right')}${side('top')}${side('bottom')}<diagonal/></border>`;
}

function xfXml(x: XfDef): string {
  let a = `numFmtId="${x.numFmtId}" fontId="${x.fontId}" fillId="${x.fillId}" borderId="${x.borderId}" xfId="0"`;
  if (x.numFmtId) a += ' applyNumberFormat="1"';
  if (x.fontId) a += ' applyFont="1"';
  if (x.fillId) a += ' applyFill="1"';
  if (x.borderId) a += ' applyBorder="1"';
  const al = x.align;
  const alAttrs = Object.entries(al).map(([k, v]) => ` ${k}="${v === true ? 1 : v}"`).join('');
  if (alAttrs) return `<xf ${a} applyAlignment="1"><alignment${alAttrs}/></xf>`;
  return `<xf ${a}/>`;
}

const STYLES_ORDER = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst'];

const STYLES_TEMPLATE = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

/** Whether anything was added since load, i.e. whether styles.xml must be rewritten. */
export function stylesChanged(ss: StyleSheet): boolean {
  return ss.fonts.length > ss.loaded.fonts || ss.fills.length > ss.loaded.fills || ss.borders.length > ss.loaded.borders
    || ss.xfs.length > ss.loaded.xfs || [...ss.numFmts.keys()].some((id) => !ss.loaded.numFmtIds.has(id));
}

/** styles.xml with the appended entries added; existing entries are left as they were. */
export function serializeStyles(ss: StyleSheet): string {
  const doc = ss.doc ?? parseXml(STYLES_TEMPLATE);
  const root = doc.documentElement as unknown as XEl;
  const section = (name: string): XEl => {
    let el = firstChild(root, name);
    if (!el) {
      el = doc.createElementNS(root.namespaceURI, root.prefix ? `${root.prefix}:${name}` : name) as unknown as XEl;
      const idx = STYLES_ORDER.indexOf(name);
      let before: XEl | null = null;
      for (const c of childElements(root)) if (STYLES_ORDER.indexOf(localName(c)) > idx) { before = c; break; }
      if (before) root.insertBefore(el, before); else root.appendChild(el);
    }
    return el;
  };
  const append = (sectionName: string, xmls: string[], count: number) => {
    const el = section(sectionName);
    for (const x of xmls) {
      const frag = parseXml(`<r xmlns="${root.namespaceURI}">${x}</r>`).documentElement as unknown as XEl;
      for (const c of childElements(frag)) el.appendChild(doc.importNode(c, true));
    }
    el.setAttribute('count', String(count));
  };
  const newNumFmts = [...ss.numFmts.entries()].filter(([id]) => !ss.loaded.numFmtIds.has(id));
  if (newNumFmts.length) {
    append('numFmts', newNumFmts.map(([id, code]) => `<numFmt numFmtId="${id}" formatCode="${escapeXml(code)}"/>`), ss.numFmts.size);
  }
  append('fonts', ss.fonts.slice(ss.loaded.fonts).map((f) => fontXml(f)), ss.fonts.length);
  append('fills', ss.fills.slice(ss.loaded.fills).map(fillXml), ss.fills.length);
  append('borders', ss.borders.slice(ss.loaded.borders).map(borderXml), ss.borders.length);
  append('cellXfs', ss.xfs.slice(ss.loaded.xfs).map(xfXml), ss.xfs.length);
  return serializeXml(doc);
}
