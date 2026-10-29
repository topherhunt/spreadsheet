import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

// Two XML tools: a DOM for the small structural parts we patch in place (so
// every element we don't understand survives untouched), and a streaming tag
// scanner for the potentially huge parts (sheetData, sharedStrings) where a
// DOM would cost hundreds of MB.

export type XDoc = ReturnType<DOMParser['parseFromString']>;
export type XEl = ReturnType<XDoc['createElement']>;

const ENTITY: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

export function decodeXml(s: string): string {
  if (s.indexOf('&') < 0 && s.indexOf('_x') < 0) return s;
  let out = s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (_m, e: string) => {
    if (e[0] === '#') return String.fromCodePoint(e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
    return ENTITY[e];
  });
  // OOXML escapes control characters as _xHHHH_ inside text.
  if (out.indexOf('_x') >= 0) out = out.replace(/_x([0-9A-Fa-f]{4})_/g, (_m, h: string) => String.fromCharCode(parseInt(h, 16)));
  return out;
}

export function escapeXml(s: string): string {
  let out = s.replace(/[&<>"]/g, (c) => (c === '&' ? '&amp;' : c === '<' ? '&lt;' : c === '>' ? '&gt;' : '&quot;'));
  // XML 1.0 can't carry most control characters; OOXML's escape for them is _xHHHH_.
  // eslint-disable-next-line no-control-regex
  out = out.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, (c) => `_x${c.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')}_`);
  return out;
}

/**
 * Refuses DTDs outright: OOXML never needs one, and they are the vector for
 * entity-expansion bombs and XXE in hostile files. Applies to scanned parts too.
 */
export function assertNoDtd(text: string): void {
  if (/<!DOCTYPE/i.test(text.slice(0, 4096)) || /<!ENTITY/i.test(text.slice(0, 65536))) {
    throw new Error('This file contains an XML DOCTYPE/ENTITY declaration, which spreadsheets never use. Refusing to open it.');
  }
}

export function parseXml(text: string): XDoc {
  assertNoDtd(text);
  const errors: string[] = [];
  const doc = new DOMParser({
    onError: (level: string, msg: string) => { if (level !== 'warning') errors.push(msg); },
  }).parseFromString(text, 'text/xml');
  if (errors.length) throw new Error(`Malformed XML: ${errors[0]}`);
  return doc;
}

export function serializeXml(doc: XDoc): string {
  let s = new XMLSerializer().serializeToString(doc);
  if (!s.startsWith('<?xml')) s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s;
  return s;
}

export function localName(el: { nodeName: string; localName?: string | null }): string {
  return el.localName ?? el.nodeName.replace(/^.*:/, '');
}

export function childElements(el: XEl, name?: string): XEl[] {
  const out: XEl[] = [];
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1 && (name === undefined || localName(n as XEl) === name)) out.push(n as XEl);
  }
  return out;
}

export function firstChild(el: XEl, name: string): XEl | null {
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1 && localName(n as XEl) === name) return n as XEl;
  }
  return null;
}

export function textOf(el: XEl): string {
  return el.textContent ?? '';
}

/** Creates an element in the parent's namespace, keeping the document's prefix convention. */
export function createLike(parent: XEl, name: string): XEl {
  const doc = parent.ownerDocument;
  if (!doc) throw new Error('createLike: element has no owner document');
  const prefix = parent.prefix;
  return doc.createElementNS(parent.namespaceURI, prefix ? `${prefix}:${name}` : name);
}

/**
 * Inserts `child` into `parent` at the position the schema's sequence demands.
 * `order` lists the local names of the parent's children in schema order.
 */
export function insertInOrder(parent: XEl, child: XEl, order: string[]): void {
  const idx = order.indexOf(localName(child));
  if (idx < 0) throw new Error(`insertInOrder: ${localName(child)} is not in the schema order list`);
  for (let n = parent.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1) continue;
    const i = order.indexOf(localName(n as XEl));
    if (i > idx) { parent.insertBefore(child, n); return; }
  }
  parent.appendChild(child);
}

// ---------------------------------------------------------------------------
// Streaming scanner

export interface Tag {
  name: string;          // local name, prefix stripped
  close: boolean;        // </x>
  selfClose: boolean;    // <x/>
  attrs: string;         // raw attribute text
  start: number;         // index of '<'
  end: number;           // index just past '>'
}

const TAG_RE = /<(\/?)([A-Za-z_][\w.:-]*)((?:\s+[^\s=>/]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>|<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[([\s\S]*?)\]\]>/g;

/** Iterates the tags in `xml` between `from` and `to`. Comments and PIs are skipped. */
export function* scanTags(xml: string, from = 0, to = xml.length): Generator<Tag> {
  const re = new RegExp(TAG_RE.source, 'g');
  re.lastIndex = from;
  for (;;) {
    const m = re.exec(xml);
    if (!m || m.index >= to) return;
    if (m[2] === undefined) continue; // comment, PI, CDATA
    yield {
      name: m[2].replace(/^.*:/, ''),
      close: m[1] === '/',
      selfClose: m[4] === '/',
      attrs: m[3],
      start: m.index,
      end: m.index + m[0].length,
    };
  }
}

const ATTR_RE = /([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

export function parseAttrs(raw: string): Record<string, string> {
  const out: Record<string, string> = {};
  if (!raw) return out;
  const re = new RegExp(ATTR_RE.source, 'g');
  for (let m = re.exec(raw); m; m = re.exec(raw)) out[m[1]] = decodeXml(m[2] ?? m[3]);
  return out;
}

export function attrsToXml(attrs: Record<string, string | number | undefined>): string {
  let s = '';
  for (const k of Object.keys(attrs)) {
    const v = attrs[k];
    if (v === undefined) continue;
    s += ` ${k}="${escapeXml(String(v))}"`;
  }
  return s;
}

/** Finds the matching close tag index for an element opened at `openEnd` (just past its '>'). */
export function findClose(xml: string, name: string, openEnd: number): { innerEnd: number; end: number } {
  let depth = 1;
  for (const t of scanTags(xml, openEnd)) {
    if (t.name !== name) continue;
    if (t.close) depth--;
    else if (!t.selfClose) depth++;
    if (depth === 0) return { innerEnd: t.start, end: t.end };
  }
  throw new Error(`Unclosed <${name}> element`);
}
