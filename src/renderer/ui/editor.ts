import { quoteSheetName, Range, rangeToA1, toA1 } from '../../core/addr';
import { tokenize } from '../../core/formula/refs';
import { displayFormula } from '../../core/formula/xlfn';
import { Cell, cellText, Run, RunFont, Workbook } from '../../core/model';
import { cellFormat, formatValue, isDateFormat } from '../../core/numfmt';
import { cssFont, escapeHtml, runStyle, textColor, XfInfo } from '../grid/cellstyle';
import { cssToHex } from './dom';

export const REF_COLORS = ['#1a73e8', '#d93025', '#7627bb', '#188038', '#e8710a', '#c5221f', '#12858a', '#b06000'];

/** The text a cell shows when you start editing it. */
export function editText(wb: Workbook, c: Cell | undefined): string {
  if (!c) return '';
  if (c.f !== undefined) return '=' + displayFormula(c.f);
  if (typeof c.v === 'number') {
    const code = cellFormat(wb, c);
    if (isDateFormat(code)) {
      const day = Math.floor(c.v), time = c.v - day;
      const fmt = day === 0 ? 'h:mm:ss AM/PM' : time ? 'm/d/yyyy h:mm:ss' : 'm/d/yyyy';
      return formatValue(c.v, fmt, wb.date1904).text;
    }
    if (/%/.test(code)) return `${parseFloat((c.v * 100).toPrecision(12))}%`;
    return String(c.v);
  }
  return cellText(c);
}

export interface EditorEvents {
  /** The user finished: `move` is where the selection goes next. */
  commit(move: { dr: number; dc: number } | null): void;
  cancel(): void;
  /** The text changed (for the formula bar mirror and reference highlighting). */
  changed(): void;
  /** Arrow keys while pointing at references in a formula: move the pointed cell. */
  pointKey(dr: number, dc: number, extend: boolean): void;
}

const POINTABLE = /(^=|[=(,+\-*/^&<>;:])\s*$/;

/**
 * The in-cell editor: a contenteditable box over the cell for rich text, with
 * the formula bar as a plain-text mirror. Either one can be where the user is
 * typing (`source`).
 */
export class CellEditor {
  readonly el: HTMLDivElement;
  active = false;
  /** Started by typing over a cell: arrow keys commit and move instead of moving the caret. */
  quick = false;
  /** Something was typed or formatted (an untouched editor commits nothing). */
  touched = false;
  source: 'cell' | 'bar' = 'cell';
  sheet = 0;
  row = 0;
  col = 0;
  /** The reference most recently inserted by pointing, as text offsets, while the caret stays after it. */
  private point: { start: number; end: number } | null = null;
  private savedRange: Range_ | null = null;
  private zoom = 1;

  constructor(private host: HTMLElement, private bar: HTMLInputElement, private events: EditorEvents) {
    this.el = document.createElement('div');
    this.el.className = 'editor qa-editor d-none';
    this.el.contentEditable = 'true';
    this.el.spellcheck = false;
    host.appendChild(this.el);
    this.el.addEventListener('input', () => this.onInput('cell'));
    this.el.addEventListener('keydown', (e) => this.onKey(e));
    this.el.addEventListener('paste', (e) => {
      // Rich HTML from other apps would bring foreign styles along: paste text only.
      e.preventDefault();
      document.execCommand('insertText', false, e.clipboardData?.getData('text/plain') ?? '');
    });
    this.el.addEventListener('drop', (e) => e.preventDefault());
    document.addEventListener('selectionchange', () => {
      const sel = window.getSelection();
      if (this.active && sel?.rangeCount && this.el.contains(sel.anchorNode)) this.savedRange = sel.getRangeAt(0).cloneRange();
    });
    bar.addEventListener('input', () => this.onInput('bar'));
  }

  /**
   * Opens the editor over a cell. `initial` replaces the content (typing over a
   * cell); otherwise the cell's current content is loaded for editing.
   */
  open(o: { wb: Workbook; sheet: number; row: number; col: number; cell: Cell | undefined; style: XfInfo; rect: DOMRect; zoom: number; initial?: string; source: 'cell' | 'bar' }): void {
    this.active = true;
    this.touched = o.initial !== undefined;
    this.quick = o.initial !== undefined;
    this.source = o.source;
    this.sheet = o.sheet; this.row = o.row; this.col = o.col;
    this.zoom = o.zoom;
    this.point = null;
    const st = o.style;
    const s = this.el.style;
    s.fontFamily = cssFont(st.fontName);
    s.fontSize = `${(st.fontSizePt * o.zoom).toFixed(2)}pt`;
    s.fontWeight = st.bold ? '700' : '400';
    s.fontStyle = st.italic ? 'italic' : 'normal';
    s.textDecoration = [st.underline ? 'underline' : '', st.strike ? 'line-through' : ''].filter(Boolean).join(' ') || 'none';
    s.color = st.fg ?? '';
    s.background = st.fill ?? '';
    s.textAlign = st.hAlign === 'center' || st.hAlign === 'right' ? st.hAlign : 'left';
    this.place(o.rect, st.wrap);

    if (o.initial !== undefined) this.el.textContent = o.initial;
    else if (o.cell?.rich && o.cell.f === undefined) this.el.innerHTML = this.runsToHtml(o.cell.rich);
    else this.el.textContent = editText(o.wb, o.cell);
    this.el.classList.remove('d-none');
    this.syncMode();
    if (o.source === 'cell') {
      this.el.focus();
      this.caretToEnd();
      this.bar.value = this.text();
    } else {
      this.bar.value = this.text();
    }
    this.events.changed();
  }

  /** Repositions over the cell (after scrolling). */
  place(rect: DOMRect, wrap: boolean): void {
    const box = this.host.getBoundingClientRect();
    const s = this.el.style;
    s.left = `${rect.left - box.left}px`;
    s.top = `${rect.top - box.top}px`;
    s.minWidth = `${rect.width}px`;
    s.minHeight = `${rect.height}px`;
    s.width = wrap ? `${rect.width}px` : 'max-content';
    s.maxWidth = `${Math.max(rect.width, box.right - rect.left - 8)}px`;
  }

  close(): void {
    this.active = false;
    this.point = null;
    this.el.classList.add('d-none');
    this.el.textContent = '';
    this.savedRange = null;
  }

  isFormula(): boolean { return this.text().startsWith('='); }

  text(): string {
    if (this.source === 'bar') return this.bar.value;
    return this.collect().map((r) => r.text).join('');
  }

  /** The content as runs, with each run's font given only where it differs from the cell's. */
  runs(): Run[] {
    const runs = this.collect();
    const out: Run[] = [];
    for (const r of runs) {
      const prev = out[out.length - 1];
      if (prev && prev.link === r.link && JSON.stringify(prev.font) === JSON.stringify(r.font)) prev.text += r.text;
      else out.push(r);
    }
    return out;
  }

  private collect(): Run[] {
    const base = this.fontOf(this.el);
    const out: Run[] = [];
    const push = (text: string, el: Element) => {
      if (!text) return;
      const f = this.fontOf(el);
      const font: RunFont = {};
      for (const k of ['bold', 'italic', 'underline', 'strike', 'size', 'name'] as const) if (f[k] !== base[k]) (font as any)[k] = f[k];
      const kept = el.closest<HTMLElement>('[data-color]');
      if (f.color !== base.color) font.color = f.color;
      else if (kept && this.el.contains(kept)) font.color = kept.dataset.color;
      const link = el.closest('a')?.getAttribute('href') ?? undefined;
      out.push({ text, font: Object.keys(font).length ? font : undefined, link: link || undefined });
    };
    const walk = (node: Node, first: { v: boolean }) => {
      for (const child of node.childNodes) {
        if (child.nodeType === Node.TEXT_NODE) { push(child.textContent ?? '', child.parentElement!); first.v = false; continue; }
        if (!(child instanceof HTMLElement)) continue;
        if (child.tagName === 'BR') {
          // A lone trailing <br> is how contenteditable keeps an empty last line open.
          if (child.nextSibling || child.parentElement !== this.el || out.length) push('\n', child.parentElement!);
          first.v = false;
          continue;
        }
        const block = child.tagName === 'DIV' || child.tagName === 'P';
        if (block && !first.v) push('\n', node instanceof Element ? node : this.el);
        walk(child, first);
      }
    };
    walk(this.el, { v: true });
    // contenteditable leaves a trailing <br> after a final line break.
    const last = out[out.length - 1];
    if (last && last.text.endsWith('\n') && this.el.lastChild?.nodeName === 'BR') last.text = last.text.slice(0, -1);
    return out.filter((r) => r.text);
  }

  private fontOf(el: Element): Required<Omit<RunFont, 'color'>> & { color?: string } {
    const cs = getComputedStyle(el);
    let underline = false, strike = false;
    for (let x: Element | null = el; x; x = x === this.el ? null : x.parentElement) {
      const d = getComputedStyle(x).textDecorationLine;
      if (d.includes('underline')) underline = true;
      if (d.includes('line-through')) strike = true;
    }
    return {
      bold: parseInt(cs.fontWeight, 10) >= 600,
      italic: cs.fontStyle === 'italic',
      underline, strike,
      size: Math.round((parseFloat(cs.fontSize) * 0.75 / this.zoom) * 2) / 2,
      name: cs.fontFamily.split(',')[0].trim().replace(/^["']|["']$/g, ''),
      color: cssToHex(cs.color),
    };
  }

  private runsToHtml(runs: Run[]): string {
    return runs.map((r) => {
      const style = runStyle(r.font, this.zoom);
      // Colors the theme overrides (black in dark mode) are kept so saving doesn't lose them.
      const kept = r.font?.color && !textColor(r.font.color) ? ` data-color="${r.font.color}"` : '';
      const inner = `<span${style ? ` style="${escapeHtml(style)}"` : ''}${kept}>${escapeHtml(r.text).replace(/\n/g, '<br>')}</span>`;
      return r.link ? `<a href="${escapeHtml(r.link)}">${inner}</a>` : inner;
    }).join('');
  }

  // -------------------------------------------------------------------------
  // Typing

  private onInput(from: 'cell' | 'bar'): void {
    if (!this.active) return;
    this.touched = true;
    this.point = null;
    if (from === 'bar') {
      this.source = 'bar';
      this.el.textContent = this.bar.value;
    } else {
      this.bar.value = this.text();
    }
    this.syncMode();
    this.events.changed();
  }

  private syncMode(): void {
    this.el.classList.toggle('formula', this.text().startsWith('='));
  }

  /** Keys for the formula bar while it's the editing source (wired up by the app). */
  onKey(e: KeyboardEvent): void {
    if (!this.active) return;
    const fromBar = e.target === this.bar;
    if (e.key === 'Enter' && (e.altKey || e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      if (fromBar) this.bar.setRangeText('\n', this.bar.selectionStart!, this.bar.selectionEnd!, 'end');
      else document.execCommand('insertLineBreak');
      return;
    }
    if (e.key === 'Enter') { e.preventDefault(); this.events.commit({ dr: e.shiftKey ? -1 : 1, dc: 0 }); return; }
    if (e.key === 'Tab') { e.preventDefault(); this.events.commit({ dr: 0, dc: e.shiftKey ? -1 : 1 }); return; }
    if (e.key === 'Escape') { e.preventDefault(); this.events.cancel(); return; }
    const arrows: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    const d = arrows[e.key];
    if (d && this.quick && !fromBar) {
      e.preventDefault();
      if (this.isFormula() && (this.point || this.canPoint())) this.events.pointKey(d[0], d[1], e.shiftKey);
      else this.events.commit({ dr: d[0], dc: d[1] });
      return;
    }
    // Any other key ends quick mode's arrow behavior once the user moves the caret by clicking.
    if (d) this.quick = false;
  }

  // -------------------------------------------------------------------------
  // Pointing at references while typing a formula

  private caret(): number {
    if (this.source === 'bar') return this.bar.selectionStart ?? this.bar.value.length;
    const sel = window.getSelection();
    if (!sel?.rangeCount || !this.el.contains(sel.anchorNode)) return this.text().length;
    const r = document.createRange();
    r.setStart(this.el, 0);
    r.setEnd(sel.anchorNode!, sel.anchorOffset);
    return r.toString().length;
  }

  /** Whether a click on a cell now should insert a reference rather than end the edit. */
  canPoint(): boolean {
    if (!this.active || !this.isFormula()) return false;
    const before = this.text().slice(0, this.caret());
    if (this.point && this.caret() === this.point.end) return true;
    return POINTABLE.test(before);
  }

  /** Inserts (or, while still pointing, replaces) a reference at the caret. */
  insertRef(range: Range, gridSheet: number, sheetName: string): void {
    let ref = range.r1 === range.r2 && range.c1 === range.c2 ? toA1(range.r1, range.c1) : rangeToA1(range);
    if (gridSheet !== this.sheet) ref = `${quoteSheetName(sheetName)}!${ref}`;
    const text = this.text();
    const at = this.caret();
    const [start, end] = this.point && this.point.end === at ? [this.point.start, this.point.end] : [at, at];
    const next = text.slice(0, start) + ref + text.slice(end);
    this.setPlain(next, start + ref.length);
    this.point = { start, end: start + ref.length };
    this.touched = true;
    this.events.changed();
  }

  /** Ends a pointing session, so the next click commits instead of replacing the reference. */
  endPoint(): void { this.point = null; }

  private setPlain(text: string, caret: number): void {
    this.el.textContent = text;
    this.bar.value = text;
    this.syncMode();
    if (this.source === 'bar') { this.bar.focus(); this.bar.setSelectionRange(caret, caret); return; }
    this.el.focus();
    const node = this.el.firstChild;
    const sel = window.getSelection()!;
    const r = document.createRange();
    if (node) r.setStart(node, Math.min(caret, node.textContent!.length)); else r.setStart(this.el, 0);
    r.collapse(true);
    sel.removeAllRanges();
    sel.addRange(r);
  }

  private caretToEnd(): void {
    const sel = window.getSelection()!;
    const r = document.createRange();
    r.selectNodeContents(this.el);
    r.collapse(false);
    sel.removeAllRanges();
    sel.addRange(r);
  }

  /** References in the formula, each with its highlight color (same sheet only). */
  refs(sheetNames: string[], gridSheet: number): { range: Range; color: string }[] {
    if (!this.isFormula()) return [];
    const out: { range: Range; color: string }[] = [];
    let i = 0;
    for (const t of tokenize(this.text().slice(1))) {
      if (t.type !== 'ref' || t.error) continue;
      const color = REF_COLORS[i++ % REF_COLORS.length];
      const si = t.sheet === undefined ? this.sheet : sheetNames.findIndex((n) => n.toLowerCase() === t.sheet!.toLowerCase());
      if (si !== gridSheet) continue;
      const b = t.b ?? t.a;
      const r1 = Math.min(t.a.row ?? 0, b.row ?? 0), r2 = Math.max(t.a.row ?? 1048575, b.row ?? 1048575);
      const c1 = Math.min(t.a.col ?? 0, b.col ?? 0), c2 = Math.max(t.a.col ?? 16383, b.col ?? 16383);
      out.push({ range: { r1, c1, r2, c2 }, color });
    }
    return out;
  }

  // -------------------------------------------------------------------------
  // Formatting part of the text

  /** Whether some text inside the cell editor is selected (formatting applies to it, not the whole cell). */
  hasTextSelection(): boolean {
    if (!this.active || this.source !== 'cell' || this.isFormula()) return false;
    const r = this.savedRange;
    return !!r && !r.collapsed && this.el.contains(r.commonAncestorContainer);
  }

  format(kind: 'bold' | 'italic' | 'underline' | 'strike' | 'color' | 'fontName' | 'fontSize' | 'link', value?: string | number): void {
    this.el.focus();
    if (this.savedRange) { const sel = window.getSelection()!; sel.removeAllRanges(); sel.addRange(this.savedRange); }
    document.execCommand('styleWithCSS', false, 'true');
    switch (kind) {
      case 'bold': document.execCommand('bold'); break;
      case 'italic': document.execCommand('italic'); break;
      case 'underline': document.execCommand('underline'); break;
      case 'strike': document.execCommand('strikeThrough'); break;
      case 'color': document.execCommand('foreColor', false, String(value)); break;
      case 'fontName': document.execCommand('fontName', false, String(value)); break;
      case 'link': document.execCommand('createLink', false, String(value)); break;
      case 'fontSize': {
        // execCommand only knows sizes 1-7: mark with 7, then swap in the real size.
        document.execCommand('fontSize', false, '7');
        for (const el of this.el.querySelectorAll<HTMLElement>('[style*="xxx-large"], font[size="7"]')) {
          el.removeAttribute('size');
          el.style.fontSize = `${(Number(value) * this.zoom).toFixed(2)}pt`;
        }
        break;
      }
    }
    this.touched = true;
    this.bar.value = this.text();
    this.events.changed();
  }
}

type Range_ = globalThis.Range;
