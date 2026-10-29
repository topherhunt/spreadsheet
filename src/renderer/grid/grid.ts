import { colToName, MAX_COLS, MAX_ROWS, Range } from '../../core/addr';
import { displayFormula } from '../../core/formula/xlfn';
import { Cell, getCell, isError, Sheet, sheetExtent, Workbook } from '../../core/model';
import { displayCell } from '../../core/numfmt';
import { cssFont, escapeHtml, runsHtml, StyleCache, XfInfo } from './cellstyle';
import { Axis, colAxis, rowAxis } from './layout';
import { expandMerges, mergeAt, Pos, Selection } from './selection';

const HEADER_H = 22;
const FILTER_BTN = 18;

export type HeaderArea = 'cell' | 'row' | 'col' | 'corner';

export interface GridEvents {
  selectionChanged(): void;
  startEdit(initial?: string): void;
  /** A mousedown while the editor is open and not pointing at references: commit first. */
  commitEdit(): void;
  /** Whether a click on a cell now inserts a reference into the formula being edited. */
  canPoint(): boolean;
  scrolled(): void;
  /** Formula editing: the user clicked or dragged over cells to insert a reference. */
  pointRef(range: Range, first: boolean): void;
  filterButton(col: number, at: { x: number; y: number }): void;
  contextMenu(area: HeaderArea, at: { x: number; y: number }): void;
  fill(src: Range, target: Range): void;
  resizeCols(cols: number[], px: number | 'auto'): void;
  resizeRows(rows: number[], px: number | 'auto'): void;
  openLink(url: string): void;
}

interface PaneCtx {
  el: HTMLElement;
  content: HTMLElement;
  lines: HTMLElement;
  cells: HTMLElement;
  overlay: HTMLElement;
  pool: Map<string, { el: HTMLElement; sig: string; gen: number }>;
}

type Drag =
  | { kind: 'select'; additive: boolean }
  | { kind: 'rows'; from: number }
  | { kind: 'cols'; from: number }
  | { kind: 'point'; start: Pos }
  | { kind: 'fill'; src: Range; target: Range | null }
  | { kind: 'resize-col'; col: number; start: number; startSize: number; px: number }
  | { kind: 'resize-row'; row: number; start: number; startSize: number; px: number };

function h<K extends keyof HTMLElementTagNameMap>(tag: K, cls: string, parent?: HTMLElement): HTMLElementTagNameMap[K] {
  const el = document.createElement(tag);
  el.className = cls;
  parent?.appendChild(el);
  return el;
}

export class Grid {
  readonly el: HTMLElement;
  readonly sel: Selection;
  zoom = 1;
  showFormulas = false;
  editing = false;
  refBoxes: { range: Range; color: string }[] = [];
  marquee: Range | null = null;

  private wb!: Workbook;
  private si = 0;
  private rows!: Axis;
  private cols!: Axis;
  private rowCount = 1000;
  private colCount = 26;
  private styles!: StyleCache;
  private scroller: HTMLElement;
  private spacer: HTMLElement;
  private panes: PaneCtx[] = [];
  private colHdr: { el: HTMLElement; content: HTMLElement }[] = [];
  private rowHdr: { el: HTMLElement; content: HTMLElement }[] = [];
  private corner: HTMLElement;
  private freezeV: HTMLElement;
  private freezeH: HTMLElement;
  private guide: HTMLElement | null = null;
  private gen = 0;
  private drag: Drag | null = null;
  private autoScroll = 0;
  private lastPointer = { x: 0, y: 0 };
  private frozenRows = 0;
  private frozenCols = 0;
  private hw = 50;

  constructor(private events: GridEvents) {
    this.el = h('div', 'grid flex-1 qa-grid');
    this.el.tabIndex = 0;
    for (let i = 0; i < 4; i++) {
      const el = h('div', 'pane', this.el);
      const content = h('div', 'pane-content', el);
      this.panes.push({ el, content, lines: h('div', '', content), cells: h('div', '', content), overlay: h('div', '', content), pool: new Map() });
    }
    for (let i = 0; i < 2; i++) {
      const el = h('div', 'hdr-layer qa-col-headers', this.el);
      this.colHdr.push({ el, content: h('div', 'pane-content', el) });
    }
    for (let i = 0; i < 2; i++) {
      const el = h('div', 'hdr-layer qa-row-headers', this.el);
      this.rowHdr.push({ el, content: h('div', 'pane-content', el) });
    }
    this.corner = h('div', 'corner qa-corner', this.el);
    this.freezeV = h('div', 'freeze-line-v', this.el);
    this.freezeH = h('div', 'freeze-line-h', this.el);
    this.scroller = h('div', 'grid-scroll', this.el);
    this.spacer = h('div', 'grid-spacer', this.scroller);
    this.sel = new Selection(() => this.sheet);

    this.scroller.addEventListener('scroll', () => { this.growIfNeeded(); this.render(); this.events.scrolled(); });
    this.scroller.addEventListener('mousedown', (e) => this.onMouseDown(e));
    this.scroller.addEventListener('dblclick', (e) => this.onDoubleClick(e));
    this.scroller.addEventListener('mousemove', (e) => { if (!this.drag) this.updateCursor(e); });
    this.scroller.addEventListener('contextmenu', (e) => this.onContextMenu(e));
    new ResizeObserver(() => this.render()).observe(this.el);
  }

  get sheet(): Sheet { return this.wb.sheets[this.si]; }
  get sheetIndex(): number { return this.si; }

  setSheet(wb: Workbook, si: number): void {
    const changed = this.wb !== wb || this.si !== si;
    this.wb = wb;
    this.si = si;
    if (changed) {
      for (const p of this.panes) { p.cells.textContent = ''; p.pool.clear(); }
      const ext = sheetExtent(this.sheet);
      this.rowCount = Math.min(MAX_ROWS, Math.max(1000, ext.rows + 200));
      this.colCount = Math.min(MAX_COLS, Math.max(26, ext.cols + 6));
      this.sel.set({ row: 0, col: 0 });
      this.scroller.scrollTo({ left: 0, top: 0, behavior: 'instant' });
    }
    this.relayout();
  }

  setZoom(z: number): void {
    this.zoom = Math.min(4, Math.max(0.25, Math.round(z * 100) / 100));
    this.relayout();
  }

  /** Recomputes sizes and styles (after structural, size, filter or style changes) and redraws. */
  relayout(): void {
    const ext = sheetExtent(this.sheet);
    this.rowCount = Math.min(MAX_ROWS, Math.max(this.rowCount, ext.rows + 50));
    this.colCount = Math.min(MAX_COLS, Math.max(this.colCount, ext.cols + 3));
    this.rows = rowAxis(this.sheet, this.rowCount, this.zoom);
    this.cols = colAxis(this.sheet, this.colCount, this.zoom);
    this.styles = new StyleCache(this.wb, this.zoom);
    this.frozenRows = Math.min(this.sheet.freeze.rows, this.rowCount);
    this.frozenCols = Math.min(this.sheet.freeze.cols, this.colCount);
    this.hw = Math.max(46, String(this.rowCount).length * 8 + 14);
    for (const p of this.panes) { p.cells.textContent = ''; p.pool.clear(); }
    this.render();
  }

  // -------------------------------------------------------------------------
  // Geometry

  private get hh(): number { return Math.round(HEADER_H * Math.max(1, this.zoom * 0.9)); }
  private get fw(): number { return this.cols.offset(this.frozenCols); }
  private get fh(): number { return this.rows.offset(this.frozenRows); }

  /** Sheet-space x of a point in the grid element, and whether it's in the frozen columns. */
  private sheetX(x: number): number { const xx = x - this.hw; return xx < this.fw ? xx : xx + this.scroller.scrollLeft; }
  private sheetY(y: number): number { const yy = y - this.hh; return yy < this.fh ? yy : yy + this.scroller.scrollTop; }

  private local(e: { clientX: number; clientY: number }): { x: number; y: number } {
    const r = this.el.getBoundingClientRect();
    return { x: e.clientX - r.left, y: e.clientY - r.top };
  }

  cellAt(x: number, y: number): Pos {
    return { row: this.rows.indexAt(this.sheetY(Math.max(this.hh, y))), col: this.cols.indexAt(this.sheetX(Math.max(this.hw, x))) };
  }

  /** Where a cell (or the merged block it's in) is drawn, in client coordinates; null when scrolled out of view. */
  cellRect(row: number, col: number): DOMRect {
    const m = mergeAt(this.sheet, row, col);
    const r = m ?? { r1: row, c1: col, r2: row, c2: col };
    const box = this.el.getBoundingClientRect();
    const sx = (c: number) => this.hw + (c < this.frozenCols ? this.cols.offset(c) : this.cols.offset(c) - this.scroller.scrollLeft);
    const sy = (rr: number) => this.hh + (rr < this.frozenRows ? this.rows.offset(rr) : this.rows.offset(rr) - this.scroller.scrollTop);
    const x = sx(r.c1), y = sy(r.r1);
    const w = this.cols.offset(r.c2 + 1) - this.cols.offset(r.c1);
    const hgt = this.rows.offset(r.r2 + 1) - this.rows.offset(r.r1);
    return new DOMRect(box.left + x, box.top + y, w, hgt);
  }

  /** Scrolls the least distance that brings the cell into view. */
  scrollToCell(row: number, col: number): void {
    const vw = this.scroller.clientWidth - this.hw - this.fw;
    const vh = this.scroller.clientHeight - this.hh - this.fh;
    let { scrollLeft: left, scrollTop: top } = this.scroller;
    if (col >= this.frozenCols) {
      const x0 = this.cols.offset(col) - this.fw, x1 = x0 + this.cols.size(col);
      if (x0 < left) left = x0; else if (x1 > left + vw) left = Math.min(x0, x1 - vw);
    }
    if (row >= this.frozenRows) {
      const y0 = this.rows.offset(row) - this.fh, y1 = y0 + this.rows.size(row);
      if (y0 < top) top = y0; else if (y1 > top + vh) top = Math.min(y0, y1 - vh);
    }
    this.ensureCapacity(row, col);
    if (left !== this.scroller.scrollLeft || top !== this.scroller.scrollTop) this.scroller.scrollTo({ left, top, behavior: 'instant' });
  }

  private ensureCapacity(row: number, col: number): void {
    if (row < this.rowCount - 20 && col < this.colCount - 2) return;
    this.rowCount = Math.min(MAX_ROWS, Math.max(this.rowCount, row + 500));
    this.colCount = Math.min(MAX_COLS, Math.max(this.colCount, col + 10));
    this.relayout();
  }

  private growIfNeeded(): void {
    const s = this.scroller;
    let grow = false;
    if (s.scrollTop + s.clientHeight * 3 > s.scrollHeight && this.rowCount < MAX_ROWS) { this.rowCount = Math.min(MAX_ROWS, this.rowCount + 1000); grow = true; }
    if (s.scrollLeft + s.clientWidth * 2 > s.scrollWidth && this.colCount < MAX_COLS) { this.colCount = Math.min(MAX_COLS, this.colCount + 10); grow = true; }
    if (grow) {
      this.rows = rowAxis(this.sheet, this.rowCount, this.zoom);
      this.cols = colAxis(this.sheet, this.colCount, this.zoom);
    }
  }

  /** The last row/col that navigation can reach. */
  limits(): Pos { return { row: this.rowCount - 1, col: this.colCount - 1 }; }

  visibleRows(): number { return Math.max(1, Math.floor((this.scroller.clientHeight - this.hh - this.fh) / this.rows.def)); }

  // -------------------------------------------------------------------------
  // Rendering

  render(): void {
    if (!this.wb) return;
    const W = this.el.clientWidth, H = this.el.clientHeight;
    const hw = this.hw, hh = this.hh, fw = this.fw, fh = this.fh;
    const sl = this.scroller.scrollLeft, st = this.scroller.scrollTop;
    this.spacer.style.width = `${hw + this.cols.total()}px`;
    this.spacer.style.height = `${hh + this.rows.total()}px`;
    this.gen++;

    const place = (el: HTMLElement, l: number, t: number, w: number, hgt: number) => {
      el.style.left = `${l}px`; el.style.top = `${t}px`; el.style.width = `${Math.max(0, w)}px`; el.style.height = `${Math.max(0, hgt)}px`;
    };
    // Panes: 0 = frozen corner, 1 = frozen rows, 2 = frozen cols, 3 = main.
    const specs = [
      { l: hw, t: hh, w: fw, h: fh, dx: 0, dy: 0 },
      { l: hw + fw, t: hh, w: W - hw - fw, h: fh, dx: fw + sl, dy: 0 },
      { l: hw, t: hh + fh, w: fw, h: H - hh - fh, dx: 0, dy: fh + st },
      { l: hw + fw, t: hh + fh, w: W - hw - fw, h: H - hh - fh, dx: fw + sl, dy: fh + st },
    ];
    specs.forEach((s, i) => {
      const p = this.panes[i];
      place(p.el, s.l, s.t, s.w, s.h);
      p.content.style.transform = `translate(${-s.dx}px, ${-s.dy}px)`;
      if (s.w <= 0 || s.h <= 0) { p.lines.textContent = ''; p.overlay.textContent = ''; this.sweep(p); return; }
      const r0 = i < 2 ? 0 : this.rows.indexAt(s.dy), r1 = i < 2 ? this.frozenRows - 1 : this.rows.indexAt(s.dy + s.h);
      const c0 = i % 2 === 0 ? 0 : this.cols.indexAt(s.dx), c1 = i % 2 === 0 ? this.frozenCols - 1 : this.cols.indexAt(s.dx + s.w);
      this.renderPane(p, r0, r1, c0, c1);
    });

    // Headers
    place(this.corner, 0, 0, hw, hh);
    place(this.colHdr[0].el, hw, 0, fw, hh);
    place(this.colHdr[1].el, hw + fw, 0, W - hw - fw, hh);
    this.colHdr[1].content.style.transform = `translateX(${-(fw + sl)}px)`;
    place(this.rowHdr[0].el, 0, hh, hw, fh);
    place(this.rowHdr[1].el, 0, hh + fh, hw, H - hh - fh);
    this.rowHdr[1].content.style.transform = `translateY(${-(fh + st)}px)`;
    this.renderColHeaders(this.colHdr[0].content, 0, this.frozenCols - 1);
    this.renderColHeaders(this.colHdr[1].content, this.cols.indexAt(fw + sl), this.cols.indexAt(fw + sl + W - hw - fw));
    this.renderRowHeaders(this.rowHdr[0].content, 0, this.frozenRows - 1);
    this.renderRowHeaders(this.rowHdr[1].content, this.rows.indexAt(fh + st), this.rows.indexAt(fh + st + H - hh - fh));
    const all = this.sel.isWholeRows() && this.sel.isWholeCols();
    this.corner.classList.toggle('hdr', false);
    this.corner.style.background = all ? 'var(--sel-border)' : '';

    this.freezeV.style.display = this.frozenCols ? '' : 'none';
    this.freezeV.style.left = `${hw + fw - 1}px`;
    this.freezeH.style.display = this.frozenRows ? '' : 'none';
    this.freezeH.style.top = `${hh + fh - 1}px`;
  }

  private sweep(p: PaneCtx): void {
    for (const [k, v] of p.pool) if (v.gen !== this.gen) { v.el.remove(); p.pool.delete(k); }
  }

  private renderPane(p: PaneCtx, r0: number, r1: number, c0: number, c1: number): void {
    const sheet = this.sheet;
    const rows = this.rows, cols = this.cols;
    const x0 = cols.offset(c0), x1 = cols.offset(c1 + 1), y0 = rows.offset(r0), y1 = rows.offset(r1 + 1);

    // Gridlines
    if (sheet.showGridLines) {
      let lines = '';
      for (let r = r0; r <= r1 + 1; r++) if (r === r1 + 1 || rows.size(r)) lines += `<div class="gl-h" style="top:${rows.offset(r) - 1}px;left:${x0}px;width:${x1 - x0}px"></div>`;
      for (let c = c0; c <= c1 + 1; c++) if (c === c1 + 1 || cols.size(c)) lines += `<div class="gl-v" style="left:${cols.offset(c) - 1}px;top:${y0}px;height:${y1 - y0}px"></div>`;
      p.lines.innerHTML = lines;
    } else p.lines.textContent = '';

    // Merged blocks overlapping the view are drawn whole, from their top-left cell.
    const covered = new Set<string>();
    const merges = sheet.merges.filter((m) => m.r2 >= r0 && m.r1 <= r1 && m.c2 >= c0 && m.c1 <= c1);
    for (const m of merges) {
      for (let r = Math.max(m.r1, r0); r <= Math.min(m.r2, r1); r++) for (let c = Math.max(m.c1, c0); c <= Math.min(m.c2, c1); c++) covered.add(`${r}:${c}`);
      this.renderCell(p, m.r1, m.c1, getCell(sheet, m.r1, m.c1), m, c1);
    }

    // Text spills right from cells left of the view, so start a few columns early.
    const spillFrom = Math.max(0, c0 - 8);
    for (let r = r0; r <= r1; r++) {
      if (!rows.size(r)) continue;
      const row = sheet.rows[r];
      if (!row) continue;
      const end = Math.min(c1, row.cells.length - 1);
      for (let c = spillFrom; c <= end; c++) {
        const cell = row.cells[c];
        if (!cell || !cols.size(c) || covered.has(`${r}:${c}`)) continue;
        if (c < c0 && (cell.v === null || cell.v === '')) continue;
        this.renderCell(p, r, c, cell, undefined, c1);
      }
    }
    this.sweep(p);
    p.overlay.innerHTML = this.overlayHtml(r0, r1, c0, c1);
  }

  private cellSig(cell: Cell): string {
    return `${cell.s ?? 0}|${typeof cell.v === 'object' && cell.v ? (cell.v as any).error : cell.v}|${this.showFormulas ? cell.f ?? '' : ''}|${cell.rich ? cell.rich.length : ''}|${cell.link ?? ''}`;
  }

  private renderCell(p: PaneCtx, r: number, c: number, cell: Cell | undefined, merge: Range | undefined, lastCol: number): void {
    if (!cell) return;
    const rows = this.rows, cols = this.cols;
    const st = this.styles.get(cell.s);
    const showF = this.showFormulas && cell.f !== undefined;
    const disp = showF ? { text: '=' + displayFormula(cell.f!), align: 'left' as const, color: undefined } : displayCell(this.wb, cell);
    const hasText = disp.text !== '';
    if (!hasText && !st.filled && !st.bordered) return;

    const left = cols.offset(merge ? merge.c1 : c), top = rows.offset(merge ? merge.r1 : r);
    const width = merge ? cols.offset(merge.c2 + 1) - left : cols.size(c);
    const height = merge ? rows.offset(merge.r2 + 1) - top : rows.size(r);
    let align = st.hAlign && st.hAlign !== 'general' ? st.hAlign : disp.align;
    if (align === 'centerContinuous' || align === 'distributed') align = 'center';
    if (align === 'fill' || align === 'justify') align = 'left';

    // Overflow into empty neighbors, as spreadsheets do for text that isn't wrapped.
    let tl = 0, tw = width;
    const spills = hasText && !st.wrap && !merge && (typeof cell.v === 'string' || showF || cell.rich);
    if (spills && align === 'left') {
      for (let n = c + 1; n <= Math.min(lastCol + 1, c + 40) && this.isEmpty(r, n); n++) tw += cols.size(n);
    } else if (spills && align === 'right') {
      for (let n = c - 1; n >= Math.max(0, c - 40) && this.isEmpty(r, n); n--) { tw += cols.size(n); tl -= cols.size(n); }
    }

    const sig = `${this.cellSig(cell)}|${left},${top},${width},${height},${tl},${tw}|${align}|${disp.text}|${disp.color ?? ''}`;
    const key = `${r}:${c}`;
    let entry = p.pool.get(key);
    if (entry && entry.sig === sig) { entry.gen = this.gen; return; }
    if (!entry) {
      const el = document.createElement('div');
      el.dataset.rc = key;
      p.cells.appendChild(el);
      entry = { el, sig: '', gen: this.gen };
      p.pool.set(key, entry);
    }
    entry.sig = sig;
    entry.gen = this.gen;
    const el = entry.el;
    el.className = `cell ${st.cls} ha-${align} va-${st.vAlign === 'center' ? 'center' : st.vAlign === 'top' ? 'top' : 'bottom'}${st.wrap ? ' wrap' : ''}`;
    el.style.cssText = `left:${left}px;top:${top}px;width:${width}px;height:${height}px`;
    if (!hasText) { el.innerHTML = ''; return; }
    let inner: string;
    if (cell.rich && !showF) inner = runsHtml(cell.rich, this.zoom);
    else inner = escapeHtml(disp.text);
    const linkCls = cell.link && !cell.rich ? ' lnk' : '';
    const color = disp.color ? `color:${disp.color};` : '';
    el.innerHTML = `<div class="t" style="left:${tl}px;width:${tw}px;height:${height}px;${color}"><span class="inner${linkCls}">${inner}</span></div>`;
  }

  private isEmpty(r: number, c: number): boolean {
    const cell = getCell(this.sheet, r, c);
    return !cell || cell.v === null || cell.v === '';
  }

  private rangeBox(r: Range): { l: number; t: number; w: number; h: number } {
    const l = this.cols.offset(r.c1), t = this.rows.offset(r.r1);
    return { l, t, w: this.cols.offset(Math.min(r.c2, this.colCount - 1) + 1) - l, h: this.rows.offset(Math.min(r.r2, this.rowCount - 1) + 1) - t };
  }

  private overlayHtml(r0: number, r1: number, c0: number, c1: number): string {
    let out = '';
    const vis = (r: Range) => r.r2 >= r0 && r.r1 <= r1 && r.c2 >= c0 && r.c1 <= c1;
    const box = (r: Range, cls: string, extra = '') => {
      const b = this.rangeBox(r);
      return `<div class="${cls}" style="left:${b.l - 1}px;top:${b.t - 1}px;width:${b.w + 1}px;height:${b.h + 1}px;${extra}"></div>`;
    };
    const f = this.sheet.filter;
    if (f && f.range.r1 >= r0 && f.range.r1 <= r1) {
      out += box({ r1: f.range.r1, r2: f.range.r1, c1: f.range.c1, c2: f.range.c2 }, 'filter-head');
      for (let c = Math.max(c0, f.range.c1); c <= Math.min(c1, f.range.c2); c++) {
        if (!this.cols.size(c)) continue;
        const b = this.filterBtnBox(c);
        out += `<div class="filter-btn qa-filter-btn qa-filter-btn-${c}${f.columns[c] ? ' on' : ''}" style="left:${b.l}px;top:${b.t}px"><i class="bi bi-funnel${f.columns[c] ? '-fill' : ''}"></i></div>`;
      }
    }
    if (this.marquee && vis(this.marquee)) out += box(this.marquee, 'copy-marquee');
    for (const rb of this.refBoxes) if (vis(rb.range)) out += box(rb.range, 'ref-box', `border-color:${rb.color};background:${rb.color}1a`);
    const ranges = this.sel.ranges();
    const single = this.sel.isSingleCell();
    for (const r of ranges) if (!single && vis(r)) out += box(r, 'sel-range qa-selection');
    const a = expandMerges(this.sheet, { r1: this.sel.anchor.row, c1: this.sel.anchor.col, r2: this.sel.anchor.row, c2: this.sel.anchor.col });
    if (vis(a)) out += box(a, 'sel-active qa-active-cell');
    if (this.drag?.kind === 'fill' && this.drag.target) out += box(this.drag.target, 'fill-preview');
    const p = this.sel.primary();
    if (!this.editing && !this.sel.extra.length && vis(p)) {
      const b = this.rangeBox(p);
      out += `<div class="fill-handle qa-fill-handle" style="left:${b.l + b.w - 4}px;top:${b.t + b.h - 4}px"></div>`;
    }
    return out;
  }

  private filterBtnBox(c: number): { l: number; t: number } {
    const f = this.sheet.filter!;
    const size = Math.round(FILTER_BTN * Math.min(1, this.zoom) );
    return { l: this.cols.offset(c + 1) - size - 3, t: this.rows.offset(f.range.r1) + Math.max(0, (this.rows.size(f.range.r1) - size) / 2) };
  }

  private renderColHeaders(el: HTMLElement, c0: number, c1: number): void {
    const p = this.sel.primary();
    const ranges = this.sel.ranges();
    const whole = this.sel.isWholeCols();
    let out = '';
    for (let c = c0; c <= c1; c++) {
      const w = this.cols.size(c);
      if (!w) continue;
      const sel = ranges.some((r) => c >= r.c1 && c <= r.c2);
      const cls = sel ? (whole && c >= p.c1 && c <= p.c2 ? ' all' : ' sel') : '';
      out += `<div class="hdr${cls}" style="left:${this.cols.offset(c)}px;top:0;width:${w}px;height:${this.hh}px">${colToName(c)}</div>`;
    }
    el.innerHTML = out;
  }

  private renderRowHeaders(el: HTMLElement, r0: number, r1: number): void {
    const p = this.sel.primary();
    const ranges = this.sel.ranges();
    const whole = this.sel.isWholeRows();
    let out = '';
    for (let r = r0; r <= r1; r++) {
      const hgt = this.rows.size(r);
      if (!hgt) continue;
      const sel = ranges.some((rg) => r >= rg.r1 && r <= rg.r2);
      const cls = sel ? (whole && r >= p.r1 && r <= p.r2 ? ' all' : ' sel') : '';
      out += `<div class="hdr${cls}" style="left:0;top:${this.rows.offset(r)}px;width:${this.hw}px;height:${hgt}px">${r + 1}</div>`;
    }
    el.innerHTML = out;
  }

  // -------------------------------------------------------------------------
  // Mouse

  private area(x: number, y: number): HeaderArea {
    if (x < this.hw && y < this.hh) return 'corner';
    if (y < this.hh) return 'col';
    if (x < this.hw) return 'row';
    return 'cell';
  }

  /** A column/row border within grabbing distance of x/y, for resizing. */
  private borderAt(area: HeaderArea, x: number, y: number): number | null {
    if (area === 'col') {
      const sx = this.sheetX(x);
      const c = this.cols.indexAt(sx);
      const right = this.cols.offset(c + 1), left = this.cols.offset(c);
      if (Math.abs(sx - right) <= 4) return c;
      if (Math.abs(sx - left) <= 4 && c > 0) return this.cols.prevVisible(c - 1);
    } else if (area === 'row') {
      const sy = this.sheetY(y);
      const r = this.rows.indexAt(sy);
      const bottom = this.rows.offset(r + 1), top = this.rows.offset(r);
      if (Math.abs(sy - bottom) <= 3) return r;
      if (Math.abs(sy - top) <= 3 && r > 0) return this.rows.prevVisible(r - 1);
    }
    return null;
  }

  private onFillHandle(x: number, y: number): boolean {
    if (this.editing || this.sel.extra.length) return false;
    const b = this.rangeBox(this.sel.primary());
    const sx = this.sheetX(x), sy = this.sheetY(y);
    return Math.abs(sx - (b.l + b.w)) <= 5 && Math.abs(sy - (b.t + b.h)) <= 5;
  }

  private filterButtonAt(x: number, y: number): number | null {
    const f = this.sheet.filter;
    if (!f) return null;
    const { row, col } = this.cellAt(x, y);
    if (row !== f.range.r1 || col < f.range.c1 || col > f.range.c2) return null;
    const b = this.filterBtnBox(col);
    const sx = this.sheetX(x), sy = this.sheetY(y);
    return sx >= b.l - 2 && sx <= b.l + FILTER_BTN + 2 && sy >= b.t - 2 && sy <= b.t + FILTER_BTN + 2 ? col : null;
  }

  private overScrollbar(e: MouseEvent): boolean {
    return e.offsetX >= this.scroller.clientWidth || e.offsetY >= this.scroller.clientHeight;
  }

  private updateCursor(e: MouseEvent): void {
    const { x, y } = this.local(e);
    const area = this.area(x, y);
    let cursor = 'cell';
    if (area === 'col') cursor = this.borderAt(area, x, y) !== null ? 'col-resize' : 'default';
    else if (area === 'row') cursor = this.borderAt(area, x, y) !== null ? 'row-resize' : 'default';
    else if (area === 'corner') cursor = 'default';
    else if (this.onFillHandle(x, y)) cursor = 'crosshair';
    else if (this.filterButtonAt(x, y) !== null) cursor = 'pointer';
    else if ((e.metaKey || e.ctrlKey) && getCell(this.sheet, this.cellAt(x, y).row, this.cellAt(x, y).col)?.link) cursor = 'pointer';
    this.scroller.style.cursor = cursor;
  }

  private onMouseDown(e: MouseEvent): void {
    if (this.overScrollbar(e) || e.button === 2) return;
    e.preventDefault();
    const { x, y } = this.local(e);
    const area = this.area(x, y);
    const pos = this.cellAt(x, y);
    this.lastPointer = { x, y };

    if (this.editing && area === 'cell' && this.events.canPoint()) {
      this.drag = { kind: 'point', start: pos };
      this.events.pointRef({ r1: pos.row, c1: pos.col, r2: pos.row, c2: pos.col }, true);
      this.startDrag();
      return;
    }
    if (this.editing) this.events.commitEdit();
    this.el.focus({ preventScroll: true });

    if (area === 'corner') { this.selectAll(); return; }
    if (area === 'col' || area === 'row') {
      const border = this.borderAt(area, x, y);
      if (border !== null) {
        const startSize = area === 'col' ? this.cols.size(border) : this.rows.size(border);
        this.drag = area === 'col'
          ? { kind: 'resize-col', col: border, start: x, startSize, px: startSize }
          : { kind: 'resize-row', row: border, start: y, startSize, px: startSize };
        this.guide = h('div', area === 'col' ? 'resize-guide-v' : 'resize-guide-h', this.el);
        this.moveGuide();
        this.startDrag();
        return;
      }
      if (area === 'col') {
        if (e.shiftKey) this.sel.extendTo({ row: MAX_ROWS - 1, col: pos.col });
        else this.sel.set({ row: 0, col: pos.col }, { row: MAX_ROWS - 1, col: pos.col }, e.metaKey || e.ctrlKey);
        this.drag = { kind: 'cols', from: this.sel.anchor.col };
      } else {
        if (e.shiftKey) this.sel.extendTo({ row: pos.row, col: MAX_COLS - 1 });
        else this.sel.set({ row: pos.row, col: 0 }, { row: pos.row, col: MAX_COLS - 1 }, e.metaKey || e.ctrlKey);
        this.drag = { kind: 'rows', from: this.sel.anchor.row };
      }
      this.changed();
      this.startDrag();
      return;
    }

    const fcol = this.filterButtonAt(x, y);
    if (fcol !== null) {
      const r = this.el.getBoundingClientRect();
      this.events.filterButton(fcol, { x: e.clientX - (e.clientX - r.left - x), y: e.clientY + 12 });
      return;
    }
    if (this.onFillHandle(x, y)) {
      this.drag = { kind: 'fill', src: this.sel.primary(), target: null };
      this.startDrag();
      return;
    }
    const link = getCell(this.sheet, pos.row, pos.col)?.link;
    if (link && (e.metaKey || e.ctrlKey)) { this.events.openLink(link); return; }

    if (e.shiftKey) this.sel.extendTo(pos);
    else this.sel.set(pos, pos, e.metaKey || e.ctrlKey);
    this.drag = { kind: 'select', additive: e.metaKey || e.ctrlKey };
    this.changed();
    this.startDrag();
  }

  private startDrag(): void {
    const move = (e: MouseEvent) => { this.lastPointer = this.local(e); this.onDragMove(); };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      cancelAnimationFrame(this.autoScroll);
      this.autoScroll = 0;
      this.endDrag();
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
    const tick = () => {
      this.edgeScroll();
      this.autoScroll = requestAnimationFrame(tick);
    };
    this.autoScroll = requestAnimationFrame(tick);
  }

  /** While dragging past the grid's edge, scroll toward the pointer at a constant step. */
  private edgeScroll(): void {
    const d = this.drag;
    if (!d || d.kind === 'resize-col' || d.kind === 'resize-row') return;
    const { x, y } = this.lastPointer;
    const W = this.el.clientWidth, H = this.el.clientHeight;
    let dx = 0, dy = 0;
    if (d.kind !== 'rows') { if (x > W - 10) dx = 20; else if (x < this.hw + this.fw && x < this.hw) dx = -20; }
    if (d.kind !== 'cols') { if (y > H - 10) dy = 20; else if (y < this.hh) dy = -20; }
    if (!dx && !dy) return;
    this.scroller.scrollTo({ left: this.scroller.scrollLeft + dx, top: this.scroller.scrollTop + dy, behavior: 'instant' });
    this.onDragMove();
  }

  private onDragMove(): void {
    const d = this.drag;
    if (!d) return;
    const { x, y } = this.lastPointer;
    const pos = this.cellAt(x, y);
    switch (d.kind) {
      case 'select':
        if (pos.row !== this.sel.head.row || pos.col !== this.sel.head.col) { this.sel.extendTo(pos); this.changed(); }
        break;
      case 'cols':
        this.sel.extendTo({ row: MAX_ROWS - 1, col: pos.col }); this.changed();
        break;
      case 'rows':
        this.sel.extendTo({ row: pos.row, col: MAX_COLS - 1 }); this.changed();
        break;
      case 'point':
        this.events.pointRef({ r1: Math.min(d.start.row, pos.row), c1: Math.min(d.start.col, pos.col), r2: Math.max(d.start.row, pos.row), c2: Math.max(d.start.col, pos.col) }, false);
        break;
      case 'fill': {
        const s = d.src;
        let t: Range | null = null;
        const down = pos.row - s.r2, up = s.r1 - pos.row, right = pos.col - s.c2, left = s.c1 - pos.col;
        const best = Math.max(down, up, right, left);
        if (best > 0) {
          if (best === down) t = { ...s, r2: pos.row };
          else if (best === up) t = { ...s, r1: pos.row };
          else if (best === right) t = { ...s, c2: pos.col };
          else t = { ...s, c1: pos.col };
        }
        d.target = t;
        this.render();
        break;
      }
      case 'resize-col':
        d.px = Math.max(4, d.startSize + x - d.start); this.moveGuide();
        break;
      case 'resize-row':
        d.px = Math.max(4, d.startSize + y - d.start); this.moveGuide();
        break;
    }
  }

  private moveGuide(): void {
    const d = this.drag;
    if (!this.guide || !d) return;
    if (d.kind === 'resize-col') {
      const x0 = d.col < this.frozenCols ? this.cols.offset(d.col) : this.cols.offset(d.col) - this.scroller.scrollLeft;
      this.guide.style.left = `${this.hw + x0 + d.px - 1}px`;
    } else if (d.kind === 'resize-row') {
      const y0 = d.row < this.frozenRows ? this.rows.offset(d.row) : this.rows.offset(d.row) - this.scroller.scrollTop;
      this.guide.style.top = `${this.hh + y0 + d.px - 1}px`;
    }
  }

  private endDrag(): void {
    const d = this.drag;
    this.drag = null;
    this.guide?.remove();
    this.guide = null;
    if (!d) return;
    if (d.kind === 'fill') { this.render(); if (d.target) this.events.fill(d.src, d.target); }
    else if (d.kind === 'resize-col' && d.px !== d.startSize) this.events.resizeCols(this.targetsFor('col', d.col), Math.round(d.px / this.zoom));
    else if (d.kind === 'resize-row' && d.px !== d.startSize) this.events.resizeRows(this.targetsFor('row', d.row), Math.round(d.px / this.zoom));
  }

  /** Resizing a selected header resizes every selected one, like Sheets. */
  private targetsFor(axis: 'row' | 'col', i: number): number[] {
    const p = this.sel.primary();
    if (axis === 'col' && this.sel.isWholeCols() && i >= p.c1 && i <= p.c2) return Array.from({ length: p.c2 - p.c1 + 1 }, (_, k) => p.c1 + k);
    if (axis === 'row' && this.sel.isWholeRows() && i >= p.r1 && i <= p.r2) return Array.from({ length: p.r2 - p.r1 + 1 }, (_, k) => p.r1 + k);
    return [i];
  }

  private onDoubleClick(e: MouseEvent): void {
    if (this.overScrollbar(e)) return;
    const { x, y } = this.local(e);
    const area = this.area(x, y);
    const border = this.borderAt(area, x, y);
    if (area === 'col' && border !== null) { this.events.resizeCols(this.targetsFor('col', border), 'auto'); return; }
    if (area === 'row' && border !== null) { this.events.resizeRows(this.targetsFor('row', border), 'auto'); return; }
    if (area !== 'cell' || this.editing || this.filterButtonAt(x, y) !== null) return;
    if (this.onFillHandle(x, y)) return;
    this.events.startEdit();
  }

  private onContextMenu(e: MouseEvent): void {
    e.preventDefault();
    const { x, y } = this.local(e);
    const area = this.area(x, y);
    const pos = this.cellAt(x, y);
    if (area === 'cell' && !this.sel.contains(pos.row, pos.col)) this.sel.set(pos);
    if (area === 'col' && !(this.sel.isWholeCols() && this.sel.contains(0, pos.col))) this.sel.set({ row: 0, col: pos.col }, { row: MAX_ROWS - 1, col: pos.col });
    if (area === 'row' && !(this.sel.isWholeRows() && this.sel.contains(pos.row, 0))) this.sel.set({ row: pos.row, col: 0 }, { row: pos.row, col: MAX_COLS - 1 });
    this.changed();
    this.events.contextMenu(area, { x: e.clientX, y: e.clientY });
  }

  // -------------------------------------------------------------------------
  // Selection and keyboard navigation

  private changed(): void {
    this.render();
    this.events.selectionChanged();
  }

  select(anchor: Pos, head: Pos = anchor, scroll = true): void {
    this.sel.set(anchor, head);
    if (scroll) this.scrollToCell(head.row, head.col);
    this.changed();
  }

  selectAll(): void {
    this.sel.set({ row: 0, col: 0 }, { row: MAX_ROWS - 1, col: MAX_COLS - 1 });
    this.changed();
  }

  /** Arrow-key movement. `jump` goes to the edge of the data (Cmd/Ctrl+arrow). */
  move(dr: number, dc: number, extend: boolean, jump: boolean): void {
    const from = extend ? this.sel.head : this.sel.anchor;
    let { row, col } = from;
    const lim = this.limits();
    if (jump) {
      ({ row, col } = this.jumpFrom(row, col, dr, dc));
    } else {
      // Step out of a merged block from its far edge.
      const m = mergeAt(this.sheet, row, col);
      if (m && !extend) { if (dr > 0) row = m.r2; if (dc > 0) col = m.c2; if (dr < 0) row = m.r1; if (dc < 0) col = m.c1; }
      row += dr; col += dc;
      while (row > 0 && row < lim.row && !this.rows.size(row)) row += dr;
      while (col > 0 && col < lim.col && !this.cols.size(col)) col += dc;
    }
    row = Math.max(0, Math.min(MAX_ROWS - 1, row));
    col = Math.max(0, Math.min(MAX_COLS - 1, col));
    if (row >= this.rowCount || col >= this.colCount) this.ensureCapacity(row, col);
    if (!extend) { const m = mergeAt(this.sheet, row, col); if (m) { row = m.r1; col = m.c1; } }
    if (extend) { this.sel.extendTo({ row, col }); this.scrollToCell(row, col); this.changed(); }
    else this.select({ row, col });
  }

  private jumpFrom(row: number, col: number, dr: number, dc: number): Pos {
    const lim = this.limits();
    const filled = (r: number, c: number) => !this.isEmpty(r, c);
    const inside = (r: number, c: number) => r >= 0 && c >= 0 && r <= lim.row && c <= lim.col;
    let r = row + dr, c = col + dc;
    if (!inside(r, c)) return { row, col };
    if (filled(row, col) && filled(r, c)) {
      while (inside(r + dr, c + dc) && filled(r + dr, c + dc)) { r += dr; c += dc; }
      return { row: r, col: c };
    }
    while (inside(r, c) && !filled(r, c)) {
      if (!inside(r + dr, c + dc)) break;
      r += dr; c += dc;
    }
    if (!filled(r, c) && dr > 0) r = Math.min(lim.row, Math.max(row, sheetExtent(this.sheet).rows - 1));
    if (!filled(r, c) && dc > 0) c = Math.min(lim.col, Math.max(col, sheetExtent(this.sheet).cols - 1));
    return { row: r, col: c };
  }

  /** Moves the active cell within the selection (Enter/Tab with a range selected), or by one cell. */
  advance(dr: number, dc: number): void {
    const p = this.sel.primary();
    if (this.sel.isSingleCell() || this.sel.extra.length) { this.move(dr, dc, false, false); return; }
    let { row, col } = this.sel.anchor;
    if (dc) { col += dc; if (col > p.c2) { col = p.c1; row = row + 1 > p.r2 ? p.r1 : row + 1; } else if (col < p.c1) { col = p.c2; row = row - 1 < p.r1 ? p.r2 : row - 1; } }
    if (dr) { row += dr; if (row > p.r2) { row = p.r1; col = col + 1 > p.c2 ? p.c1 : col + 1; } else if (row < p.r1) { row = p.r2; col = col - 1 < p.c1 ? p.c2 : col - 1; } }
    this.sel.anchor = { row, col };
    this.scrollToCell(row, col);
    this.changed();
  }

  page(dir: number, extend: boolean): void {
    const n = this.visibleRows();
    this.scroller.scrollTo({ top: this.scroller.scrollTop + dir * n * this.rows.def, behavior: 'instant' });
    this.move(dir * n, 0, extend, false);
  }

  /** Width that fits each column's widest displayed value (for double-clicking a header border). */
  fitWidth(col: number): number {
    const ctx = document.createElement('canvas').getContext('2d')!;
    let max = 0;
    const sheet = this.sheet;
    sheet.rows.forEach((row, r) => {
      const cell = row?.cells[col];
      if (!cell || row!.hidden || row!.filtered || mergeAt(sheet, r, col)) return;
      const text = displayCell(this.wb, cell).text;
      if (!text) return;
      const st = this.styles.get(cell.s);
      ctx.font = `${st.italic ? 'italic ' : ''}${st.bold ? '700 ' : ''}${st.fontSizePt}pt ${cssFont(st.fontName)}`;
      const longest = st.wrap ? text.split('\n').reduce((a, b) => (a.length > b.length ? a : b)) : text;
      max = Math.max(max, ctx.measureText(longest).width);
    });
    if (sheet.filter && sheet.filter.range.c1 <= col && col <= sheet.filter.range.c2) max += FILTER_BTN + 4;
    return Math.min(1000, Math.ceil(max + 10));
  }

  fitHeight(row: number): number {
    let max = 0;
    for (const cell of this.sheet.rows[row]?.cells ?? []) if (cell) max = Math.max(max, this.styles.get(cell.s).fontSizePt);
    return Math.max(this.rows.def / this.zoom, Math.round(((max || 10) * 96) / 72 * 1.45));
  }

  isErrorAt(row: number, col: number): boolean { return isError(getCell(this.sheet, row, col)?.v); }

  style(xf: number | undefined): XfInfo { return this.styles.get(xf); }
}
