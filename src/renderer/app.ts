import Papa from 'papaparse';
import { colToName, MAX_COLS, MAX_ROWS, parseRange, Range, rangeToA1, toA1 } from '../core/addr';
import { newWorkbook } from '../core/csv';
import { Doc } from '../core/doc';
import { dataRegion } from '../core/filter';
import { formatOf, readWorkbook, saveLosses, withExtension, writeWorkbook } from '../core/io';
import { Cell, cloneCell, FileFormat, getCell, Sheet, sheetExtent } from '../core/model';
import { cellFormat, displayCell, isDateFormat } from '../core/numfmt';
import { StylePatch } from '../core/styles';
import type { InitialDoc, MenuCommand } from '../shared/api';
import { dark, escapeHtml } from './grid/cellstyle';
import { Grid, HeaderArea } from './grid/grid';
import { pxToColWidth, pxToRowHeight } from './grid/layout';
import { mergeAt, Pos } from './grid/selection';
import { CellEditor, editText } from './ui/editor';
import { h, icon, menu, MenuEntry, prompt, toast } from './ui/dom';
import { openFilterMenu } from './ui/filtermenu';
import { FindBar } from './ui/find';
import { ToolAction, Toolbar } from './ui/toolbar';

const ZOOMS = [0.5, 0.67, 0.75, 0.8, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3];
const THIN = { style: 'thin' };

/** Copied cells, kept so a paste inside the app brings formulas and formatting along. */
interface Clip { text: string; sheet: number; range: Range; cells: (Cell | undefined)[][]; cut: boolean }

function basename(p: string): string { return p.split(/[\\/]/).pop()!; }

function isTextTarget(t: EventTarget | null): boolean {
  return t instanceof HTMLElement && (t.isContentEditable || t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT');
}

/** A number format with one more or fewer decimal places; General takes its count from the value. */
export function changeDecimals(code: string, value: unknown, delta: number): string {
  if (code === 'General' || code === '@') {
    const have = typeof value === 'number' ? Math.min(10, (String(value).split('.')[1] ?? '').length) : 0;
    const n = Math.max(0, have + delta);
    return n ? `0.${'0'.repeat(n)}` : '0';
  }
  if (isDateFormat(code)) return code;
  return code.split(';').map((sec) => sec.replace(/0(\.0+)?/, (_m, frac: string | undefined) => {
    const n = Math.max(0, (frac ? frac.length - 1 : 0) + delta);
    return n ? `0.${'0'.repeat(n)}` : '0';
  })).join(';');
}

/** Border patches for a toolbar preset over one range; later parts apply on top of earlier ones. */
export function borderParts(r: Range, preset: string): { ranges: Range[]; patch: StylePatch }[] {
  const side = (s: 'left' | 'right' | 'top' | 'bottom', rg: Range) => ({ ranges: [rg], patch: { border: { [s]: THIN } } });
  const top = { ...r, r2: r.r1 }, bottom = { ...r, r1: r.r2 }, left = { ...r, c2: r.c1 }, right = { ...r, c1: r.c2 };
  const outer = [side('top', top), side('bottom', bottom), side('left', left), side('right', right)];
  const inner: { ranges: Range[]; patch: StylePatch }[] = [];
  if (r.r2 > r.r1) inner.push(side('bottom', { ...r, r2: r.r2 - 1 }), side('top', { ...r, r1: r.r1 + 1 }));
  if (r.c2 > r.c1) inner.push(side('right', { ...r, c2: r.c2 - 1 }), side('left', { ...r, c1: r.c1 + 1 }));
  switch (preset) {
    case 'all': return [{ ranges: [r], patch: { border: { left: THIN, right: THIN, top: THIN, bottom: THIN } } }];
    case 'none': return [{ ranges: [r], patch: { border: { left: null, right: null, top: null, bottom: null } } }];
    case 'outer': return outer;
    case 'inner': return inner;
    case 'top': return [outer[0]];
    case 'bottom': return [outer[1]];
    case 'left': return [outer[2]];
    case 'right': return [outer[3]];
  }
  throw new Error(`Unknown border preset ${preset}`);
}

/** The window's controller: owns the open document and wires the grid, editor and chrome to it. */
export class App {
  doc!: Doc;
  path: string | null = null;
  readonly grid: Grid;
  private editor: CellEditor;
  private toolbar: Toolbar;
  private find: FindBar;
  private title: HTMLElement;
  private badge: HTMLElement;
  private dirtyDot: HTMLElement;
  private nameBox: HTMLInputElement;
  private bar: HTMLInputElement;
  private banner: HTMLElement;
  private bannerText: HTMLElement;
  private bannerClosed = false;
  private main: HTMLElement;
  private placeholder: HTMLElement;
  private linkChip: HTMLElement;
  private tabs: HTMLElement;
  private status: HTMLElement;
  private zoomLabel: HTMLElement;
  private clip: Clip | null = null;
  /** Keyboard pointing while typing a formula: where the pointed range starts and its moving end. */
  private pointStart: Pos | null = null;
  private pointCur: Pos | null = null;
  private pointing = false;

  constructor(root: HTMLElement) {
    const titlebar = h('div', 'titlebar d-flex align-center gap-2 bg-chrome', root);
    if (!/Mac/.test(navigator.platform)) titlebar.style.paddingLeft = '12px';
    this.title = h('span', 'doc-title ellipsis qa-doc-title', titlebar);
    this.badge = h('span', 'format-badge qa-format-badge', titlebar);
    this.dirtyDot = h('span', 'dirty-dot d-none qa-dirty', titlebar);
    this.dirtyDot.title = 'Unsaved changes';

    this.toolbar = new Toolbar((a) => this.guard(() => this.tool(a)));
    root.append(this.toolbar.el);

    const fbar = h('div', 'formula-bar d-flex align-center border-b', root);
    this.nameBox = h('input', 'name-box mono small qa-name-box', fbar);
    this.nameBox.spellcheck = false;
    h('span', 'fx', fbar, 'fx');
    this.bar = h('input', 'formula-input px-2 qa-formula-input', fbar);
    this.bar.spellcheck = false;

    this.banner = h('div', 'banner d-flex align-center gap-2 d-none qa-banner', root);
    this.banner.append(icon('exclamation-triangle-fill'));
    this.bannerText = h('span', 'flex-1', this.banner);
    const closeBanner = h('button', 'btn-icon qa-banner-close', this.banner);
    closeBanner.append(icon('x-lg'));
    closeBanner.title = 'Dismiss';
    closeBanner.onclick = () => { this.bannerClosed = true; this.banner.classList.add('d-none'); };

    this.main = h('div', 'd-flex flex-1 relative', root);
    this.grid = new Grid({
      selectionChanged: () => this.selectionChanged(),
      startEdit: (initial) => this.guard(() => this.startEdit(initial, 'cell')),
      commitEdit: () => this.guard(() => this.commit(null)),
      canPoint: () => this.editor.canPoint(),
      scrolled: () => this.placeOverlays(),
      pointRef: (range) => this.pointAt(range, { row: range.r1, col: range.c1 }, { row: range.r2, col: range.c2 }),
      filterButton: (col, at) => this.guard(() => openFilterMenu(this.doc, this.si, col, at)),
      contextMenu: (area, at) => this.contextMenu(area, at),
      fill: (src, target) => this.guard(() => { this.doc.fillSeries(this.si, src, target); this.grid.select({ row: target.r1, col: target.c1 }, { row: target.r2, col: target.c2 }, false); }),
      resizeCols: (cols, px) => this.guard(() => this.doc.setColWidth(this.si, cols, px === 'auto' ? cols.map((c) => pxToColWidth(this.grid.fitWidth(c))) : pxToColWidth(px))),
      resizeRows: (rows, px) => this.guard(() => this.doc.setRowHeight(this.si, rows, px === 'auto' ? rows.map((r) => pxToRowHeight(this.grid.fitHeight(r))) : pxToRowHeight(px))),
      openLink: (url) => this.guard(() => this.openLink(url)),
    });
    this.main.append(this.grid.el);
    this.placeholder = h('div', 'flex-1 d-flex align-center justify-center muted d-none qa-other-sheet', this.main, 'This sheet (a chart or dialog sheet) can’t be shown here. It’s kept as is when you save.');
    this.editor = new CellEditor(this.main, this.bar, {
      commit: (move) => this.guard(() => this.commit(move)),
      cancel: () => this.endEdit(),
      changed: () => this.editorChanged(),
      pointKey: (dr, dc, extend) => this.pointKey(dr, dc, extend),
    });
    this.find = new FindBar({
      doc: () => this.doc,
      sheet: () => this.si,
      active: () => this.grid.sel.anchor,
      select: (hit) => this.grid.select(hit),
      close: () => this.grid.el.focus(),
    });
    this.main.append(this.find.el);
    this.linkChip = h('div', 'link-chip d-flex align-center gap-2 d-none qa-link-chip', this.main);

    const bottom = h('div', 'bottombar d-flex align-center border-t bg-chrome', root);
    const add = h('button', 'btn-icon mx-1 qa-add-sheet', bottom);
    add.append(icon('plus-lg'));
    add.title = 'Add sheet';
    add.onclick = () => this.guard(() => { this.commit(null); this.doc.addSheet(); });
    this.tabs = h('div', 'tabs d-flex flex-1 qa-tabs', bottom);
    this.status = h('div', 'status nowrap qa-status', bottom);
    this.zoomLabel = h('div', 'status nowrap qa-zoom', bottom);

    this.wire();
  }

  get si(): number { return this.grid.sheetIndex; }
  get sheet(): Sheet { return this.doc.wb.sheets[this.si]; }

  /** Runs a user action, showing its error (a bad sheet name, nothing to filter) instead of failing silently. */
  private guard(fn: () => void | Promise<unknown>): void {
    try {
      const p = fn();
      if (p instanceof Promise) p.catch((e) => this.fail(e));
    } catch (e) { this.fail(e); }
  }

  private fail(e: unknown): void {
    console.error(e);
    toast(e instanceof Error ? e.message : String(e), true);
  }

  // -------------------------------------------------------------------------
  // Loading and saving

  /** Shows a file (or a new workbook when the doc has no path). */
  open(d: InitialDoc): void {
    const format: FileFormat = d.path ? formatOf(d.path) : 'xlsx';
    const wb = d.bytes ? readWorkbook(d.bytes, format) : newWorkbook('xlsx');
    if (this.editor.active) this.endEdit();
    this.path = d.path;
    this.doc = new Doc(wb);
    this.doc.onChange((kind) => this.docChanged(kind));
    this.clip = null;
    this.grid.marquee = null;
    this.bannerClosed = false;
    const first = wb.sheets[wb.activeSheet]?.state === 'visible' ? wb.activeSheet : wb.sheets.findIndex((s) => s.state === 'visible');
    wb.activeSheet = Math.max(0, first);
    this.showSheet();
    this.updateBanner();
    this.grid.el.focus();
  }

  private showSheet(): void {
    this.grid.setSheet(this.doc.wb, this.doc.wb.activeSheet);
    const other = this.sheet.kind === 'other';
    this.grid.el.classList.toggle('d-none', other);
    this.placeholder.classList.toggle('d-none', !other);
    this.renderTabs();
    this.updateChrome();
  }

  private updateBanner(): void {
    const wb = this.doc.wb;
    const features = new Set(wb.features);
    for (const s of wb.sheets) for (const f of s.features) features.add(f);
    const parts: string[] = [];
    if (features.size) parts.push(`Kept but not active here: ${[...features].join(', ')}. They’re saved back unchanged.`);
    if (wb.unsupportedFunctions.size) parts.push(`Formulas using ${[...wb.unsupportedFunctions].join(', ')} show their last saved results.`);
    this.bannerText.textContent = parts.join(' ');
    this.banner.classList.toggle('d-none', !parts.length || this.bannerClosed);
  }

  /** Saves, asking where when there's no file yet (or `as`). Resolves to whether the file was written. */
  async save(as: boolean): Promise<boolean> {
    this.commit(null);
    const wb = this.doc.wb;
    let path = this.path;
    let format = wb.format;
    if (!as && path) {
      const losses = saveLosses(wb, format);
      if (losses.length) {
        const kind = format.toUpperCase();
        const choice = await window.api.message({
          type: 'warning',
          message: `A ${kind} file can’t keep ${losses.join(', ')}.`,
          detail: `Save as an Excel workbook (.xlsx) to keep everything, or keep the ${kind} and save only the values on the current sheet.`,
          buttons: ['Save as XLSX…', `Keep ${kind}`, 'Cancel'], defaultId: 0, cancelId: 2,
        });
        if (choice === 2) return false;
        if (choice === 0) { as = true; path = withExtension(path, 'xlsx'); format = 'xlsx'; }
      }
    }
    if (as || !path) {
      const chosen = await window.api.saveDialog(path ?? `Untitled.${format}`, format);
      if (!chosen) return false;
      path = chosen;
      format = formatOf(chosen);
      const losses = saveLosses(wb, format);
      if (losses.length && format !== wb.format) {
        const choice = await window.api.message({
          type: 'warning', message: `A ${format.toUpperCase()} file can’t keep ${losses.join(', ')}.`,
          detail: 'Only the values on the current sheet are saved.', buttons: ['Save', 'Cancel'], defaultId: 0, cancelId: 1,
        });
        if (choice === 1) return false;
      }
    }
    if (wb.features.has('Macros (VBA)') && format !== 'xlsm') {
      await window.api.message({ type: 'error', message: 'This workbook has macros.', detail: 'Save it as a macro-enabled workbook (.xlsm) so the macros aren’t lost.', buttons: ['OK'] });
      return false;
    }
    const prev = wb.format;
    wb.format = format;
    try {
      await window.api.writeFile(path, writeWorkbook(wb, format));
    } catch (e) {
      wb.format = prev;
      throw e;
    }
    this.path = path;
    this.doc.markSaved();
    this.updateChrome();
    return true;
  }

  // -------------------------------------------------------------------------
  // Change notifications

  private docChanged(kind: 'cells' | 'layout' | 'sheets'): void {
    if (kind === 'sheets') this.showSheet();
    else if (kind === 'layout') this.grid.relayout();
    else this.grid.render();
    if (kind !== 'sheets') this.updateChrome();
  }

  private selectionChanged(): void {
    if (this.editor.active) return;
    this.updateChrome();
  }

  /** Everything outside the grid that reflects the document or the active cell. */
  private updateChrome(): void {
    const wb = this.doc.wb;
    const title = this.path ? basename(this.path) : 'Untitled';
    this.title.textContent = title;
    this.badge.textContent = wb.format.toUpperCase();
    this.dirtyDot.classList.toggle('d-none', !this.doc.dirty);
    window.api.setState({ path: this.path, title, dirty: this.doc.dirty });

    const { row, col } = this.grid.sel.anchor;
    const cell = getCell(this.sheet, row, col);
    const st = this.grid.style(cell?.s);
    this.toolbar.update({
      canUndo: this.doc.canUndo(), canRedo: this.doc.canRedo(),
      bold: st.bold, italic: st.italic, underline: st.underline, strike: st.strike,
      fontName: st.fontName, fontSize: st.fontSizePt, color: st.color, fill: st.fill,
      hAlign: st.hAlign, vAlign: st.vAlign, wrap: st.wrap,
      merged: !!mergeAt(this.sheet, row, col), filter: !!this.sheet.filter,
    });
    if (!this.editor.active) {
      this.bar.value = editText(wb, cell);
      if (document.activeElement !== this.nameBox) this.nameBox.value = this.selectionName();
    }
    this.status.textContent = this.stats();
    this.zoomLabel.textContent = `${Math.round(this.grid.zoom * 100)}%`;
    this.placeOverlays();
  }

  private selectionName(): string {
    const p = this.grid.sel.primary();
    if (this.grid.sel.isSingleCell()) return toA1(this.grid.sel.anchor.row, this.grid.sel.anchor.col);
    if (this.grid.sel.isWholeCols(p) && !this.grid.sel.isWholeRows(p)) return `${colToName(p.c1)}:${colToName(p.c2)}`;
    if (this.grid.sel.isWholeRows(p) && !this.grid.sel.isWholeCols(p)) return `${p.r1 + 1}:${p.r2 + 1}`;
    return rangeToA1(this.clamp(p));
  }

  /** Sum, average and count of the selected numbers, like the Sheets status bar. */
  private stats(): string {
    if (this.grid.sel.isSingleCell()) return '';
    let sum = 0, count = 0, filled = 0;
    for (const r of this.grid.sel.ranges().map((x) => this.clamp(x))) {
      for (let ri = r.r1; ri <= r.r2; ri++) {
        const row = this.sheet.rows[ri];
        if (!row || row.hidden || row.filtered) continue;
        for (let c = r.c1; c <= Math.min(r.c2, row.cells.length - 1); c++) {
          const v = row.cells[c]?.v;
          if (v === undefined || v === null || v === '') continue;
          filled++;
          if (typeof v === 'number') { sum += v; count++; }
        }
      }
    }
    if (!count) return filled ? `Count: ${filled}` : '';
    const fmt = (n: number) => parseFloat(n.toPrecision(12)).toLocaleString(undefined, { maximumFractionDigits: 10 });
    return `Sum: ${fmt(sum)}   Average: ${fmt(sum / count)}   Count: ${filled}`;
  }

  /** The link chip under the active cell, and the editor, follow the cell as the grid scrolls. */
  private placeOverlays(): void {
    if (this.editor.active) {
      this.editor.place(this.grid.cellRect(this.editor.row, this.editor.col), this.grid.style(getCell(this.sheet, this.editor.row, this.editor.col)?.s).wrap);
    }
    const { row, col } = this.grid.sel.anchor;
    const link = getCell(this.sheet, row, col)?.link;
    const show = !!link && !this.editor.active && this.grid.sel.isSingleCell() && this.sheet.kind === 'worksheet';
    this.linkChip.classList.toggle('d-none', !show);
    if (!show) return;
    this.linkChip.textContent = '';
    this.linkChip.append(icon('link-45deg'));
    const a = h('a', 'ellipsis qa-link-open', this.linkChip, link);
    a.href = '#';
    a.onclick = (e) => { e.preventDefault(); this.guard(() => this.openLink(link)); };
    const edit = h('button', 'btn-icon qa-link-edit', this.linkChip);
    edit.title = 'Edit link';
    edit.append(icon('pencil'));
    edit.onclick = () => this.guard(() => this.insertLink());
    const remove = h('button', 'btn-icon qa-link-remove', this.linkChip);
    remove.title = 'Remove link';
    remove.append(icon('x-lg'));
    remove.onclick = () => this.guard(() => this.doc.setLink(this.si, row, col, undefined));
    const rect = this.grid.cellRect(row, col);
    const box = this.main.getBoundingClientRect();
    this.linkChip.style.left = `${rect.left - box.left}px`;
    this.linkChip.style.top = `${rect.bottom - box.top + 4}px`;
  }

  // -------------------------------------------------------------------------
  // Sheet tabs

  private renderTabs(): void {
    const wb = this.doc.wb;
    this.tabs.textContent = '';
    wb.sheets.forEach((s, i) => {
      if (s.state !== 'visible') return;
      const tab = h('button', `tab qa-tab${i === wb.activeSheet ? ' active' : ''}`, this.tabs, s.name);
      tab.dataset.index = String(i);
      if (s.tabColor) tab.style.borderBottom = `3px solid ${s.tabColor}`;
      tab.draggable = true;
      tab.onclick = () => this.guard(() => this.switchSheet(i));
      tab.ondblclick = () => this.guard(() => this.renameSheet(i));
      tab.oncontextmenu = (e) => { e.preventDefault(); this.tabMenu(i, { x: e.clientX, y: e.clientY }); };
      tab.ondragstart = (e) => e.dataTransfer!.setData('text/x-sheet', String(i));
      tab.ondragover = (e) => { if (e.dataTransfer!.types.includes('text/x-sheet')) e.preventDefault(); };
      tab.ondrop = (e) => {
        e.preventDefault();
        const from = Number(e.dataTransfer!.getData('text/x-sheet'));
        if (from !== i) this.guard(() => this.doc.moveSheet(from, i));
      };
    });
  }

  private switchSheet(i: number): void {
    this.commit(null);
    this.doc.setActiveSheet(i);
  }

  private async renameSheet(i: number): Promise<void> {
    const name = await prompt('Rename sheet', this.doc.wb.sheets[i].name, { ok: 'Rename', qa: 'qa-rename-sheet' });
    if (name !== null && name !== this.doc.wb.sheets[i].name) this.doc.renameSheet(i, name.trim());
  }

  private tabMenu(i: number, at: { x: number; y: number }): void {
    const wb = this.doc.wb;
    const visible = wb.sheets.filter((s) => s.state === 'visible').length;
    menu([
      { label: 'Rename…', qa: 'qa-tab-rename', run: () => this.guard(() => this.renameSheet(i)) },
      { label: 'Delete', qa: 'qa-tab-delete', disabled: visible < 2, run: () => this.guard(() => this.deleteSheet(i)) },
      'sep',
      { label: 'Move left', qa: 'qa-tab-left', disabled: i === 0, run: () => this.guard(() => this.doc.moveSheet(i, i - 1)) },
      { label: 'Move right', qa: 'qa-tab-right', disabled: i === wb.sheets.length - 1, run: () => this.guard(() => this.doc.moveSheet(i, i + 1)) },
      'sep',
      { label: 'Insert sheet', icon: 'plus-lg', qa: 'qa-tab-insert', run: () => this.guard(() => this.doc.addSheet(i + 1)) },
    ], at);
  }

  private async deleteSheet(i: number): Promise<void> {
    const name = this.doc.wb.sheets[i].name;
    const choice = await window.api.message({ type: 'warning', message: `Delete the sheet “${name}”?`, detail: 'You can undo this.', buttons: ['Delete', 'Cancel'], defaultId: 0, cancelId: 1 });
    if (choice === 0) { this.commit(null); this.doc.deleteSheet(i); }
  }

  // -------------------------------------------------------------------------
  // Editing

  private startEdit(initial: string | undefined, source: 'cell' | 'bar'): void {
    if (this.sheet.kind !== 'worksheet') return;
    const { row, col } = this.grid.sel.anchor;
    this.grid.scrollToCell(row, col);
    const cell = getCell(this.sheet, row, col);
    if (cell?.spill) { toast('This value comes from an array formula in another cell.', true); return; }
    this.grid.editing = true;
    this.linkChip.classList.add('d-none');
    this.editor.open({
      wb: this.doc.wb, sheet: this.si, row, col, cell, style: this.grid.style(cell?.s),
      rect: this.grid.cellRect(row, col), zoom: this.grid.zoom, initial, source,
    });
    this.grid.render();
  }

  /** Writes the editor's content into its cell; `move` then steps the selection (Enter, Tab). */
  private commit(move: { dr: number; dc: number } | null): void {
    if (!this.editor.active) return;
    const e = this.editor;
    const { sheet, row, col } = e;
    const touched = e.touched;
    const plain = e.isFormula() || e.source === 'bar';
    const text = e.text();
    const runs = plain ? [] : e.runs();
    this.endEdit();
    if (touched) {
      if (plain) this.doc.setInput(sheet, row, col, text);
      else this.doc.setRich(sheet, row, col, runs);
    }
    if (move) this.grid.advance(move.dr, move.dc);
  }

  private endEdit(): void {
    this.editor.close();
    this.grid.editing = false;
    this.grid.refBoxes = [];
    this.pointStart = this.pointCur = null;
    this.grid.el.focus({ preventScroll: true });
    this.grid.render();
    this.updateChrome();
  }

  private editorChanged(): void {
    if (!this.pointing) this.pointStart = this.pointCur = null;
    this.grid.refBoxes = this.editor.refs(this.doc.wb.sheets.map((s) => s.name), this.si);
    this.grid.render();
    const { row, col } = this.editor;
    this.nameBox.value = toA1(row, col);
  }

  private pointAt(range: Range, start: Pos, cur: Pos): void {
    this.pointing = true;
    try {
      this.editor.insertRef(range, this.si, this.sheet.name);
    } finally {
      this.pointing = false;
    }
    this.pointStart = start;
    this.pointCur = cur;
  }

  /** Arrow keys while typing a formula after an operator point at cells, as in Sheets. */
  private pointKey(dr: number, dc: number, extend: boolean): void {
    const from = this.pointCur ?? { row: this.editor.row, col: this.editor.col };
    const cur = { row: Math.max(0, Math.min(MAX_ROWS - 1, from.row + dr)), col: Math.max(0, Math.min(MAX_COLS - 1, from.col + dc)) };
    const start = extend && this.pointStart ? this.pointStart : cur;
    this.grid.scrollToCell(cur.row, cur.col);
    this.pointAt({ r1: Math.min(start.row, cur.row), c1: Math.min(start.col, cur.col), r2: Math.max(start.row, cur.row), c2: Math.max(start.col, cur.col) }, start, cur);
  }

  // -------------------------------------------------------------------------
  // Selection helpers

  /** A whole-row/column selection reduced to the used part of the sheet. */
  private clamp(r: Range): Range {
    const ext = sheetExtent(this.sheet);
    const out = { ...r };
    if (r.r2 >= MAX_ROWS - 1) out.r2 = Math.max(r.r1, ext.rows - 1);
    if (r.c2 >= MAX_COLS - 1) out.c2 = Math.max(r.c1, ext.cols - 1);
    return out;
  }

  private ranges(): Range[] { return this.grid.sel.ranges().map((r) => this.clamp(r)); }

  /** The rows an insert/delete acts on: the selected rows, or the active cell's when whole columns are selected. */
  private rowSpan(): [number, number] {
    const p = this.grid.sel.primary();
    return p.r2 >= MAX_ROWS - 1 ? [this.grid.sel.anchor.row, this.grid.sel.anchor.row] : [p.r1, p.r2];
  }

  private colSpan(): [number, number] {
    const p = this.grid.sel.primary();
    return p.c2 >= MAX_COLS - 1 ? [this.grid.sel.anchor.col, this.grid.sel.anchor.col] : [p.c1, p.c2];
  }

  private activeCell(): Cell | undefined { return getCell(this.sheet, this.grid.sel.anchor.row, this.grid.sel.anchor.col); }

  /** Jumps to "A1", "B2:C9" or "Sheet 2!A1" (the name box, internal links). */
  private goTo(ref: string): void {
    const m = /^(?:'((?:[^']|'')+)'|([^!']+))!(.+)$/.exec(ref.trim());
    let addr = ref.trim();
    if (m) {
      const name = (m[1] ?? m[2]).replace(/''/g, "'");
      const i = this.doc.wb.sheets.findIndex((s) => s.name.toLowerCase() === name.toLowerCase());
      if (i < 0) throw new Error(`There’s no sheet named “${name}”.`);
      addr = m[3];
      this.doc.setActiveSheet(i);
    }
    const r = parseRange(addr.replace(/\$/g, ''));
    this.grid.select({ row: r.r1, col: r.c1 }, { row: r.r2, col: r.c2 });
    this.grid.el.focus();
  }

  private openLink(url: string): void {
    if (url.startsWith('#')) this.goTo(url.slice(1));
    else window.api.openExternal(url);
  }

  private async insertLink(): Promise<void> {
    const { row, col } = this.grid.sel.anchor;
    const current = this.activeCell()?.link ?? '';
    const withSel = this.editor.hasTextSelection();
    const url = await prompt('Link', current, { placeholder: 'https://example.com', ok: 'Apply', qa: 'qa-link-prompt' });
    if (url === null) return;
    const clean = url.trim();
    if (withSel && this.editor.active) { if (clean) this.editor.format('link', clean); return; }
    this.commit(null);
    this.doc.setLink(this.si, row, col, clean || undefined);
  }

  // -------------------------------------------------------------------------
  // Toolbar and formatting

  private style(patch: StylePatch, label?: string): void {
    this.commit(null);
    this.doc.setStyle(this.si, this.ranges(), patch, label);
  }

  private tool(a: ToolAction): void {
    const st = this.grid.style(this.activeCell()?.s);
    const inText = this.editor.hasTextSelection();
    switch (a.kind) {
      case 'undo': this.command('undo'); return;
      case 'redo': this.command('redo'); return;
      case 'numFmt': this.style({ numFmt: a.code }, 'Number format'); return;
      case 'decimals': {
        const cell = this.activeCell();
        this.style({ numFmt: changeDecimals(cellFormat(this.doc.wb, cell), cell?.v, a.delta) }, 'Number format');
        return;
      }
      case 'bold': case 'italic': case 'underline': case 'strike':
        if (inText) this.editor.format(a.kind);
        else this.style({ [a.kind]: !st[a.kind] });
        return;
      case 'fontName':
        if (inText) this.editor.format('fontName', a.name); else this.style({ fontName: a.name });
        return;
      case 'fontSize':
        if (inText) this.editor.format('fontSize', a.size); else this.style({ fontSize: a.size });
        return;
      case 'color':
        if (inText && a.color) this.editor.format('color', a.color); else this.style({ color: a.color });
        return;
      case 'fill': this.style({ fill: a.color }); return;
      case 'borders':
        this.commit(null);
        this.doc.setStyles(this.si, this.ranges().flatMap((r) => borderParts(r, a.preset)), 'Borders');
        return;
      case 'merge': {
        this.commit(null);
        const { row, col } = this.grid.sel.anchor;
        const m = mergeAt(this.sheet, row, col);
        if (m) { this.doc.setMerge(this.si, m, false); return; }
        const p = this.clamp(this.grid.sel.primary());
        if (p.r1 === p.r2 && p.c1 === p.c2) throw new Error('Select more than one cell to merge.');
        this.doc.setMerge(this.si, p, true);
        return;
      }
      case 'hAlign': this.style({ horizontal: a.value }); return;
      case 'vAlign': this.style({ vertical: a.value }); return;
      case 'wrap': this.style({ wrapText: !st.wrap }); return;
      case 'link': this.guard(() => this.insertLink()); return;
      case 'clearFormat': this.commit(null); this.doc.clearFormat(this.si, this.ranges()); return;
      case 'filter': this.command('toggle-filter'); return;
      case 'autoTable': this.command('auto-table'); return;
    }
  }

  // -------------------------------------------------------------------------
  // Clipboard

  private copy(e: ClipboardEvent, cut: boolean): void {
    if (isTextTarget(e.target) || !this.doc) return;
    e.preventDefault();
    if (this.grid.sel.extra.length) throw new Error('Copying several ranges at once isn’t supported.');
    const r = this.clamp(this.grid.sel.primary());
    const cells: (Cell | undefined)[][] = [];
    const text: string[][] = [];
    for (let row = r.r1; row <= r.r2; row++) {
      const line: (Cell | undefined)[] = [];
      const tline: string[] = [];
      for (let col = r.c1; col <= r.c2; col++) {
        const c = getCell(this.sheet, row, col);
        line.push(c && cloneCell(c));
        tline.push(displayCell(this.doc.wb, c).text);
      }
      cells.push(line);
      text.push(tline);
    }
    const plain = Papa.unparse(text, { delimiter: '\t', newline: '\n', quotes: false, escapeFormulae: false });
    const html = `<table>${text.map((l) => `<tr>${l.map((t) => `<td>${escapeHtml(t)}</td>`).join('')}</tr>`).join('')}</table>`;
    e.clipboardData!.setData('text/plain', plain);
    e.clipboardData!.setData('text/html', html);
    this.clip = { text: plain, sheet: this.si, range: r, cells, cut };
    this.grid.marquee = r;
    this.grid.render();
  }

  private paste(text: string, valuesOnly: boolean): void {
    this.commit(null);
    const p = this.grid.sel.primary();
    const clip = this.clip;
    let size: { rows: number; cols: number };
    if (clip && clip.text === text.replace(/\r\n/g, '\n').replace(/\n$/, '')) {
      const h = clip.range.r2 - clip.range.r1 + 1, w = clip.range.c2 - clip.range.c1 + 1;
      const sr = p.r2 - p.r1 + 1, sc = p.c2 - p.c1 + 1;
      const tile = !this.grid.sel.isWholeRows() && !this.grid.sel.isWholeCols() && sr % h === 0 && sc % w === 0;
      size = tile ? { rows: sr, cols: sc } : { rows: h, cols: w };
      if (clip.cut) {
        if (clip.sheet === this.si) this.doc.moveCells(this.si, clip.range, p.r1, p.c1);
        else {
          this.doc.pasteCells(this.si, p.r1, p.c1, clip.cells, { row: clip.range.r1, col: clip.range.c1 });
          this.doc.clear(clip.sheet, [clip.range]);
        }
        this.clip = null;
        this.grid.marquee = null;
        size = { rows: h, cols: w };
      } else {
        this.doc.pasteCells(this.si, p.r1, p.c1, clip.cells, { row: clip.range.r1, col: clip.range.c1 }, valuesOnly, size);
      }
    } else {
      const parsed = Papa.parse<string[]>(text.replace(/\r?\n$/, ''), { delimiter: '\t', newline: text.includes('\r\n') ? '\r\n' : '\n' });
      const grid = parsed.data.length ? parsed.data : [['']];
      this.doc.pasteText(this.si, p.r1, p.c1, grid);
      size = { rows: grid.length, cols: Math.max(...grid.map((l) => l.length)) };
    }
    this.grid.select({ row: p.r1, col: p.c1 }, { row: p.r1 + size.rows - 1, col: p.c1 + size.cols - 1 });
  }

  private async pasteFromClipboard(valuesOnly: boolean): Promise<void> {
    this.paste(await navigator.clipboard.readText(), valuesOnly);
  }

  // -------------------------------------------------------------------------
  // Menus

  private contextMenu(area: HeaderArea, at: { x: number; y: number }): void {
    if (this.editor.active) this.commit(null);
    const run = (fn: () => void | Promise<unknown>) => () => this.guard(fn);
    const [r1, r2] = this.rowSpan();
    const [c1, c2] = this.colSpan();
    const nr = r2 - r1 + 1, nc = c2 - c1 + 1;
    const rows = (n: number) => (n === 1 ? '1 row' : `${n} rows`);
    const cols = (n: number) => (n === 1 ? '1 column' : `${n} columns`);
    const clipboard: MenuEntry[] = [
      { label: 'Cut', icon: 'scissors', hint: '⌘X', qa: 'qa-ctx-cut', run: run(() => { this.grid.el.focus(); document.execCommand('cut'); }) },
      { label: 'Copy', icon: 'copy', hint: '⌘C', qa: 'qa-ctx-copy', run: run(() => { this.grid.el.focus(); document.execCommand('copy'); }) },
      { label: 'Paste', icon: 'clipboard', hint: '⌘V', qa: 'qa-ctx-paste', run: run(() => this.pasteFromClipboard(false)) },
      { label: 'Paste values only', hint: '⇧⌘V', qa: 'qa-ctx-paste-values', run: run(() => this.pasteFromClipboard(true)) },
      'sep',
    ];
    const rowItems: MenuEntry[] = [
      { label: `Insert ${rows(nr)} above`, icon: 'plus-lg', qa: 'qa-ctx-row-above', run: run(() => this.command('insert-row-above')) },
      { label: `Insert ${rows(nr)} below`, icon: 'plus-lg', qa: 'qa-ctx-row-below', run: run(() => this.command('insert-row-below')) },
      { label: `Delete ${rows(nr)}`, icon: 'trash', qa: 'qa-ctx-delete-rows', run: run(() => this.command('delete-rows')) },
    ];
    const colItems: MenuEntry[] = [
      { label: `Insert ${cols(nc)} left`, icon: 'plus-lg', qa: 'qa-ctx-col-left', run: run(() => this.command('insert-col-left')) },
      { label: `Insert ${cols(nc)} right`, icon: 'plus-lg', qa: 'qa-ctx-col-right', run: run(() => this.command('insert-col-right')) },
      { label: `Delete ${cols(nc)}`, icon: 'trash', qa: 'qa-ctx-delete-cols', run: run(() => this.command('delete-cols')) },
    ];
    const span = (a: number, b: number) => Array.from({ length: b - a + 1 }, (_, k) => a + k);
    let items: MenuEntry[];
    if (area === 'row') {
      items = [...clipboard, ...rowItems, 'sep',
        { label: `Hide ${rows(nr)}`, icon: 'eye-slash', qa: 'qa-ctx-hide-rows', run: run(() => this.doc.setHidden(this.si, 'row', span(r1, r2), true)) },
        { label: 'Unhide rows', icon: 'eye', qa: 'qa-ctx-unhide-rows', run: run(() => this.doc.setHidden(this.si, 'row', span(Math.max(0, r1 - 1), r2 + 1), false)) },
        { label: 'Fit row height', qa: 'qa-ctx-fit-rows', run: run(() => this.doc.setRowHeight(this.si, span(r1, r2), undefined)) },
      ];
    } else if (area === 'col') {
      items = [...clipboard, ...colItems, 'sep',
        { label: `Hide ${cols(nc)}`, icon: 'eye-slash', qa: 'qa-ctx-hide-cols', run: run(() => this.doc.setHidden(this.si, 'col', span(c1, c2), true)) },
        { label: 'Unhide columns', icon: 'eye', qa: 'qa-ctx-unhide-cols', run: run(() => this.doc.setHidden(this.si, 'col', span(Math.max(0, c1 - 1), c2 + 1), false)) },
        { label: 'Fit to data', qa: 'qa-ctx-fit-cols', run: run(() => this.doc.setColWidth(this.si, span(c1, c2), span(c1, c2).map((c) => pxToColWidth(this.grid.fitWidth(c))))) },
        'sep',
        { label: 'Sort sheet A → Z', icon: 'sort-alpha-down', qa: 'qa-ctx-sort-asc', run: run(() => this.sortColumn(false)) },
        { label: 'Sort sheet Z → A', icon: 'sort-alpha-down-alt', qa: 'qa-ctx-sort-desc', run: run(() => this.sortColumn(true)) },
      ];
    } else {
      items = [...clipboard, ...rowItems, ...colItems, 'sep',
        { label: 'Clear contents', icon: 'backspace', hint: '⌫', qa: 'qa-ctx-clear', run: run(() => this.doc.clear(this.si, this.ranges())) },
        { label: 'Insert link', icon: 'link-45deg', hint: '⌘K', qa: 'qa-ctx-link', run: run(() => this.insertLink()) },
        'sep',
        { label: 'Sort range A → Z', icon: 'sort-alpha-down', qa: 'qa-ctx-sort-range-asc', run: run(() => this.command('sort-asc')) },
        { label: 'Sort range Z → A', icon: 'sort-alpha-down-alt', qa: 'qa-ctx-sort-range-desc', run: run(() => this.command('sort-desc')) },
      ];
    }
    menu(items, at, 'qa-context-menu');
  }

  /** Sorts the data (below any frozen header) by the active column. */
  private sortColumn(desc: boolean): void {
    if (this.sheet.filter) { this.doc.sortFilterColumn(this.si, this.grid.sel.anchor.col, desc); return; }
    const region = dataRegion(this.sheet);
    if (!region) throw new Error('There’s no data to sort.');
    const header = this.sheet.freeze.rows > region.r1;
    this.doc.sort(this.si, region, [{ col: this.grid.sel.anchor.col, desc }], header);
  }

  private sortRange(desc: boolean): void {
    if (this.grid.sel.isSingleCell()) { this.sortColumn(desc); return; }
    const p = this.clamp(this.grid.sel.primary());
    this.doc.sort(this.si, p, [{ col: this.grid.sel.anchor.col, desc }], false);
  }

  private zoomBy(dir: number): void {
    const z = this.grid.zoom;
    const next = dir > 0 ? ZOOMS.find((x) => x > z + 0.001) ?? ZOOMS[ZOOMS.length - 1] : [...ZOOMS].reverse().find((x) => x < z - 0.001) ?? ZOOMS[0];
    this.setZoom(next);
  }

  private setZoom(z: number): void {
    this.commit(null);
    this.grid.setZoom(z);
    this.updateChrome();
  }

  /** Menu bar commands (and the toolbar's and keyboard's equivalents). */
  command(cmd: MenuCommand): void {
    this.guard(() => this.run(cmd));
  }

  private run(cmd: MenuCommand): void | Promise<unknown> {
    const si = this.si;
    const sel = this.grid.sel;
    const editing = this.editor.active;
    switch (cmd) {
      case 'new': return window.api.newWindow();
      case 'open': return window.api.openDialog(!this.path && !this.doc.dirty);
      case 'save': return this.save(false);
      case 'save-as': return this.save(true);
      case 'save-then-close': return this.save(false).then((ok) => { if (ok) window.api.closeConfirmed(); });
      case 'undo':
        if (editing || isTextTarget(document.activeElement)) { document.execCommand('undo'); return; }
        this.doc.undo();
        return;
      case 'redo':
        if (editing || isTextTarget(document.activeElement)) { document.execCommand('redo'); return; }
        this.doc.redo();
        return;
      case 'find': case 'replace':
        this.commit(null);
        this.find.show(cmd === 'replace');
        return;
      case 'select-all':
        if (editing || isTextTarget(document.activeElement)) { document.execCommand('selectAll'); return; }
        this.grid.selectAll();
        return;
      case 'paste-values': return this.pasteFromClipboard(true);
      case 'bold': case 'italic': case 'underline': case 'strike':
        this.tool({ kind: cmd });
        return;
      case 'clear-format': this.tool({ kind: 'clearFormat' }); return;
      case 'insert-link': return this.insertLink();
      case 'auto-table': this.commit(null); this.doc.autoTable(si); return;
      case 'toggle-filter': {
        this.commit(null);
        const p = this.clamp(sel.primary());
        this.doc.toggleFilter(si, sel.isSingleCell() ? undefined : p);
        return;
      }
      case 'clear-filters': this.doc.clearFilters(si); return;
      case 'sort-asc': case 'sort-desc': this.commit(null); this.sortRange(cmd === 'sort-desc'); return;
      case 'freeze-none': this.doc.setFreeze(si, 0, 0); return;
      case 'freeze-row': this.doc.setFreeze(si, 1, this.sheet.freeze.cols); return;
      case 'freeze-col': this.doc.setFreeze(si, this.sheet.freeze.rows, 1); return;
      case 'freeze-selection': this.doc.setFreeze(si, sel.anchor.row, sel.anchor.col); return;
      case 'zoom-in': this.zoomBy(1); return;
      case 'zoom-out': this.zoomBy(-1); return;
      case 'zoom-reset': this.setZoom(1); return;
      case 'toggle-gridlines': this.doc.setGridLines(si, !this.sheet.showGridLines); return;
      case 'toggle-formulas':
        this.grid.showFormulas = !this.grid.showFormulas;
        this.grid.relayout();
        return;
      case 'insert-row-above': case 'insert-row-below': case 'delete-rows': {
        this.commit(null);
        const [r1, r2] = this.rowSpan();
        const n = r2 - r1 + 1;
        if (cmd === 'delete-rows') this.doc.insertRows(si, r1, -n);
        else this.doc.insertRows(si, cmd === 'insert-row-above' ? r1 : r2 + 1, n);
        return;
      }
      case 'insert-col-left': case 'insert-col-right': case 'delete-cols': {
        this.commit(null);
        const [c1, c2] = this.colSpan();
        const n = c2 - c1 + 1;
        if (cmd === 'delete-cols') this.doc.insertCols(si, c1, -n);
        else this.doc.insertCols(si, cmd === 'insert-col-left' ? c1 : c2 + 1, n);
        return;
      }
      case 'add-sheet': this.commit(null); this.doc.addSheet(); return;
      case 'fill-down': case 'fill-right': {
        this.commit(null);
        const p = this.clamp(sel.primary());
        const down = cmd === 'fill-down';
        // With one row (column) selected, fill copies from the cell above (left), as in Sheets.
        const r = down && p.r1 === p.r2 ? { ...p, r1: Math.max(0, p.r1 - 1) } : !down && p.c1 === p.c2 ? { ...p, c1: Math.max(0, p.c1 - 1) } : p;
        this.doc.fill(si, r, down ? 'down' : 'right');
        return;
      }
    }
  }

  // -------------------------------------------------------------------------
  // Keyboard and window events

  private wire(): void {
    this.grid.el.addEventListener('keydown', (e) => this.guard(() => this.gridKey(e)));
    this.grid.el.addEventListener('wheel', (e) => {
      if (!e.ctrlKey) return;
      e.preventDefault();
      this.setZoom(this.grid.zoom * (e.deltaY < 0 ? 1.1 : 1 / 1.1));
    }, { passive: false });

    this.bar.addEventListener('focus', () => this.guard(() => { if (!this.editor.active) this.startEdit(undefined, 'bar'); }));
    this.bar.addEventListener('keydown', (e) => { e.stopPropagation(); this.editor.onKey(e); });

    this.nameBox.addEventListener('focus', () => this.nameBox.select());
    this.nameBox.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Enter') { e.preventDefault(); this.guard(() => { this.commit(null); this.goTo(this.nameBox.value); }); }
      if (e.key === 'Escape') { e.preventDefault(); this.nameBox.value = this.selectionName(); this.grid.el.focus(); }
    });
    this.nameBox.addEventListener('blur', () => { this.nameBox.value = this.editor.active ? toA1(this.editor.row, this.editor.col) : this.selectionName(); });

    document.addEventListener('copy', (e) => this.guard(() => this.copy(e, false)));
    document.addEventListener('cut', (e) => this.guard(() => this.copy(e, true)));
    document.addEventListener('paste', (e) => {
      if (isTextTarget(e.target) || !this.doc) return;
      e.preventDefault();
      this.guard(() => this.paste(e.clipboardData!.getData('text/plain'), false));
    });

    dark.addEventListener('change', () => this.grid.relayout());
  }

  private gridKey(e: KeyboardEvent): void {
    if (e.isComposing || this.sheet.kind !== 'worksheet') return;
    const g = this.grid;
    const mod = e.metaKey || e.ctrlKey;
    const arrows: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    const d = arrows[e.key];
    if (d) { e.preventDefault(); g.move(d[0], d[1], e.shiftKey, mod); return; }
    switch (e.key) {
      case 'Enter': e.preventDefault(); this.startEdit(undefined, 'cell'); return;
      case 'F2': e.preventDefault(); this.startEdit(undefined, 'cell'); return;
      case 'Tab': e.preventDefault(); g.advance(0, e.shiftKey ? -1 : 1); return;
      case 'Delete': case 'Backspace':
        e.preventDefault();
        this.doc.clear(this.si, this.ranges());
        return;
      case 'PageDown': e.preventDefault(); g.page(1, e.shiftKey); return;
      case 'PageUp': e.preventDefault(); g.page(-1, e.shiftKey); return;
      case 'Home':
        e.preventDefault();
        g.select({ row: mod ? 0 : g.sel.anchor.row, col: 0 });
        return;
      case 'End': {
        e.preventDefault();
        const ext = sheetExtent(this.sheet);
        g.select({ row: mod ? Math.max(0, ext.rows - 1) : g.sel.anchor.row, col: Math.max(0, ext.cols - 1) });
        return;
      }
      case 'Escape':
        if (this.clip) { this.clip = null; g.marquee = null; g.render(); }
        return;
    }
    if (e.key.length === 1 && !mod) {
      e.preventDefault();
      this.startEdit(e.key, 'cell');
    }
  }
}
