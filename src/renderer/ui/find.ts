import { Doc } from '../../core/doc';
import { Cell } from '../../core/model';
import { displayCell } from '../../core/numfmt';
import { editText } from './editor';
import { h, icon } from './dom';

export interface Hit { row: number; col: number }

export interface FindHost {
  doc(): Doc;
  sheet(): number;
  /** Where searching starts from. */
  active(): Hit;
  select(hit: Hit): void;
  close(): void;
}

/** Find (and replace) in the current sheet, matching what cells show, or their formulas. */
export class FindBar {
  readonly el: HTMLElement;
  private query: HTMLInputElement;
  private replacement: HTMLInputElement;
  private replaceRow: HTMLElement;
  private matchCase: HTMLInputElement;
  private inFormulas: HTMLInputElement;
  private count: HTMLElement;
  private hits: Hit[] = [];
  private index = -1;

  constructor(private host: FindHost) {
    this.el = h('div', 'find-bar popover d-flex flex-col gap-2 d-none qa-find-bar');
    this.el.style.position = 'absolute';
    const row1 = h('div', 'd-flex align-center gap-1', this.el);
    this.query = h('input', 'flex-1 qa-find-input', row1);
    this.query.placeholder = 'Find in sheet';
    this.count = h('span', 'small muted nowrap qa-find-count', row1);
    const btn = (ic: string, title: string, qa: string, fn: () => void, parent = row1) => {
      const b = h('button', `btn-icon ${qa}`, parent);
      b.title = title;
      b.append(icon(ic));
      b.onclick = fn;
      return b;
    };
    btn('chevron-up', 'Previous (⇧↵)', 'qa-find-prev', () => this.step(-1));
    btn('chevron-down', 'Next (↵)', 'qa-find-next', () => this.step(1));
    btn('x-lg', 'Close (Esc)', 'qa-find-close', () => this.hide());
    this.replaceRow = h('div', 'd-flex align-center gap-1', this.el);
    this.replacement = h('input', 'flex-1 qa-replace-input', this.replaceRow);
    this.replacement.placeholder = 'Replace with';
    const one = h('button', 'btn qa-replace-one', this.replaceRow, 'Replace');
    const all = h('button', 'btn qa-replace-all', this.replaceRow, 'All');
    one.onclick = () => this.replaceOne();
    all.onclick = () => this.replaceAll();
    const opts = h('div', 'd-flex gap-3 small muted', this.el);
    const check = (label: string, qa: string) => {
      const l = h('label', 'd-flex align-center gap-1', opts);
      const c = h('input', qa, l);
      c.type = 'checkbox';
      l.append(label);
      c.onchange = () => this.search();
      return c;
    };
    this.matchCase = check('Match case', 'qa-find-case');
    this.inFormulas = check('Search in formulas', 'qa-find-formulas');

    this.query.addEventListener('input', () => this.search());
    this.el.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Escape') { e.preventDefault(); this.hide(); }
      if (e.key === 'Enter' && e.target === this.query) { e.preventDefault(); this.step(e.shiftKey ? -1 : 1); }
      if (e.key === 'Enter' && e.target === this.replacement) { e.preventDefault(); this.replaceOne(); }
    });
  }

  get open(): boolean { return !this.el.classList.contains('d-none'); }

  show(replace: boolean): void {
    this.el.classList.remove('d-none');
    this.replaceRow.classList.toggle('d-none', !replace);
    this.query.focus();
    this.query.select();
    this.search();
  }

  hide(): void {
    this.el.classList.add('d-none');
    this.hits = [];
    this.host.close();
  }

  private text(c: Cell): string {
    const doc = this.host.doc();
    return this.inFormulas.checked && c.f !== undefined ? editText(doc.wb, c) : displayCell(doc.wb, c).text;
  }

  private matches(s: string): boolean {
    const q = this.query.value;
    return this.matchCase.checked ? s.includes(q) : s.toLowerCase().includes(q.toLowerCase());
  }

  /** Recomputes the hits in reading order (row by row) and jumps to the first one at or after the active cell. */
  search(): void {
    this.hits = [];
    this.index = -1;
    const q = this.query.value;
    if (q) {
      const sheet = this.host.doc().wb.sheets[this.host.sheet()];
      sheet.rows.forEach((row, r) => {
        if (!row || row.hidden || row.filtered) return;
        row.cells.forEach((c, col) => { if (c && this.matches(this.text(c))) this.hits.push({ row: r, col }); });
      });
      const a = this.host.active();
      const at = this.hits.findIndex((hh) => hh.row > a.row || (hh.row === a.row && hh.col >= a.col));
      if (this.hits.length) { this.index = at < 0 ? 0 : at; this.host.select(this.hits[this.index]); }
    }
    this.showCount();
  }

  private showCount(): void {
    this.count.textContent = !this.query.value ? '' : this.hits.length ? `${this.index + 1} of ${this.hits.length}` : 'No results';
  }

  private step(d: number): void {
    if (!this.hits.length) return;
    this.index = (this.index + d + this.hits.length) % this.hits.length;
    this.host.select(this.hits[this.index]);
    this.showCount();
  }

  /** The cell's input text with every match replaced. */
  private replaced(c: Cell): string {
    const doc = this.host.doc();
    const src = editText(doc.wb, c);
    const q = this.query.value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    return src.replace(new RegExp(q, this.matchCase.checked ? 'g' : 'gi'), () => this.replacement.value);
  }

  private replaceOne(): void {
    const hit = this.hits[this.index];
    if (!hit) return;
    const doc = this.host.doc();
    const si = this.host.sheet();
    const c = doc.wb.sheets[si].rows[hit.row]!.cells[hit.col]!;
    doc.setInput(si, hit.row, hit.col, this.replaced(c));
    const at = this.index;
    this.search();
    if (this.hits.length) { this.index = Math.min(at, this.hits.length - 1); this.host.select(this.hits[this.index]); this.showCount(); }
  }

  private replaceAll(): void {
    if (!this.hits.length) return;
    const doc = this.host.doc();
    const si = this.host.sheet();
    const sheet = doc.wb.sheets[si];
    const cells = this.hits.map((hit) => ({
      row: hit.row, col: hit.col,
      cell: doc.typedCell(si, hit.row, hit.col, this.replaced(sheet.rows[hit.row]!.cells[hit.col]!)),
    }));
    doc.putCells(si, cells, 'Replace all');
    this.count.textContent = `Replaced ${cells.length}`;
    this.hits = [];
  }
}
