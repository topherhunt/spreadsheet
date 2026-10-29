import { colToName } from '../../core/addr';
import { Doc } from '../../core/doc';
import { columnValues, extendFilterRange } from '../../core/filter';
import { ColumnFilter, FilterOp, getCell } from '../../core/model';
import { displayCell } from '../../core/numfmt';
import { h, icon, popover } from './dom';

const OPS: { op: FilterOp | ''; label: string; args: number }[] = [
  { op: '', label: 'None', args: 0 },
  { op: 'empty', label: 'Is empty', args: 0 },
  { op: 'notEmpty', label: 'Is not empty', args: 0 },
  { op: 'contains', label: 'Text contains', args: 1 },
  { op: 'notContains', label: 'Text does not contain', args: 1 },
  { op: 'startsWith', label: 'Text starts with', args: 1 },
  { op: 'endsWith', label: 'Text ends with', args: 1 },
  { op: 'equals', label: 'Is equal to', args: 1 },
  { op: 'notEquals', label: 'Is not equal to', args: 1 },
  { op: 'gt', label: 'Greater than', args: 1 },
  { op: 'gte', label: 'Greater than or equal to', args: 1 },
  { op: 'lt', label: 'Less than', args: 1 },
  { op: 'lte', label: 'Less than or equal to', args: 1 },
  { op: 'between', label: 'Is between', args: 2 },
  { op: 'notBetween', label: 'Is not between', args: 2 },
];

/** The menu under a filter button: sort, filter by condition, filter by values. */
export function openFilterMenu(doc: Doc, si: number, col: number, at: { x: number; y: number }): void {
  const sheet = doc.wb.sheets[si];
  const f = sheet.filter;
  if (!f) throw new Error('Sheet has no filter');
  extendFilterRange(sheet);
  const current = f.columns[col];
  const values = columnValues(doc.wb, sheet, col);
  const checked = new Set(current?.values ?? values.map((v) => v.text));

  const el = h('div', 'filter-menu d-flex flex-col p-2 gap-2 qa-filter-menu');
  const header = displayCell(doc.wb, getCell(sheet, f.range.r1, col)).text || `Column ${colToName(col)}`;
  h('div', 'bold ellipsis', el, header);

  const sorts = h('div', 'd-flex flex-col', el);
  const sortItem = (label: string, ic: string, desc: boolean, qa: string) => {
    const it = h('div', `menu-item ${qa}`, sorts);
    it.style.padding = '4px 6px';
    it.append(icon(ic), label);
    it.onclick = () => { close(); doc.sortFilterColumn(si, col, desc); };
  };
  sortItem('Sort A → Z', 'sort-alpha-down', false, 'qa-sort-asc');
  sortItem('Sort Z → A', 'sort-alpha-down-alt', true, 'qa-sort-desc');
  h('div', 'menu-sep', el);

  // By condition
  h('div', 'section-head', el, 'Filter by condition');
  const opSel = h('select', 'w-100 qa-filter-op', el);
  for (const o of OPS) { const opt = h('option', '', opSel, o.label); opt.value = o.op; }
  opSel.value = current?.condition?.op ?? '';
  const argRow = h('div', 'd-flex gap-1', el);
  const argA = h('input', 'flex-1 qa-filter-arg-a', argRow);
  const argB = h('input', 'flex-1 qa-filter-arg-b', argRow);
  argA.placeholder = 'Value'; argB.placeholder = 'and';
  argA.value = current?.condition?.a ?? '';
  argB.value = current?.condition?.b ?? '';
  const syncArgs = () => {
    const n = OPS.find((o) => o.op === opSel.value)!.args;
    argRow.classList.toggle('d-none', n === 0);
    argB.classList.toggle('d-none', n < 2);
  };
  opSel.onchange = () => { syncArgs(); if (!argRow.classList.contains('d-none')) argA.focus(); };
  syncArgs();

  // By values
  h('div', 'section-head mt-1', el, 'Filter by values');
  const links = h('div', 'd-flex gap-2 small', el);
  const all = h('button', 'link-btn qa-filter-select-all', links, `Select all ${values.length}`);
  const none = h('button', 'link-btn qa-filter-clear', links, 'Clear');
  const search = h('input', 'w-100 qa-filter-search', el);
  search.type = 'search';
  search.placeholder = 'Search values';
  const list = h('div', 'filter-values qa-filter-values', el);
  const boxes: { text: string; box: HTMLInputElement; row: HTMLElement }[] = [];
  for (const v of values) {
    const row = h('label', 'qa-filter-value', list);
    const box = h('input', '', row);
    box.type = 'checkbox';
    box.checked = checked.has(v.text);
    box.dataset.value = v.text;
    box.onchange = () => { if (box.checked) checked.add(v.text); else checked.delete(v.text); };
    h('span', `ellipsis${v.text ? '' : ' muted'}`, row, v.text || '(Blanks)');
    h('span', 'count', row, String(v.count));
    boxes.push({ text: v.text, box, row });
  }
  const visible = () => boxes.filter((b) => !b.row.classList.contains('d-none'));
  all.onclick = () => { for (const b of visible()) { b.box.checked = true; checked.add(b.text); } };
  none.onclick = () => { for (const b of visible()) { b.box.checked = false; checked.delete(b.text); } };
  search.oninput = () => {
    const q = search.value.toLowerCase();
    for (const b of boxes) b.row.classList.toggle('d-none', !!q && !b.text.toLowerCase().includes(q));
  };

  const buttons = h('div', 'd-flex justify-end gap-2 mt-1', el);
  const cancel = h('button', 'btn qa-filter-cancel', buttons, 'Cancel');
  const ok = h('button', 'btn btn-primary qa-filter-ok', buttons, 'OK');
  cancel.onclick = () => close();
  ok.onclick = () => {
    const cf: ColumnFilter = {};
    if (checked.size < values.length) cf.values = values.map((v) => v.text).filter((t) => checked.has(t));
    if (opSel.value) {
      const n = OPS.find((o) => o.op === opSel.value)!.args;
      cf.condition = { op: opSel.value as FilterOp };
      if (n >= 1) cf.condition.a = argA.value;
      if (n >= 2) cf.condition.b = argB.value;
    }
    close();
    doc.setColumnFilter(si, col, cf.values || cf.condition ? cf : null);
  };
  el.addEventListener('keydown', (e) => {
    e.stopPropagation();
    if (e.key === 'Enter' && (e.target as HTMLElement).tagName === 'INPUT' && (e.target as HTMLInputElement).type !== 'checkbox') ok.click();
  });

  const close = popover(el, at);
  search.focus();
}
