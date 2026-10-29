import { h, icon, menu, MenuEntry, popover, popoverBelow } from './dom';

export type ToolAction =
  | { kind: 'undo' } | { kind: 'redo' }
  | { kind: 'numFmt'; code: string } | { kind: 'decimals'; delta: number }
  | { kind: 'fontName'; name: string } | { kind: 'fontSize'; size: number }
  | { kind: 'bold' } | { kind: 'italic' } | { kind: 'underline' } | { kind: 'strike' }
  | { kind: 'color'; color: string | null } | { kind: 'fill'; color: string | null }
  | { kind: 'borders'; preset: 'all' | 'outer' | 'inner' | 'top' | 'bottom' | 'left' | 'right' | 'none' }
  | { kind: 'merge' }
  | { kind: 'hAlign'; value: string | null } | { kind: 'vAlign'; value: string | null } | { kind: 'wrap' }
  | { kind: 'link' } | { kind: 'clearFormat' }
  | { kind: 'filter' } | { kind: 'autoTable' };

export interface ToolState {
  canUndo: boolean; canRedo: boolean;
  bold: boolean; italic: boolean; underline: boolean; strike: boolean;
  fontName: string; fontSize: number;
  color?: string; fill?: string;
  hAlign?: string; vAlign: string; wrap: boolean;
  merged: boolean; filter: boolean;
}

export const NUMBER_FORMATS: { label: string; code: string; example: string }[] = [
  { label: 'Automatic', code: 'General', example: '' },
  { label: 'Plain text', code: '@', example: '' },
  { label: 'Number', code: '#,##0.00', example: '1,000.12' },
  { label: 'Percent', code: '0.00%', example: '10.12%' },
  { label: 'Scientific', code: '0.00E+00', example: '1.01E+03' },
  { label: 'Currency', code: '"$"#,##0.00', example: '$1,000.12' },
  { label: 'Currency rounded', code: '"$"#,##0', example: '$1,000' },
  { label: 'Date', code: 'm/d/yyyy', example: '9/26/2008' },
  { label: 'Time', code: 'h:mm:ss AM/PM', example: '3:59:00 PM' },
  { label: 'Date time', code: 'm/d/yyyy h:mm:ss', example: '9/26/2008 15:59:00' },
  { label: 'Duration', code: '[h]:mm:ss', example: '24:01:00' },
];

const FONTS = ['Arial', 'Calibri', 'Cambria', 'Comic Sans MS', 'Courier New', 'Georgia', 'Helvetica', 'Roboto', 'Times New Roman', 'Trebuchet MS', 'Verdana'];

const PALETTE = [
  '#000000', '#434343', '#666666', '#999999', '#B7B7B7', '#CCCCCC', '#D9D9D9', '#EFEFEF', '#F3F3F3', '#FFFFFF',
  '#980000', '#FF0000', '#FF9900', '#FFFF00', '#00FF00', '#00FFFF', '#4A86E8', '#0000FF', '#9900FF', '#FF00FF',
  '#E6B8AF', '#F4CCCC', '#FCE5CD', '#FFF2CC', '#D9EAD3', '#D0E0E3', '#C9DAF8', '#CFE2F3', '#D9D2E9', '#EAD1DC',
  '#DD7E6B', '#EA9999', '#F9CB9C', '#FFE599', '#B6D7A8', '#A2C4C9', '#A4C2F4', '#9FC5E8', '#B4A7D6', '#D5A6BD',
  '#CC4125', '#E06666', '#F6B26B', '#FFD966', '#93C47D', '#76A5AF', '#6D9EEB', '#6FA8DC', '#8E7CC3', '#C27BA0',
  '#A61C00', '#CC0000', '#E69138', '#F1C232', '#6AA84F', '#45818E', '#3C78D8', '#3D85C6', '#674EA7', '#A64D79',
  '#5B0F00', '#660000', '#783F04', '#7F6000', '#274E13', '#0C343D', '#1C4587', '#073763', '#20124D', '#4C1130',
];

function paletteMenu(anchor: HTMLElement, current: string | undefined, resetLabel: string, qa: string, pick: (c: string | null) => void): void {
  const el = h('div', `p-2 d-flex flex-col gap-2 ${qa}`);
  const reset = h('button', 'btn w-100 qa-color-reset', el);
  reset.append(icon('slash-circle'), resetLabel);
  const grid = h('div', 'palette', el);
  let close = () => {};
  for (const c of PALETTE) {
    const sw = h('div', `swatch qa-swatch${current?.toUpperCase() === c ? ' current' : ''}`, grid);
    sw.style.background = c;
    sw.title = c;
    sw.dataset.color = c;
    sw.onclick = () => { close(); pick(c); };
  }
  const custom = h('label', 'd-flex align-center gap-2 small muted', el);
  const input = h('input', '', custom);
  input.type = 'color';
  input.value = current ?? '#000000';
  custom.append('Custom…');
  input.onchange = () => { close(); pick(input.value.toUpperCase()); };
  reset.onclick = () => { close(); pick(null); };
  close = popover(el, popoverBelow(anchor));
}

/** The formatting toolbar. It reports clicks as actions and shows the active cell's formatting. */
export class Toolbar {
  readonly el: HTMLElement;
  private btns = new Map<string, HTMLButtonElement>();
  private fontSel: HTMLSelectElement;
  private sizeInput: HTMLInputElement;
  private colorBar: HTMLElement;
  private fillBar: HTMLElement;
  private state: ToolState | null = null;

  constructor(private act: (a: ToolAction) => void) {
    this.el = h('div', 'toolbar d-flex align-center border-b bg-chrome qa-toolbar');
    const btn = (name: string, ic: string, title: string, a: () => void, text?: string) => {
      const b = h('button', `btn-icon qa-tb-${name}`, this.el);
      b.title = title;
      b.append(icon(ic));
      if (text) h('span', 'btn-text', b, text);
      // Keep focus (and the text selection) in the cell editor while clicking toolbar buttons.
      b.addEventListener('mousedown', (e) => e.preventDefault());
      b.onclick = () => a();
      this.btns.set(name, b);
      return b;
    };
    const sep = () => h('div', 'tb-sep', this.el);

    btn('undo', 'arrow-counterclockwise', 'Undo (⌘Z)', () => act({ kind: 'undo' }));
    btn('redo', 'arrow-clockwise', 'Redo (⇧⌘Z)', () => act({ kind: 'redo' }));
    sep();
    btn('currency', 'currency-dollar', 'Format as currency', () => act({ kind: 'numFmt', code: '"$"#,##0.00' }));
    btn('percent', 'percent', 'Format as percent', () => act({ kind: 'numFmt', code: '0.00%' }));
    btn('dec-less', 'chevron-left', 'Decrease decimal places', () => act({ kind: 'decimals', delta: -1 }), '.0');
    btn('dec-more', 'chevron-right', 'Increase decimal places', () => act({ kind: 'decimals', delta: 1 }), '.00');
    const fmt = btn('numfmt', '123', 'More formats', () => {
      const items: MenuEntry[] = NUMBER_FORMATS.map((f) => ({ label: f.label, hint: f.example, qa: `qa-numfmt-${f.label.toLowerCase().replace(/\s+/g, '-')}`, run: () => act({ kind: 'numFmt', code: f.code }) }));
      menu(items, popoverBelow(fmt));
    });
    fmt.append(h('i', 'bi bi-caret-down-fill caret'));
    sep();

    this.fontSel = h('select', 'font-select qa-tb-font', this.el);
    this.fontSel.title = 'Font';
    this.fontSel.onchange = () => act({ kind: 'fontName', name: this.fontSel.value });
    this.sizeInput = h('input', 'size-input qa-tb-size', this.el);
    this.sizeInput.title = 'Font size';
    this.sizeInput.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Enter') { e.preventDefault(); this.commitSize(); }
      if (e.key === 'Escape') { e.preventDefault(); this.sizeInput.blur(); }
    });
    this.sizeInput.addEventListener('change', () => this.commitSize());
    sep();

    btn('bold', 'type-bold', 'Bold (⌘B)', () => act({ kind: 'bold' }));
    btn('italic', 'type-italic', 'Italic (⌘I)', () => act({ kind: 'italic' }));
    btn('underline', 'type-underline', 'Underline (⌘U)', () => act({ kind: 'underline' }));
    btn('strike', 'type-strikethrough', 'Strikethrough (⇧⌘X)', () => act({ kind: 'strike' }));
    const color = btn('color', 'fonts', 'Text color', () => paletteMenu(color, this.state?.color, 'Automatic', 'qa-palette-color', (c) => act({ kind: 'color', color: c })));
    color.classList.add('flex-col');
    this.colorBar = h('span', 'color-bar', color);
    const fill = btn('fill', 'paint-bucket', 'Fill color', () => paletteMenu(fill, this.state?.fill, 'No fill', 'qa-palette-fill', (c) => act({ kind: 'fill', color: c })));
    fill.classList.add('flex-col');
    this.fillBar = h('span', 'color-bar', fill);
    sep();

    const borders = btn('borders', 'border-all', 'Borders', () => {
      const presets: [string, string, ToolAction & { kind: 'borders' }][] = [
        ['All borders', 'border-all', { kind: 'borders', preset: 'all' }],
        ['Inner borders', 'border-inner', { kind: 'borders', preset: 'inner' }],
        ['Outer borders', 'border-outer', { kind: 'borders', preset: 'outer' }],
        ['Top border', 'border-top', { kind: 'borders', preset: 'top' }],
        ['Bottom border', 'border-bottom', { kind: 'borders', preset: 'bottom' }],
        ['Left border', 'border-left', { kind: 'borders', preset: 'left' }],
        ['Right border', 'border-right', { kind: 'borders', preset: 'right' }],
        ['Clear borders', 'border', { kind: 'borders', preset: 'none' }],
      ];
      menu(presets.map(([label, ic, a]) => ({ label, icon: ic, qa: `qa-border-${a.preset}`, run: () => act(a) })), popoverBelow(borders));
    });
    btn('merge', 'arrows-angle-expand', 'Merge cells', () => act({ kind: 'merge' }));
    sep();
    const align = btn('align', 'text-left', 'Horizontal align', () => {
      const cur = this.state?.hAlign;
      menu([
        { label: 'Left', icon: 'text-left', checked: cur === 'left', qa: 'qa-align-left', run: () => act({ kind: 'hAlign', value: 'left' }) },
        { label: 'Center', icon: 'text-center', checked: cur === 'center', qa: 'qa-align-center', run: () => act({ kind: 'hAlign', value: 'center' }) },
        { label: 'Right', icon: 'text-right', checked: cur === 'right', qa: 'qa-align-right', run: () => act({ kind: 'hAlign', value: 'right' }) },
        { label: 'Automatic', checked: !cur || cur === 'general', qa: 'qa-align-auto', run: () => act({ kind: 'hAlign', value: null }) },
      ], popoverBelow(align));
    });
    const valign = btn('valign', 'align-bottom', 'Vertical align', () => {
      const cur = this.state?.vAlign;
      menu([
        { label: 'Top', icon: 'align-top', checked: cur === 'top', qa: 'qa-valign-top', run: () => act({ kind: 'vAlign', value: 'top' }) },
        { label: 'Middle', icon: 'align-middle', checked: cur === 'center', qa: 'qa-valign-middle', run: () => act({ kind: 'vAlign', value: 'center' }) },
        { label: 'Bottom', icon: 'align-bottom', checked: cur === 'bottom', qa: 'qa-valign-bottom', run: () => act({ kind: 'vAlign', value: null }) },
      ], popoverBelow(valign));
    });
    btn('wrap', 'text-wrap', 'Wrap text', () => act({ kind: 'wrap' }));
    sep();
    btn('link', 'link-45deg', 'Insert link (⌘K)', () => act({ kind: 'link' }));
    btn('clear-format', 'eraser', 'Clear formatting (⌘\\)', () => act({ kind: 'clearFormat' }));
    btn('filter', 'funnel', 'Create a filter (⇧⌘L)', () => act({ kind: 'filter' }));

    const auto = h('button', 'btn-accent ms-auto qa-tb-auto-table', this.el);
    auto.title = 'Bold the header row, freeze it, and add a filter (⌥⌘T)';
    auto.append(icon('table'), 'Auto table');
    auto.addEventListener('mousedown', (e) => e.preventDefault());
    auto.onclick = () => act({ kind: 'autoTable' });
  }

  private commitSize(): void {
    const n = Number(this.sizeInput.value);
    if (n >= 1 && n <= 400) this.act({ kind: 'fontSize', size: n });
    else if (this.state) this.sizeInput.value = String(this.state.fontSize);
  }

  update(s: ToolState): void {
    this.state = s;
    const set = (name: string, on: boolean) => this.btns.get(name)!.classList.toggle('active', on);
    this.btns.get('undo')!.disabled = !s.canUndo;
    this.btns.get('redo')!.disabled = !s.canRedo;
    set('bold', s.bold); set('italic', s.italic); set('underline', s.underline); set('strike', s.strike);
    set('wrap', s.wrap); set('merge', s.merged); set('filter', s.filter);
    this.btns.get('filter')!.title = s.filter ? 'Remove filter (⇧⌘L)' : 'Create a filter (⇧⌘L)';
    const alignIcon = s.hAlign === 'center' ? 'text-center' : s.hAlign === 'right' ? 'text-right' : 'text-left';
    this.btns.get('align')!.querySelector('i')!.className = `bi bi-${alignIcon}`;
    const vIcon = s.vAlign === 'top' ? 'align-top' : s.vAlign === 'center' ? 'align-middle' : 'align-bottom';
    this.btns.get('valign')!.querySelector('i')!.className = `bi bi-${vIcon}`;
    this.colorBar.style.background = s.color ?? 'var(--fg)';
    this.fillBar.style.background = s.fill ?? 'transparent';
    this.fillBar.style.boxShadow = s.fill ? '' : 'inset 0 0 0 1px var(--border-strong)';
    const fonts = FONTS.includes(s.fontName) ? FONTS : [s.fontName, ...FONTS];
    if (this.fontSel.options.length !== fonts.length || this.fontSel.options[0].value !== fonts[0]) {
      this.fontSel.textContent = '';
      for (const f of fonts) { const o = h('option', '', this.fontSel, f); o.value = f; }
    }
    this.fontSel.value = s.fontName;
    if (document.activeElement !== this.sizeInput) this.sizeInput.value = String(s.fontSize);
  }
}
