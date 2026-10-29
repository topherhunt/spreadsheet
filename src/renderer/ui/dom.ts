export function h<K extends keyof HTMLElementTagNameMap>(tag: K, cls = '', parent?: HTMLElement, text?: string): HTMLElementTagNameMap[K] {
  const el = document.createElement(tag);
  if (cls) el.className = cls;
  if (text !== undefined) el.textContent = text;
  parent?.appendChild(el);
  return el;
}

export function icon(name: string): HTMLElement {
  return h('i', `bi bi-${name}`);
}

let openPopover: { el: HTMLElement; close: () => void } | null = null;

export function closePopover(): void {
  openPopover?.close();
}

/**
 * Shows `el` as a floating popover at a client point (kept on screen), closing
 * any other. Clicking outside or pressing Escape closes it.
 */
export function popover(el: HTMLElement, at: { x: number; y: number }, onClose?: () => void): () => void {
  closePopover();
  el.classList.add('popover');
  document.body.appendChild(el);
  const r = el.getBoundingClientRect();
  el.style.left = `${Math.max(4, Math.min(at.x, window.innerWidth - r.width - 6))}px`;
  el.style.top = `${at.y + r.height > window.innerHeight - 6 ? Math.max(4, at.y - r.height - 4) : at.y}px`;
  const outside = (e: MouseEvent) => { if (!el.contains(e.target as Node)) close(); };
  const key = (e: KeyboardEvent) => { if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); close(); } };
  let closed = false;
  const close = () => {
    if (closed) return;
    closed = true;
    el.remove();
    window.removeEventListener('mousedown', outside, true);
    window.removeEventListener('keydown', key, true);
    if (openPopover?.el === el) openPopover = null;
    onClose?.();
  };
  // Wait a frame so the click that opened this doesn't close it.
  requestAnimationFrame(() => { if (!closed) window.addEventListener('mousedown', outside, true); });
  window.addEventListener('keydown', key, true);
  openPopover = { el, close };
  return close;
}

export function popoverBelow(anchor: HTMLElement): { x: number; y: number } {
  const r = anchor.getBoundingClientRect();
  return { x: r.left, y: r.bottom + 4 };
}

export interface MenuItem {
  label: string;
  icon?: string;
  hint?: string;
  qa?: string;
  disabled?: boolean;
  checked?: boolean;
  run: () => void;
}

export type MenuEntry = MenuItem | 'sep';

export function menu(items: MenuEntry[], at: { x: number; y: number }, cls = ''): () => void {
  const el = h('div', `menu ${cls}`);
  let close = () => {};
  for (const it of items) {
    if (it === 'sep') { h('div', 'menu-sep', el); continue; }
    const row = h('div', `menu-item${it.disabled ? ' disabled' : ''}${it.qa ? ' ' + it.qa : ''}`, el);
    // An empty <i class="bi"> keeps labels aligned when the item has no icon.
    row.appendChild(it.checked ? icon('check2') : it.icon ? icon(it.icon) : h('i', 'bi'));
    h('span', '', row, it.label);
    if (it.hint) h('span', 'hint', row, it.hint);
    row.addEventListener('click', () => { close(); it.run(); });
  }
  close = popover(el, at);
  return close;
}

let toastTimer = 0;

export function toast(message: string, error = false): void {
  document.querySelector('.toast')?.remove();
  const el = h('div', `toast qa-toast${error ? ' error' : ''}`, document.body, message);
  clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => el.remove(), error ? 5000 : 2500);
}

/** A small modal with a text field; resolves to the entered text, or null when cancelled. */
export function prompt(title: string, value: string, opts: { placeholder?: string; ok?: string; qa?: string } = {}): Promise<string | null> {
  return new Promise((resolve) => {
    const back = h('div', 'modal-backdrop', document.body);
    const box = h('div', `modal ${opts.qa ?? ''}`, back);
    h('div', 'bold mb-2', box, title);
    const input = h('input', 'w-100 qa-prompt-input', box);
    input.value = value;
    input.placeholder = opts.placeholder ?? '';
    const row = h('div', 'd-flex justify-end gap-2 mt-2', box);
    const cancel = h('button', 'btn qa-prompt-cancel', row, 'Cancel');
    const ok = h('button', 'btn btn-primary qa-prompt-ok', row, opts.ok ?? 'OK');
    const done = (v: string | null) => { back.remove(); resolve(v); };
    cancel.onclick = () => done(null);
    ok.onclick = () => done(input.value);
    back.addEventListener('mousedown', (e) => { if (e.target === back) done(null); });
    input.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Enter') { e.preventDefault(); done(input.value); }
      if (e.key === 'Escape') { e.preventDefault(); done(null); }
    });
    input.focus();
    input.select();
  });
}

/** "#RRGGBB" for a CSS color string like "rgb(1, 2, 3)". */
export function cssToHex(css: string): string | undefined {
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)/.exec(css);
  if (!m || (m[4] !== undefined && Number(m[4]) === 0)) return undefined;
  return '#' + [m[1], m[2], m[3]].map((n) => Number(n).toString(16).padStart(2, '0')).join('').toUpperCase();
}
