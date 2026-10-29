import { expect, Spreadsheet, test } from './harness';

const ref = (row: number, col: number) => `${String.fromCharCode(65 + col)}${row + 1}`;

test('clicking a cell after scrolling down selects that cell', async ({ launch }) => {
  const s = await launch();
  await s.wheel(0, 4000);
  const row = await s.middleRow();
  expect(row).toBeGreaterThan(100);
  const at = await s.center(row, 2);
  await s.page.mouse.click(at.x, at.y);
  await expect(s.page.locator('.qa-name-box')).toHaveValue(ref(row, 2));
});

test('dragging across cells after scrolling down selects the range', async ({ launch }) => {
  const s = await launch();
  await s.wheel(0, 4000);
  const row = await s.middleRow();
  const from = await s.center(row, 1), to = await s.center(row + 3, 3);
  await s.page.mouse.move(from.x, from.y);
  await s.page.mouse.down();
  await s.page.mouse.move(to.x, to.y, { steps: 5 });
  await s.page.mouse.up();
  await expect(s.page.locator('.qa-name-box')).toHaveValue(`${ref(row, 1)}:${ref(row + 3, 3)}`);
});

test('clicking a cell after scrolling right selects that cell', async ({ launch }) => {
  const s = await launch();
  await s.wheel(1500, 0);
  const col = await s.middleCol();
  expect(col).toBeGreaterThan(10);
  const at = await s.center(3, col);
  await s.page.mouse.click(at.x, at.y);
  await expect(s.page.locator('.qa-name-box')).toHaveValue(ref(3, col));
});

test('double-clicking a cell after scrolling down edits that cell', async ({ launch }) => {
  const s = await launch();
  await s.wheel(0, 4000);
  const row = await s.middleRow();
  const at = await s.center(row, 0);
  await s.page.mouse.dblclick(at.x, at.y);
  await expect(s.page.locator('.qa-editor')).toBeVisible();
  await s.page.keyboard.type('deep');
  await s.page.keyboard.press('Enter');
  expect(await s.value(ref(row, 0))).toBe('deep');
});

test('dragging the fill handle after scrolling down fills the cells below', async ({ launch }) => {
  const s = await launch();
  await s.wheel(0, 4000);
  const row = await s.middleRow();
  await s.enter(ref(row, 0), '7');
  await s.selectCell(ref(row, 0));
  const handle = (await s.page.locator('.qa-fill-handle').boundingBox())!;
  const to = await s.center(row + 2, 0);
  await s.page.mouse.move(handle.x + 3, handle.y + 3);
  await s.page.mouse.down();
  await s.page.mouse.move(to.x, to.y, { steps: 5 });
  await s.page.mouse.up();
  expect([await s.value(ref(row + 1, 0)), await s.value(ref(row + 2, 0))]).toEqual([7, 7]);
});

test('clicking a row header after scrolling down selects that row', async ({ launch }) => {
  const s = await launch();
  await s.wheel(0, 4000);
  const row = await s.middleRow();
  const at = await s.center(row, 0);
  const g = (await s.page.locator('.qa-grid').boundingBox())!;
  await s.page.mouse.click(g.x + 10, at.y);
  await expect(s.page.locator('.qa-name-box')).toHaveValue(`${row + 1}:${row + 1}`);
});

test('selecting a whole column from the name box keeps the view at the top', async ({ launch }) => {
  const s = await launch();
  await s.selectCell('B:B');
  expect(await s.page.locator('.qa-grid-scroll').evaluate((el) => el.scrollTop)).toBe(0);
});

test('zooming in after scrolling keeps the active cell in view', async ({ launch }) => {
  const s = await launch();
  await s.wheel(0, 4000);
  const row = await s.middleRow();
  const at = await s.center(row, 2);
  await s.page.mouse.click(at.x, at.y);
  for (let i = 0; i < 3; i++) await s.command('zoom-in');
  await expect(s.page.locator('.qa-zoom')).toHaveText('150%');
  const g = (await s.page.locator('.qa-grid').boundingBox())!;
  const c = await s.center(row, 2);
  expect(c.y).toBeGreaterThan(g.y);
  expect(c.y).toBeLessThan(g.y + g.height);
});

test('arrow keys and Page Down step over hidden rows onto the next visible one', async ({ launch }) => {
  const s = await launch();
  await s.headerMenu('row', 0);
  await s.page.locator('.qa-ctx-hide-rows').click();
  await s.selectCell('A2');
  await s.page.keyboard.press('ArrowUp');
  await expect(s.page.locator('.qa-name-box')).toHaveValue('A2');
  const n = await s.model<number>('(app) => app.grid.visibleRows()');
  await s.model('(app, i) => app.doc.setHidden(app.si, \'row\', [i], true)', n + 1);
  await s.selectCell('A2');
  await s.page.keyboard.press('PageDown');
  await expect(s.page.locator('.qa-name-box')).toHaveValue(`A${n + 3}`);
});

test('clicking anywhere in a merged block makes its top-left cell active', async ({ launch }) => {
  const s = await launch();
  await s.selectCell('A1:C3');
  await s.page.locator('.qa-tb-merge').click();
  const b = await s.model<{ right: number; bottom: number }>('(app) => { const r = app.grid.cellRect(0, 0); return { right: r.right, bottom: r.bottom }; }');
  await s.page.mouse.click(b.right - 5, b.bottom - 5);
  await expect(s.page.locator('.qa-name-box')).toHaveValue('A1');
  await s.page.keyboard.type('hello');
  await s.page.keyboard.press('Enter');
  expect(await s.value('A1')).toBe('hello');
  await expect(await s.cellAt('A1')).toHaveText('hello');
});

test('dragging up into a frozen row while scrolled scrolls back instead of jumping to the frozen row', async ({ launch }) => {
  const s = await launch();
  await s.command('freeze-row');
  await s.wheel(0, 4000);
  const scrollTop = () => s.page.locator('.qa-grid-scroll').evaluate((el) => el.scrollTop);
  const start = await scrollTop();
  const row = await s.middleRow();
  const from = await s.center(row, 1), frozen = await s.center(0, 1);
  await s.page.mouse.move(from.x, from.y);
  await s.page.mouse.down();
  await s.page.mouse.move(frozen.x, frozen.y, { steps: 5 });
  await expect.poll(scrollTop).toBeLessThan(start - 100);
  await s.page.mouse.up();
  const sel = await s.model<{ r1: number; r2: number }>('(app) => app.grid.sel.primary()');
  expect(sel.r1).toBeGreaterThan(0);
  expect(sel.r2).toBe(row);
});

test('filter buttons stay inside their cell when zoomed out', async ({ dir, launch }) => {
  const s = await launch([Spreadsheet.write(dir, 'people.csv', 'name,age\nAda,36\nAlan,41\n')]);
  await s.page.locator('.qa-tb-auto-table').click();
  for (let i = 0; i < 5; i++) await s.command('zoom-out');
  await expect(s.page.locator('.qa-zoom')).toHaveText('50%');
  const btn = (await s.page.locator('.qa-filter-btn-0').boundingBox())!;
  const cell = await s.model<{ left: number; top: number; right: number; bottom: number }>('(app) => app.grid.cellRect(0, 0).toJSON()');
  expect(btn.x + btn.width).toBeLessThanOrEqual(cell.right + 0.5);
  expect(btn.y + btn.height).toBeLessThanOrEqual(cell.bottom + 0.5);
  expect(btn.x).toBeGreaterThanOrEqual(cell.left);
});

test('dragging the top edge of the first row header after scrolling does not resize a row scrolled out of view', async ({ launch }) => {
  const s = await launch();
  const h = await s.model<number>('(app) => app.grid.cellRect(0, 0).height');
  await s.page.locator('.qa-grid-scroll').evaluate((el, top) => el.scrollTo({ top, behavior: 'instant' }), 190 * h);
  const top = await s.model<number>('(app) => app.grid.cellRect(190, 0).top');
  const g = (await s.page.locator('.qa-grid').boundingBox())!;
  await s.page.mouse.move(g.x + 10, top + 1);
  await s.page.mouse.down();
  await s.page.mouse.move(g.x + 10, top + 31, { steps: 5 });
  await s.page.mouse.up();
  expect(await s.model<number>('(app) => app.grid.cellRect(189, 0).height')).toBe(h);
});

test('freezing at a cell scrolled far down is refused and leaves the freeze alone', async ({ launch }) => {
  const s = await launch();
  await s.wheel(0, 4000);
  const row = await s.middleRow();
  const at = await s.center(row, 1);
  await s.page.mouse.click(at.x, at.y);
  await s.command('freeze-selection');
  await expect(s.page.locator('.qa-toast.error')).toBeVisible();
  expect(await s.model<{ rows: number; cols: number }>('(app) => app.sheet.freeze')).toEqual({ rows: 0, cols: 0 });
});

test('freezing at a cell near the top-left still works', async ({ launch }) => {
  const s = await launch();
  const at = await s.center(2, 1);
  await s.page.mouse.click(at.x, at.y);
  await s.command('freeze-selection');
  await expect.poll(() => s.model<{ rows: number; cols: number }>('(app) => app.sheet.freeze')).toEqual({ rows: 2, cols: 1 });
});
