import { join } from 'node:path';
import { expect, readBack, Sheeyit, test } from './harness';

const PEOPLE = 'name,team,score\nAna,red,90\nBen,blue,72\nCid,red,85\nDee,green,60\n';

test('Auto table bolds the header row, freezes it and adds a filter', async ({ dir, launch }) => {
  const s = await launch([Sheeyit.write(dir, 'people.csv', PEOPLE)]);
  await s.page.locator('.qa-tb-auto-table').click();
  await expect(s.page.locator('.qa-filter-btn')).toHaveCount(3);
  const state = await s.model<{ bold: boolean[]; freeze: unknown; filter: unknown }>(`(app) => ({
    bold: [0, 1, 2].map((c) => app.grid.style(app.sheet.rows[0].cells[c].s).bold),
    freeze: app.sheet.freeze,
    filter: app.sheet.filter.range,
  })`);
  expect(state.bold).toEqual([true, true, true]);
  expect(state.freeze).toEqual({ rows: 1, cols: 0 });
  expect(state.filter).toEqual({ r1: 0, c1: 0, r2: 4, c2: 2 });
  await s.selectCell('A1');
  await expect(s.page.locator('.qa-tb-bold')).toHaveClass(/active/);
});

test('filtering by values hides the unchecked rows, and clearing shows them again', async ({ dir, launch }) => {
  const s = await launch([Sheeyit.write(dir, 'people.csv', PEOPLE)]);
  await s.page.locator('.qa-tb-auto-table').click();
  await s.openFilter(1);
  const menu = s.page.locator('.qa-filter-menu');
  await expect(menu.locator('.qa-filter-value')).toHaveCount(3);
  await menu.locator('.qa-filter-clear').click();
  await menu.locator('.qa-filter-value input[data-value="red"]').check();
  await menu.locator('.qa-filter-ok').click();
  await expect(s.page.locator('.qa-filter-btn-1')).toHaveClass(/on/);
  await expect(await s.cellAt('A2')).toHaveText('Ana');
  await expect(await s.cellAt('A3')).toHaveCount(0);
  await expect(await s.cellAt('A4')).toHaveText('Cid');
  await expect(await s.cellAt('A5')).toHaveCount(0);
  await s.command('clear-filters');
  await expect(await s.cellAt('A3')).toHaveText('Ben');
});

test('filtering by condition keeps only matching rows', async ({ dir, launch }) => {
  const s = await launch([Sheeyit.write(dir, 'people.csv', PEOPLE)]);
  await s.page.locator('.qa-tb-auto-table').click();
  await s.openFilter(2);
  const menu = s.page.locator('.qa-filter-menu');
  await menu.locator('.qa-filter-op').selectOption('gte');
  await menu.locator('.qa-filter-arg-a').fill('80');
  await menu.locator('.qa-filter-ok').click();
  const shown = await s.model<string[]>('(app) => app.sheet.rows.slice(1).filter((r) => !r.filtered).map((r) => r.cells[0].v)');
  expect(shown).toEqual(['Ana', 'Cid']);
});

test('sorting from the filter menu reorders the rows under the header', async ({ dir, launch }) => {
  const s = await launch([Sheeyit.write(dir, 'people.csv', PEOPLE)]);
  await s.page.locator('.qa-tb-auto-table').click();
  await s.openFilter(2);
  await s.page.locator('.qa-filter-menu .qa-sort-desc').click();
  const names = await s.model<string[]>('(app) => app.sheet.rows.map((r) => r.cells[0].v)');
  expect(names).toEqual(['name', 'Ana', 'Cid', 'Ben', 'Dee']);
});

test('a filter saved to XLSX reads back with its range and criteria', async ({ dir, launch }) => {
  const s = await launch();
  await s.enter('A1', 'k');
  await s.enter('A2', 'x');
  await s.enter('A3', 'y');
  await s.page.locator('.qa-tb-auto-table').click();
  await s.openFilter(0);
  await s.page.locator('.qa-filter-menu .qa-filter-value input[data-value="y"]').uncheck();
  await s.page.locator('.qa-filter-menu .qa-filter-ok').click();
  const out = join(dir, 'filtered.xlsx');
  await s.answerSave(out);
  await s.save();
  const sheet = readBack(out).sheets[0];
  expect(sheet.filter!.range).toEqual({ r1: 0, c1: 0, r2: 2, c2: 0 });
  expect(sheet.filter!.columns[0]!.values).toEqual(['x']);
  expect(sheet.rows[2]!.filtered).toBe(true);
  expect(sheet.freeze).toEqual({ rows: 1, cols: 0 });
});

test('bolding part of a cell’s text makes rich text that survives saving', async ({ dir, launch }) => {
  const s = await launch();
  await s.enter('A1', 'hello world');
  await s.selectCell('A1');
  await s.page.keyboard.press('F2');
  for (let i = 0; i < 5; i++) await s.page.keyboard.press('Shift+ArrowLeft');
  await s.page.locator('.qa-tb-bold').click();
  await s.page.keyboard.press('End');
  await s.page.keyboard.press('Enter');
  const runs = await s.model<{ text: string; bold: boolean }[]>('(app) => app.sheet.rows[0].cells[0].rich.map((r) => ({ text: r.text, bold: !!r.font?.bold }))');
  expect(runs).toEqual([{ text: 'hello ', bold: false }, { text: 'world', bold: true }]);
  await expect((await s.cellAt('A1')).locator('b, [style*="bold"], [style*="700"]')).toHaveText('world');

  const out = join(dir, 'rich.xlsx');
  await s.answerSave(out);
  await s.save();
  const cell = readBack(out).sheets[0].rows[0]!.cells[0]!;
  expect(cell.v).toBe('hello world');
  expect(cell.rich!.map((r) => [r.text, !!r.font?.bold])).toEqual([['hello ', false], ['world', true]]);
});

test('formatting the whole cell from the toolbar applies to the selection', async ({ launch }) => {
  const s = await launch();
  await s.enter('A1', '1234.5');
  await s.enter('A2', '0.25');
  await s.selectCell('A1');
  await s.page.locator('.qa-tb-currency').click();
  await expect(await s.cellAt('A1')).toHaveText('$1,234.50');
  await s.selectCell('A2');
  await s.page.locator('.qa-tb-percent').click();
  await expect(await s.cellAt('A2')).toHaveText('25.00%');
  await s.page.locator('.qa-tb-dec-less').click();
  await expect(await s.cellAt('A2')).toHaveText('25.0%');
  await s.selectCell('A1:A2');
  await s.page.locator('.qa-tb-italic').click();
  const italic = await s.model<boolean[]>('(app) => [0, 1].map((r) => app.grid.style(app.sheet.rows[r].cells[0].s).italic)');
  expect(italic).toEqual([true, true]);
});

test('links can be added to a cell from the link prompt', async ({ launch }) => {
  const s = await launch();
  await s.enter('A1', 'site');
  await s.selectCell('A1');
  await s.page.locator('.qa-tb-link').click();
  await s.page.locator('.qa-prompt-input').fill('https://example.com');
  await s.page.locator('.qa-prompt-ok').click();
  await expect(s.page.locator('.qa-link-chip .qa-link-open')).toHaveText('https://example.com');
  await expect((await s.cellAt('A1')).locator('.lnk')).toHaveText('site');
});


test('find and replace changes every match on the sheet', async ({ dir, launch }) => {
  const s = await launch([Sheeyit.write(dir, 'people.csv', PEOPLE)]);
  await s.command('replace');
  await s.page.locator('.qa-find-input').fill('red');
  await expect(s.page.locator('.qa-find-count')).toHaveText('1 of 2');
  await s.page.locator('.qa-replace-input').fill('crimson');
  await s.page.locator('.qa-replace-all').click();
  await expect(await s.cellAt('B2')).toHaveText('crimson');
  await expect(await s.cellAt('B4')).toHaveText('crimson');
  await s.command('undo');
  await expect(await s.cellAt('B2')).toHaveText('red');
});

test('sheets can be added, renamed and deleted from the tabs', async ({ launch }) => {
  const s = await launch();
  await s.page.locator('.qa-add-sheet').click();
  await expect(s.page.locator('.qa-tab')).toHaveText(['Sheet1', 'Sheet2']);
  await expect(s.page.locator('.qa-tab.active')).toHaveText('Sheet2');
  await s.page.locator('.qa-tab.active').dblclick();
  await s.page.locator('.qa-prompt-input').fill('Totals');
  await s.page.locator('.qa-prompt-ok').click();
  await expect(s.page.locator('.qa-tab')).toHaveText(['Sheet1', 'Totals']);
  await s.enter('A1', '=Sheet1!A1+1');
  await s.page.locator('.qa-tab', { hasText: 'Sheet1' }).click();
  await s.enter('A1', '41');
  await s.page.locator('.qa-tab', { hasText: 'Totals' }).click();
  await expect(await s.cellAt('A1')).toHaveText('42');
  await s.page.locator('.qa-tab', { hasText: 'Totals' }).click({ button: 'right' });
  await s.answerMessage(0);
  await s.page.locator('.qa-tab-delete').click();
  await expect(s.page.locator('.qa-tab')).toHaveText(['Sheet1']);
});

test('inserting and deleting rows shifts cells and formulas', async ({ launch }) => {
  const s = await launch();
  await s.enter('A1', '1');
  await s.enter('A2', '2');
  await s.enter('A3', '=SUM(A1:A2)');
  await s.selectCell('A2');
  await s.command('insert-row-above');
  await expect(await s.cellAt('A3')).toHaveText('2');
  await s.selectCell('A4');
  expect(await s.formulaBar()).toBe('=SUM(A1:A3)');
  await s.selectCell('A1');
  await s.command('delete-rows');
  await s.selectCell('A3');
  expect(await s.formulaBar()).toBe('=SUM(A1:A2)');
  await expect(await s.cellAt('A3')).toHaveText('2');
});

test('typing a formula and clicking a cell inserts a reference to it', async ({ launch }) => {
  const s = await launch();
  await s.enter('C3', '5');
  await s.selectCell('A1');
  await s.page.keyboard.type('=1+');
  const rect = await s.model<{ x: number; y: number }>('(app) => { const r = app.grid.cellRect(2, 2); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; }');
  await s.page.mouse.click(rect.x, rect.y);
  await s.page.keyboard.press('Enter');
  await s.selectCell('A1');
  expect(await s.formulaBar()).toBe('=1+C3');
  await expect(await s.cellAt('A1')).toHaveText('6');
});

test('dragging the fill handle continues a number series', async ({ launch }) => {
  const s = await launch();
  await s.enter('A1', '2');
  await s.enter('A2', '4');
  await s.selectCell('A1:A2');
  const handle = await s.page.locator('.qa-fill-handle').boundingBox();
  const target = await s.model<{ x: number; y: number }>('(app) => { const r = app.grid.cellRect(4, 0); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; }');
  await s.page.mouse.move(handle!.x + 3, handle!.y + 3);
  await s.page.mouse.down();
  await s.page.mouse.move(target.x, target.y, { steps: 5 });
  await s.page.mouse.up();
  expect([await s.value('A3'), await s.value('A4'), await s.value('A5')]).toEqual([6, 8, 10]);
});

test('the theme follows the system’s dark mode', async ({ launch }) => {
  const s = await launch();
  await s.enter('A1', 'text');
  const bg = () => s.page.evaluate(() => getComputedStyle(document.body).backgroundColor);
  const fg = async () => (await s.cellAt('A1')).evaluate((el) => getComputedStyle(el.querySelector('.inner')!).color);
  await s.page.emulateMedia({ colorScheme: 'light' });
  const lightBg = await bg();
  const lightFg = await fg();
  await s.page.emulateMedia({ colorScheme: 'dark' });
  await expect.poll(bg).not.toBe(lightBg);
  await expect.poll(fg).not.toBe(lightFg);
  const lum = (c: string) => c.match(/\d+/g)!.slice(0, 3).map(Number).reduce((a, b) => a + b, 0);
  expect(lum(await bg())).toBeLessThan(lum(lightBg));
});

test('zoom changes cell size and shows the level', async ({ launch }) => {
  const s = await launch();
  const width = () => s.model<number>('(app) => app.grid.cellRect(0, 0).width');
  const w1 = await width();
  await s.command('zoom-in');
  await expect(s.page.locator('.qa-zoom')).toHaveText('110%');
  expect(await width()).toBeGreaterThan(w1);
  await s.command('zoom-reset');
  await expect(s.page.locator('.qa-zoom')).toHaveText('100%');
});
