import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { expect, readBack, Sheeyit, test } from './harness';

test('a new window shows an empty untitled XLSX workbook', async ({ launch }) => {
  const s = await launch();
  await expect(s.page.locator('.qa-doc-title')).toHaveText('Untitled');
  await expect(s.page.locator('.qa-format-badge')).toHaveText('XLSX');
  await expect(s.page.locator('.qa-tab')).toHaveText(['Sheet1']);
  await expect(s.page.locator('.qa-name-box')).toHaveValue('A1');
  await expect(s.page.locator('.qa-dirty')).toBeHidden();
});

test('typing values and formulas computes results and marks the doc dirty', async ({ launch }) => {
  const s = await launch();
  await s.enter('A1', '2');
  await s.enter('A2', '3');
  await s.enter('A3', '=A1*A2+1');
  await expect(await s.cellAt('A3')).toHaveText('7');
  await expect(s.page.locator('.qa-dirty')).toBeVisible();
  await s.selectCell('A3');
  expect(await s.formulaBar()).toBe('=A1*A2+1');
  await s.enter('A1', '10');
  await expect(await s.cellAt('A3')).toHaveText('31');
});

test('undo and redo step through edits', async ({ launch }) => {
  const s = await launch();
  await s.enter('B2', 'hello');
  await s.enter('B2', 'world');
  await expect(await s.cellAt('B2')).toHaveText('world');
  await s.command('undo');
  await expect(await s.cellAt('B2')).toHaveText('hello');
  await s.command('undo');
  await expect(await s.cellAt('B2')).toHaveCount(0);
  await s.command('redo');
  await expect(await s.cellAt('B2')).toHaveText('hello');
});

test('Escape cancels an edit without changing the cell', async ({ launch }) => {
  const s = await launch();
  await s.enter('A1', 'keep');
  await s.selectCell('A1');
  await s.page.keyboard.type('discard');
  await expect(s.page.locator('.qa-editor')).toBeVisible();
  await s.page.keyboard.press('Escape');
  await expect(s.page.locator('.qa-editor')).toBeHidden();
  await expect(await s.cellAt('A1')).toHaveText('keep');
});

test('arrow keys, Tab and Enter move the active cell', async ({ launch }) => {
  const s = await launch();
  const box = s.page.locator('.qa-name-box');
  await s.selectCell('B2');
  await s.page.keyboard.press('ArrowDown');
  await expect(box).toHaveValue('B3');
  await s.page.keyboard.press('ArrowRight');
  await expect(box).toHaveValue('C3');
  await s.page.keyboard.press('Tab');
  await expect(box).toHaveValue('D3');
  await s.page.keyboard.type('x');
  await s.page.keyboard.press('Enter');
  await expect(box).toHaveValue('D4');
  await s.page.keyboard.press('Shift+ArrowUp');
  await expect(box).toHaveValue('D3:D4');
});

test('Delete clears the selected cells', async ({ launch }) => {
  const s = await launch();
  await s.enter('A1', 'one');
  await s.enter('A2', 'two');
  await s.selectCell('A1:A2');
  await s.page.keyboard.press('Delete');
  expect(await s.value('A1')).toBeUndefined();
  expect(await s.value('A2')).toBeUndefined();
});

test('the status bar sums the selected numbers', async ({ dir, launch }) => {
  const s = await launch([Sheeyit.write(dir, 'nums.csv', 'a\n1\n2\n3.5\n')]);
  await s.selectCell('A1:A4');
  await expect(s.page.locator('.qa-status')).toHaveText(/Sum: 6\.5\s+Average: 2\.1666666667\s+Count: 4/);
});

test('saving a new workbook asks for a path and writes a readable XLSX', async ({ dir, launch }) => {
  const s = await launch();
  await s.enter('A1', 'name');
  await s.enter('B1', '=LEN(A1)');
  const out = join(dir, 'new.xlsx');
  await s.answerSave(out);
  await s.save();
  await expect(s.page.locator('.qa-doc-title')).toHaveText('new.xlsx');
  const wb = readBack(out);
  expect(wb.sheets[0].rows[0]!.cells[0]!.v).toBe('name');
  expect(wb.sheets[0].rows[0]!.cells[1]!.f).toBe('LEN(A1)');
  expect(wb.sheets[0].rows[0]!.cells[1]!.v).toBe(4);
  expect(readFileSync(out).subarray(0, 2).toString()).toBe('PK');
});

test('the name box jumps to a range and to another sheet', async ({ dir, launch }) => {
  const s = await launch([Sheeyit.fixture(dir, '55906-MultiSheetRefs.xlsx')]);
  const names = await s.page.locator('.qa-tab').allTextContents();
  expect(names.length).toBeGreaterThan(1);
  await s.page.locator('.qa-name-box').fill(`'${names[1]}'!C5`);
  await s.page.locator('.qa-name-box').press('Enter');
  await expect(s.page.locator('.qa-tab.active')).toHaveText(names[1]);
  await expect(s.page.locator('.qa-name-box')).toHaveValue('C5');
});

test('closing a window with unsaved changes asks, and Save writes the file first', async ({ dir, launch }) => {
  const file = Sheeyit.write(dir, 'close.csv', 'a,b\n1,2\n');
  const s = await launch([file]);
  await s.enter('A2', '42');
  await s.answerMessage(0);
  const closed = s.app.waitForEvent('close');
  await s.app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].close());
  await closed;
  expect(readFileSync(file, 'utf8')).toBe('a,b\n42,2\n');
});
