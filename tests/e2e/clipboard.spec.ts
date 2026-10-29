import { expect, test } from './harness';

// These share the system clipboard, so they run one at a time in a single worker.
test.describe.configure({ mode: 'default' });

test('copy and paste inside the app moves formulas relative to the new spot', async ({ launch }) => {
  const s = await launch();
  await s.enter('A1', '1');
  await s.enter('B1', '=A1*10');
  await s.selectCell('A1:B1');
  await s.app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].webContents.copy());
  await expect(s.page.locator('.copy-marquee')).toHaveCount(1);
  await s.enter('A3', '5');
  await s.selectCell('A4');
  await s.app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].webContents.paste());
  await expect(await s.cellAt('B4')).toHaveText('10');
  await s.selectCell('B4');
  expect(await s.formulaBar()).toBe('=A4*10');
});

test('cut and paste moves cells and keeps references to them pointing at the new place', async ({ launch }) => {
  const s = await launch();
  await s.enter('A1', '7');
  await s.enter('C1', '=A1+1');
  await s.selectCell('A1');
  await s.app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].webContents.cut());
  await s.selectCell('A5');
  await s.app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].webContents.paste());
  await expect(await s.cellAt('A5')).toHaveText('7');
  expect(await s.value('A1')).toBeUndefined();
  await s.selectCell('C1');
  expect(await s.formulaBar()).toBe('=A5+1');
  await expect(await s.cellAt('C1')).toHaveText('8');
});

test('pasting tab-separated text from another app fills a block of cells', async ({ launch }) => {
  const s = await launch();
  await s.app.evaluate(({ clipboard }) => clipboard.writeText('a\tb\n1\t2\n'));
  await s.selectCell('B2');
  await s.app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].webContents.paste());
  await expect(await s.cellAt('C3')).toHaveText('2');
  expect(await s.value('C3')).toBe(2);
  await expect(s.page.locator('.qa-name-box')).toHaveValue('B2:C3');
});
