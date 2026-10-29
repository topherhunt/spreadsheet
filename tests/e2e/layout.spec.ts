import { join } from 'node:path';
import { expect, readBack, test } from './harness';

test('freezing the top row draws the freeze line and saves the pane', async ({ dir, launch }) => {
  const s = await launch();
  await s.enter('A1', 'head');
  await s.command('freeze-row');
  await expect(s.page.locator('.freeze-line-h')).toBeVisible();
  await expect(s.page.locator('.freeze-line-v')).toBeHidden();
  const out = join(dir, 'freeze.xlsx');
  await s.answerSave(out);
  await s.save();
  expect(readBack(out).sheets[0].freeze).toEqual({ rows: 1, cols: 0 });
  await s.command('freeze-none');
  await expect(s.page.locator('.freeze-line-h')).toBeHidden();
});

test('columns hide and unhide from the header menu, and hidden columns save', async ({ dir, launch }) => {
  const s = await launch();
  await s.enter('B1', 'secret');
  await s.headerMenu('col', 1);
  await s.page.locator('.qa-ctx-hide-cols').click();
  await expect(await s.cellAt('B1')).toBeHidden();
  const out = join(dir, 'hidden.xlsx');
  await s.answerSave(out);
  await s.save();
  expect(readBack(out).sheets[0].cols[1]!.hidden).toBe(true);
  await s.headerMenu('col', 0);
  await s.page.locator('.qa-ctx-unhide-cols').click();
  await expect(await s.cellAt('B1')).toHaveText('secret');
});

test('rows insert above from the row header menu and formulas follow', async ({ launch }) => {
  const s = await launch();
  await s.enter('A1', '3');
  await s.enter('A2', '=A1*2');
  await s.headerMenu('row', 0);
  await s.page.locator('.qa-ctx-row-above').click();
  expect(await s.value('A2')).toBe(3);
  await s.selectCell('A3');
  expect(await s.formulaBar()).toBe('=A2*2');
  await expect(await s.cellAt('A3')).toHaveText('6');
});

test('merging a block shows one cell spanning it and saves the merge', async ({ dir, launch }) => {
  const s = await launch();
  await s.enter('A1', 'title');
  await s.selectCell('A1:C2');
  await s.page.locator('.qa-tb-merge').click();
  const a1 = await (await s.cellAt('A1')).boundingBox();
  const d1 = await s.model<{ width: number; height: number }>('(app) => { const r = app.grid.cellRect(0, 3); return { width: r.width, height: r.height }; }');
  expect(a1!.width).toBeGreaterThan(d1.width * 2.5);
  expect(a1!.height).toBeGreaterThan(d1.height * 1.5);
  const out = join(dir, 'merge.xlsx');
  await s.answerSave(out);
  await s.save();
  expect(readBack(out).sheets[0].merges).toEqual([{ r1: 0, c1: 0, r2: 1, c2: 2 }]);
});

test('outer borders go on the edges of the selection only', async ({ dir, launch }) => {
  const s = await launch();
  await s.enter('A1', '1');
  await s.enter('B2', '2');
  await s.selectCell('A1:B2');
  await s.page.locator('.qa-tb-borders').click();
  await s.page.locator('.qa-border-outer').click();
  const out = join(dir, 'borders.xlsx');
  await s.answerSave(out);
  await s.save();
  const wb = readBack(out);
  const sides = (r: number, c: number) => {
    const b = wb.styles.borders[wb.styles.xfs[wb.sheets[0].rows[r]!.cells[c]!.s!]!.borderId]!;
    return (['top', 'right', 'bottom', 'left'] as const).filter((k) => b[k]?.style);
  };
  expect(sides(0, 0)).toEqual(['top', 'left']);
  expect(sides(1, 1)).toEqual(['right', 'bottom']);
});

test('Auto table on a workbook bolds the header, freezes it, adds a filter, and saves all three', async ({ dir, launch }) => {
  const s = await launch();
  await s.enter('A1', 'Name');
  await s.enter('B1', 'Score');
  await s.enter('A2', 'Ann');
  await s.enter('B2', '9');
  await s.page.locator('.qa-tb-auto-table').click();
  await expect(s.page.locator('.freeze-line-h')).toBeVisible();
  await expect(s.page.locator('.qa-filter-btn-1')).toBeAttached();
  const out = join(dir, 'table.xlsx');
  await s.answerSave(out);
  await s.save();
  const wb = readBack(out);
  const sh = wb.sheets[0];
  expect(sh.freeze.rows).toBe(1);
  expect(sh.filter!.range).toEqual({ r1: 0, c1: 0, r2: 1, c2: 1 });
  for (const c of [0, 1]) expect(wb.styles.fonts[wb.styles.xfs[sh.rows[0]!.cells[c]!.s!]!.fontId]!.bold).toBe(true);
});
