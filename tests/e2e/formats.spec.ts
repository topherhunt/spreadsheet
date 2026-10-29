import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { expect, readBack, Spreadsheet, test } from './harness';

const CSV = 'id,name,amount,when\n007,"Smith, Jo",1200.50,2024-01-15\n008,Lee,99,2024-02-01\n';

test('a CSV opens with its cells and saves edits without rewriting untouched text', async ({ dir, launch }) => {
  const file = Spreadsheet.write(dir, 'people.csv', CSV);
  const s = await launch([file]);
  await expect(s.page.locator('.qa-format-badge')).toHaveText('CSV');
  await expect(s.page.locator('.qa-doc-title')).toHaveText('people.csv');
  await expect(await s.cellAt('B2')).toHaveText('Smith, Jo');
  await s.enter('B3', 'Lee Ann');
  await s.save();
  expect(await s.messages()).toEqual([]);
  expect(readFileSync(file, 'utf8')).toBe('id,name,amount,when\n007,"Smith, Jo",1200.50,2024-01-15\n008,Lee Ann,99,2024-02-01\n');
});

test('a TSV opens and saves back as tab-separated text', async ({ dir, launch }) => {
  const file = Spreadsheet.write(dir, 'data.tsv', 'a\tb\n1\thello, world\n');
  const s = await launch([file]);
  await expect(s.page.locator('.qa-format-badge')).toHaveText('TSV');
  await expect(await s.cellAt('B2')).toHaveText('hello, world');
  await s.enter('A2', '2');
  await s.save();
  expect(readFileSync(file, 'utf8')).toBe('a\tb\n2\thello, world\n');
});

test('saving a CSV with new formatting offers XLSX and writes a workbook next to it', async ({ dir, launch }) => {
  const file = Spreadsheet.write(dir, 'fmt.csv', CSV);
  const s = await launch([file]);
  await s.selectCell('B2');
  await s.page.locator('.qa-tb-bold').click();
  await s.answerMessage(0);
  const out = join(dir, 'fmt.xlsx');
  await s.answerSave(out);
  await s.save();
  const msgs = await s.messages();
  expect(msgs[0].message).toContain('formatting');
  expect(msgs[0].buttons[0]).toBe('Save as XLSX…');
  expect(msgs[1].message).toBe(`save:${out}`);
  await expect(s.page.locator('.qa-format-badge')).toHaveText('XLSX');
  await expect(s.page.locator('.qa-doc-title')).toHaveText('fmt.xlsx');
  expect(readFileSync(file, 'utf8')).toBe(CSV);
  const wb = readBack(out);
  const cell = wb.sheets[0].rows[1]!.cells[1]!;
  expect(cell.v).toBe('Smith, Jo');
  expect(wb.styles.fonts[wb.styles.xfs[cell.s!].fontId].bold).toBe(true);
});

test('saving a CSV with a formula can keep the CSV, writing the computed value', async ({ dir, launch }) => {
  const file = Spreadsheet.write(dir, 'calc.csv', 'a,b\n2,3\n');
  const s = await launch([file]);
  await s.enter('C2', '=A2*B2');
  await s.answerMessage(1);
  await s.save();
  expect((await s.messages())[0].message).toContain('formulas');
  await expect(s.page.locator('.qa-format-badge')).toHaveText('CSV');
  expect(readFileSync(file, 'utf8')).toBe('a,b,\n2,3,6\n');
});

test('cancelling the CSV formatting prompt writes nothing and stays dirty', async ({ dir, launch }) => {
  const file = Spreadsheet.write(dir, 'cancel.csv', CSV);
  const s = await launch([file]);
  await s.selectCell('A2');
  await s.page.locator('.qa-tb-italic').click();
  await s.answerMessage(2);
  await s.command('save');
  await expect.poll(async () => (await s.messages()).length).toBe(1);
  await expect(s.page.locator('.qa-dirty')).toBeVisible();
  expect(readFileSync(file, 'utf8')).toBe(CSV);
  expect(existsSync(join(dir, 'cancel.xlsx'))).toBe(false);
});

test('Auto table on a CSV saves without asking, since it only changes the view', async ({ dir, launch }) => {
  const file = Spreadsheet.write(dir, 'auto.csv', CSV);
  const s = await launch([file]);
  await s.page.locator('.qa-tb-auto-table').click();
  await expect(s.page.locator('.qa-filter-btn')).toHaveCount(4);
  await s.enter('C3', '100');
  await s.save();
  expect(await s.messages()).toEqual([]);
  expect(readFileSync(file, 'utf8')).toContain('008,Lee,100,2024-02-01');
});

test('Save As converts an XLSX to CSV after warning what a CSV drops', async ({ dir, launch }) => {
  const s = await launch([Spreadsheet.fixture(dir, '55906-MultiSheetRefs.xlsx')]);
  const out = join(dir, 'export.csv');
  await s.answerSave(out);
  await s.answerMessage(0);
  await s.command('save-as');
  await expect(s.page.locator('.qa-format-badge')).toHaveText('CSV');
  const msgs = await s.messages();
  expect(msgs[1].message).toContain('multiple sheets');
  expect(readFileSync(out, 'utf8').length).toBeGreaterThan(0);
});

test('a macro-enabled workbook saves as XLSM with its macros, and refuses Save As XLSX', async ({ dir, launch }) => {
  const file = Spreadsheet.fixture(dir, 'SimpleMacro.xlsm');
  const s = await launch([file]);
  await expect(s.page.locator('.qa-format-badge')).toHaveText('XLSM');
  await expect(s.page.locator('.qa-banner')).toContainText('Macros (VBA)');
  await s.enter('H20', 'edited');
  await s.save();
  const wb = readBack(file);
  expect(wb.features.has('Macros (VBA)')).toBe(true);
  expect(wb.sheets[wb.activeSheet].rows[19]!.cells[7]!.v).toBe('edited');

  const out = join(dir, 'nomacros.xlsx');
  await s.answerSave(out);
  await s.answerMessage(0);
  await s.command('save-as');
  await expect.poll(async () => (await s.messages()).length).toBe(2);
  expect((await s.messages())[1].message).toBe('This workbook has macros.');
  expect(existsSync(out)).toBe(false);
});
