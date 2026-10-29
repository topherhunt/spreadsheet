import { strFromU8, unzipSync } from 'fflate';
import { readFileSync } from 'node:fs';
import { expect, readBack, Sheeyit, test } from './harness';

function parts(path: string): Record<string, Uint8Array> {
  return unzipSync(new Uint8Array(readFileSync(path)));
}

test('a workbook with charts, pivots and conditional formatting keeps them all after an edit', async ({ dir, launch }) => {
  const file = Sheeyit.fixture(dir, 'WithChartSheet.xlsx');
  const before = parts(file);
  const s = await launch([file]);
  await expect(s.page.locator('.qa-banner')).toContainText('Pivot tables');
  await expect(s.page.locator('.qa-banner')).toContainText('Chart sheet');
  await s.enter('F2', 'note');
  await s.save();
  const after = parts(file);
  const kept = Object.keys(before).filter((n) => /chart|pivot|drawing/i.test(n));
  expect(kept.length).toBeGreaterThan(3);
  for (const name of kept) expect(after[name], name).toEqual(before[name]);
  expect(readBack(file).sheets[0].rows[1]!.cells[5]!.v).toBe('note');
});

test('conditional formatting and data validation are listed in the banner and kept on save', async ({ dir, launch }) => {
  const file = Sheeyit.fixture(dir, 'ConditionalFormattingSamples.xlsx');
  const s = await launch([file]);
  const banner = s.page.locator('.qa-banner');
  await expect(banner).toContainText('Conditional formatting');
  await expect(banner).toContainText('Data validation');
  await s.page.locator('.qa-tab', { hasText: /^Products1$/ }).click();
  await s.enter('F3', '1');
  await s.save();
  const sheetXml = Object.entries(parts(file)).filter(([n]) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)).map(([, b]) => strFromU8(b));
  expect(sheetXml.filter((x) => x.includes('<conditionalFormatting')).length).toBeGreaterThan(10);
  await s.page.locator('.qa-banner-close').click();
  await expect(banner).toBeHidden();
});

test('a chart sheet shows a placeholder tab instead of a grid', async ({ dir, launch }) => {
  const s = await launch([Sheeyit.fixture(dir, 'chart_sheet.xlsx')]);
  await expect(s.page.locator('.qa-tab.active')).toHaveText('Chart1');
  await expect(s.page.locator('.qa-other-sheet')).toBeVisible();
  await expect(s.page.locator('.qa-grid')).toBeHidden();
  await s.page.locator('.qa-tab', { hasText: 'Sheet1' }).click();
  await expect(s.page.locator('.qa-grid')).toBeVisible();
  await expect(await s.cellAt('A1')).toHaveText('soap');
});

test('hidden sheets get no tab', async ({ dir, launch }) => {
  const s = await launch([Sheeyit.fixture(dir, 'TwoSheetsOneHidden.xlsx')]);
  await expect(s.page.locator('.qa-tab')).toHaveText(['Sheet2']);
  await expect(await s.cellAt('A1')).toHaveText('Sheet2A1');
});

test('a hyperlinked cell shows its link, and the link can be removed', async ({ dir, launch }) => {
  const s = await launch([Sheeyit.fixture(dir, 'sharedhyperlink.xlsx')]);
  await s.selectCell('A1');
  const chip = s.page.locator('.qa-link-chip');
  await expect(chip).toBeVisible();
  await expect(chip.locator('.qa-link-open')).toHaveText('http://www.apache.org/');
  await chip.locator('.qa-link-remove').click();
  await expect(chip).toBeHidden();
  expect(await s.model('(app) => app.sheet.rows[0].cells[0].link')).toBeUndefined();
});

test('a sheet with 50,000 merged ranges opens and scrolls', async ({ dir, launch }) => {
  const s = await launch([Sheeyit.fixture(dir, '57893-many-merges.xlsx')]);
  await s.selectCell('A5000');
  await s.page.keyboard.press('ArrowDown');
  await expect(s.page.locator('.qa-name-box')).toHaveValue(/^A500\d$/);
});

test('XLOOKUP formulas from Excel compute', async ({ dir, launch }) => {
  const s = await launch([Sheeyit.fixture(dir, 'xlookup.xlsx')]);
  const mismatches = await s.model<string[]>(`(app) => {
    const out = [];
    app.doc.wb.sheets.forEach((sh, si) => sh.rows.forEach((r, ri) => r && r.cells.forEach((c, ci) => {
      if (c && c.f && /XLOOKUP/i.test(c.f) && c.v && typeof c.v === 'object') out.push(si + ':' + ri + ':' + ci + ' ' + c.f + ' -> ' + JSON.stringify(c.v));
    })));
    return out;
  }`);
  expect(mismatches).toEqual([]);
});

test('a hostile file is refused with a message and leaves an untitled workbook', async ({ dir, launch }) => {
  const s = await launch([Sheeyit.fixture(dir, 'poc-xmlbomb.xlsx')]);
  await expect(s.page.locator('.qa-toast')).toContainText('Couldn’t open poc-xmlbomb.xlsx');
  await expect(s.page.locator('.qa-doc-title')).toHaveText('Untitled');
});
