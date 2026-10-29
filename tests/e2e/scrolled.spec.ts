import { expect, test } from './harness';

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
