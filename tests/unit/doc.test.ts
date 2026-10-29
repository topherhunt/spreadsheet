import { describe, expect, it } from 'vitest';
import { MAX_ROWS } from '../../src/core/addr';
import { newWorkbook, readDelimited, writeDelimited } from '../../src/core/csv';
import { Doc } from '../../src/core/doc';
import { columnValues, compareValues, dataRegion } from '../../src/core/filter';
import { parseTyped } from '../../src/core/input';
import { getCell, Workbook } from '../../src/core/model';
import { cellFormat, displayCell } from '../../src/core/numfmt';
import { shiftSpan } from '../../src/core/structure';
import { readXlsx } from '../../src/core/xlsx/read';
import { writeXlsx } from '../../src/core/xlsx/write';

const v = (wb: Workbook, a1: [number, number], sheet = 0) => getCell(wb.sheets[sheet], a1[0], a1[1])?.v ?? null;
const f = (wb: Workbook, a1: [number, number], sheet = 0) => getCell(wb.sheets[sheet], a1[0], a1[1])?.f;

function report(): Doc {
  return new Doc(readDelimited('Name,Region,Sales\nAnn,East,10\nBob,West,30\nCat,East,20\nDan,North,5\n', 'csv'));
}

describe('typed input', () => {
  it('recognizes dates, times, currency, percents and grouped numbers with matching formats', () => {
    expect(parseTyped('2024-01-15')).toEqual({ v: 45306, fmt: 'm/d/yyyy' });
    expect(parseTyped('1/15/2024')).toEqual({ v: 45306, fmt: 'm/d/yyyy' });
    expect(parseTyped('Jan 15, 2024')).toEqual({ v: 45306, fmt: 'm/d/yyyy' });
    expect(parseTyped('12:30')).toEqual({ v: 0.5208333333333334, fmt: 'h:mm' });
    expect(parseTyped('1:00 PM').v).toBeCloseTo(13 / 24);
    expect(parseTyped('$1,234.5')).toEqual({ v: 1234.5, fmt: '"$"#,##0.0' });
    expect(parseTyped('$12')).toEqual({ v: 12, fmt: '"$"#,##0.00' });
    expect(parseTyped('15%')).toEqual({ v: 0.15, fmt: '0%' });
    expect(parseTyped('1,000')).toEqual({ v: 1000, fmt: '#,##0' });
    expect(parseTyped('42')).toEqual({ v: 42 });
  });

  it('keeps invalid dates, apostrophe-prefixed text and formulas apart', () => {
    expect(parseTyped('2/30/2024')).toEqual({ v: '2/30/2024' });
    expect(parseTyped("'123")).toEqual({ v: '123' });
    expect(parseTyped('=SUM(A1:A3)')).toEqual({ v: null, formula: 'SUM(A1:A3)' });
    expect(parseTyped('=')).toEqual({ v: '=' });
  });
});

describe('structure helpers', () => {
  it('shifts and truncates spans on insert and delete', () => {
    expect(shiftSpan(2, 5, 3, 2)).toEqual([2, 7]);
    expect(shiftSpan(2, 5, 0, -1)).toEqual([1, 4]);
    expect(shiftSpan(2, 5, 2, -4)).toBeNull();
    expect(shiftSpan(2, 5, 4, -5)).toEqual([2, 3]);
  });
});

describe('Doc editing', () => {
  it('typed input recalculates formulas and undo/redo restores both', () => {
    const d = new Doc(newWorkbook());
    d.setInput(0, 0, 0, '2');
    d.setInput(0, 0, 1, '=A1*3');
    expect(v(d.wb, [0, 1])).toBe(6);
    d.setInput(0, 0, 0, '5');
    expect(v(d.wb, [0, 1])).toBe(15);
    d.undo();
    expect(v(d.wb, [0, 0])).toBe(2);
    expect(v(d.wb, [0, 1])).toBe(6);
    d.redo();
    expect(v(d.wb, [0, 1])).toBe(15);
    expect(d.dirty).toBe(true);
  });

  it('typed dates and date formulas get a date format', () => {
    const d = new Doc(newWorkbook());
    d.setInput(0, 0, 0, '2024-03-01');
    expect(displayCell(d.wb, getCell(d.wb.sheets[0], 0, 0)).text).toBe('3/1/2024');
    d.setInput(0, 0, 1, '=A1+1');
    expect(displayCell(d.wb, getCell(d.wb.sheets[0], 0, 1)).text).toBe('3/2/2024');
    d.setInput(0, 1, 0, '=DATE(2020,1,1)');
    expect(cellFormat(d.wb, getCell(d.wb.sheets[0], 1, 0))).toBe('m/d/yyyy');
    d.undo(); d.redo();
    expect(cellFormat(d.wb, getCell(d.wb.sheets[0], 1, 0))).toBe('m/d/yyyy');
  });

  it('fill down translates relative references', () => {
    const d = new Doc(newWorkbook());
    d.pasteText(0, 0, 0, [['1', '=A1*2'], ['2', ''], ['3', '']]);
    d.fill(0, { r1: 0, c1: 1, r2: 2, c2: 1 }, 'down');
    expect(f(d.wb, [2, 1])).toBe('A3*2');
    expect(v(d.wb, [2, 1])).toBe(6);
  });

  it('drag-fill continues numeric series, repeats text and translates formulas, in both directions', () => {
    const d = new Doc(newWorkbook());
    d.pasteText(0, 2, 0, [['1', 'a', '=A3*10'], ['3', 'b', '=A4*10']]);
    d.fillSeries(0, { r1: 2, c1: 0, r2: 3, c2: 2 }, { r1: 2, c1: 0, r2: 6, c2: 2 });
    expect([4, 5, 6].map((r) => v(d.wb, [r, 0]))).toEqual([5, 7, 9]);
    expect([4, 5, 6].map((r) => v(d.wb, [r, 1]))).toEqual(['a', 'b', 'a']);
    expect(f(d.wb, [5, 2])).toBe('A6*10');
    expect(v(d.wb, [5, 2])).toBe(70);
    d.fillSeries(0, { r1: 2, c1: 0, r2: 3, c2: 0 }, { r1: 0, c1: 0, r2: 3, c2: 0 });
    expect([0, 1].map((r) => v(d.wb, [r, 0]))).toEqual([-3, -1]);
  });

  it('moving cells keeps formula references, clears the source, and undoes in one step', () => {
    const d = new Doc(newWorkbook());
    d.pasteText(0, 0, 0, [['4', '=A1+1']]);
    d.moveCells(0, { r1: 0, c1: 1, r2: 0, c2: 1 }, 3, 3);
    expect(f(d.wb, [3, 3])).toBe('A1+1');
    expect(v(d.wb, [3, 3])).toBe(5);
    expect(getCell(d.wb.sheets[0], 0, 1)).toBeUndefined();
    d.undo();
    expect(f(d.wb, [0, 1])).toBe('A1+1');
    expect(getCell(d.wb.sheets[0], 3, 3)).toBeUndefined();
  });

  it('moving cells rewrites references to them on every sheet, and undo restores those formulas', () => {
    const d = new Doc(newWorkbook());
    d.pasteText(0, 0, 0, [['4', '=A1+1', '=SUM(A1:B1)', '=$A$1:A2']]);
    d.addSheet(1);
    d.setInput(1, 0, 0, '=Sheet1!A1*2');
    d.moveCells(0, { r1: 0, c1: 0, r2: 0, c2: 1 }, 4, 0);
    expect(f(d.wb, [4, 1])).toBe('A5+1');
    expect(v(d.wb, [4, 1])).toBe(5);
    expect(f(d.wb, [0, 2])).toBe('SUM(A5:B5)');
    expect(v(d.wb, [0, 2])).toBe(9);
    expect(f(d.wb, [0, 3])).toBe('$A$1:A2');
    expect(f(d.wb, [0, 0], 1)).toBe('Sheet1!A5*2');
    expect(v(d.wb, [0, 0], 1)).toBe(8);
    d.undo();
    expect(f(d.wb, [0, 2])).toBe('SUM(A1:B1)');
    expect(f(d.wb, [0, 0], 1)).toBe('Sheet1!A1*2');
    expect(v(d.wb, [0, 1])).toBe(5);
  });

  it('pasting copied cells moves formulas relative to their origin', () => {
    const d = new Doc(newWorkbook());
    d.pasteText(0, 0, 0, [['4', '=A1+1']]);
    const src = [[getCell(d.wb.sheets[0], 0, 1)]];
    d.pasteCells(0, 5, 3, src, { row: 0, col: 1 });
    expect(f(d.wb, [5, 3])).toBe('C6+1');
    d.pasteCells(0, 6, 1, src, { row: 0, col: 1 }, true);
    expect(f(d.wb, [6, 1])).toBeUndefined();
    expect(v(d.wb, [6, 1])).toBe(5);
  });

  it('formatting rich text overrides the runs’ own property', () => {
    const d = new Doc(newWorkbook());
    d.setRich(0, 0, 0, [{ text: 'big', font: { bold: true, size: 20 } }, { text: ' link', link: 'https://example.com' }]);
    const c = getCell(d.wb.sheets[0], 0, 0)!;
    expect(c.v).toBe('big link');
    expect(c.link).toBe('https://example.com');
    d.setStyle(0, [{ r1: 0, c1: 0, r2: 0, c2: 0 }], { bold: false });
    expect(getCell(d.wb.sheets[0], 0, 0)!.rich![0].font).toEqual({ size: 20 });
  });

  it('inserting a row shifts formulas and undo restores them', () => {
    const d = new Doc(newWorkbook());
    d.pasteText(0, 0, 0, [['1'], ['2'], ['=SUM(A1:A2)']]);
    d.insertRows(0, 1, 1);
    expect(f(d.wb, [3, 0])).toBe('SUM(A1:A3)');
    expect(v(d.wb, [3, 0])).toBe(3);
    d.setInput(0, 1, 0, '10');
    expect(v(d.wb, [3, 0])).toBe(13);
    d.undo(); d.undo();
    expect(f(d.wb, [2, 0])).toBe('SUM(A1:A2)');
    expect(v(d.wb, [2, 0])).toBe(3);
  });

  it('bolds, clears and undoes 100,000 cells in a few seconds at most', () => {
    const rows = Array.from({ length: 20000 }, (_, i) => `${i},a${i},${i * 2},b,${i % 7}`).join('\n');
    const d = new Doc(readDelimited(`${rows}\n`, 'csv'));
    d.setInput(0, 0, 6, '=SUM(A1:E20000)');
    let sum = 0;
    for (let i = 0; i < 20000; i++) sum += i + i * 2 + (i % 7);
    expect(v(d.wb, [0, 6])).toBe(sum);
    const all = [{ r1: 0, c1: 0, r2: 19999, c2: 4 }];
    const timed = (step: () => void) => { const t = performance.now(); step(); expect(performance.now() - t).toBeLessThan(3000); };
    timed(() => d.setStyle(0, all, { bold: true }));
    timed(() => d.clear(0, all));
    expect(v(d.wb, [0, 6])).toBe(0);
    timed(() => d.undo());
    expect(v(d.wb, [0, 6])).toBe(sum);
    expect(v(d.wb, [19999, 0])).toBe(19999);
  });

  it('refuses to insert rows that would push cells off the sheet, and keeps working', () => {
    const d = new Doc(newWorkbook());
    d.setInput(0, MAX_ROWS - 1, 0, 'last');
    expect(() => d.insertRows(0, 1, 1)).toThrow(/off the end of the sheet/);
    expect(v(d.wb, [MAX_ROWS - 1, 0])).toBe('last');
    d.setInput(0, 0, 0, '=1+1');
    expect(v(d.wb, [0, 0])).toBe(2);
  });

  it('inserting rows keeps a range that runs to the last row, and moves internal links', () => {
    const d = new Doc(newWorkbook());
    d.pasteText(0, 0, 0, [['=SUM(B2:B1048576)']]);
    d.setInput(0, 0, 1, 'go');
    d.setLink(0, 0, 1, '#Sheet1!B6000');
    d.insertRows(0, 4, 1);
    expect(f(d.wb, [0, 0])).toBe('SUM(B2:B1048576)');
    expect(getCell(d.wb.sheets[0], 0, 1)!.link).toBe('#Sheet1!B6001');
  });

  it('deleting a referenced column turns references into #REF!', () => {
    const d = new Doc(newWorkbook());
    d.pasteText(0, 0, 0, [['1', '2', '=B1']]);
    d.insertCols(0, 1, -1);
    expect(f(d.wb, [0, 1])).toBe('#REF!');
    expect(v(d.wb, [0, 1])).toEqual({ error: '#REF!' });
  });

  it('renaming a sheet updates formulas that reference it', () => {
    const d = new Doc(newWorkbook());
    d.addSheet();
    d.setInput(1, 0, 0, '7');
    d.setInput(0, 0, 0, '=Sheet2!A1*2');
    d.renameSheet(1, 'Data Set');
    expect(f(d.wb, [0, 0])).toBe("'Data Set'!A1*2");
    expect(v(d.wb, [0, 0])).toBe(14);
    expect(() => d.renameSheet(1, 'Sheet1')).toThrow('Another sheet already has that name.');
  });
});

describe('Doc filtering and sorting', () => {
  it('auto-table bolds the header, freezes it and turns on a filter over the data', () => {
    const d = report();
    d.autoTable(0);
    const s = d.wb.sheets[0];
    expect(s.freeze).toEqual({ rows: 1, cols: 0 });
    expect(s.filter!.range).toEqual({ r1: 0, c1: 0, r2: 4, c2: 2 });
    for (let c = 0; c < 3; c++) expect(d.wb.styles.fonts[d.wb.styles.xfs[getCell(s, 0, c)!.s!].fontId].bold).toBe(true);
    d.undo();
    expect(d.wb.sheets[0].freeze).toEqual({ rows: 0, cols: 0 });
    expect(d.wb.sheets[0].filter).toBeNull();
  });

  it('a value filter hides rows and lists counts respecting other columns', () => {
    const d = report();
    d.toggleFilter(0);
    expect(columnValues(d.wb, d.wb.sheets[0], 1)).toEqual([{ text: 'East', count: 2 }, { text: 'North', count: 1 }, { text: 'West', count: 1 }]);
    d.setColumnFilter(0, 1, { values: ['East'] });
    const hidden = d.wb.sheets[0].rows.map((r, i) => (r?.filtered ? i : -1)).filter((i) => i >= 0);
    expect(hidden).toEqual([2, 4]);
    expect(columnValues(d.wb, d.wb.sheets[0], 0).map((x) => x.text)).toEqual(['Ann', 'Cat']);
    d.undo();
    expect(d.wb.sheets[0].rows.some((r) => r?.filtered)).toBe(false);
  });

  it('a condition filter compares numbers numerically', () => {
    const d = report();
    d.toggleFilter(0);
    d.setColumnFilter(0, 2, { condition: { op: 'gt', a: '9' } });
    expect(d.wb.sheets[0].rows.map((r, i) => (r?.filtered ? i : -1)).filter((i) => i >= 0)).toEqual([4]);
  });

  it('sorting by a filter column keeps the header and moves whole rows', () => {
    const d = report();
    d.toggleFilter(0);
    d.sortFilterColumn(0, 2, true);
    expect([1, 2, 3, 4].map((r) => v(d.wb, [r, 0]))).toEqual(['Bob', 'Cat', 'Ann', 'Dan']);
    expect(v(d.wb, [0, 0])).toBe('Name');
    d.undo();
    expect([1, 2, 3, 4].map((r) => v(d.wb, [r, 0]))).toEqual(['Ann', 'Bob', 'Cat', 'Dan']);
  });

  it('sorting keeps formulas pointing at their own row', () => {
    const d = new Doc(newWorkbook());
    d.pasteText(0, 0, 0, [['3', '=A1*10'], ['1', '=A2*10'], ['2', '=A3*10']]);
    d.sort(0, { r1: 0, c1: 0, r2: 2, c2: 1 }, [{ col: 0, desc: false }], false);
    expect([0, 1, 2].map((r) => v(d.wb, [r, 1]))).toEqual([10, 20, 30]);
  });

  it('orders numbers before text before booleans before blanks', () => {
    expect([null, true, 'b', 2, 'a10', 'a9', 1].sort(compareValues)).toEqual([1, 2, 'a9', 'a10', 'b', true, null]);
  });

  it('finds the data region below leading blank rows', () => {
    const d = new Doc(newWorkbook());
    d.pasteText(0, 2, 1, [['h1', 'h2'], ['1', '2']]);
    expect(dataRegion(d.wb.sheets[0])).toEqual({ r1: 2, c1: 1, r2: 3, c2: 2 });
  });

  it('filter, freeze and bold header survive an XLSX round trip', () => {
    const d = report();
    d.autoTable(0);
    d.setColumnFilter(0, 1, { values: ['West'] });
    const back = readXlsx(writeXlsx(d.wb));
    const s = back.sheets[0];
    expect(s.freeze).toEqual({ rows: 1, cols: 0 });
    expect(s.filter!.columns[1]).toEqual({ values: ['West'] });
    expect(s.rows[1]!.filtered).toBe(true);
    expect(s.rows[2]!.filtered).toBeFalsy();
    expect(back.styles.fonts[back.styles.xfs[getCell(s, 0, 0)!.s!].fontId].bold).toBe(true);
  });
});

describe('delimited files', () => {
  it.each([['csv', 'a, b, 12\n"x,y","say ""hi""",3\n'], ['tsv', 'a\t b \t12\n']] as const)('an untouched %s saves exactly as read', (format, text) => {
    const wb = readDelimited(text, format);
    expect(writeDelimited(wb, format, (c) => displayCell(wb, c).text)).toBe(text);
  });
});
