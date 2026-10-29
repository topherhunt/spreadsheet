import { describe, expect, it } from 'vitest';
import { newWorkbook } from '../../src/core/csv';
import { Engine } from '../../src/core/formula/engine';
import { displayFormula, fileFormula } from '../../src/core/formula/xlfn';
import { Cell, getCell, setCell, Workbook } from '../../src/core/model';
import { writeXlsx } from '../../src/core/xlsx/write';
import { readXlsx } from '../../src/core/xlsx/read';
import { loadFixture, poiFixtures } from './helpers';

function book(rows: (Cell | null)[][]): Workbook {
  const wb = newWorkbook();
  rows.forEach((r, ri) => r.forEach((c, ci) => { if (c) setCell(wb.sheets[0], ri, ci, c); }));
  return wb;
}
const n = (v: number): Cell => ({ v });
const f = (formula: string): Cell => ({ v: null, f: fileFormula(formula) });

describe('formula engine', () => {
  it('recalculates dependents when a value changes', () => {
    const wb = book([[n(1), n(2), f('A1+B1')], [f('C1*10')]]);
    const e = new Engine(wb);
    e.apply([{ sheet: 0, row: 0, col: 0, cell: n(5) }]);
    expect(getCell(wb.sheets[0], 0, 2)!.v).toBe(7);
    expect(getCell(wb.sheets[0], 1, 0)!.v).toBe(70);
  });

  it('computes a newly entered formula and reports it changed', () => {
    const wb = book([[n(3), n(4)]]);
    const e = new Engine(wb);
    const changed = e.apply([{ sheet: 0, row: 1, col: 0, cell: f('SUM(A1:B1)') }]);
    expect(getCell(wb.sheets[0], 1, 0)!.v).toBe(7);
    expect(changed).toContainEqual({ sheet: 0, row: 1, col: 0 });
  });

  it('keeps text that looks like a number or formula as text', () => {
    const wb = book([[{ v: '123' }, { v: '=1+1' }, f('ISTEXT(A1)'), f('B1')]]);
    new Engine(wb).apply([{ sheet: 0, row: 1, col: 0, cell: f('ISTEXT(A1)&"|"&B1') }]);
    expect(getCell(wb.sheets[0], 1, 0)!.v).toBe('TRUE|=1+1');
  });

  it('treats TRUE and FALSE as constants inside functions', () => {
    const wb = book([[{ v: 'a' }, { v: 'b' }]]);
    new Engine(wb).apply([{ sheet: 0, row: 1, col: 0, cell: f('TEXTJOIN("-",TRUE,A1:B1)&IF(FALSE,"x","y")') }]);
    expect(getCell(wb.sheets[0], 1, 0)!.v).toBe('a-by');
  });

  it('matches Excel date serials', () => {
    const wb = book([[]]);
    const e = new Engine(wb);
    expect(e.evaluate(0, '=DATE(2024,1,1)')).toBe(45292);
    expect(e.evaluate(0, '=TEXT(45292,"yyyy-mm-dd")')).toBe('2024-01-01');
    expect(e.evaluate(0, '=DATE(1900,3,1)')).toBe(61);
    expect(e.evaluate(0, '=DATE(1900,1,1)')).toBe(1);
    expect(e.evaluate(0, '=DAY(60)')).toBe(29); // Excel's phantom Feb 29, 1900
    expect(e.evaluate(0, '=DATEVALUE("3/15/2024")')).toBe(45366);
    expect(e.evaluate(0, '=DATEVALUE("2024-03-15")')).toBe(45366);
  });

  it('formats with TEXT the way Excel does', () => {
    const e = new Engine(book([[]]));
    expect(e.evaluate(0, '=TEXT(0.256,"0%")')).toBe('26%');
    expect(e.evaluate(0, '=TEXT(1234.5,"$#,##0.00")')).toBe('$1,234.50');
    expect(e.evaluate(0, '=TEXT(1234,"£#,##0")')).toBe('£1,234');
    expect(e.evaluate(0, '=TEXT(45292,"dddd, mmmm d")')).toBe('Monday, January 1');
    expect(e.evaluate(0, '=TEXT(100000000000,"General")')).toBe('1E+11');
    expect(e.evaluate(0, '=TEXT("12","0.0")')).toBe('12.0');
    expect(e.evaluate(0, '=INT(-1.5)')).toBe(-2);
  });

  it('spills dynamic arrays into neighbouring cells and records the extent', () => {
    const wb = book([[n(1)], [n(5)], [n(3)], [n(8)]]);
    const e = new Engine(wb);
    e.apply([{ sheet: 0, row: 0, col: 2, cell: f('FILTER(A1:A4,A1:A4>2)') }]);
    const s = wb.sheets[0];
    expect(getCell(s, 0, 2)!.f).toBe('_xlfn._xlws.FILTER(A1:A4,A1:A4>2)');
    expect([0, 1, 2].map((r) => getCell(s, r, 2)!.v)).toEqual([5, 3, 8]);
    expect(getCell(s, 1, 2)!.spill).toBe(true);
    expect(getCell(s, 0, 2)!.arrayRef).toBe('C1:C3');

    e.apply([{ sheet: 0, row: 3, col: 0, cell: n(0) }]);
    expect(getCell(s, 2, 2)).toBeUndefined();
    expect(getCell(s, 0, 2)!.arrayRef).toBe('C1:C2');
  });

  it('saves a spilled formula as an array formula Excel can read', () => {
    const wb = book([[n(2)], [n(1)]]);
    new Engine(wb).apply([{ sheet: 0, row: 0, col: 1, cell: f('SORT(A1:A2)') }]);
    const again = readXlsx(writeXlsx(wb));
    const anchor = getCell(again.sheets[0], 0, 1)!;
    expect(anchor.f).toBe('_xlfn._xlws.SORT(A1:A2)');
    expect(anchor.arrayRef).toBe('B1:B2');
    expect(getCell(again.sheets[0], 1, 1)).toMatchObject({ v: 2, spill: true });
  });

  it('evaluates functions added for Excel and Sheets compatibility', () => {
    const wb = book([[{ v: 'apple' }, { v: 'banana' }, { v: 'cherry' }]]);
    const e = new Engine(wb);
    expect(e.evaluate(0, '=CONCAT(A1:C1,"!")')).toBe('applebananacherry!');
    expect(e.evaluate(0, '=XMATCH("banana",A1:C1)')).toBe(2);
    expect(e.evaluate(0, '=REGEXMATCH(A1,"^ap")')).toBe(true);
    expect(e.evaluate(0, '=REGEXEXTRACT("order-1234","\\d+")')).toBe('1234');
    expect(e.evaluate(0, '=REGEXREPLACE("a1b2","\\d","#")')).toBe('a#b#');
    expect(e.evaluate(0, '=TEXTBEFORE("a.b.c",".",-1)')).toBe('a.b');
    expect(e.evaluate(0, '=TEXTAFTER("a.b.c",".")')).toBe('b.c');
    expect(e.evaluate(0, '=XLOOKUP("cherry",A1:C1,A1:C1)')).toBe('cherry');
  });

  it('evaluates Excel statistical, lookup and text functions HyperFormula lacks', () => {
    const wb = book([[n(1), n(10), { v: 'a' }], [n(2), n(20), { v: 'b' }], [n(2), n(30), { v: 'a' }], [n(4), n(40), { v: 'b' }], [n(5), n(50), { v: 'a' }]]);
    const e = new Engine(wb);
    expect(e.evaluate(0, '=AVERAGEIFS(B1:B5,C1:C5,"a",A1:A5,">1")')).toBe(40);
    expect(e.evaluate(0, '=LOOKUP(3,A1:A5,B1:B5)')).toBe(30);
    expect(e.evaluate(0, '=LOOKUP(2,1/(C1:C5="b"),B1:B5)')).toBe(40);
    expect(e.evaluate(0, '=MODE(A1:A5)')).toBe(2);
    expect(e.evaluate(0, '=RANK(4,A1:A5)')).toBe(2);
    expect(e.evaluate(0, '=RANK.AVG(2,A1:A5,1)')).toBe(2.5);
    expect(e.evaluate(0, '=INTERCEPT(B1:B5,A1:B5)')).toEqual({ error: '#N/A' });
    expect(e.evaluate(0, '=FORECAST(3,{2,4,6},{1,2,3})')).toBe(6);
    expect(e.evaluate(0, '=PERCENTRANK(A1:A5,4)')).toBe(0.75);
    expect(e.evaluate(0, '=TRIMMEAN({1,2,3,4,100},0.4)')).toBe(3);
    expect(e.evaluate(0, '=TYPE(C1)*100+TYPE(A1:A2)')).toBe(264);
    expect(e.evaluate(0, '=ERROR.TYPE(1/0)')).toBe(2);
    expect(e.evaluate(0, '=NUMBERVALUE("1.234,5",",",".")')).toBe(1234.5);
    expect(e.evaluate(0, '=FIXED(1234.567,1)&" "&DOLLAR(-1234.567)')).toBe('1,234.6 ($1,234.57)');
    expect(e.evaluate(0, '=PERMUT(5,2)+PERMUTATIONA(3,2)+GESTEP(5,4)')).toBe(30);
    expect(e.evaluate(0, '=MDETERM({1,2;3,4})')).toBe(-2);
  });

  it('keeps the saved value of a legacy formula Excel evaluated by implicit intersection', () => {
    const wb = book([[n(1)], [n(2), { v: 20, f: 'A1:A3*10' }], [n(3)]]);
    const e = new Engine(wb);
    expect(getCell(wb.sheets[0], 1, 1)).toMatchObject({ v: 20, frozen: true });
    expect(e.value(0, 2, 1)).toBeNull();
    expect([...wb.unsupportedFunctions]).toEqual(['implicit intersection']);
  });

  it('keeps the saved values of formulas that span sheets', () => {
    const wb = loadFixture('55906-MultiSheetRefs.xlsx');
    const e = new Engine(wb);
    const cells = wb.sheets.flatMap((s) => s.rows.flatMap((r) => (r?.cells ?? []).filter((c) => c?.f)));
    expect(cells.length).toBeGreaterThan(0);
    expect(cells.every((c) => c!.frozen)).toBe(true);
    expect(wb.unsupportedFunctions.has('3-D sheet references')).toBe(true);
    e.destroy();
  });

  it('freezes formulas with unknown functions at their cached value and lists them', () => {
    const wb = book([[n(2), { v: 42, f: 'MYADDIN(A1)' }, f('B1+1')]]);
    const e = new Engine(wb);
    expect(getCell(wb.sheets[0], 0, 1)).toMatchObject({ v: 42, frozen: true });
    expect([...wb.unsupportedFunctions]).toEqual(['MYADDIN']);
    expect(e.value(0, 0, 2)).toBe(43);
  });

  it('builds quickly when one stray cell sits far from the rest', () => {
    const wb = book([[n(1)]]);
    setCell(wb.sheets[0], 10000, 16383, n(7));
    const t = performance.now();
    const e = new Engine(wb);
    expect(performance.now() - t).toBeLessThan(200);
    expect(e.evaluate(0, '=A1+XFD10001')).toBe(8);
  });

  it('resolves workbook defined names', () => {
    const wb = book([[n(10), n(20)]]);
    wb.definedNames.push({ name: 'Total', ref: 'Sheet1!$A$1:$B$1' });
    expect(new Engine(wb).evaluate(0, '=SUM(Total)')).toBe(30);
  });

  it('strips and restores future-function prefixes', () => {
    expect(displayFormula('_xlfn.XLOOKUP(A1,B:B,C:C)+_xlfn._xlws.FILTER(A1:A2,TRUE)')).toBe('XLOOKUP(A1,B:B,C:C)+FILTER(A1:A2,TRUE)');
    expect(fileFormula('xlookup(A1,B:B,C:C)+SUM(1)+_xlfn.IFS(1,2)')).toBe('_xlfn.xlookup(A1,B:B,C:C)+SUM(1)+_xlfn.IFS(1,2)');
  });

  it('loads every sample workbook and evaluates its formulas without crashing', () => {
    for (const name of poiFixtures()) {
      const wb = loadFixture(name);
      const e = new Engine(wb);
      wb.sheets.forEach((s, si) => {
        if (s.kind !== 'worksheet') return;
        s.rows.forEach((r, ri) => r?.cells.forEach((c, ci) => { if (c?.f) e.value(si, ri, ci); }));
      });
      e.destroy();
    }
  });
});
