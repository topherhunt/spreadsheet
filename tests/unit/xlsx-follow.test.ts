import { describe, expect, it } from 'vitest';
import { Doc } from '../../src/core/doc';
import { getCell } from '../../src/core/model';
import { readXlsx } from '../../src/core/xlsx/read';
import { writeXlsx } from '../../src/core/xlsx/write';
import { childElements, firstChild, parseXml, XEl } from '../../src/core/xlsx/xml';
import { entries, fixtureBytes, loadFixture, same, text } from './helpers';

// Parts the app doesn't model (tables, comments, drawings, charts, pivots, conditional formats,
// validations) move with inserted and deleted rows and columns, as in Excel.

const open = (name: string) => new Doc(loadFixture(name));
const save = (d: Doc) => entries(writeXlsx(d.wb));
const root = (files: Record<string, Uint8Array>, path: string) => parseXml(text(files[path])).documentElement as unknown as XEl;
const all = (el: XEl, name: string): XEl[] => childElements(el).flatMap((c) => [...(c.localName === name ? [c] : []), ...all(c, name)]);

describe('tables follow row and column edits', () => {
  it('a column inserted inside a table becomes a named table column with its header written', () => {
    const d = open('WithTable.xlsx');
    d.insertCols(0, 1, 1);
    const files = save(d);
    const t = root(files, 'xl/tables/table1.xml');
    expect(t.getAttribute('ref')).toBe('A1:C2');
    expect(firstChild(t, 'autoFilter')!.getAttribute('ref')).toBe('A1:C2');
    const cols = firstChild(t, 'tableColumns')!;
    expect(cols.getAttribute('count')).toBe('3');
    expect(childElements(cols).map((c) => [c.getAttribute('id'), c.getAttribute('name')])).toEqual([['1', 'a'], ['3', 'Column1'], ['2', 'b']]);
    expect(getCell(readXlsx(writeXlsx(d.wb)).sheets[0], 0, 1)?.v).toBe('Column1');
  });

  it('rows inserted above a table move it down, and deleting a data row shrinks it', () => {
    const d = open('WithTable.xlsx');
    d.insertRows(0, 0, 2);
    expect(root(save(d), 'xl/tables/table1.xml').getAttribute('ref')).toBe('A3:B4');
    const s = open('StructuredReferences.xlsx');
    s.insertRows(1, 3, -1);
    expect(root(save(s), 'xl/tables/table1.xml').getAttribute('ref')).toBe('A1:C6');
  });

  it('deleting one of a table’s columns drops that table column', () => {
    const d = open('WithTable.xlsx');
    d.insertCols(0, 0, -1);
    const t = root(save(d), 'xl/tables/table1.xml');
    expect(t.getAttribute('ref')).toBe('A1:A2');
    expect(childElements(firstChild(t, 'tableColumns')!).map((c) => c.getAttribute('name'))).toEqual(['b']);
  });

  it('deleting every column of a table removes the table, and undo after saving brings it back exactly', () => {
    const d = open('WithTable.xlsx');
    d.insertCols(0, 0, -2);
    const gone = save(d);
    expect(gone['xl/tables/table1.xml']).toBeUndefined();
    expect(text(gone['xl/worksheets/sheet1.xml'])).not.toContain('tablePart');
    expect(text(gone['xl/worksheets/_rels/sheet1.xml.rels'])).not.toContain('table1.xml');
    expect(text(gone['[Content_Types].xml'])).not.toContain('table1.xml');

    d.undo();
    const back = save(d);
    const orig = entries(fixtureBytes('WithTable.xlsx'));
    expect(same(back['xl/tables/table1.xml'], orig['xl/tables/table1.xml'])).toBe(true);
    expect(same(back['xl/worksheets/_rels/sheet1.xml.rels'], orig['xl/worksheets/_rels/sheet1.xml.rels'])).toBe(true);
    expect(text(back['xl/worksheets/sheet1.xml'])).toContain('tablePart');
    expect(text(back['[Content_Types].xml'])).toContain('/xl/tables/table1.xml');
  });

  it('deleting a table’s header row on its own is refused and changes nothing', () => {
    const d = open('StructuredReferences.xlsx');
    expect(() => d.insertRows(1, 0, -1)).toThrow(/header row/);
    expect(d.wb.edits).toEqual([]);
    expect(getCell(d.wb.sheets[1], 0, 1)?.v).toBe('Name');
  });
});

describe('comments and notes follow row and column edits', () => {
  it('deleting a commented row removes its comment and note, and moves the ones below up', () => {
    const d = open('SimpleWithComments.xlsx');
    d.insertRows(0, 1, -1);
    const files = save(d);
    const comments = all(root(files, 'xl/comments1.xml'), 'comment');
    expect(comments.map((c) => c.getAttribute('ref'))).toEqual(['B1', 'B2']);
    expect(comments[1].textContent).toContain('third cell');
    const vml = text(files['xl/drawings/vmlDrawing1.vml']);
    expect([...vml.matchAll(/<x:Row>(\d+)<\/x:Row>/g)].map((m) => m[1])).toEqual(['0', '1']);
  });

  it('deleting a column to the left of comments moves them and their note boxes left with their cells', () => {
    const d = open('SimpleWithComments.xlsx');
    d.insertCols(0, 0, -1);
    const files = save(d);
    const comments = all(root(files, 'xl/comments1.xml'), 'comment');
    expect(comments.map((c) => c.getAttribute('ref'))).toEqual(['A1', 'A2', 'A3']);
    expect(comments[0].textContent).toContain('first cell');
    const vml = text(files['xl/drawings/vmlDrawing1.vml']);
    expect([...vml.matchAll(/<x:Column>(\d+)<\/x:Column>/g)].map((m) => m[1])).toEqual(['0', '0', '0']);
    expect(vml).toContain('<x:Anchor>1, 15, 0, 2, 3, 15, 4, 8</x:Anchor>');
    const back = readXlsx(writeXlsx(d.wb));
    expect(back.sheets[0].features.has('Notes/comments')).toBe(true);
  });

  it('inserting a column moves comments and their note boxes right', () => {
    const d = open('SimpleWithComments.xlsx');
    d.insertCols(0, 0, 1);
    const files = save(d);
    expect(all(root(files, 'xl/comments1.xml'), 'comment').map((c) => c.getAttribute('ref'))).toEqual(['C1', 'C2', 'C3']);
    const vml = text(files['xl/drawings/vmlDrawing1.vml']);
    expect(vml).toContain('<x:Anchor>3, 15, 0, 2, 5, 15, 4, 8</x:Anchor>');
    expect([...vml.matchAll(/<x:Column>(\d+)<\/x:Column>/g)].map((m) => m[1])).toEqual(['2', '2', '2']);
  });
});

describe('charts and drawings follow edits', () => {
  it('a chart’s data ranges follow rows inserted on its data sheet, and a rename of that sheet', () => {
    const d = open('WithChart.xlsx');
    d.insertRows(0, 0, 1);
    expect(text(save(d)['xl/charts/chart1.xml'])).toContain('<c:f>Sheet1!$A$2:$A$7</c:f>');
    d.renameSheet(0, 'Data');
    expect(text(save(d)['xl/charts/chart1.xml'])).toContain('<c:f>Data!$B$2:$B$7</c:f>');
  });

  it('a chart’s anchor moves with columns inserted before it', () => {
    const d = open('WithChart.xlsx');
    d.insertCols(1, 0, 1);
    const anchor = all(root(save(d), 'xl/drawings/drawing1.xml'), 'twoCellAnchor')[0];
    const col = (m: string) => firstChild(firstChild(anchor, m)!, 'col')!.textContent;
    expect([col('from'), col('to')]).toEqual(['1', '9']);
  });

  it('an edit on another sheet leaves charts and drawings byte-identical', () => {
    const d = open('WithChart.xlsx');
    d.insertRows(2, 0, 3);
    const files = save(d);
    const orig = entries(fixtureBytes('WithChart.xlsx'));
    for (const p of ['xl/charts/chart1.xml', 'xl/drawings/drawing1.xml']) expect(same(files[p], orig[p]), p).toBe(true);
  });
});

describe('conditional formatting and data validation follow edits', () => {
  it('inserting a row moves conditional formats and the references in their formulas', () => {
    const d = open('WithConditionalFormatting.xlsx');
    d.insertRows(0, 0, 1);
    const cfs = all(root(save(d), 'xl/worksheets/sheet1.xml'), 'conditionalFormatting');
    expect(cfs.map((c) => c.getAttribute('sqref'))).toEqual(['A2:A9', 'B10', 'B2:B8']);
    expect(all(cfs[1], 'formula')[0].textContent).toBe('$A$9>5');
  });

  it('deleting the column a conditional format covers removes it, and references to that column become #REF!', () => {
    const d = open('WithConditionalFormatting.xlsx');
    d.insertCols(0, 0, -1);
    const cfs = all(root(save(d), 'xl/worksheets/sheet1.xml'), 'conditionalFormatting');
    expect(cfs.map((c) => c.getAttribute('sqref'))).toEqual(['A9', 'A1:A7']);
    expect(all(cfs[0], 'formula')[0].textContent).toBe('#REF!>5');
  });

  it('validation formulas follow a deleted column, and validations on a deleted row are removed', () => {
    const d = open('DataValidations-49244.xlsx');
    const before = all(root(entries(fixtureBytes('DataValidations-49244.xlsx')), 'xl/worksheets/sheet1.xml'), 'dataValidation').length;
    d.insertCols(0, 3, -1);
    d.insertRows(0, 5, -1);
    const sheet = root(save(d), 'xl/worksheets/sheet1.xml');
    const dvs = all(sheet, 'dataValidation');
    // Row 6 holds two of them, in B6 and C6.
    expect(dvs.length).toBe(before - 2);
    expect(all(sheet, 'dataValidations')[0].getAttribute('count')).toBe(String(before - 2));
    expect([dvs[0].getAttribute('sqref'), firstChild(dvs[0], 'formula1')!.textContent]).toEqual(['C7:C8', 'D7']);
  });
});

describe('pivot tables follow edits', () => {
  it('a pivot table moves with rows inserted above it, and its source range follows its data sheet', () => {
    const d = open('ExcelPivotTableSample.xlsx');
    d.insertRows(1, 0, 1);
    d.insertRows(0, 0, 1);
    d.renameSheet(0, 'Source');
    const files = save(d);
    expect(firstChild(root(files, 'xl/pivotTables/pivotTable1.xml'), 'location')!.getAttribute('ref')).toBe('A4:D8');
    const ws = all(root(files, 'xl/pivotCache/pivotCacheDefinition1.xml'), 'worksheetSource')[0];
    expect([ws.getAttribute('sheet'), ws.getAttribute('ref')]).toEqual(['Source', 'A2:C5']);
  });

  it('inserting rows through the middle of a pivot table is refused', () => {
    const d = open('ExcelPivotTableSample.xlsx');
    expect(() => d.insertRows(1, 4, 1)).toThrow(/pivot table/);
    expect(d.wb.edits).toEqual([]);
  });
});

describe('undo across a save', () => {
  it('undoing row inserts after saving restores every followed part as read', () => {
    const d = open('SimpleWithComments.xlsx');
    d.insertRows(0, 0, 2);
    save(d);
    d.undo();
    const files = save(d);
    const orig = entries(fixtureBytes('SimpleWithComments.xlsx'));
    for (const p of ['xl/comments1.xml', 'xl/drawings/vmlDrawing1.vml', 'xl/worksheets/_rels/sheet1.xml.rels']) expect(same(files[p], orig[p]), p).toBe(true);
  });
});
