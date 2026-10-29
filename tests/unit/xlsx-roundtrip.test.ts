import { describe, expect, it } from 'vitest';
import { readXlsx } from '../../src/core/xlsx/read';
import { writeXlsx } from '../../src/core/xlsx/write';
import { parseXml } from '../../src/core/xlsx/xml';
import { contentOf, entries, fixtureBytes, loadFixture, poiFixtures, same, text } from './helpers';

const fixtures = poiFixtures();

describe('XLSX round trip of every sample file', () => {
  it.each(fixtures)('%s saves unchanged with every part byte-identical', (name) => {
    const orig = entries(fixtureBytes(name));
    const out = entries(writeXlsx(loadFixture(name)));
    expect(Object.keys(out).sort()).toEqual(Object.keys(orig).sort());
    for (const p of Object.keys(orig)) expect(same(out[p], orig[p]), p).toBe(true);
  });

  it.each(fixtures)('%s keeps all content when every sheet is regenerated', (name) => {
    const wb = loadFixture(name);
    const before = contentOf(loadFixture(name));
    for (const s of wb.sheets) if (s.kind === 'worksheet') s.pristine = false;
    const bytes = writeXlsx(wb);
    const orig = entries(fixtureBytes(name));
    const out = entries(bytes);

    // Only the sheets, the workbook, calc-chain bookkeeping and appended shared strings (from inline strings) may differ.
    const sheetParts = new Set(wb.sheets.filter((s) => s.kind === 'worksheet').map((s) => s.src!.path));
    const allowed = (p: string) => sheetParts.has(p) || p === 'xl/workbook.xml' || p === '[Content_Types].xml'
      || p.endsWith('workbook.xml.rels') || /calcChain|sharedStrings/.test(p) || /worksheets\/_rels\//.test(p);
    for (const p of Object.keys(orig)) {
      if (allowed(p)) continue;
      expect(same(out[p], orig[p]), `${p} changed`).toBe(true);
    }
    for (const [p, b] of Object.entries(out)) if (/\.(xml|rels)$/.test(p)) expect(() => parseXml(text(b)), p).not.toThrow();
    expect(out['xl/calcChain.xml']).toBeUndefined();

    const again = readXlsx(bytes, name.endsWith('.xlsm') ? 'xlsm' : 'xlsx');
    expect(contentOf(again)).toEqual(before);
  });
});
