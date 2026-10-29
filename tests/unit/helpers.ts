import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { strFromU8, unzipSync } from 'fflate';
import { readXlsx } from '../../src/core/xlsx/read';
import { Workbook } from '../../src/core/model';

export const POI_DIR = join(__dirname, '../fixtures/poi');

/** Fixtures that are deliberately broken or hostile; tested separately. */
export const BAD_FIXTURES = ['deep-data.xlsx', 'poc-xmlbomb.xlsx'];

export function poiFixtures(): string[] {
  return readdirSync(POI_DIR).filter((f) => /\.xls[xm]$/.test(f) && !BAD_FIXTURES.includes(f)).sort();
}

export function fixtureBytes(name: string): Uint8Array {
  return new Uint8Array(readFileSync(join(POI_DIR, name)));
}

export function loadFixture(name: string): Workbook {
  return readXlsx(fixtureBytes(name), name.endsWith('.xlsm') ? 'xlsm' : 'xlsx');
}

export function entries(bytes: Uint8Array): Record<string, Uint8Array> {
  return unzipSync(bytes);
}

export function same(a: Uint8Array | undefined, b: Uint8Array | undefined): boolean {
  if (!a || !b || a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

export function text(bytes: Uint8Array): string {
  return strFromU8(bytes);
}

/** Everything the app models about a workbook's content, as plain data for deep comparison. */
export function contentOf(wb: Workbook) {
  return {
    sheets: wb.sheets.map((s) => ({
      name: s.name, kind: s.kind, state: s.state,
      freeze: s.freeze, merges: s.merges, filter: s.filter, tabColor: s.tabColor,
      cols: s.cols.map((c) => c && { width: c.width, hidden: c.hidden, style: c.style }),
      rows: s.rows.map((r) => r && {
        height: r.height, hidden: r.hidden, filtered: r.filtered,
        cells: r.cells.map((c) => c && { v: c.v, f: c.f, s: c.s, link: c.link, arrayRef: c.arrayRef, rich: c.rich?.map((x) => [x.text, x.font]) }),
      }),
    })),
    names: wb.definedNames.map((n) => [n.name, n.localSheet, n.ref]),
    date1904: wb.date1904,
  };
}
