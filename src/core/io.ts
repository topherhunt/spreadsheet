import { delimitedLosses, readDelimited, writeDelimited } from './csv';
import { FileFormat, Workbook } from './model';
import { displayCell } from './numfmt';
import { readXlsx } from './xlsx/read';
import { writeXlsx } from './xlsx/write';

export function formatOf(path: string): FileFormat {
  const ext = /\.([^./\\]+)$/.exec(path)?.[1].toLowerCase();
  if (ext === 'xlsx' || ext === 'xlsm' || ext === 'csv' || ext === 'tsv') return ext;
  if (ext === 'txt' || ext === 'tab') return 'tsv';
  throw new Error(`Can’t open .${ext ?? ''} files. Spreadsheet opens .xlsx, .xlsm, .csv and .tsv.`);
}

export function readWorkbook(bytes: Uint8Array, format: FileFormat): Workbook {
  if (format === 'xlsx' || format === 'xlsm') return readXlsx(bytes, format);
  return readDelimited(new TextDecoder('utf-8').decode(bytes), format);
}

export function writeWorkbook(wb: Workbook, format: FileFormat): Uint8Array {
  if (format === 'xlsx' || format === 'xlsm') return writeXlsx(wb);
  return new TextEncoder().encode(writeDelimited(wb, format, (c) => displayCell(wb, c).text));
}

/** What saving `wb` back as CSV/TSV would drop; empty for spreadsheet formats. */
export function saveLosses(wb: Workbook, format: FileFormat): string[] {
  return format === 'csv' || format === 'tsv' ? delimitedLosses(wb) : [];
}

export function withExtension(path: string, ext: string): string {
  return path.replace(/\.[^./\\]*$/, '') + '.' + ext;
}
