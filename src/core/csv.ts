import Papa from 'papaparse';
import { Cell, newSheet, Scalar, sheetExtent, Workbook } from './model';
import { emptyStyleSheet } from './styles';

export function newWorkbook(format: Workbook['format'] = 'xlsx'): Workbook {
  return {
    sheets: [newSheet('Sheet1')], styles: emptyStyleSheet(), definedNames: [], date1904: false, activeSheet: 0, format,
    features: new Set(), unsupportedFunctions: new Set(), viewStyles: new Set(), edits: [],
  };
}

const NUMBER = /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/;
const GROUPED = /^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/;
const CURRENCY = /^([+-]?)[$€£¥]\s?(.+)$/;

/**
 * The value a typed or imported string stands for, the way Sheets reads it:
 * plain, grouped, currency and percent numbers, and TRUE/FALSE. Numbers with
 * leading zeros (ZIP codes, IDs) stay text.
 */
export function parseInput(s: string): Scalar {
  const t = s.trim();
  if (t === '') return s === '' ? null : s;
  if (/^(true|false)$/i.test(t)) return t.toLowerCase() === 'true';
  let body = t, scale = 1, sign = '';
  const cur = CURRENCY.exec(body);
  if (cur) { sign = cur[1]; body = cur[2]; }
  if (body.endsWith('%')) { body = body.slice(0, -1).trimEnd(); scale = 0.01; }
  if (GROUPED.test(body)) body = body.replace(/,/g, '');
  if (!NUMBER.test(body)) return s;
  if (/^[+-]?0\d/.test(body)) return s;
  const n = Number(sign + body) * scale;
  return Number.isFinite(n) ? n : s;
}

export function readDelimited(text: string, format: 'csv' | 'tsv'): Workbook {
  const bom = text.charCodeAt(0) === 0xfeff;
  if (bom) text = text.slice(1);
  const parsed = Papa.parse<string[]>(text, {
    delimiter: format === 'tsv' ? '\t' : '',
    delimitersToGuess: [',', ';', '\t', '|'],
    skipEmptyLines: false,
  });
  const fatal = parsed.errors.find((e) => e.type === 'Quotes' && e.code === 'MissingQuotes');
  if (fatal) throw new Error(`Couldn't parse this ${format.toUpperCase()}: ${fatal.message} (row ${(fatal.row ?? 0) + 1}).`);
  const rows = parsed.data;
  // Papaparse reports a trailing newline as one empty final row.
  if (rows.length && rows[rows.length - 1].length === 1 && rows[rows.length - 1][0] === '') rows.pop();
  // Trailing rows and columns holding only empty or whitespace fields are not part of the sheet.
  const blank = (s: string) => s.trim() === '';
  while (rows.length && rows[rows.length - 1].every(blank)) rows.pop();
  const width = rows.reduce((w, fields) => Math.max(w, fields.findLastIndex((s) => !blank(s)) + 1), 0);
  for (let i = 0; i < rows.length; i++) rows[i] = rows[i].slice(0, width);

  const wb = newWorkbook(format);
  wb.csv = { delimiter: parsed.meta.delimiter, newline: parsed.meta.linebreak || '\n', bom };
  const sheet = wb.sheets[0];
  rows.forEach((fields, ri) => {
    const cells: (Cell | undefined)[] = [];
    fields.forEach((raw, ci) => {
      if (raw === '') return;
      const v = parseInput(raw);
      cells[ci] = typeof v === 'string' ? { v } : { v, raw };
    });
    sheet.rows[ri] = { cells };
  });
  return wb;
}

/** Why saving this workbook as CSV/TSV would lose something the user added. Empty means it's safe. */
export function delimitedLosses(wb: Workbook): string[] {
  const out = new Set<string>();
  if (wb.sheets.length > 1) out.add('multiple sheets');
  for (const s of wb.sheets) {
    if (s.merges.length) out.add('merged cells');
    for (const r of s.rows) {
      if (!r) continue;
      for (const c of r.cells) {
        if (!c) continue;
        if (c.f !== undefined) out.add('formulas');
        if (c.s && !wb.viewStyles.has(c.s)) out.add('formatting');
        if (c.rich) out.add('formatting');
        if (c.link) out.add('links');
      }
    }
  }
  return [...out];
}

/** Text written for a cell: its original text while the value is unchanged, otherwise the displayed value. */
export function writeDelimited(wb: Workbook, format: 'csv' | 'tsv', display: (c: Cell) => string): string {
  const sheet = wb.sheets[wb.activeSheet];
  const opts = wb.csv ?? { delimiter: format === 'tsv' ? '\t' : ',', newline: '\n', bom: false };
  const delimiter = format === 'tsv' ? '\t' : wb.format === 'csv' ? opts.delimiter : ',';
  const ext = sheetExtent(sheet);
  const data: string[][] = [];
  for (let r = 0; r < ext.rows; r++) {
    const row = sheet.rows[r];
    const fields: string[] = [];
    for (let c = 0; c < ext.cols; c++) {
      const cell = row?.cells[c];
      fields.push(!cell ? '' : cell.raw !== undefined ? cell.raw : display(cell));
    }
    data.push(fields);
  }
  // Quote only what must be quoted: Papa.unparse also quotes leading and trailing spaces, which changes untouched fields.
  const quote = (s: string) => (s.includes(delimiter) || /["\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s);
  const text = data.map((fields) => fields.map(quote).join(delimiter)).join(opts.newline);
  return (opts.bom ? '﻿' : '') + text + (data.length ? opts.newline : '');
}
