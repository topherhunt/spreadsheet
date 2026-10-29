import * as SSF from 'ssf';
import { Cell, isError, Scalar, Workbook } from './model';
import { numFmtCode } from './styles';

export interface Display { text: string; color?: string; align: 'left' | 'right' | 'center' }

const NAMED_COLORS: Record<string, string> = {
  black: '#000000', blue: '#0000FF', cyan: '#00FFFF', green: '#00FF00', magenta: '#FF00FF', red: '#FF0000', white: '#FFFFFF', yellow: '#FFFF00',
};

/** Splits a format code into its ;-separated sections, ignoring ; inside quotes, brackets and escapes. */
function sections(code: string): string[] {
  const out: string[] = [];
  let cur = '', q = false, br = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '\\' && !q) { cur += ch + (code[i + 1] ?? ''); i++; continue; }
    if (ch === '"') q = !q;
    else if (!q && ch === '[') br = true;
    else if (!q && ch === ']') br = false;
    if (ch === ';' && !q && !br) { out.push(cur); cur = ''; continue; }
    cur += ch;
  }
  out.push(cur);
  return out;
}

const colorCache = new Map<string, (string | undefined)[]>();

/** The [Color] of the section a number displays with. Conditional sections ([>100]) aren't evaluated for color. */
function sectionColor(code: string, v: number | string): string | undefined {
  let cols = colorCache.get(code);
  if (!cols) {
    cols = sections(code).map((s) => {
      const m = /\[(black|blue|cyan|green|magenta|red|white|yellow)\]/i.exec(s);
      return m ? NAMED_COLORS[m[1].toLowerCase()] : undefined;
    });
    colorCache.set(code, cols);
  }
  if (typeof v === 'string') return cols.length >= 4 ? cols[3] : undefined;
  if (v < 0 && cols.length >= 2) return cols[1];
  if (v === 0 && cols.length >= 3) return cols[2];
  return cols[0];
}

/** Formats a number like Sheets' "Automatic": up to 10 significant digits, scientific when very large or small. */
export function generalNumber(n: number): string {
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
  const abs = Math.abs(n);
  if (abs !== 0 && (abs >= 1e15 || abs < 1e-9)) return n.toExponential(Math.min(5, String(n).length)).replace(/\.?0+e/, 'E').replace('e', 'E');
  return String(parseFloat(n.toPrecision(10)));
}

export function formatValue(v: Scalar, code: string, date1904 = false): Display {
  if (v === null) return { text: '', align: 'left' };
  if (isError(v)) return { text: v.error, align: 'center' };
  if (typeof v === 'boolean') return { text: v ? 'TRUE' : 'FALSE', align: 'center' };
  if (typeof v === 'string') {
    if (code === 'General' || code === '@' || sections(code).length < 4) return { text: v, align: 'left' };
    return { text: SSF.format(code, v), color: sectionColor(code, v), align: 'left' };
  }
  if (code === 'General') return { text: generalNumber(v), align: 'right' };
  try {
    return { text: SSF.format(code, v, { date1904 }), color: sectionColor(code, v), align: 'right' };
  } catch {
    // A format code ssf can't handle (rare locale or Excel-only syntax): show the number plainly rather than nothing.
    return { text: generalNumber(v), align: 'right' };
  }
}

export function cellFormat(wb: Workbook, c: Cell | undefined): string {
  if (!c || !c.s) return 'General';
  const xf = wb.styles.xfs[c.s];
  if (!xf) throw new Error(`Cell style ${c.s} is out of range`);
  return numFmtCode(wb.styles, xf.numFmtId);
}

/** What a cell shows. CSV cells show their original text until they're edited. */
export function displayCell(wb: Workbook, c: Cell | undefined): Display {
  if (!c) return { text: '', align: 'left' };
  if (c.raw !== undefined && !c.s) return { text: c.raw, align: typeof c.v === 'number' ? 'right' : 'center' };
  return formatValue(c.v, cellFormat(wb, c), wb.date1904);
}

export function isDateFormat(code: string): boolean {
  if (code === 'General') return false;
  try { return SSF.is_date(code); } catch { return false; }
}
