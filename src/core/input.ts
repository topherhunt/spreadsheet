import { parseInput } from './csv';
import { serialFromMs } from './dates';
import { Scalar } from './model';

export interface Typed {
  v: Scalar;
  /** Formula text in display syntax, without '='. */
  formula?: string;
  /** Number format the input implies (a date, percent or currency), for cells still formatted General. */
  fmt?: string;
}

const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];

function timeFraction(h: number, m: number, s: number, ampm?: string): number | null {
  if (ampm) {
    if (h < 1 || h > 12) return null;
    h = (h % 12) + (/p/i.test(ampm) ? 12 : 0);
  }
  if (h > 23 || m > 59 || s > 59) return null;
  return (h * 3600 + m * 60 + s) / 86400;
}

function dateSerial(y: number, mo: number, d: number, date1904: boolean): number | null {
  if (y < 100) y += y < 30 ? 2000 : 1900;
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  const ms = Date.UTC(y, mo - 1, d);
  const back = new Date(ms);
  if (back.getUTCDate() !== d) return null; // Feb 30th and friends
  return serialFromMs(ms, date1904);
}

const TIME = String.raw`(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([ap]\.?m\.?)?`;

/** Recognizes dates and times the way Sheets does for en-US input. */
export function parseDateTime(s: string, date1904 = false): Typed | null {
  const t = s.trim();
  let m = new RegExp(`^${TIME}$`, 'i').exec(t);
  if (m) {
    const f = timeFraction(+m[1], +m[2], +(m[3] ?? 0), m[4]);
    return f === null ? null : { v: f, fmt: m[4] ? 'h:mm AM/PM' : m[3] ? 'h:mm:ss' : 'h:mm' };
  }
  let date: number | null = null;
  let rest = '';
  if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](.*))?$/.exec(t))) { date = dateSerial(+m[1], +m[2], +m[3], date1904); rest = m[4] ?? ''; }
  else if ((m = /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})(?:\s+(.*))?$/.exec(t))) { date = dateSerial(+m[3], +m[1], +m[2], date1904); rest = m[4] ?? ''; }
  else if ((m = /^([a-z]{3})[a-z]*\.?\s+(\d{1,2}),?\s+(\d{4})$/i.exec(t)) && MONTHS.includes(m[1].toLowerCase())) {
    date = dateSerial(+m[3], MONTHS.indexOf(m[1].toLowerCase()) + 1, +m[2], date1904);
  }
  if (date === null) return null;
  if (!rest) return { v: date, fmt: 'm/d/yyyy' };
  const tm = new RegExp(`^${TIME}$`, 'i').exec(rest.trim());
  if (!tm) return null;
  const f = timeFraction(+tm[1], +tm[2], +(tm[3] ?? 0), tm[4]);
  return f === null ? null : { v: date + f, fmt: 'm/d/yyyy h:mm:ss' };
}

/** What text typed into a cell means. */
export function parseTyped(text: string, date1904 = false): Typed {
  if (text.startsWith('=') && text.length > 1) return { v: null, formula: text.slice(1) };
  if (text.startsWith("'")) return { v: text.slice(1) };
  const dt = parseDateTime(text, date1904);
  if (dt) return dt;
  const v = parseInput(text);
  if (typeof v !== 'number') return { v };
  const t = text.trim();
  const decimals = (/\.(\d+)/.exec(t)?.[1].length) ?? 0;
  const dec = decimals ? '.' + '0'.repeat(Math.min(decimals, 2)) : '';
  if (/^[+-]?[$€£¥]/.test(t)) {
    const sym = /[$€£¥]/.exec(t)![0];
    return { v, fmt: `"${sym}"#,##0${dec || '.00'}` };
  }
  if (t.endsWith('%')) return { v, fmt: `0${dec}%` };
  if (/\d,\d{3}/.test(t)) return { v, fmt: `#,##0${dec}` };
  return { v };
}
