import { CellError, EmptyValue, ErrorType, FunctionArgumentType, FunctionPlugin, HyperFormula, SimpleRangeValue } from 'hyperformula';
import * as SSF from 'ssf';
import { formatValue, generalNumber } from '../numfmt';

// Scalar Excel/Sheets functions HyperFormula lacks, plus TEXT and INT, which it
// gets wrong. Array-returning ones (TEXTSPLIT, TOCOL, ...) aren't here:
// formulas using them keep the file's cached value instead (see engine.ts).

type Arg = unknown;
type Ast = { args: unknown[] };
type State = unknown;

const { STRING, NUMBER, SCALAR, NOERROR, RANGE, BOOLEAN, ANY } = FunctionArgumentType;

function text(v: Arg): string {
  if (typeof v === 'number') return generalNumber(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (v === null || v === undefined) return '';
  if (typeof v === 'object' && 'val' in (v as object)) return text((v as { val: unknown }).val); // rich number types
  return String(v);
}

/** The plain number behind HyperFormula's date/percent/currency number types. */
function raw(v: Arg): Arg {
  return typeof v === 'object' && v !== null && 'val' in v ? (v as { val: unknown }).val : v;
}

/** The numbers in range/scalar arguments, skipping text, booleans and blanks as Excel's statistical functions do. Returns the first error. */
function numbers(args: Arg[]): number[] | CellError {
  const out: number[] = [];
  for (const a of args) {
    const vals = a instanceof SimpleRangeValue ? a.valuesFromTopLeftCorner() : [a];
    for (const v of vals.map(raw)) {
      if (v instanceof CellError) return v;
      if (typeof v === 'number') out.push(v);
    }
  }
  return out;
}

/** Number pairs from two equal-sized ranges, keeping positions where both are numbers. */
function pairs(ys: SimpleRangeValue, xs: SimpleRangeValue): [number, number][] | CellError {
  const y = ys.valuesFromTopLeftCorner().map(raw), x = xs.valuesFromTopLeftCorner().map(raw);
  if (y.length !== x.length) return new CellError(ErrorType.NA, 'Ranges differ in size');
  const out: [number, number][] = [];
  for (let i = 0; i < y.length; i++) {
    if (y[i] instanceof CellError) return y[i] as CellError;
    if (x[i] instanceof CellError) return x[i] as CellError;
    if (typeof y[i] === 'number' && typeof x[i] === 'number') out.push([y[i] as number, x[i] as number]);
  }
  return out;
}

function linear(ys: SimpleRangeValue, xs: SimpleRangeValue): { slope: number; intercept: number } | CellError {
  const p = pairs(ys, xs);
  if (p instanceof CellError) return p;
  if (!p.length) return new CellError(ErrorType.DIV_BY_ZERO, 'No data');
  const my = p.reduce((a, [y]) => a + y, 0) / p.length, mx = p.reduce((a, [, x]) => a + x, 0) / p.length;
  let sxy = 0, sxx = 0;
  for (const [y, x] of p) { sxy += (x - mx) * (y - my); sxx += (x - mx) ** 2; }
  if (sxx === 0) return new CellError(ErrorType.DIV_BY_ZERO, 'No variance in x');
  const slope = sxy / sxx;
  return { slope, intercept: my - slope * mx };
}

function rank(v: number, range: SimpleRangeValue, ascending: number, average: boolean): number | CellError {
  const ns = numbers([range]);
  if (ns instanceof CellError) return ns;
  const before = ns.filter((n) => (ascending ? n < v : n > v)).length;
  const ties = ns.filter((n) => n === v).length;
  if (!ties) return new CellError(ErrorType.NA, 'Value not in list');
  return average ? before + (ties + 1) / 2 : before + 1;
}

function percentrank(range: SimpleRangeValue, x: number, sig: number, exclusive: boolean): number | CellError {
  const ns = numbers([range]);
  if (ns instanceof CellError) return ns;
  sig = Math.trunc(sig);
  if (!ns.length || sig < 1) return new CellError(ErrorType.NUM, 'Bad arguments');
  ns.sort((a, b) => a - b);
  if (x < ns[0] || x > ns[ns.length - 1]) return new CellError(ErrorType.NA, 'Out of range');
  const n = ns.length;
  let pos: number; // 0-based fractional rank of x
  const at = ns.indexOf(x);
  if (at >= 0) pos = at;
  else {
    const i = ns.findIndex((v) => v > x) - 1;
    pos = i + (x - ns[i]) / (ns[i + 1] - ns[i]);
  }
  const r = exclusive ? (pos + 1) / (n + 1) : n === 1 ? 1 : pos / (n - 1);
  const f = 10 ** sig;
  return Math.floor(r * f + 1e-9) / f;
}

function fixed(n: number, decimals: number, commas: boolean): string {
  decimals = Math.trunc(decimals);
  const f = 10 ** decimals;
  const rounded = Math.sign(n) * Math.round(Math.abs(n) * f) / f;
  return rounded.toLocaleString('en-US', { minimumFractionDigits: Math.max(0, decimals), maximumFractionDigits: Math.max(0, decimals), useGrouping: commas });
}

const ERROR_NUMBERS: Record<string, number> = { [ErrorType.DIV_BY_ZERO]: 2, [ErrorType.VALUE]: 3, [ErrorType.REF]: 4, [ErrorType.NAME]: 5, [ErrorType.NUM]: 6, [ErrorType.NA]: 7 };

function regex(pattern: string, flags = ''): RegExp | CellError {
  try { return new RegExp(pattern, flags); } catch { return new CellError(ErrorType.VALUE, `Invalid regular expression: ${pattern}`); }
}

function eq(a: Arg, b: Arg): boolean {
  if (typeof a === 'string' && typeof b === 'string') return a.toLowerCase() === b.toLowerCase();
  return text(a) === text(b) && typeof a === typeof b;
}

function wildcard(pattern: string): RegExp {
  const src = pattern.replace(/~([*?~])|([*?])|([.+^${}()|[\]\\])/g, (_m, lit, wild, meta) =>
    lit ? '\\' + lit : wild === '*' ? '.*' : wild === '?' ? '.' : '\\' + meta);
  return new RegExp(`^${src}$`, 'i');
}

class ExtraFunctions extends FunctionPlugin {
  static implementedFunctions = {
    CONCAT: { method: 'concat', parameters: [{ argumentType: SCALAR }], repeatLastArgs: 1, expandRanges: true },
    REGEXMATCH: { method: 'regexmatch', parameters: [{ argumentType: STRING }, { argumentType: STRING }] },
    REGEXTEST: { method: 'regextest', parameters: [{ argumentType: STRING }, { argumentType: STRING }, { argumentType: NUMBER, defaultValue: 0 }] },
    REGEXEXTRACT: { method: 'regexextract', parameters: [{ argumentType: STRING }, { argumentType: STRING }] },
    REGEXREPLACE: { method: 'regexreplace', parameters: [{ argumentType: STRING }, { argumentType: STRING }, { argumentType: STRING }] },
    TEXTBEFORE: { method: 'textbefore', parameters: [{ argumentType: STRING }, { argumentType: STRING }, { argumentType: NUMBER, defaultValue: 1 }, { argumentType: NUMBER, defaultValue: 0 }] },
    TEXTAFTER: { method: 'textafter', parameters: [{ argumentType: STRING }, { argumentType: STRING }, { argumentType: NUMBER, defaultValue: 1 }, { argumentType: NUMBER, defaultValue: 0 }] },
    XMATCH: { method: 'xmatch', parameters: [{ argumentType: NOERROR }, { argumentType: RANGE }, { argumentType: NUMBER, defaultValue: 0 }, { argumentType: NUMBER, defaultValue: 1 }] },
    TEXT: { method: 'text', parameters: [{ argumentType: SCALAR }, { argumentType: STRING }] },
    INT: { method: 'int', parameters: [{ argumentType: NUMBER }] },
    LOOKUP: { method: 'lookup', parameters: [{ argumentType: NOERROR }, { argumentType: RANGE }, { argumentType: RANGE, optionalArg: true }] },
    MODE: { method: 'mode', parameters: [{ argumentType: RANGE }], repeatLastArgs: 1 },
    'MODE.SNGL': { method: 'mode', parameters: [{ argumentType: RANGE }], repeatLastArgs: 1 },
    RANK: { method: 'rankEq', parameters: [{ argumentType: NUMBER }, { argumentType: RANGE }, { argumentType: NUMBER, defaultValue: 0 }] },
    'RANK.EQ': { method: 'rankEq', parameters: [{ argumentType: NUMBER }, { argumentType: RANGE }, { argumentType: NUMBER, defaultValue: 0 }] },
    'RANK.AVG': { method: 'rankAvg', parameters: [{ argumentType: NUMBER }, { argumentType: RANGE }, { argumentType: NUMBER, defaultValue: 0 }] },
    FORECAST: { method: 'forecast', parameters: [{ argumentType: NUMBER }, { argumentType: RANGE }, { argumentType: RANGE }] },
    'FORECAST.LINEAR': { method: 'forecast', parameters: [{ argumentType: NUMBER }, { argumentType: RANGE }, { argumentType: RANGE }] },
    INTERCEPT: { method: 'intercept', parameters: [{ argumentType: RANGE }, { argumentType: RANGE }] },
    TYPE: { method: 'type', parameters: [{ argumentType: ANY }] },
    'ERROR.TYPE': { method: 'errorType', parameters: [{ argumentType: SCALAR }] },
    NUMBERVALUE: { method: 'numbervalue', parameters: [{ argumentType: STRING }, { argumentType: STRING, defaultValue: '.' }, { argumentType: STRING, defaultValue: ',' }] },
    FIXED: { method: 'fixed', parameters: [{ argumentType: NUMBER }, { argumentType: NUMBER, defaultValue: 2 }, { argumentType: BOOLEAN, defaultValue: false }] },
    DOLLAR: { method: 'dollar', parameters: [{ argumentType: NUMBER }, { argumentType: NUMBER, defaultValue: 2 }] },
    PERCENTRANK: { method: 'percentrankInc', parameters: [{ argumentType: RANGE }, { argumentType: NUMBER }, { argumentType: NUMBER, defaultValue: 3 }] },
    'PERCENTRANK.INC': { method: 'percentrankInc', parameters: [{ argumentType: RANGE }, { argumentType: NUMBER }, { argumentType: NUMBER, defaultValue: 3 }] },
    'PERCENTRANK.EXC': { method: 'percentrankExc', parameters: [{ argumentType: RANGE }, { argumentType: NUMBER }, { argumentType: NUMBER, defaultValue: 3 }] },
    TRIMMEAN: { method: 'trimmean', parameters: [{ argumentType: RANGE }, { argumentType: NUMBER }] },
    KURT: { method: 'kurt', parameters: [{ argumentType: RANGE }], repeatLastArgs: 1 },
    PERMUT: { method: 'permut', parameters: [{ argumentType: NUMBER }, { argumentType: NUMBER }] },
    PERMUTATIONA: { method: 'permutationa', parameters: [{ argumentType: NUMBER }, { argumentType: NUMBER }] },
    GESTEP: { method: 'gestep', parameters: [{ argumentType: NUMBER }, { argumentType: NUMBER, defaultValue: 0 }] },
    MDETERM: { method: 'mdeterm', parameters: [{ argumentType: RANGE }] },
    ISBETWEEN: { method: 'isbetween', parameters: [{ argumentType: NUMBER }, { argumentType: NUMBER }, { argumentType: NUMBER }, { argumentType: BOOLEAN, defaultValue: true }, { argumentType: BOOLEAN, defaultValue: true }] },
  };

  private run(name: string, ast: Ast, state: State, fn: (...args: any[]) => unknown): any {
    return (this as any).runFunction(ast.args, state, (this as any).metadata(name), fn);
  }

  concat(ast: Ast, state: State) {
    return this.run('CONCAT', ast, state, (...args: Arg[]) => {
      let out = '';
      for (const a of args) {
        if (a instanceof CellError) return a;
        out += text(a);
      }
      return out;
    });
  }

  regexmatch(ast: Ast, state: State) {
    return this.run('REGEXMATCH', ast, state, (s: string, p: string) => {
      const re = regex(p);
      return re instanceof CellError ? re : re.test(s);
    });
  }

  regextest(ast: Ast, state: State) {
    return this.run('REGEXTEST', ast, state, (s: string, p: string, caseMode: number) => {
      const re = regex(p, caseMode === 1 ? 'i' : '');
      return re instanceof CellError ? re : re.test(s);
    });
  }

  regexextract(ast: Ast, state: State) {
    return this.run('REGEXEXTRACT', ast, state, (s: string, p: string) => {
      const re = regex(p);
      if (re instanceof CellError) return re;
      const m = re.exec(s);
      if (!m) return new CellError(ErrorType.NA, 'No match');
      return m.length > 1 ? m[1] : m[0];
    });
  }

  regexreplace(ast: Ast, state: State) {
    return this.run('REGEXREPLACE', ast, state, (s: string, p: string, r: string) => {
      const re = regex(p, 'g');
      return re instanceof CellError ? re : s.replace(re, r.replace(/\\(\d)/g, '$$$1'));
    });
  }

  textbefore(ast: Ast, state: State) {
    return this.run('TEXTBEFORE', ast, state, (s: string, d: string, n: number, mode: number) => {
      const at = nthIndex(s, d, n, mode === 1);
      return at < 0 ? new CellError(ErrorType.NA, 'Delimiter not found') : s.slice(0, at);
    });
  }

  textafter(ast: Ast, state: State) {
    return this.run('TEXTAFTER', ast, state, (s: string, d: string, n: number, mode: number) => {
      const at = nthIndex(s, d, n, mode === 1);
      return at < 0 ? new CellError(ErrorType.NA, 'Delimiter not found') : s.slice(at + d.length);
    });
  }

  xmatch(ast: Ast, state: State) {
    return this.run('XMATCH', ast, state, (needle: Arg, range: SimpleRangeValue, matchMode: number, searchMode: number) => {
      const values = range.valuesFromTopLeftCorner();
      const order = values.map((_, i) => i);
      if (searchMode < 0) order.reverse();
      const wild = matchMode === 2 && typeof needle === 'string' ? wildcard(needle) : null;
      let best = -1;
      for (const i of order) {
        const v = values[i];
        if (wild ? typeof v === 'string' && wild.test(v) : eq(v, needle)) return i + 1;
        if (matchMode === -1 || matchMode === 1) {
          if (typeof v !== typeof needle || v === null) continue;
          const below = (v as number | string) < (needle as number | string);
          if (below !== (matchMode === -1)) continue;
          if (best < 0 || (matchMode === -1 ? (v as number) > (values[best] as number) : (v as number) < (values[best] as number))) best = i;
        }
      }
      return best >= 0 ? best + 1 : new CellError(ErrorType.NA, 'Value not found');
    });
  }

  isbetween(ast: Ast, state: State) {
    return this.run('ISBETWEEN', ast, state, (v: number, lo: number, hi: number, loIn: boolean, hiIn: boolean) =>
      (loIn ? v >= lo : v > lo) && (hiIn ? v <= hi : v < hi));
  }

  text(ast: Ast, state: State) {
    return this.run('TEXT', ast, state, (v: Arg, code: string) => {
      v = raw(v);
      if (v instanceof CellError) return v;
      if (typeof v === 'string' && v.trim() !== '' && Number.isFinite(Number(v))) v = Number(v);
      if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
      if (v === EmptyValue || v === null || v === undefined) v = 0;
      if (typeof v === 'number' && code.toLowerCase() === 'general') return SSF.format('General', v); // Excel's General, not the grid's
      const date1904 = (this as any).config.nullDate.year === 1904;
      return formatValue(v as number | string, code, date1904).text;
    });
  }

  int(ast: Ast, state: State) {
    return this.run('INT', ast, state, (n: number) => Math.floor(n));
  }

  lookup(ast: Ast, state: State) {
    return this.run('LOOKUP', ast, state, (needle: Arg, range: SimpleRangeValue, result?: SimpleRangeValue) => {
      needle = raw(needle);
      const rows = range.data.map((r) => r.map(raw));
      const byRow = !result && rows[0].length > rows.length; // array form: search the first row when wider than tall
      const keys = result ? rows.flat() : byRow ? rows[0] : rows.map((r) => r[0]);
      let found = -1;
      for (let i = 0; i < keys.length; i++) {
        const k = keys[i];
        if (typeof k !== typeof needle || k === null) continue;
        const cmp = typeof k === 'string' ? k.toLowerCase().localeCompare((needle as string).toLowerCase()) : (k as number) - (needle as number);
        if (cmp <= 0) found = i; else break;
      }
      if (found < 0) return new CellError(ErrorType.NA, 'Value not found');
      if (result) {
        const out = result.valuesFromTopLeftCorner();
        return found < out.length ? out[found] : new CellError(ErrorType.NA, 'Result range too short');
      }
      return byRow ? rows[rows.length - 1][found] : rows[found][rows[found].length - 1];
    });
  }

  mode(ast: Ast, state: State) {
    return this.run('MODE', ast, state, (...args: Arg[]) => {
      const ns = numbers(args);
      if (ns instanceof CellError) return ns;
      const counts = new Map<number, number>();
      let best: number | undefined, most = 1;
      for (const n of ns) {
        const c = (counts.get(n) ?? 0) + 1;
        counts.set(n, c);
        if (c > most) { most = c; best = n; }
      }
      // Ties go to the value whose repeats came first, as in Excel.
      if (best === undefined) return new CellError(ErrorType.NA, 'No repeated value');
      return ns.find((n) => counts.get(n) === most)!;
    });
  }

  rankEq(ast: Ast, state: State) {
    return this.run('RANK', ast, state, (v: number, range: SimpleRangeValue, order: number) => rank(v, range, order, false));
  }

  rankAvg(ast: Ast, state: State) {
    return this.run('RANK.AVG', ast, state, (v: number, range: SimpleRangeValue, order: number) => rank(v, range, order, true));
  }

  forecast(ast: Ast, state: State) {
    return this.run('FORECAST', ast, state, (x: number, ys: SimpleRangeValue, xs: SimpleRangeValue) => {
      const l = linear(ys, xs);
      return l instanceof CellError ? l : l.intercept + l.slope * x;
    });
  }

  intercept(ast: Ast, state: State) {
    return this.run('INTERCEPT', ast, state, (ys: SimpleRangeValue, xs: SimpleRangeValue) => {
      const l = linear(ys, xs);
      return l instanceof CellError ? l : l.intercept;
    });
  }

  type(ast: Ast, state: State) {
    return this.run('TYPE', ast, state, (v: Arg) => {
      if (v instanceof SimpleRangeValue) return v.width() === 1 && v.height() === 1 ? typeOf(raw(v.data[0][0])) : 64;
      return typeOf(raw(v));
    });
  }

  errorType(ast: Ast, state: State) {
    return this.run('ERROR.TYPE', ast, state, (v: Arg) => {
      const n = v instanceof CellError ? ERROR_NUMBERS[v.type] : undefined;
      return n === undefined ? new CellError(ErrorType.NA, 'Not an error') : n;
    });
  }

  numbervalue(ast: Ast, state: State) {
    return this.run('NUMBERVALUE', ast, state, (s: string, dec: string, group: string) => {
      if (!dec || !group || dec[0] === group[0]) return new CellError(ErrorType.VALUE, 'Bad separators');
      let t = s.replace(/\s+/g, '').split(group[0]).join('');
      const percents = /%*$/.exec(t)![0].length;
      t = t.slice(0, t.length - percents).replace(dec[0], '.');
      if (t === '') return 0;
      const n = Number(t);
      return Number.isFinite(n) ? n / 100 ** percents : new CellError(ErrorType.VALUE, 'Not a number');
    });
  }

  fixed(ast: Ast, state: State) {
    return this.run('FIXED', ast, state, (n: number, decimals: number, noCommas: boolean) => fixed(n, decimals, !noCommas));
  }

  dollar(ast: Ast, state: State) {
    return this.run('DOLLAR', ast, state, (n: number, decimals: number) => {
      const s = '$' + fixed(Math.abs(n), decimals, true);
      return n < 0 && fixed(n, decimals, true).startsWith('-') ? `(${s})` : s;
    });
  }

  percentrankInc(ast: Ast, state: State) {
    return this.run('PERCENTRANK.INC', ast, state, (range: SimpleRangeValue, x: number, sig: number) => percentrank(range, x, sig, false));
  }

  percentrankExc(ast: Ast, state: State) {
    return this.run('PERCENTRANK.EXC', ast, state, (range: SimpleRangeValue, x: number, sig: number) => percentrank(range, x, sig, true));
  }

  trimmean(ast: Ast, state: State) {
    return this.run('TRIMMEAN', ast, state, (range: SimpleRangeValue, pct: number) => {
      const ns = numbers([range]);
      if (ns instanceof CellError) return ns;
      if (!ns.length || pct < 0 || pct >= 1) return new CellError(ErrorType.NUM, 'Bad arguments');
      const cut = Math.floor((ns.length * pct) / 2);
      const kept = ns.sort((a, b) => a - b).slice(cut, ns.length - cut);
      return kept.reduce((a, b) => a + b, 0) / kept.length;
    });
  }

  kurt(ast: Ast, state: State) {
    return this.run('KURT', ast, state, (...args: Arg[]) => {
      const ns = numbers(args);
      if (ns instanceof CellError) return ns;
      const n = ns.length;
      if (n < 4) return new CellError(ErrorType.DIV_BY_ZERO, 'Needs four values');
      const mean = ns.reduce((a, b) => a + b, 0) / n;
      const sd = Math.sqrt(ns.reduce((a, b) => a + (b - mean) ** 2, 0) / (n - 1));
      if (sd === 0) return new CellError(ErrorType.DIV_BY_ZERO, 'No variance');
      const s4 = ns.reduce((a, b) => a + ((b - mean) / sd) ** 4, 0);
      return (n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3)) * s4 - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3));
    });
  }

  permut(ast: Ast, state: State) {
    return this.run('PERMUT', ast, state, (n: number, k: number) => {
      n = Math.trunc(n); k = Math.trunc(k);
      if (n < 0 || k < 0 || n < k) return new CellError(ErrorType.NUM, 'Bad arguments');
      let out = 1;
      for (let i = 0; i < k; i++) out *= n - i;
      return out;
    });
  }

  permutationa(ast: Ast, state: State) {
    return this.run('PERMUTATIONA', ast, state, (n: number, k: number) => {
      n = Math.trunc(n); k = Math.trunc(k);
      return n < 0 || k < 0 ? new CellError(ErrorType.NUM, 'Bad arguments') : n ** k;
    });
  }

  gestep(ast: Ast, state: State) {
    return this.run('GESTEP', ast, state, (n: number, step: number) => (n >= step ? 1 : 0));
  }

  mdeterm(ast: Ast, state: State) {
    return this.run('MDETERM', ast, state, (range: SimpleRangeValue) => {
      const m = range.data.map((r) => r.map(raw));
      if (m.length !== m[0].length) return new CellError(ErrorType.VALUE, 'Not square');
      for (const r of m) for (const v of r) if (typeof v !== 'number') return v instanceof CellError ? v : new CellError(ErrorType.VALUE, 'Not a number');
      return determinant(m as number[][]);
    });
  }
}

function typeOf(v: Arg): number {
  if (v instanceof CellError) return 16;
  if (typeof v === 'string') return 2;
  if (typeof v === 'boolean') return 4;
  return 1;
}

/** Gaussian elimination with partial pivoting. */
function determinant(m: number[][]): number {
  const a = m.map((r) => [...r]);
  const n = a.length;
  let det = 1;
  for (let c = 0; c < n; c++) {
    let p = c;
    for (let r = c + 1; r < n; r++) if (Math.abs(a[r][c]) > Math.abs(a[p][c])) p = r;
    if (a[p][c] === 0) return 0;
    if (p !== c) { [a[p], a[c]] = [a[c], a[p]]; det = -det; }
    det *= a[c][c];
    for (let r = c + 1; r < n; r++) {
      const f = a[r][c] / a[c][c];
      for (let k = c; k < n; k++) a[r][k] -= f * a[c][k];
    }
  }
  return det;
}

type Aggregator = { computeConditionalAggregationFunction(values: SimpleRangeValue, conditions: Arg[], name: string, init: unknown, compose: (a: any, b: any) => unknown, map: (v: Arg) => unknown): unknown };
// HyperFormula doesn't export its plugin classes; SUMIFS's plugin owns the criteria matching and caching AVERAGEIFS needs.
const ConditionalPlugin = HyperFormula.getFunctionPlugin('SUMIFS') as unknown as typeof FunctionPlugin | undefined;
if (!ConditionalPlugin || typeof (ConditionalPlugin.prototype as unknown as Aggregator).computeConditionalAggregationFunction !== 'function') {
  throw new Error('HyperFormula conditional aggregation is unavailable');
}

class ConditionalFunctions extends (ConditionalPlugin as typeof FunctionPlugin) {
  static implementedFunctions = {
    AVERAGEIFS: { method: 'averageifs', parameters: [{ argumentType: RANGE }, { argumentType: RANGE }, { argumentType: NOERROR }], repeatLastArgs: 2 },
  };

  averageifs(ast: Ast, state: State) {
    return (this as any).runFunction(ast.args, state, (this as any).metadata('AVERAGEIFS'), (values: SimpleRangeValue, ...conditions: Arg[]) => {
      const sum = (this as unknown as Aggregator).computeConditionalAggregationFunction(values, conditions, 'AVERAGEIFS', [0, 0],
        (a, b) => (a instanceof CellError ? a : b instanceof CellError ? b : [a[0] + b[0], a[1] + b[1]]),
        (v) => { v = raw(v); return v instanceof CellError ? v : typeof v === 'number' ? [v, 1] : [0, 0]; });
      if (sum instanceof CellError) return sum;
      const [total, count] = sum as [number, number];
      return count ? total / count : new CellError(ErrorType.DIV_BY_ZERO, 'No matching values');
    });
  }
}

function nthIndex(s: string, d: string, n: number, ignoreCase: boolean): number {
  if (n === 0 || d === '') return -1;
  const hay = ignoreCase ? s.toLowerCase() : s;
  const needle = ignoreCase ? d.toLowerCase() : d;
  if (n > 0) {
    let at = -1;
    for (let k = 0; k < n; k++) { at = hay.indexOf(needle, at + 1); if (at < 0) return -1; }
    return at;
  }
  let at = hay.length + 1;
  for (let k = 0; k < -n; k++) { at = hay.lastIndexOf(needle, at - 1); if (at < 0) return -1; }
  return at;
}

HyperFormula.unregisterFunction('TEXT');
HyperFormula.unregisterFunction('INT');
for (const plugin of [ExtraFunctions, ConditionalFunctions]) {
  const names = Object.keys(plugin.implementedFunctions);
  HyperFormula.registerFunctionPlugin(plugin as any, { enGB: Object.fromEntries(names.map((n) => [n, n])) });
}
