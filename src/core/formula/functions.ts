import { CellError, ErrorType, FunctionArgumentType, FunctionPlugin, HyperFormula, SimpleRangeValue } from 'hyperformula';
import { generalNumber } from '../numfmt';

// Scalar Excel/Sheets functions HyperFormula lacks. Array-returning ones
// (TEXTSPLIT, TOCOL, ...) aren't here: formulas using them keep the file's
// cached value instead (see engine.ts).

type Arg = unknown;
type Ast = { args: unknown[] };
type State = unknown;

const { STRING, NUMBER, SCALAR, NOERROR, RANGE, BOOLEAN } = FunctionArgumentType;

function text(v: Arg): string {
  if (typeof v === 'number') return generalNumber(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (v === null || v === undefined) return '';
  if (typeof v === 'object' && 'val' in (v as object)) return text((v as { val: unknown }).val); // rich number types
  return String(v);
}

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

const names = Object.keys(ExtraFunctions.implementedFunctions);
HyperFormula.registerFunctionPlugin(ExtraFunctions as any, { enGB: Object.fromEntries(names.map((n) => [n, n])) });
