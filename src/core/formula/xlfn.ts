import { tokenize, untokenize } from './refs';

// Functions added to Excel after 2007 are stored with an _xlfn. prefix (and a
// few also _xlws.); Excel shows #NAME? for them if the prefix is missing.
const XLFN = new Set(`ACOT ACOTH AGGREGATE ARABIC ARRAYTOTEXT BASE BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE BINOM.INV BITAND
BITLSHIFT BITOR BITRSHIFT BITXOR BYCOL BYROW CEILING.MATH CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST
CHOOSECOLS CHOOSEROWS COMBINA CONCAT CONFIDENCE.NORM CONFIDENCE.T COT COTH COVARIANCE.P COVARIANCE.S CSC CSCH DAYS DECIMAL DROP
ECMA.CEILING ENCODEURL ERF.PRECISE ERFC.PRECISE EXPAND EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FIELDVALUE FILTERXML
FLOOR.MATH FLOOR.PRECISE FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FORECAST.LINEAR FORMULATEXT
GAMMA GAMMA.DIST GAMMA.INV GAMMALN.PRECISE GAUSS GROUPBY HSTACK HYPGEOM.DIST IFNA IFS IMAGE IMCOSH IMCOT IMCSC IMCSCH IMSEC IMSECH
IMSINH IMTAN ISFORMULA ISOMITTED ISOWEEKNUM LAMBDA LET LOGNORM.DIST LOGNORM.INV MAKEARRAY MAP MAXIFS MINIFS MODE.MULT MODE.SNGL MUNIT
NEGBINOM.DIST NETWORKDAYS.INTL NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV NUMBERVALUE PDURATION PERCENTILE.EXC PERCENTILE.INC
PERCENTOF PERCENTRANK.EXC PERCENTRANK.INC PERMUTATIONA PHI PIVOTBY POISSON.DIST QUARTILE.EXC QUARTILE.INC RANDARRAY RANK.AVG RANK.EQ
REDUCE REGEXEXTRACT REGEXREPLACE REGEXTEST RRI SCAN SEC SECH SEQUENCE SHEET SHEETS SKEW.P SORTBY STDEV.P STDEV.S STOCKHISTORY SWITCH
T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST TAKE TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT TOCOL TOROW TRIMRANGE UNICHAR UNICODE
UNIQUE VALUETOTEXT VAR.P VAR.S VSTACK WEBSERVICE WEIBULL.DIST WORKDAY.INTL WRAPCOLS WRAPROWS XLOOKUP XMATCH XOR Z.TEST`.split(/\s+/));
const XLWS = new Set(['FILTER', 'SORT']);

function bareName(text: string): string {
  return text.replace(/^_xlfn\./i, '').replace(/^_xlws\./i, '');
}

function mapFuncs(formula: string, fn: (text: string) => string): string {
  if (!formula.includes('(')) return formula;
  const toks = tokenize(formula);
  let changed = false;
  for (const t of toks) {
    if (t.type !== 'func') continue;
    const text = fn(t.text);
    if (text !== t.text) { t.text = text; t.name = text.toUpperCase(); changed = true; }
  }
  return changed ? untokenize(toks) : formula;
}

/** File syntax to what people read and type: future-function prefixes removed. */
export function displayFormula(f: string): string {
  return /_xl/i.test(f) ? mapFuncs(f, bareName) : f;
}

/** Typed formula to file syntax: newer functions get the prefixes Excel expects. */
export function fileFormula(f: string): string {
  return mapFuncs(f, (text) => {
    const name = bareName(text).toUpperCase();
    if (XLWS.has(name)) return `_xlfn._xlws.${bareName(text)}`;
    if (XLFN.has(name)) return `_xlfn.${bareName(text)}`;
    return text;
  });
}
