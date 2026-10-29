import type { Range } from './addr';
import type { StyleSheet } from './styles';
import type { XlsxSource, SheetSource } from './xlsx/read';

export interface ErrorValue { error: string }
export type Scalar = number | string | boolean | ErrorValue | null;

export function isError(v: unknown): v is ErrorValue {
  return typeof v === 'object' && v !== null && 'error' in v;
}

/** Font properties of a rich-text run. Unset means "inherit from the cell style". */
export interface RunFont {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;   // #RRGGBB
  size?: number;    // points
  name?: string;
}

export interface Run {
  text: string;
  font?: RunFont;
  link?: string;
  /** Original <rPr> XML, written back verbatim while `font` is unchanged. */
  rPr?: string;
  rPrFont?: RunFont;
}

export interface Cell {
  /** Displayed/calculated value. For formulas this is the latest result. */
  v: Scalar;
  /** Formula in file syntax, without the leading '='. */
  f?: string;
  /** Rich text runs; `v` holds their concatenated text. */
  rich?: Run[];
  /** Index into the workbook's cellXfs. */
  s?: number;
  /** Cell-level hyperlink target (URL, or "#Sheet!A1" for an internal location). */
  link?: string;
  /** Shared-string index this value came from; dropped as soon as the value changes. */
  sst?: number;
  /** Original CSV text, written back verbatim while the value is unchanged. */
  raw?: string;
  /** Array formula extent in A1 (for t="array" formulas). */
  arrayRef?: string;
  /** Formula the engine can't evaluate: the file's cached value is shown and preserved. */
  frozen?: boolean;
  /** Value produced by another cell's array formula (spilled), not user input. */
  spill?: boolean;
  /** Attributes and children of <c>/<f> we don't model, written back as found. */
  keep?: { cAttrs?: Record<string, string>; fAttrs?: Record<string, string>; extra?: string };
}

export interface Row {
  cells: (Cell | undefined)[];
  /** Height in points, when set explicitly. */
  height?: number;
  hidden?: boolean;
  /** Hidden by the active filter (separate from `hidden`, which the user set). */
  filtered?: boolean;
  /** Row attributes we don't model (s, customFormat, outlineLevel, ...). */
  keep?: Record<string, string>;
}

export type FilterOp =
  | 'empty' | 'notEmpty'
  | 'contains' | 'notContains' | 'startsWith' | 'endsWith' | 'equals' | 'notEquals'
  | 'gt' | 'gte' | 'lt' | 'lte' | 'between' | 'notBetween';

export interface FilterCondition { op: FilterOp; a?: string; b?: string }

export interface ColumnFilter {
  /** Display strings to show. Blank cells are represented by ''. Undefined: no value filter. */
  values?: string[];
  condition?: FilterCondition;
}

export interface SortSpec { col: number; desc: boolean }

export interface AutoFilter {
  range: Range;
  /** Keyed by absolute column index. */
  columns: Record<number, ColumnFilter>;
}

export interface ColInfo {
  /** Width in Excel character units. */
  width?: number;
  hidden?: boolean;
  /** Default style for empty cells in the column. */
  style?: number;
  keep?: Record<string, string>;
}

export interface Sheet {
  name: string;
  /** Worksheets are editable; chart/dialog/macro sheets are carried through untouched. */
  kind: 'worksheet' | 'other';
  rows: (Row | undefined)[];
  cols: (ColInfo | undefined)[];
  defaultColWidth: number;   // character units
  defaultRowHeight: number;  // points
  merges: Range[];
  freeze: { rows: number; cols: number };
  filter: AutoFilter | null;
  state: 'visible' | 'hidden' | 'veryHidden';
  tabColor?: string;
  rightToLeft?: boolean;
  showGridLines: boolean;
  /** Features present in the file that are preserved but not shown/editable. */
  features: Set<string>;
  /** Sheet came from a file and has not been touched: written back byte-for-byte. */
  pristine: boolean;
  src?: SheetSource;
}

export interface DefinedName {
  name: string;
  ref: string;
  localSheet?: number;
  hidden?: boolean;
  keep?: Record<string, string>;
}

export type FileFormat = 'xlsx' | 'xlsm' | 'csv' | 'tsv';

export interface Workbook {
  sheets: Sheet[];
  styles: StyleSheet;
  definedNames: DefinedName[];
  date1904: boolean;
  activeSheet: number;
  /** Format the document was opened from (or will be saved as). */
  format: FileFormat;
  /** The original package, kept so saving can patch it instead of regenerating it. */
  src?: XlsxSource;
  /** Workbook-level features preserved but not supported (charts, pivots, macros...). */
  features: Set<string>;
  /** Functions used by formulas that the engine can't evaluate. */
  unsupportedFunctions: Set<string>;
  /** Delimited-text details, to write a CSV/TSV back the way it came. */
  csv?: { delimiter: string; newline: string; bom: boolean };
  /** Styles applied as view setup (the auto-table header), which don't make a CSV need XLSX. */
  viewStyles: Set<number>;
}

export function newSheet(name: string): Sheet {
  return {
    name, kind: 'worksheet', rows: [], cols: [], defaultColWidth: 13.57, defaultRowHeight: 15.75,
    merges: [], freeze: { rows: 0, cols: 0 }, filter: null, state: 'visible', showGridLines: true,
    features: new Set(), pristine: false,
  };
}

export function getCell(sheet: Sheet, row: number, col: number): Cell | undefined {
  const r = sheet.rows[row];
  return r ? r.cells[col] : undefined;
}

export function ensureRow(sheet: Sheet, row: number): Row {
  let r = sheet.rows[row];
  if (!r) { r = { cells: [] }; sheet.rows[row] = r; }
  return r;
}

export function setCell(sheet: Sheet, row: number, col: number, cell: Cell | undefined): void {
  if (cell === undefined) {
    const r = sheet.rows[row];
    if (r) delete r.cells[col];
    return;
  }
  ensureRow(sheet, row).cells[col] = cell;
}

/** Last used row/col index (-1 when the sheet is empty). */
export function sheetExtent(sheet: Sheet): { rows: number; cols: number } {
  let maxRow = -1, maxCol = -1;
  sheet.rows.forEach((r, ri) => {
    if (!r) return;
    let rowHas = false;
    r.cells.forEach((c, ci) => {
      if (c && (c.v !== null || c.f !== undefined || c.s !== undefined)) { rowHas = true; if (ci > maxCol) maxCol = ci; }
    });
    if (rowHas) maxRow = ri;
  });
  return { rows: maxRow + 1, cols: maxCol + 1 };
}

export function cellText(c: Cell | undefined): string {
  if (!c || c.v === null) return '';
  if (isError(c.v)) return c.v.error;
  if (typeof c.v === 'boolean') return c.v ? 'TRUE' : 'FALSE';
  return String(c.v);
}

export function cloneCell(c: Cell): Cell {
  const out: Cell = { ...c };
  if (c.rich) out.rich = c.rich.map((r) => ({ ...r, font: r.font ? { ...r.font } : undefined }));
  if (c.keep) out.keep = { ...c.keep };
  return out;
}
