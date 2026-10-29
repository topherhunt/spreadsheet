// The bridge the preload script exposes to the renderer as `window.api`.

export interface InitialDoc {
  /** Absolute path, or null for a new untitled workbook. */
  path: string | null;
  bytes: Uint8Array | null;
}

export interface MessageOptions {
  type?: 'none' | 'info' | 'error' | 'question' | 'warning';
  message: string;
  detail?: string;
  buttons: string[];
  defaultId?: number;
  cancelId?: number;
}

export type MenuCommand =
  | 'new' | 'open' | 'save' | 'save-as' | 'save-then-close'
  | 'undo' | 'redo' | 'find' | 'replace' | 'select-all' | 'paste-values'
  | 'bold' | 'italic' | 'underline' | 'strike' | 'clear-format' | 'insert-link'
  | 'auto-table' | 'toggle-filter' | 'clear-filters' | 'sort-asc' | 'sort-desc'
  | 'freeze-none' | 'freeze-row' | 'freeze-col' | 'freeze-selection'
  | 'zoom-in' | 'zoom-out' | 'zoom-reset' | 'toggle-gridlines' | 'toggle-formulas'
  | 'insert-row-above' | 'insert-row-below' | 'insert-col-left' | 'insert-col-right' | 'delete-rows' | 'delete-cols'
  | 'add-sheet' | 'fill-down' | 'fill-right';

export interface Api {
  initialDoc(): Promise<InitialDoc>;
  /** Asks for a file to open and opens it, reusing this window when it's an untouched new workbook. */
  openDialog(reuseWindow: boolean): Promise<void>;
  newWindow(): Promise<void>;
  saveDialog(defaultPath: string, format: 'xlsx' | 'xlsm' | 'csv' | 'tsv'): Promise<string | null>;
  writeFile(path: string, bytes: Uint8Array): Promise<void>;
  message(opts: MessageOptions): Promise<number>;
  /** Tells the window its file, title and whether it has unsaved changes. */
  setState(state: { path: string | null; title: string; dirty: boolean }): void;
  /** Closes this window without asking about unsaved changes. */
  closeConfirmed(): void;
  openExternal(url: string): void;
  onMenu(fn: (cmd: MenuCommand) => void): void;
  /** Called to load a file into this window (it was untouched when the user opened another file). */
  onLoad(fn: (doc: InitialDoc) => void): void;
}

declare global {
  interface Window { api: Api }
}
