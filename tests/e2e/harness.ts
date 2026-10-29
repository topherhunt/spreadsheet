import { _electron as electron, ElectronApplication, expect, Page, test as base } from '@playwright/test';
import type { ChildProcess } from 'node:child_process';
import { copyFileSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import type { MenuCommand } from '../../src/shared/api';
import { readWorkbook } from '../../src/core/io';
import { formatOf } from '../../src/core/io';
import type { Workbook } from '../../src/core/model';

export const ROOT = resolve(__dirname, '../..');
export const POI = join(ROOT, 'tests/fixtures/poi');

/** A running app with one window, plus a scratch directory for the files it opens and saves. */
export class Sheeyit {
  private proc: ChildProcess;
  private exited = false;

  constructor(readonly app: ElectronApplication, readonly page: Page, readonly dir: string) {
    this.proc = app.process();
    this.proc.once('exit', () => { this.exited = true; });
  }

  static async launch(dir: string, files: string[] = [], opts: { dark?: boolean } = {}): Promise<Sheeyit> {
    const app = await electron.launch({
      args: [join(ROOT, 'dist/main/main.cjs'), '--test-headless', `--test-user-data=${join(dir, '.profile')}`, ...files],
      cwd: ROOT,
      colorScheme: opts.dark ? 'dark' : 'light',
    });
    // A launch that fails partway would otherwise leave an app nothing tracks, alive in the Dock.
    try {
      const page = await app.firstWindow();
      await page.locator('.qa-grid').waitFor({ state: 'attached' });
      await page.waitForFunction(() => !!(window as any).sheeyit?.doc);
      const s = new Sheeyit(app, page, dir);
      await s.stubDialogs();
      return s;
    } catch (e) {
      app.process().kill('SIGKILL');
      throw e;
    }
  }

  /** Copies a fixture into the scratch directory so saving never touches the original. */
  static fixture(dir: string, name: string, from = POI): string {
    const dest = join(dir, name);
    copyFileSync(join(from, name), dest);
    return dest;
  }

  static write(dir: string, name: string, content: string): string {
    const dest = join(dir, name);
    writeFileSync(dest, content);
    return dest;
  }

  /**
   * Native dialogs can't be clicked in tests, so main answers them from queues the
   * test fills: `answerSave` paths and `answerMessage` button indexes. Each dialog
   * shown is recorded for assertions; an unexpected one fails loudly.
   */
  private async stubDialogs(): Promise<void> {
    await this.app.evaluate(({ dialog }) => {
      const g = globalThis as any;
      g.__saves = [] as string[];
      g.__answers = [] as number[];
      g.__messages = [] as { message: string; buttons: string[] }[];
      dialog.showSaveDialog = (async (...args: any[]) => {
        const opts = args[args.length - 1];
        g.__messages.push({ message: `save:${opts.defaultPath}`, buttons: [] });
        if (!g.__saves.length) throw new Error(`Unexpected save dialog for ${opts.defaultPath}`);
        const filePath = g.__saves.shift();
        return { canceled: filePath === null, filePath: filePath ?? undefined };
      }) as any;
      dialog.showMessageBox = (async (...args: any[]) => {
        const opts = args[args.length - 1];
        g.__messages.push({ message: opts.message, buttons: opts.buttons });
        if (!g.__answers.length) throw new Error(`Unexpected message box: ${opts.message}`);
        return { response: g.__answers.shift(), checkboxChecked: false };
      }) as any;
      dialog.showMessageBoxSync = ((...args: any[]) => {
        const opts = args[args.length - 1];
        g.__messages.push({ message: opts.message, buttons: opts.buttons });
        if (!g.__answers.length) throw new Error(`Unexpected message box: ${opts.message}`);
        return g.__answers.shift();
      }) as any;
    });
  }

  async answerSave(path: string | null): Promise<void> {
    await this.app.evaluate((_e, p) => { (globalThis as any).__saves.push(p); }, path);
  }

  async answerMessage(button: number): Promise<void> {
    await this.app.evaluate((_e, b) => { (globalThis as any).__answers.push(b); }, button);
  }

  async messages(): Promise<{ message: string; buttons: string[] }[]> {
    return this.app.evaluate(() => (globalThis as any).__messages);
  }

  /** Ends the app without the unsaved-changes prompt; tests that check that prompt close the window themselves. */
  async close(): Promise<void> {
    if (this.exited) return;
    const exited = new Promise((r) => this.proc.once('exit', r));
    // SIGTERM would quit gracefully, which asks about unsaved changes.
    this.proc.kill('SIGKILL');
    await exited;
  }

  /** Runs a menu command the way the menu bar does. */
  async command(cmd: MenuCommand): Promise<void> {
    await this.app.evaluate(({ BrowserWindow }, c) => BrowserWindow.getAllWindows()[0].webContents.send('menu', c), cmd);
  }

  /** Waits until a save finished: the dirty dot is gone and nothing is pending. */
  async save(): Promise<void> {
    await this.command('save');
    await expect(this.page.locator('.qa-dirty')).toBeHidden();
  }

  async selectCell(ref: string): Promise<void> {
    const box = this.page.locator('.qa-name-box');
    await box.click();
    await box.fill(ref);
    await box.press('Enter');
    await expect(box).toHaveValue(ref.toUpperCase());
  }

  /** Types into a cell as a user would: select, type, Enter. */
  async enter(ref: string, text: string): Promise<void> {
    await this.selectCell(ref);
    await this.page.keyboard.type(text);
    await this.page.keyboard.press('Enter');
  }

  /** The text the grid shows for a cell (0-based row and column). */
  cell(row: number, col: number) {
    return this.page.locator(`.qa-grid [data-rc="${row}:${col}"]`).first();
  }

  async cellAt(ref: string) {
    const { row, col } = await this.page.evaluate((r) => {
      const m = /^([A-Z]+)(\d+)$/.exec(r)!;
      let col = 0;
      for (const ch of m[1]) col = col * 26 + ch.charCodeAt(0) - 64;
      return { row: Number(m[2]) - 1, col: col - 1 };
    }, ref);
    return this.cell(row, col);
  }

  /** Data from the in-memory model, for checks the grid can't show (formulas, styles). */
  async model<T>(fn: string, arg?: unknown): Promise<T> {
    return this.page.evaluate(([src, a]) => {
      const app = (window as any).sheeyit;
      return new Function('app', 'arg', `return (${src})(app, arg)`)(app, a);
    }, [fn, arg] as const) as Promise<T>;
  }

  /** The value a cell holds in the model; `ref` is A1 on the active sheet. */
  async value(ref: string): Promise<unknown> {
    return this.model(`(app, ref) => {
      const m = /^([A-Z]+)(\\d+)$/.exec(ref);
      let col = 0;
      for (const ch of m[1]) col = col * 26 + ch.charCodeAt(0) - 64;
      const c = app.sheet.rows[Number(m[2]) - 1]?.cells[col - 1];
      return c ? (c.v && typeof c.v === 'object' ? c.v.error : c.v) : undefined;
    }`, ref);
  }

  /** Filter buttons are drawn under the grid's scroll layer, which hit-tests clicks by position. */
  async openFilter(col: number): Promise<void> {
    await this.page.locator(`.qa-filter-btn-${col}`).click({ force: true });
    await expect(this.page.locator('.qa-filter-menu')).toBeVisible();
  }

  /** Right-clicks a column header (`col`) or row header (`row`), 0-based, and waits for the menu. */
  async headerMenu(area: 'col' | 'row', index: number): Promise<void> {
    const at = await this.model<{ x: number; y: number }>(`(app, [area, i]) => {
      const box = app.grid.el.getBoundingClientRect();
      const r = app.grid.cellRect(area === 'row' ? i : 0, area === 'col' ? i : 0);
      return area === 'col' ? { x: r.left + r.width / 2, y: box.top + 4 } : { x: box.left + 4, y: r.top + r.height / 2 };
    }`, [area, index]);
    await this.page.mouse.click(at.x, at.y, { button: 'right' });
    await expect(this.page.locator('.qa-context-menu')).toBeVisible();
  }

  async formulaBar(): Promise<string> {
    return this.page.locator('.qa-formula-input').inputValue();
  }
}

/** Reads a saved file back with the app's own reader. */
export function readBack(path: string): Workbook {
  return readWorkbook(new Uint8Array(readFileSync(path)), formatOf(path));
}

export const test = base.extend<{ dir: string; launch: (files?: string[], opts?: { dark?: boolean }) => Promise<Sheeyit> }>({
  dir: async ({}, use) => {
    const dir = mkdtempSync(join(tmpdir(), 'sheeyit-e2e-'));
    await use(dir);
    rmSync(dir, { recursive: true, force: true });
  },
  /** Starts the app; every app a test starts is ended afterwards, pass or fail. */
  launch: async ({ dir }, use) => {
    const apps: Sheeyit[] = [];
    await use(async (files = [], opts = {}) => {
      const s = await Sheeyit.launch(dir, files, opts);
      apps.push(s);
      return s;
    });
    await Promise.all(apps.map((s) => s.close()));
  },
});

export { expect };
