import { app, BrowserWindow, dialog, ipcMain, Menu, MenuItemConstructorOptions, nativeTheme, shell } from 'electron';
import fs from 'node:fs';
import path from 'node:path';
import type { InitialDoc, MenuCommand, MessageOptions } from '../shared/api';

const flag = (name: string) => process.argv.includes(`--${name}`);
const flagValue = (name: string) => process.argv.find((a) => a.startsWith(`--${name}=`))?.slice(name.length + 3);

// Tests run windows hidden and in their own profile so parallel runs never meet.
const headless = flag('test-headless');
const testUserData = flagValue('test-user-data');
if (testUserData) app.setPath('userData', testUserData);
const devUrl = process.env.SHEEYIT_DEV_URL;

const OPENABLE = /\.(xlsx|xlsm|csv|tsv|txt|tab)$/i;

interface WinState { path: string | null; dirty: boolean; pending: InitialDoc; closing: boolean }
const windows = new Map<number, { win: BrowserWindow; state: WinState }>();
let quitting = false;

function readDoc(file: string | null): InitialDoc {
  if (!file) return { path: null, bytes: null };
  return { path: file, bytes: new Uint8Array(fs.readFileSync(file)) };
}

function createWindow(file: string | null): BrowserWindow {
  const offset = windows.size * 24;
  const win = new BrowserWindow({
    width: 1280, height: 820, minWidth: 560, minHeight: 360,
    x: headless ? undefined : 80 + offset, y: headless ? undefined : 60 + offset,
    show: false,
    title: 'Sheeyit',
    titleBarStyle: process.platform === 'darwin' ? 'hiddenInset' : 'default',
    backgroundColor: nativeTheme.shouldUseDarkColors ? '#1e1f22' : '#ffffff',
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      spellcheck: false,
    },
  });
  const state: WinState = { path: file, dirty: false, pending: readDoc(file), closing: false };
  const id = win.webContents.id;
  windows.set(id, { win, state });
  if (file) app.addRecentDocument(file);

  if (!headless) win.once('ready-to-show', () => win.show());
  win.on('close', (e) => {
    if (!state.dirty || state.closing) return;
    e.preventDefault();
    const name = state.path ? path.basename(state.path) : 'Untitled';
    const choice = dialog.showMessageBoxSync(win, {
      type: 'warning', message: `Do you want to save the changes you made to “${name}”?`,
      detail: 'Your changes will be lost if you don’t save them.',
      buttons: ['Save', 'Cancel', 'Don’t Save'], defaultId: 0, cancelId: 1,
    });
    if (choice === 1) { quitting = false; return; }
    if (choice === 2) { state.closing = true; win.close(); return; }
    win.webContents.send('menu', 'save-then-close' satisfies MenuCommand);
  });
  win.on('closed', () => {
    windows.delete(id);
    if (quitting && !windows.size) app.quit();
  });
  win.webContents.setWindowOpenHandler(({ url }) => { openExternal(url); return { action: 'deny' }; });
  win.webContents.on('will-navigate', (e) => e.preventDefault());

  if (devUrl) win.loadURL(devUrl);
  else win.loadFile(path.join(__dirname, '../renderer/index.html'));
  return win;
}

function entry(sender: Electron.WebContents) {
  const e = windows.get(sender.id);
  if (!e) throw new Error('Unknown window');
  return e;
}

/** Brings the window already showing `file` forward, or opens it in a new one. */
function openFile(file: string): void {
  const resolved = path.resolve(file);
  for (const { win, state } of windows.values()) {
    if (state.path === resolved) { win.show(); win.focus(); return; }
  }
  createWindow(resolved);
}

function openExternal(url: string): void {
  if (/^(https?|mailto):/i.test(url)) shell.openExternal(url);
}

const FILTERS: Record<string, Electron.FileFilter> = {
  xlsx: { name: 'Excel Workbook', extensions: ['xlsx'] },
  xlsm: { name: 'Excel Macro-Enabled Workbook', extensions: ['xlsm'] },
  csv: { name: 'CSV (comma-separated)', extensions: ['csv'] },
  tsv: { name: 'TSV (tab-separated)', extensions: ['tsv'] },
};

ipcMain.handle('initial-doc', (e) => entry(e.sender).state.pending);

ipcMain.handle('open-dialog', async (e, reuseWindow: boolean) => {
  const { win, state } = entry(e.sender);
  const res = await dialog.showOpenDialog(win, {
    properties: ['openFile', 'multiSelections'],
    filters: [{ name: 'Spreadsheets', extensions: ['xlsx', 'xlsm', 'csv', 'tsv', 'txt', 'tab'] }],
  });
  if (res.canceled) return;
  const files = res.filePaths.map((f) => path.resolve(f));
  if (reuseWindow && !state.path && !state.dirty) {
    const first = files.shift()!;
    if ([...windows.values()].some((w) => w.state.path === first)) openFile(first);
    else {
      state.path = first;
      state.pending = readDoc(first);
      app.addRecentDocument(first);
      win.webContents.send('load', state.pending);
    }
  }
  for (const f of files) openFile(f);
});

ipcMain.handle('new-window', () => { createWindow(null); });

ipcMain.handle('save-dialog', async (e, defaultPath: string, format: string) => {
  const { win } = entry(e.sender);
  // The current format first; the rest let Save As convert.
  const filters = [FILTERS[format], ...Object.entries(FILTERS).filter(([k]) => k !== format).map(([, f]) => f)];
  const res = await dialog.showSaveDialog(win, { defaultPath, filters });
  return res.canceled || !res.filePath ? null : res.filePath;
});

// Write to a temp file and rename, so a failed save never leaves a half-written file.
ipcMain.handle('write-file', (e, file: string, bytes: Uint8Array) => {
  const tmp = path.join(path.dirname(file), `.${path.basename(file)}.sheeyit-${process.pid}.tmp`);
  try {
    fs.writeFileSync(tmp, bytes);
    fs.renameSync(tmp, file);
  } catch (err) {
    fs.rmSync(tmp, { force: true });
    throw err;
  }
  const { state } = entry(e.sender);
  state.path = file;
  app.addRecentDocument(file);
});

ipcMain.handle('message', (e, opts: MessageOptions) => dialog.showMessageBox(entry(e.sender).win, opts).then((r) => r.response));

ipcMain.on('set-state', (e, s: { path: string | null; title: string; dirty: boolean }) => {
  const { win, state } = entry(e.sender);
  state.path = s.path;
  state.dirty = s.dirty;
  win.setTitle(s.title);
  win.setDocumentEdited(s.dirty);
  if (s.path) win.setRepresentedFilename(s.path);
});

ipcMain.on('close-confirmed', (e) => {
  const { win, state } = entry(e.sender);
  state.closing = true;
  win.close();
});

ipcMain.on('open-external', (_e, url: string) => openExternal(url));

// ---------------------------------------------------------------------------
// Menu

function send(cmd: MenuCommand) {
  return () => BrowserWindow.getFocusedWindow()?.webContents.send('menu', cmd);
}

function item(label: string, cmd: MenuCommand, accelerator?: string): MenuItemConstructorOptions {
  return { label, accelerator, click: send(cmd) };
}

function buildMenu(): void {
  const mac = process.platform === 'darwin';
  const template: MenuItemConstructorOptions[] = [
    ...(mac ? [{ role: 'appMenu' as const }] : []),
    {
      label: 'File',
      submenu: [
        item('New', 'new', 'CmdOrCtrl+N'),
        item('Open…', 'open', 'CmdOrCtrl+O'),
        ...(mac ? [{ role: 'recentDocuments' as const, submenu: [{ role: 'clearRecentDocuments' as const }] }] : []),
        { type: 'separator' },
        item('Save', 'save', 'CmdOrCtrl+S'),
        item('Save As…', 'save-as', 'Shift+CmdOrCtrl+S'),
        { type: 'separator' },
        mac ? { role: 'close' } : { role: 'quit' },
      ],
    },
    {
      label: 'Edit',
      submenu: [
        item('Undo', 'undo', 'CmdOrCtrl+Z'),
        item('Redo', 'redo', mac ? 'Shift+Cmd+Z' : 'Ctrl+Y'),
        { type: 'separator' },
        { role: 'cut' }, { role: 'copy' }, { role: 'paste' },
        item('Paste Values Only', 'paste-values', 'Shift+CmdOrCtrl+V'),
        item('Select All', 'select-all', 'CmdOrCtrl+A'),
        { type: 'separator' },
        item('Fill Down', 'fill-down', 'CmdOrCtrl+D'),
        item('Fill Right', 'fill-right', 'CmdOrCtrl+R'),
        { type: 'separator' },
        item('Find…', 'find', 'CmdOrCtrl+F'),
        item('Find and Replace…', 'replace', 'Shift+CmdOrCtrl+H'),
      ],
    },
    {
      label: 'View',
      submenu: [
        item('Zoom In', 'zoom-in', 'CmdOrCtrl+='),
        item('Zoom Out', 'zoom-out', 'CmdOrCtrl+-'),
        item('Actual Size', 'zoom-reset', 'CmdOrCtrl+0'),
        { type: 'separator' },
        item('Show Formulas', 'toggle-formulas', 'Ctrl+`'),
        item('Gridlines', 'toggle-gridlines'),
        { type: 'separator' },
        item('Unfreeze', 'freeze-none'),
        item('Freeze Top Row', 'freeze-row'),
        item('Freeze First Column', 'freeze-col'),
        item('Freeze Up to Selection', 'freeze-selection'),
        { type: 'separator' },
        { role: 'togglefullscreen' },
        { role: 'toggleDevTools' },
      ],
    },
    {
      label: 'Insert',
      submenu: [
        item('Row Above', 'insert-row-above'),
        item('Row Below', 'insert-row-below'),
        item('Column Left', 'insert-col-left'),
        item('Column Right', 'insert-col-right'),
        { type: 'separator' },
        item('Delete Rows', 'delete-rows'),
        item('Delete Columns', 'delete-cols'),
        { type: 'separator' },
        item('Sheet', 'add-sheet', 'Shift+F11'),
        item('Link…', 'insert-link', 'CmdOrCtrl+K'),
      ],
    },
    {
      label: 'Format',
      submenu: [
        item('Bold', 'bold', 'CmdOrCtrl+B'),
        item('Italic', 'italic', 'CmdOrCtrl+I'),
        item('Underline', 'underline', 'CmdOrCtrl+U'),
        item('Strikethrough', 'strike', 'Shift+CmdOrCtrl+X'),
        { type: 'separator' },
        item('Clear Formatting', 'clear-format', 'CmdOrCtrl+\\'),
      ],
    },
    {
      label: 'Data',
      submenu: [
        item('Auto Table (Bold, Freeze and Filter Header)', 'auto-table', 'Alt+CmdOrCtrl+T'),
        item('Create / Remove Filter', 'toggle-filter', 'Shift+CmdOrCtrl+L'),
        item('Clear All Filters', 'clear-filters'),
        { type: 'separator' },
        item('Sort Range A → Z', 'sort-asc'),
        item('Sort Range Z → A', 'sort-desc'),
      ],
    },
    { role: 'windowMenu' },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

// ---------------------------------------------------------------------------
// App lifecycle

const argvFiles = () => process.argv.slice(app.isPackaged ? 1 : 2).filter((a) => !a.startsWith('-') && OPENABLE.test(a) && fs.existsSync(a));
const queued: string[] = [];

if (!headless && !app.requestSingleInstanceLock()) app.quit();
else {
  app.on('second-instance', (_e, argv) => {
    const files = argv.filter((a) => OPENABLE.test(a) && fs.existsSync(a));
    if (files.length) files.forEach(openFile);
    else BrowserWindow.getAllWindows()[0]?.focus();
  });

  // macOS delivers files opened from Finder this way, possibly before ready.
  app.on('open-file', (e, file) => {
    e.preventDefault();
    if (app.isReady()) openFile(file); else queued.push(file);
  });

  app.on('before-quit', () => { quitting = true; });
  app.on('window-all-closed', () => { if (process.platform !== 'darwin' || headless) app.quit(); });
  app.on('activate', () => { if (!windows.size) createWindow(null); });

  app.whenReady().then(() => {
    buildMenu();
    const files = [...queued, ...argvFiles()];
    if (files.length) files.forEach(openFile); else createWindow(null);
  });
}
