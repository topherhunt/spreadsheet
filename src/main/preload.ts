import { contextBridge, ipcRenderer } from 'electron';
import type { Api } from '../shared/api';

const api: Api = {
  initialDoc: () => ipcRenderer.invoke('initial-doc'),
  openDialog: (reuseWindow) => ipcRenderer.invoke('open-dialog', reuseWindow),
  newWindow: () => ipcRenderer.invoke('new-window'),
  saveDialog: (defaultPath, format) => ipcRenderer.invoke('save-dialog', defaultPath, format),
  writeFile: (path, bytes) => ipcRenderer.invoke('write-file', path, bytes),
  message: (opts) => ipcRenderer.invoke('message', opts),
  setState: (state) => ipcRenderer.send('set-state', state),
  closeConfirmed: () => ipcRenderer.send('close-confirmed'),
  openExternal: (url) => ipcRenderer.send('open-external', url),
  onMenu: (fn) => { ipcRenderer.on('menu', (_e, cmd) => fn(cmd)); },
  onLoad: (fn) => { ipcRenderer.on('load', (_e, doc) => fn(doc)); },
};

contextBridge.exposeInMainWorld('api', api);
