import './styles.css';
import type { InitialDoc } from '../shared/api';
import { App } from './app';
import { toast } from './ui/dom';

const app = new App(document.getElementById('app')!);
(window as unknown as { sheeyit: App }).sheeyit = app;

/** A file that can't be read leaves the window on a new workbook, so saving can't overwrite it. */
function load(d: InitialDoc): void {
  try {
    app.open(d);
  } catch (e) {
    console.error(e);
    app.open({ path: null, bytes: null });
    const name = d.path!.split(/[\\/]/).pop();
    toast(`Couldn’t open ${name}: ${e instanceof Error ? e.message : String(e)}`, true);
  }
}

window.api.onMenu((cmd) => app.command(cmd));
window.api.onLoad(load);
load(await window.api.initialDoc());
