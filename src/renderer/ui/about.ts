import appIcon from '../../../build/icon.png';
import { h } from './dom';

const REPO = 'https://github.com/topherhunt/spreadsheet';

export function showAbout(): void {
  if (document.querySelector('.qa-about')) return;
  const back = h('div', 'modal-backdrop', document.body);
  const box = h('div', 'modal d-flex flex-col align-center gap-2 text-center qa-about', back);
  const img = h('img', '', box);
  img.src = appIcon;
  img.width = img.height = 80;
  h('div', 'bold large', box, 'Spreadsheet');
  h('div', 'muted qa-about-version', box, `Version ${__APP_VERSION__} (${__BUILD_COMMIT__})`);
  h('div', '', box, 'Made with 💙 by Topher Hunt');
  const link = h('a', 'qa-about-link', box, 'github.com/topherhunt/spreadsheet');
  link.href = REPO;
  link.onclick = (e) => { e.preventDefault(); window.api.openExternal(REPO); };
  const close = () => { back.remove(); document.removeEventListener('keydown', onKey, true); };
  const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape' || e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); close(); } };
  document.addEventListener('keydown', onKey, true);
  back.addEventListener('mousedown', (e) => { if (e.target === back) close(); });
}
