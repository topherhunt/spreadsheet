// `npm start`: Vite dev server for the renderer (in this process, so it can't
// outlive us), esbuild watching main + preload, and Electron, restarted
// whenever main or preload rebuild. Quitting the app ends the session.
import { context } from 'esbuild';
import { createServer } from 'vite';
import { spawn } from 'node:child_process';
import { createRequire } from 'node:module';
import { mainOptions } from './build.mjs';

const electronPath = createRequire(import.meta.url)('electron');
const PORT = 5373;

const server = await createServer({ configFile: 'vite.config.mts', server: { port: PORT, strictPort: true } });
await server.listen();
const url = `http://localhost:${PORT}/`;

let electron = null;
let restarting = false;
function launch() {
  electron = spawn(electronPath, ['.', ...process.argv.slice(2)], { stdio: 'inherit', env: { ...process.env, SHEEYIT_DEV_URL: url } });
  electron.on('exit', (code) => {
    if (restarting) { restarting = false; launch(); return; }
    shutdown(code ?? 0);
  });
}

const ctx = await context({
  ...mainOptions,
  plugins: [{
    name: 'restart-electron',
    setup(build) {
      build.onEnd((res) => {
        if (res.errors.length) return;
        if (!electron) { launch(); return; }
        console.log('[dev] main process rebuilt; restarting Electron');
        restarting = true;
        electron.kill('SIGTERM');
      });
    },
  }],
});
await ctx.watch();

async function shutdown(code) {
  await ctx.dispose();
  await server.close();
  process.exit(code);
}
for (const sig of ['SIGINT', 'SIGTERM', 'SIGHUP']) process.once(sig, () => { electron?.kill('SIGTERM'); shutdown(130); });
