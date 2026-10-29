// Production build: main + preload with esbuild, renderer with Vite, into dist/.
import { build as esbuild } from 'esbuild';
import { build as vite } from 'vite';
import { rmSync } from 'node:fs';

export const mainOptions = {
  entryPoints: { main: 'src/main/main.ts', preload: 'src/main/preload.ts' },
  outdir: 'dist/main',
  outExtension: { '.js': '.cjs' },
  bundle: true,
  platform: 'node',
  format: 'cjs',
  target: 'node22',
  external: ['electron'],
  sourcemap: true,
  logLevel: 'warning',
};

if (import.meta.url === `file://${process.argv[1]}`) {
  rmSync('dist', { recursive: true, force: true });
  await esbuild(mainOptions);
  await vite({ configFile: 'vite.config.mts', logLevel: 'warn' });
  console.log('Built dist/');
}
