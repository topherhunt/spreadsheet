import { chromium } from '@playwright/test';
import { spawn, spawnSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { join } from 'node:path';
import { expect, ROOT, Sheeyit, test } from './harness';

// Covers the packaged .app rather than the dev build: macOS refuses to run a bundle
// whose signature doesn't seal its contents, which source-based tests can't see.
// Skipped until `npm run release` has built one.
const APP = join(ROOT, 'release/mac-arm64/Sheeyit.app');

test.describe('packaged app', () => {
  test.skip(process.platform !== 'darwin' || !existsSync(APP), 'No packaged build; run `npm run release` first');

  test('has a code signature that seals the whole bundle', () => {
    const verify = spawnSync('codesign', ['--verify', '--deep', '--strict', APP], { encoding: 'utf8' });
    expect(verify.stderr.trim()).toBe('');
    expect(verify.status).toBe(0);
    const details = spawnSync('codesign', ['-dvvv', APP], { encoding: 'utf8' });
    const out = details.stdout + details.stderr;
    expect(out).toContain('Identifier=com.topherhunt.sheeyit');
    expect(out).not.toContain('Info.plist=not bound');
    expect(out).not.toContain('Sealed Resources=none');
  });

  // The bundle's fuses turn off Node's --inspect, which Playwright's Electron launcher
  // needs, so this attaches to the renderer over Chromium's debugging port instead.
  test('launches from the bundle and opens a CSV', async ({ dir }) => {
    const file = Sheeyit.write(dir, 'people.csv', 'name,age\nAnn,30\n');
    const port = 9300 + Math.floor(Math.random() * 500);
    const proc = spawn(join(APP, 'Contents/MacOS/Sheeyit'), [`--remote-debugging-port=${port}`, '--test-headless', `--test-user-data=${join(dir, '.profile')}`, file], { stdio: 'ignore' });
    try {
      let browser;
      for (let i = 0; !browser; i++) {
        try { browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`); } catch (e) {
          if (i > 50) throw e;
          await new Promise((r) => setTimeout(r, 200));
        }
      }
      const context = browser.contexts()[0];
      const page = context.pages()[0] ?? await context.waitForEvent('page');
      await expect(page.locator('.qa-doc-title')).toHaveText(/people/);
      await expect(page.locator('.qa-grid [data-rc="1:0"]')).toHaveText('Ann');
    } finally {
      proc.kill('SIGKILL');
    }
  });
});
