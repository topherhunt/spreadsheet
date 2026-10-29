// electron-builder afterSign hook: fails the build unless the packaged .app carries a
// valid signature that seals the whole bundle.
//
// If electron-builder skips signing, the app keeps the stock Electron binary's ad-hoc
// signature while Info.plist and Resources change around it. macOS then calls the app
// "damaged and can't be opened", with no "Open Anyway". `mac.identity: "-"` makes
// electron-builder sign the finished bundle ad-hoc; this catches it if that stops.

const { execFileSync, spawnSync } = require('child_process');
const path = require('path');

function verifyAppSignature(appPath) {
  execFileSync('codesign', ['--verify', '--deep', '--strict', appPath], { stdio: 'inherit' });

  // --verify alone passes on a bundle whose Info.plist and resources were never sealed.
  // `codesign -d` reports on stderr.
  const result = spawnSync('codesign', ['-dvvv', appPath], { encoding: 'utf8' });
  if (result.status !== 0) throw new Error(`codesign -dvvv failed for ${appPath}: ${result.stderr}`);
  const details = result.stdout + result.stderr;
  if (details.includes('Info.plist=not bound')) throw new Error(`${appPath}: Info.plist is not bound to the code signature; macOS will report this app as damaged.`);
  if (details.includes('Sealed Resources=none')) throw new Error(`${appPath}: bundle resources are not sealed; macOS will report this app as damaged.`);
}

function afterSign(context) {
  if (context.electronPlatformName !== 'darwin') return;
  const appPath = path.join(context.appOutDir, `${context.packager.appInfo.productFilename}.app`);
  verifyAppSignature(appPath);
  console.log(`  • code signature verified  app=${appPath}`);
}

module.exports = { afterSign, verifyAppSignature };
