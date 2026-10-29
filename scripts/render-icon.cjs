// Renders build/icon.svg to build/icon.png (1024px), from which electron-builder makes
// the .icns and .ico. Run after editing the SVG:  npx electron scripts/render-icon.cjs
const { app, BrowserWindow } = require('electron');
const fs = require('fs');
const path = require('path');

const dir = path.join(__dirname, '../build');
setTimeout(() => { console.error('Timed out rendering the icon'); app.exit(1); }, 20000);

app.whenReady().then(() => {
  const win = new BrowserWindow({ width: 1024, height: 1024, show: false, transparent: true, frame: false, useContentSize: true, webPreferences: { offscreen: true } });
  let loaded = false;
  // Offscreen windows hand over frames through 'paint'; take the first one after load.
  win.webContents.on('paint', (_e, _dirty, image) => {
    if (!loaded || image.isEmpty() || image.getSize().width < 1024) return;
    const png = image.resize({ width: 1024, height: 1024 }).toPNG();
    fs.writeFileSync(path.join(dir, 'icon.png'), png);
    console.log(`Wrote build/icon.png (${png.length} bytes)`);
    app.exit(0);
  });
  win.webContents.once('did-finish-load', () => { loaded = true; win.webContents.invalidate(); });
  win.loadFile(path.join(dir, 'icon.svg'));
});
