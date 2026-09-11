// Run with the pinned Electron runtime against a finished app's resources folder.
const { app, BrowserWindow, session } = require('electron');
const fs = require('node:fs');
const path = require('node:path');
const { createRequire } = require('node:module');
const assert = require('node:assert/strict');
const { once } = require('node:events');

const resources = path.resolve(process.argv[2]);
const archive = path.join(resources, 'app.asar');
const isolatedData = process.env.SUBMINER_PACKAGE_SMOKE_DATA;
assert(
  isolatedData && fs.existsSync(isolatedData),
  'Use bun run test:package to create an isolated profile',
);
app.setPath('userData', isolatedData);
app.disableHardwareAcceleration();
app.on('window-all-closed', () => {});
const timeout = setTimeout(() => {
  console.error('Package smoke timed out');
  app.exit(1);
}, 60_000);

async function smoke() {
  await app.whenReady();
  const packagedRequire = createRequire(path.join(archive, 'package.json'));
  const Database = packagedRequire('libsql');
  const database = new Database(':memory:');
  assert.equal(database.prepare('select 42 as answer').get().answer, 42);
  database.close();
  if (process.platform === 'win32') {
    const win32 = packagedRequire('./dist/window-trackers/win32.js');
    assert(Array.isArray(win32.findMpvWindows().matches));
  }
  const { Texthooker } = packagedRequire('./dist/core/services/texthooker.js');
  const texthooker = new Texthooker();
  const server = texthooker.start(0);
  assert(server, 'Packaged texthooker assets could not be found');
  try {
    await once(server, 'listening');
    const response = await fetch(`http://127.0.0.1:${server.address().port}/`);
    assert.equal(response.status, 200);
    assert((await response.text()).includes('<html'));
  } finally {
    texthooker.stop();
  }
  const extension = await session.defaultSession.extensions.loadExtension(
    path.join(resources, 'yomitan'),
    { allowFileAccess: true },
  );
  assert(extension.id, 'Yomitan extension failed to load');
  const failedRequests = [];
  session.defaultSession.webRequest.onErrorOccurred({ urls: ['file://*/*'] }, (details) => {
    if (details.error !== 'net::ERR_ABORTED')
      failedRequests.push(`${details.url}: ${details.error}`);
  });
  for (const ui of ['renderer', 'settings', 'syncui', 'stats']) {
    const win = new BrowserWindow({
      show: false,
      webPreferences: {
        sandbox: false,
        preload: path.join(archive, 'dist', ui === 'renderer' ? 'preload.js' : `preload-${ui}.js`),
      },
    });
    try {
      await win.loadFile(
        path.join(archive, ui === 'stats' ? 'stats/dist/index.html' : `dist/${ui}/index.html`),
      );
      if (ui !== 'stats') {
        const loaded = await win.webContents.executeJavaScript(
          `document.fonts.load('400 16px "M PLUS 1"', '日本語').then(fonts => fonts.length > 0 && fonts.every(font => font.status === 'loaded'))`,
        );
        assert(loaded, `${ui}: shared Japanese font failed to load`);
      }
    } finally {
      win.destroy();
    }
  }
  assert.deepEqual(failedRequests, [], 'Packaged UI resources failed to load');
  console.log(
    'Package smoke passed: SQLite, platform FFI, texthooker, Yomitan loading, UI pages, shared Japanese font.',
  );
}

smoke()
  .then(() => {
    clearTimeout(timeout);
    app.exit(0);
  })
  .catch((error) => {
    console.error(error);
    clearTimeout(timeout);
    app.exit(1);
  });
