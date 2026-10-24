// Run with the pinned Electron runtime against a finished app's resources folder.
const { app, BrowserWindow, session } = require('electron');
const fs = require('node:fs');
const http = require('node:http');
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
  const hachidori = await session
    .fromPartition('persist:hachidori')
    .extensions.loadExtension(path.join(resources, 'hachidori'), { allowFileAccess: true });
  assert(hachidori.id, 'Hachidori extension failed to load');
  const failedRequests = [];
  session.defaultSession.webRequest.onErrorOccurred(
    { urls: ['file://*/*', 'http://127.0.0.1/*'] },
    (details) => {
      // Chromium probes the cache before fetching @font-face fonts; an uncached
      // font reports ERR_CACHE_MISS and is then fetched normally.
      if (!['net::ERR_ABORTED', 'net::ERR_CACHE_MISS'].includes(details.error))
        failedRequests.push(`${details.url}: ${details.error}`);
    },
  );
  for (const ui of ['renderer', 'settings', 'syncui']) {
    const win = new BrowserWindow({
      show: false,
      webPreferences: {
        sandbox: false,
        preload: path.join(archive, 'dist', ui === 'renderer' ? 'preload.js' : `preload-${ui}.js`),
      },
    });
    try {
      await win.loadFile(path.join(archive, `dist/${ui}/index.html`));
      const loaded = await win.webContents.executeJavaScript(
        `document.fonts.load('400 16px "M PLUS 1"', '日本語').then(fonts => fonts.length > 0 && fonts.every(font => font.status === 'loaded'))`,
      );
      assert(loaded, `${ui}: shared Japanese font failed to load`);
    } finally {
      win.destroy();
    }
  }
  // The stats dashboard uses HTTP for both assets and API requests in the app.
  const { ImmersionTrackerService } = packagedRequire(
    './dist/core/services/immersion-tracker-service.js',
  );
  const { createStatsApp, startNodeHttpServer } = packagedRequire(
    './dist/core/services/stats-server.js',
  );
  const tracker = new ImmersionTrackerService({ dbPath: path.join(isolatedData, 'stats.db') });
  const statsConfig = { port: 0, staticDir: path.join(archive, 'stats/dist'), tracker };
  let statsHttp;
  const statsServer = await startNodeHttpServer(
    createStatsApp(tracker, statsConfig),
    statsConfig,
    (listener) => (statsHttp = http.createServer(listener)),
  );
  const statsWindow = new BrowserWindow({ show: false });
  try {
    const url = `http://127.0.0.1:${statsHttp.address().port}`;
    session.defaultSession.webRequest.onCompleted({ urls: [`${url}/*`] }, (details) => {
      if (details.statusCode >= 400) failedRequests.push(`${details.url}: ${details.statusCode}`);
    });
    await statsWindow.loadURL(url);
    // Let in-flight font requests settle before the window goes away.
    await statsWindow.webContents.executeJavaScript('document.fonts.ready.then(() => true)');
    for (const endpoint of ['overview', 'sessions']) {
      const response = await fetch(`${url}/api/stats/${endpoint}`);
      assert.equal(response.status, 200, `Stats ${endpoint} request failed`);
      await response.json();
    }
  } finally {
    statsWindow.destroy();
    await statsServer.close();
    tracker.destroy();
  }
  assert.deepEqual(failedRequests, [], 'Packaged UI resources failed to load');
  console.log(
    'Package smoke passed: SQLite, platform FFI, texthooker, both dictionary extensions, UI pages, stats HTTP, shared Japanese font.',
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
