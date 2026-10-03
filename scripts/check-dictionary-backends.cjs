// Exercises the real entry point with disposable Linux config and data directories.
const { app, BrowserWindow, session, shell } = require('electron');
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
if (process.platform !== 'linux')
  throw new Error('This app-entry smoke requires Linux XDG isolation.');
const root = process.cwd();
const backend = process.argv.includes('--backend=yomitan') ? 'yomitan' : 'hachidori';
const externalYomitanProfile =
  backend === 'hachidori' && !process.argv.includes('--bundled-yomitan');
const profile = process.env.SUBMINER_DICTIONARY_SMOKE_DATA;
assert(
  profile && fs.existsSync(profile),
  'Use bun run test:dictionary:electron for profile isolation',
);
process.env.XDG_CONFIG_HOME = profile;
process.env.XDG_DATA_HOME = path.join(profile, 'data');
fs.mkdirSync(path.join(profile, 'SubMiner'));
fs.writeFileSync(
  path.join(profile, 'SubMiner', 'config.json'),
  JSON.stringify({
    dictionaryBackend: backend,
    // External Yomitan profiles must not block Hachidori settings.
    yomitan: {
      externalProfilePath: externalYomitanProfile ? path.join(profile, 'external-yomitan') : '',
    },
    mpv: { socketPath: path.join(profile, 'missing-mpv.sock') },
    ankiConnect: { enabled: false },
    startupWarmups: { lowPowerMode: true },
    discordPresence: { enabled: false },
    updates: { enabled: false },
  }),
);
const openedLinks = [];
shell.openExternal = async (url) => {
  openedLinks.push(url);
};
app.setAppPath(root);
app.getVersion = () => require(path.join(root, 'package.json')).version;
process.env.SUBMINER_APP_LOG = path.join(profile, 'app.log');
app.commandLine.appendSwitch('ozone-platform', 'x11');
app.commandLine.appendSwitch('disable-gpu');
app.commandLine.appendSwitch('disable-dev-shm-usage');
process.argv = [process.execPath, root, '--hachidori', '--log-level', 'debug'];
require(path.join(root, 'dist/main-entry.js'));
const deadline = setTimeout(() => {
  console.error(
    'FAIL timeout',
    BrowserWindow.getAllWindows().map((w) => w.webContents.getURL()),
  );
  finish(1);
}, 60000);
function finish(exitCode) {
  clearTimeout(deadline);
  app.exit(exitCode);
}
(async () => {
  await app.whenReady();
  let window;
  for (let i = 0; i < 300; i++) {
    window = BrowserWindow.getAllWindows().find(
      (w) => w.getTitle().includes('Hachidori') && w.isVisible(),
    );
    if (window) break;
    await new Promise((r) => setTimeout(r, 100));
  }
  assert.ok(window, 'Hachidori settings window opens from actual app flag');
  assert.equal(window.webContents.session, session.fromPartition('persist:hachidori'));
  assert.match(window.webContents.getURL(), /settings.html/);
  const status = await window.webContents.executeJavaScript(
    `chrome.runtime.sendMessage({target:'hoshidicts-offscreen',type:'hd_status',requestId:'app-settings-smoke'})`,
  );
  assert.equal(status.ok, true);
  assert.equal(status.ready, true);

  console.log(
    'PASS actual --hachidori startup, visible settings, isolated backend session, native engine ready',
  );
  app.emit('second-instance', {}, [process.execPath, root, '--yomitan'], root);
  if (externalYomitanProfile) {
    const blockedMessage =
      'Yomitan settings unavailable while using read-only external-profile mode.';
    for (let i = 0; i < 100; i++) {
      if (fs.readFileSync(process.env.SUBMINER_APP_LOG, 'utf8').includes(blockedMessage)) break;
      await new Promise((r) => setTimeout(r, 100));
    }
    assert.ok(
      fs.readFileSync(process.env.SUBMINER_APP_LOG, 'utf8').includes(blockedMessage),
      '--yomitan is rejected when an external Yomitan profile is configured',
    );
    assert.equal(session.defaultSession.extensions.getAllExtensions().length, 0);
    assert.ok(!BrowserWindow.getAllWindows().some((w) => w.getTitle().includes('Yomitan')));
    console.log('PASS external Yomitan settings blocked while Hachidori settings remain available');
  } else {
    let yomi;
    for (let i = 0; i < 200; i++) {
      yomi = BrowserWindow.getAllWindows().find(
        (w) => w.getTitle().includes('Yomitan') && w.isVisible(),
      );
      if (yomi) break;
      await new Promise((r) => setTimeout(r, 100));
    }
    assert.ok(yomi, 'inactive --yomitan settings opens');
    assert.equal(yomi.webContents.session, session.defaultSession);
  }
  assert.equal(window.webContents.session, session.fromPartition('persist:hachidori'));
  app.emit('second-instance', {}, [process.execPath, root, '--toggle-visible-overlay'], root);
  let overlay;
  for (let i = 0; i < 200; i++) {
    overlay = BrowserWindow.getAllWindows().find((w) =>
      w.webContents.getURL().includes('/renderer/index.html?'),
    );
    if (overlay && !overlay.webContents.isLoading()) break;
    await new Promise((r) => setTimeout(r, 100));
  }
  assert.ok(overlay, 'actual overlay initialized');
  assert.equal(
    overlay.webContents.session,
    backend === 'hachidori' ? session.fromPartition('persist:hachidori') : session.defaultSession,
  );
  const requestLink = (url) =>
    overlay.webContents.executeJavaScript(`new Promise((resolve,reject)=>{
  const timer=setTimeout(()=>reject(Error('No external link acknowledgment')),3000);
  window.addEventListener('hachidori-open-external-result',e=>{clearTimeout(timer);resolve(e.detail);},{once:true});
  window.dispatchEvent(new CustomEvent('hachidori-open-external',{detail:{requestId:'smoke-link',url:${JSON.stringify(url)}}}));
 })`);
  assert.equal((await requestLink('file:///tmp/private')).ok, false);
  assert.equal((await requestLink('https://example.com/word')).ok, backend === 'hachidori');
  assert.deepEqual(openedLinks, backend === 'hachidori' ? ['https://example.com/word'] : []);
  console.log(
    `PASS ${backend} overlay session, independent settings windows, real preload external link bridge`,
  );

  finish(0);
})().catch((error) => {
  console.error(error);
  finish(1);
});
