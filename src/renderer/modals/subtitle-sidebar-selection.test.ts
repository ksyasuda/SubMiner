import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { promisify } from 'node:util';
import test from 'node:test';
import { build } from 'esbuild';

// This check opens Electron and uses the clipboard. Keep it out of normal code-only lanes.
const electronTest =
  process.env.SUBMINER_ELECTRON_TESTS === '1' &&
  (process.platform !== 'linux' || process.env.DISPLAY)
    ? test
    : test.skip;

electronTest(
  'sidebar selection copies clean chronological text without seeking or losing context',
  {
    timeout: 30_000,
  },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'subminer-sidebar-selection-'));
    await build({
      entryPoints: [resolve('src/renderer/modals/subtitle-sidebar-selection.electron-fixture.ts')],
      bundle: true,
      platform: 'browser',
      format: 'iife',
      globalName: 'sidebarTest',
      outfile: join(dir, 'fixture.js'),
    });
    const html = (await readFile('src/renderer/index.html', 'utf8'))
      .replace(
        '<script type="module" src="renderer.js"></script>',
        '<script src="fixture.js"></script>',
      )
      .replace(
        'href="style.css"',
        `href="${new URL(`file://${resolve('src/renderer/style.css')}`).href}"`,
      );
    await writeFile(join(dir, 'index.html'), html);
    await writeFile(
      join(dir, 'clipboard.cjs'),
      'const { clipboard, contextBridge } = require("electron"); contextBridge.exposeInMainWorld("copyTestSelection", text => clipboard.writeText(text));',
    );
    await build({
      entryPoints: [resolve('src/core/services/overlay-window-input.ts')],
      bundle: true,
      platform: 'node',
      format: 'cjs',
      outfile: join(dir, 'input.cjs'),
    });
    await writeFile(
      join(dir, 'run.cjs'),
      `
const { app, BrowserWindow, clipboard } = require('electron');
const assert = require('node:assert/strict');
const { handleOverlayWindowBeforeInputEvent } = require('./input.cjs');
app.setPath('userData', ${JSON.stringify(join(dir, 'user-data'))});
app.whenReady().then(async () => {
  const window = new BrowserWindow({ width: 900, height: 700, show: false, webPreferences: { preload: ${JSON.stringify(join(dir, 'clipboard.cjs'))}, sandbox: false } });
  window.webContents.on('console-message', (_event, details) => {
    if (details.level === 'error') console.error(details.message);
  });
  let intercepted = 0;
  window.webContents.on('before-input-event', (event, input) => handleOverlayWindowBeforeInputEvent({
    kind: 'visible', windowVisible: true, input,
    preventDefault: () => event.preventDefault(),
    sendKeyboardModeToggleRequested() {}, sendLookupWindowToggleRequested() {}, forwardTabToMpv() {},
    tryHandleOverlayShortcutLocalFallback() { intercepted++; return true; },
  }));
  await window.loadFile(${JSON.stringify(join(dir, 'index.html'))});
  const run = (code) => window.webContents.executeJavaScript(code, true);
  await run('sidebarTest.setup().then(checks => { window.checks = checks; })');
  const expected = { text: 'の台詞\\n\\n同じ台詞\\n二行目\\n\\n同じ', cueCount: 3 };
  assert.deepEqual(await run('checks.select()'), expected);
  assert.deepEqual(await run('checks.select(true)'), expected);
  assert.equal(await run('checks.buttonVisible()'), true);
  assert.equal(await run('checks.clickCue()'), 0);
  assert.equal(await run('checks.updatePlayback()'), 0);
  assert.deepEqual(await run('checks.selected()'), expected);
  const previousClipboard = clipboard.readText();
  try {
    clipboard.writeText('sentinel');
    window.show(); app.focus({ steal: true }); window.focus(); window.webContents.focus();
    await new Promise(resolve => setTimeout(resolve, 100));
    await run('checks.clear()');
    const [start, end] = await run('checks.dragPoints()');
    window.webContents.sendInputEvent({ type: 'mouseDown', ...start, button: 'left', clickCount: 1 });
    window.webContents.sendInputEvent({ type: 'mouseMove', ...end, button: 'left' });
    window.webContents.sendInputEvent({ type: 'mouseUp', ...end, button: 'left', clickCount: 1 });
    await new Promise(resolve => setTimeout(resolve, 100));
    assert.deepEqual(await run('checks.selected()'), expected);
    window.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'C', modifiers: [process.platform === 'darwin' ? 'meta' : 'control'] });
    window.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'C', modifiers: [process.platform === 'darwin' ? 'meta' : 'control'] });
    await new Promise(resolve => setTimeout(resolve, 100));
    assert.equal(intercepted, 0);
    assert.equal(await run('checks.fallbackCopies()'), 0);
    assert.equal(clipboard.readText() === expected.text, true, 'Keyboard copies the selected excerpt');
    clipboard.writeText('sentinel');
    await run('checks.clickCopy()');
    await new Promise(resolve => setTimeout(resolve, 100));
    assert.equal(clipboard.readText() === expected.text, true, 'Button copies the selected excerpt');
  } finally { clipboard.writeText(previousClipboard); }
  await run('document.dispatchEvent(new KeyboardEvent("keydown", {key:"Escape", bubbles:true}))');
  assert.equal(await run('checks.selected()'), null);
  assert.equal(await run('checks.buttonVisible()'), false);
  assert.equal(await run('checks.clickCue()'), 1);
  await run('document.dispatchEvent(new KeyboardEvent("keydown", {key:"c", ctrlKey:true, bubbles:true}))');
  assert.equal(await run('checks.fallbackCopies()'), 1);
  await run('checks.select()');
  assert.equal(await run('checks.changeSource()'), null);
  await run('checks.select(); checks.close()');
  assert.equal(await run('checks.selected()'), null);
  window.destroy();
  console.log('SIDEBAR_SELECTION_OK');
  app.quit();
}).catch(error => { console.error(error); app.exit(1); });
`,
    );
    const env = { ...process.env };
    delete env.ELECTRON_RUN_AS_NODE;
    const { stdout } = await promisify(execFile)(
      resolve('node_modules/.bin/electron'),
      [join(dir, 'run.cjs')],
      { env, timeout: 25_000 },
    );
    assert.match(stdout, /SIDEBAR_SELECTION_OK/);
  },
);
