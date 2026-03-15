import assert from 'node:assert/strict';
import test from 'node:test';
import {
  buildStatsWindowLoadFileOptions,
  buildStatsWindowOptions,
  shouldHideStatsWindowForInput,
} from './stats-window-runtime';

test('buildStatsWindowOptions uses tracked overlay bounds and preload-friendly web preferences', () => {
  const options = buildStatsWindowOptions({
    preloadPath: '/tmp/preload-stats.js',
    bounds: {
      x: 120,
      y: 80,
      width: 1440,
      height: 900,
    },
  });

  assert.equal(options.x, 120);
  assert.equal(options.y, 80);
  assert.equal(options.width, 1440);
  assert.equal(options.height, 900);
  assert.equal(options.frame, false);
  assert.equal(options.transparent, true);
  assert.equal(options.resizable, false);
  assert.equal(options.webPreferences?.preload, '/tmp/preload-stats.js');
  assert.equal(options.webPreferences?.contextIsolation, true);
  assert.equal(options.webPreferences?.nodeIntegration, false);
  assert.equal(options.webPreferences?.sandbox, true);
});

test('shouldHideStatsWindowForInput matches Escape and configured bare toggle key', () => {
  assert.equal(
    shouldHideStatsWindowForInput(
      {
        type: 'keyDown',
        key: 'Escape',
        code: 'Escape',
      } as Electron.Input,
      'Backquote',
    ),
    true,
  );

  assert.equal(
    shouldHideStatsWindowForInput(
      {
        type: 'keyDown',
        key: '`',
        code: 'Backquote',
      } as Electron.Input,
      'Backquote',
    ),
    true,
  );

  assert.equal(
    shouldHideStatsWindowForInput(
      {
        type: 'keyDown',
        key: '`',
        code: 'Backquote',
        control: true,
      } as Electron.Input,
      'Backquote',
    ),
    false,
  );

  assert.equal(
    shouldHideStatsWindowForInput(
      {
        type: 'keyDown',
        key: '`',
        code: 'Backquote',
        alt: true,
      } as Electron.Input,
      'Backquote',
    ),
    false,
  );

  assert.equal(
    shouldHideStatsWindowForInput(
      {
        type: 'keyDown',
        key: '`',
        code: 'Backquote',
        meta: true,
      } as Electron.Input,
      'Backquote',
    ),
    false,
  );

  assert.equal(
    shouldHideStatsWindowForInput(
      {
        type: 'keyDown',
        key: '`',
        code: 'Backquote',
        isAutoRepeat: true,
      } as Electron.Input,
      'Backquote',
    ),
    false,
  );

  assert.equal(
    shouldHideStatsWindowForInput(
      {
        type: 'keyDown',
        key: '`',
        code: 'Backquote',
        shift: true,
      } as Electron.Input,
      'Backquote',
    ),
    false,
  );

  assert.equal(
    shouldHideStatsWindowForInput(
      {
        type: 'keyUp',
        key: '`',
        code: 'Backquote',
      } as Electron.Input,
      'Backquote',
    ),
    false,
  );
});

test('buildStatsWindowLoadFileOptions enables overlay rendering mode', () => {
  assert.deepEqual(buildStatsWindowLoadFileOptions(), {
    query: {
      overlay: '1',
    },
  });
});
