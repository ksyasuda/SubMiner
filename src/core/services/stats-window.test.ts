import assert from 'node:assert/strict';
import test from 'node:test';
import {
  buildStatsWindowLoadFileOptions,
  buildStatsWindowOptions,
  presentStatsWindow,
  promoteVisibleStatsWindowAboveOverlay,
  promoteStatsWindowLevel,
  resolveStatsWindowOuterBoundsForContent,
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

  assert.equal(options.title, 'SubMiner Stats');
  assert.equal(options.x, 120);
  assert.equal(options.y, 80);
  assert.equal(options.width, 1440);
  assert.equal(options.height, 900);
  assert.equal(options.frame, false);
  assert.equal(options.transparent, false);
  assert.equal(options.backgroundColor, '#24273a');
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

test('buildStatsWindowLoadFileOptions includes provided stats API base URL', () => {
  assert.deepEqual(buildStatsWindowLoadFileOptions('http://127.0.0.1:6123'), {
    query: {
      overlay: '1',
      apiBase: 'http://127.0.0.1:6123',
    },
  });
});

test('resolveStatsWindowOuterBoundsForContent compensates for Wayland content insets', () => {
  assert.deepEqual(
    resolveStatsWindowOuterBoundsForContent(
      {
        getBounds: () => ({ x: 0, y: 0, width: 3440, height: 1440 }),
        getContentBounds: () => ({ x: 0, y: 14, width: 3440, height: 1426 }),
      },
      { x: 0, y: 0, width: 3440, height: 1440 },
    ),
    { x: 0, y: -14, width: 3440, height: 1454 },
  );
});

test('resolveStatsWindowOuterBoundsForContent ignores invalid inset geometry', () => {
  const target = { x: 0, y: 0, width: 3440, height: 1440 };
  assert.deepEqual(
    resolveStatsWindowOuterBoundsForContent(
      {
        getBounds: () => ({ x: 0, y: 0, width: 3440, height: 1440 }),
        getContentBounds: () => ({ x: -1, y: 0, width: 3440, height: 1440 }),
      },
      target,
    ),
    target,
  );
});

test('promoteStatsWindowLevel raises stats above overlay level on macOS', () => {
  const calls: string[] = [];
  promoteStatsWindowLevel(
    {
      setAlwaysOnTop: (flag: boolean, level?: string, relativeLevel?: number) => {
        calls.push(`always-on-top:${flag}:${level ?? 'none'}:${relativeLevel ?? 0}`);
      },
      setVisibleOnAllWorkspaces: (
        visible: boolean,
        options?: { visibleOnFullScreen?: boolean },
      ) => {
        calls.push(
          `all-workspaces:${visible}:${options?.visibleOnFullScreen === true ? 'fullscreen' : 'plain'}`,
        );
      },
      setFullScreenable: (fullscreenable: boolean) => {
        calls.push(`fullscreenable:${fullscreenable}`);
      },
      moveTop: () => {
        calls.push('move-top');
      },
    } as never,
    'darwin',
  );

  assert.deepEqual(calls, [
    'always-on-top:true:screen-saver:2',
    'all-workspaces:true:fullscreen',
    'fullscreenable:false',
    'move-top',
  ]);
});

test('promoteStatsWindowLevel raises stats above overlay level on Windows', () => {
  const calls: string[] = [];
  promoteStatsWindowLevel(
    {
      setAlwaysOnTop: (flag: boolean, level?: string, relativeLevel?: number) => {
        calls.push(`always-on-top:${flag}:${level ?? 'none'}:${relativeLevel ?? 0}`);
      },
      moveTop: () => {
        calls.push('move-top');
      },
    } as never,
    'win32',
  );

  assert.deepEqual(calls, ['always-on-top:true:screen-saver:2', 'move-top']);
});

test('promoteVisibleStatsWindowAboveOverlay reasserts stats above overlay on Hyprland', () => {
  const calls: string[] = [];
  const promoted = promoteVisibleStatsWindowAboveOverlay(
    {
      isDestroyed: () => false,
      isVisible: () => true,
      setAlwaysOnTop: (flag: boolean, level?: string, relativeLevel?: number) => {
        calls.push(`always-on-top:${flag}:${level ?? 'none'}:${relativeLevel ?? 0}`);
      },
      moveTop: () => {
        calls.push('move-top');
      },
    } as never,
    {
      platform: 'linux',
      promoteHyprlandWindow: () => calls.push('hyprland-top'),
    },
  );

  assert.equal(promoted, true);
  assert.deepEqual(calls, ['always-on-top:true:none:0', 'move-top', 'hyprland-top']);
});

test('promoteVisibleStatsWindowAboveOverlay skips hidden stats windows', () => {
  const calls: string[] = [];
  const promoted = promoteVisibleStatsWindowAboveOverlay(
    {
      isDestroyed: () => false,
      isVisible: () => false,
      setAlwaysOnTop: () => calls.push('always-on-top'),
      moveTop: () => calls.push('move-top'),
    } as never,
    {
      promoteHyprlandWindow: () => calls.push('hyprland-top'),
    },
  );

  assert.equal(promoted, false);
  assert.deepEqual(calls, []);
});

test('presentStatsWindow shows inactive on macOS to stay on the fullscreen mpv Space', () => {
  const calls: string[] = [];

  presentStatsWindow(
    {
      show: () => {
        calls.push('show');
      },
      showInactive: () => {
        calls.push('show-inactive');
      },
      focus: () => {
        calls.push('focus');
      },
    } as never,
    'darwin',
  );

  assert.deepEqual(calls, ['show-inactive']);
});

test('presentStatsWindow shows and focuses on non-macOS platforms', () => {
  const calls: string[] = [];

  presentStatsWindow(
    {
      show: () => {
        calls.push('show');
      },
      showInactive: () => {
        calls.push('show-inactive');
      },
      focus: () => {
        calls.push('focus');
      },
    } as never,
    'linux',
  );

  assert.deepEqual(calls, ['show', 'focus']);
});
