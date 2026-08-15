import assert from 'node:assert/strict';
import test from 'node:test';
import {
  buildStatsWindowLoadFileOptions,
  buildStatsWindowOptions,
  buildStatsNativeConfirmDialogOptions,
  demoteVisibleStatsWindowBelowDialogs,
  presentStatsWindow,
  promoteVisibleStatsWindowAboveOverlay,
  promoteStatsWindowLevel,
  resolveStatsWindowOuterBoundsForContent,
  scheduleStatsWindowPostShowReconciles,
  showStatsNativeConfirmDialog,
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

test('buildStatsWindowOptions uses a fullscreen auxiliary panel on macOS', () => {
  const options = buildStatsWindowOptions({
    preloadPath: '/tmp/preload-stats.js',
    platform: 'darwin',
  });

  assert.equal(options.type, 'panel');
});

test('buildStatsWindowOptions remains a regular window off macOS', () => {
  const options = buildStatsWindowOptions({
    preloadPath: '/tmp/preload-stats.js',
    platform: 'linux',
  });

  assert.equal(options.type, undefined);
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

test('demoteVisibleStatsWindowBelowDialogs lowers visible stats below native dialogs', () => {
  const calls: string[] = [];
  const demoted = demoteVisibleStatsWindowBelowDialogs({
    isDestroyed: () => false,
    isVisible: () => true,
    setAlwaysOnTop: (flag: boolean, level?: string, relativeLevel?: number) => {
      calls.push(`always-on-top:${flag}:${level ?? 'none'}:${relativeLevel ?? 0}`);
    },
  } as never);

  assert.equal(demoted, true);
  assert.deepEqual(calls, ['always-on-top:false:none:0']);
});

test('demoteVisibleStatsWindowBelowDialogs skips hidden stats windows', () => {
  const calls: string[] = [];
  const demoted = demoteVisibleStatsWindowBelowDialogs({
    isDestroyed: () => false,
    isVisible: () => false,
    setAlwaysOnTop: () => calls.push('always-on-top'),
  } as never);

  assert.equal(demoted, false);
  assert.deepEqual(calls, []);
});

test('buildStatsNativeConfirmDialogOptions makes delete the explicit destructive action', () => {
  assert.deepEqual(buildStatsNativeConfirmDialogOptions('Delete this session?'), {
    type: 'warning',
    message: 'Delete this session?',
    buttons: ['Delete', 'Cancel'],
    defaultId: 1,
    cancelId: 1,
    noLink: true,
  });
});

test('showStatsNativeConfirmDialog parents the native dialog to live stats windows', () => {
  const calls: string[] = [];
  const parent = { isDestroyed: () => false };

  const confirmed = showStatsNativeConfirmDialog(parent, 'Delete this session?', {
    showWithParent: (window, options) => {
      assert.equal(window, parent);
      calls.push(`${options.message}:${options.defaultId}:${options.cancelId}`);
      return 0;
    },
    showWithoutParent: () => {
      calls.push('unparented');
      return 1;
    },
  });

  assert.equal(confirmed, true);
  assert.deepEqual(calls, ['Delete this session?:1:1']);
});

test('showStatsNativeConfirmDialog treats cancel as not confirmed', () => {
  const confirmed = showStatsNativeConfirmDialog({ isDestroyed: () => false }, 'Delete?', {
    showWithParent: () => 1,
    showWithoutParent: () => 0,
  });

  assert.equal(confirmed, false);
});

test('showStatsNativeConfirmDialog falls back to an unparented dialog without a live stats window', () => {
  const calls: string[] = [];

  const confirmed = showStatsNativeConfirmDialog({ isDestroyed: () => true }, 'Delete?', {
    showWithParent: () => {
      calls.push('parented');
      return 0;
    },
    showWithoutParent: (options) => {
      calls.push(options.message);
      return 0;
    },
  });

  assert.equal(confirmed, true);
  assert.deepEqual(calls, ['Delete?']);
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

test('scheduleStatsWindowPostShowReconciles retries placement after a reused hidden window is remapped', () => {
  const calls: string[] = [];

  scheduleStatsWindowPostShowReconciles(
    () => {
      calls.push('reconcile');
    },
    {
      setTimeout: (callback, delayMs) => {
        calls.push(`timer:${delayMs}`);
        callback();
        return {
          unref: () => {
            calls.push(`unref:${delayMs}`);
          },
        };
      },
    },
  );

  assert.deepEqual(calls, [
    'timer:50',
    'reconcile',
    'unref:50',
    'timer:150',
    'reconcile',
    'unref:150',
    'timer:300',
    'reconcile',
    'unref:300',
    'timer:600',
    'reconcile',
    'unref:600',
  ]);
});
