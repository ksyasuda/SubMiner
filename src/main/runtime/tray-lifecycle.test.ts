import test from 'node:test';
import assert from 'node:assert/strict';
import { createDestroyTrayHandler, createEnsureTrayHandler } from './tray-lifecycle';

test('ensure tray updates menu when tray already exists', () => {
  const calls: string[] = [];
  const tray = {
    setContextMenu: () => calls.push('set-menu'),
    setToolTip: () => calls.push('set-tooltip'),
    on: () => calls.push('bind-click'),
    destroy: () => calls.push('destroy'),
  };

  const ensureTray = createEnsureTrayHandler({
    getTray: () => tray,
    setTray: () => calls.push('set-tray'),
    buildTrayMenu: () => ({}),
    resolveTrayIconPath: () => null,
    createImageFromPath: () =>
      ({
        isEmpty: () => false,
        resize: () => {
          throw new Error('should not resize');
        },
        setTemplateImage: () => {},
      }) as never,
    createEmptyImage: () =>
      ({
        isEmpty: () => true,
        resize: () => {
          throw new Error('should not resize');
        },
        setTemplateImage: () => {},
      }) as never,
    createTray: () => tray as never,
    trayTooltip: 'SubMiner',
    platform: 'darwin',
    logWarn: () => calls.push('warn'),
    ensureOverlayVisibleFromTrayClick: () => calls.push('show-overlay'),
  });

  ensureTray();
  assert.deepEqual(calls, ['set-menu']);
});

test('ensure tray refreshes existing tray menu on linux with setContextMenu', () => {
  const calls: string[] = [];
  let trayRef: unknown = {
    setContextMenu: () => calls.push('old-set-menu'),
    setToolTip: () => calls.push('old-set-tooltip'),
    on: () => calls.push('old-bind-click'),
    destroy: () => calls.push('old-destroy'),
  };

  const ensureTray = createEnsureTrayHandler({
    getTray: () => trayRef as never,
    setTray: (tray) => {
      trayRef = tray;
      calls.push(tray ? 'set-new-tray' : 'clear-tray');
    },
    buildTrayMenu: () => ({ id: 'menu' }),
    resolveTrayIconPath: () => '/tmp/icon.png',
    createImageFromPath: () =>
      ({
        isEmpty: () => false,
        resize: (options: { width: number; height: number }) => {
          calls.push(`resize:${options.width}x${options.height}`);
          return {
            isEmpty: () => false,
            resize: () => {
              throw new Error('unexpected');
            },
            setTemplateImage: () => {},
          };
        },
        setTemplateImage: () => {},
      }) as never,
    createEmptyImage: () =>
      ({
        isEmpty: () => true,
        resize: () => {
          throw new Error('unexpected');
        },
        setTemplateImage: () => {},
      }) as never,
    createTray: () =>
      ({
        setContextMenu: () => calls.push('new-set-menu'),
        setToolTip: () => calls.push('new-set-tooltip'),
        on: () => calls.push('new-bind-click'),
        destroy: () => calls.push('new-destroy'),
      }) as never,
    trayTooltip: 'SubMiner',
    platform: 'linux',
    logWarn: () => calls.push('warn'),
    ensureOverlayVisibleFromTrayClick: () => calls.push('show-overlay'),
  });

  ensureTray();
  assert.deepEqual(calls, ['old-set-menu']);
});

test('ensure tray creates new tray and binds click handler', () => {
  const calls: string[] = [];
  let trayRef: unknown = null;

  const ensureTray = createEnsureTrayHandler({
    getTray: () => null,
    setTray: (tray) => {
      trayRef = tray;
      calls.push('set-tray');
    },
    buildTrayMenu: () => ({ id: 'menu' }),
    resolveTrayIconPath: () => '/tmp/icon.png',
    createImageFromPath: () =>
      ({
        isEmpty: () => false,
        resize: (options: {
          width: number;
          height: number;
          quality?: 'best' | 'better' | 'good';
        }) => {
          calls.push(`resize:${options.width}x${options.height}`);
          return {
            isEmpty: () => false,
            resize: () => {
              throw new Error('unexpected');
            },
            setTemplateImage: () => calls.push('template'),
          };
        },
        setTemplateImage: () => calls.push('template'),
      }) as never,
    createEmptyImage: () =>
      ({
        isEmpty: () => true,
        resize: () => {
          throw new Error('unexpected');
        },
        setTemplateImage: () => {},
      }) as never,
    createTray: () =>
      ({
        setContextMenu: () => calls.push('set-menu'),
        setToolTip: () => calls.push('set-tooltip'),
        on: (_event: 'click', _handler: () => void) => {
          calls.push('bind-click');
        },
        destroy: () => calls.push('destroy'),
      }) as never,
    trayTooltip: 'SubMiner',
    platform: 'darwin',
    logWarn: () => calls.push('warn'),
    ensureOverlayVisibleFromTrayClick: () => calls.push('show-overlay'),
  });

  ensureTray();
  assert.ok(trayRef);
  assert.ok(calls.includes('set-tray'));
  assert.ok(calls.includes('set-menu'));
  assert.ok(calls.includes('bind-click'));
});

test('ensure tray logs Linux tray registration failures without crashing startup', () => {
  const calls: string[] = [];
  let trayRef: unknown = null;

  const ensureTray = createEnsureTrayHandler({
    getTray: () => null,
    setTray: (tray) => {
      trayRef = tray;
      calls.push('set-tray');
    },
    buildTrayMenu: () => ({ id: 'menu' }),
    resolveTrayIconPath: () => '/tmp/icon.png',
    createImageFromPath: () =>
      ({
        isEmpty: () => false,
        resize: () => ({
          isEmpty: () => false,
          resize: () => {
            throw new Error('unexpected');
          },
          setTemplateImage: () => {},
        }),
        setTemplateImage: () => {},
      }) as never,
    createEmptyImage: () =>
      ({
        isEmpty: () => true,
        resize: () => {
          throw new Error('unexpected');
        },
        setTemplateImage: () => {},
      }) as never,
    createTray: () => {
      throw new Error('StatusNotifier watcher unavailable');
    },
    trayTooltip: 'SubMiner',
    platform: 'linux',
    logWarn: (message) => calls.push(`warn:${message}`),
    ensureOverlayVisibleFromTrayClick: () => calls.push('show-overlay'),
  });

  ensureTray();

  assert.equal(trayRef, null);
  assert.deepEqual(calls, [
    'warn:Unable to create Linux tray icon. Ensure your desktop has a StatusNotifier/AppIndicator tray host. StatusNotifier watcher unavailable',
  ]);
});

test('destroy tray handler destroys active tray and clears ref', () => {
  const calls: string[] = [];
  let tray: { destroy: () => void } | null = {
    destroy: () => calls.push('destroy'),
  };
  const destroyTray = createDestroyTrayHandler({
    getTray: () => tray as never,
    setTray: (next) => {
      tray = next as never;
      calls.push('set-null');
    },
  });

  destroyTray();
  assert.deepEqual(calls, ['destroy', 'set-null']);
  assert.equal(tray, null);
});
