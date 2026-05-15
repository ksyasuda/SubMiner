import assert from 'node:assert/strict';
import test from 'node:test';

import {
  buildYomitanSettingsWindowMenuTemplate,
  buildYomitanSettingsUrl,
  configureYomitanSettingsWindowChrome,
  destroyYomitanSettingsWindow,
  showYomitanSettingsWindow,
} from './yomitan-settings';

test('yomitan settings window uses a close-only menu without app quit', () => {
  const calls: string[] = [];

  configureYomitanSettingsWindowChrome({
    isDestroyed: () => false,
    close: () => calls.push('close'),
    setAutoHideMenuBar: (hide: boolean) => calls.push(`auto-hide:${hide}`),
    setMenu: (menu: unknown) => calls.push(`menu:${menu === null ? 'null' : 'custom'}`),
  } as never, (template) => {
    calls.push(`menu-label:${template[0]?.label ?? ''}`);
    const submenu = template[0]?.submenu;
    assert.ok(Array.isArray(submenu));
    const closeItem = submenu[0];
    assert.equal(closeItem?.label, 'Close');
    assert.notEqual(closeItem?.role, 'quit');
    closeItem?.click?.({} as never, {} as never, {} as never);
    return { id: 'settings-menu' } as never;
  });

  assert.deepEqual(calls, ['auto-hide:false', 'menu-label:File', 'close', 'menu:custom']);
});

test('yomitan settings close menu skips destroyed windows', () => {
  const calls: string[] = [];
  const template = buildYomitanSettingsWindowMenuTemplate({
    isDestroyed: () => true,
    close: () => calls.push('close'),
  } as never);
  const submenu = template[0]?.submenu;
  assert.ok(Array.isArray(submenu));
  submenu[0]?.click?.({} as never, {} as never, {} as never);
  assert.deepEqual(calls, []);
});

test('yomitan settings URL disables the embedded popup preview', () => {
  assert.equal(
    buildYomitanSettingsUrl('abc123'),
    'chrome-extension://abc123/settings.html?popup-preview=false',
  );
});

test('showYomitanSettingsWindow restores, repaints, shows, and focuses an existing window', () => {
  const calls: string[] = [];

  showYomitanSettingsWindow({
    isDestroyed: () => false,
    isMinimized: () => true,
    restore: () => calls.push('restore'),
    getSize: () => [1200, 800],
    setSize: (width: number, height: number) => calls.push(`set-size:${width}x${height}`),
    webContents: {
      invalidate: () => calls.push('invalidate'),
    },
    show: () => calls.push('show'),
    focus: () => calls.push('focus'),
  } as never);

  assert.deepEqual(calls, ['restore', 'set-size:1200x800', 'invalidate', 'show', 'focus']);
});

test('destroyYomitanSettingsWindow destroys a live settings window before app quit', () => {
  const calls: string[] = [];

  const destroyed = destroyYomitanSettingsWindow({
    isDestroyed: () => false,
    destroy: () => calls.push('destroy'),
  } as never);

  assert.equal(destroyed, true);
  assert.deepEqual(calls, ['destroy']);
});

test('destroyYomitanSettingsWindow skips missing or already destroyed settings windows', () => {
  assert.equal(destroyYomitanSettingsWindow(null), false);
  assert.equal(
    destroyYomitanSettingsWindow({
      isDestroyed: () => true,
      destroy: () => {
        throw new Error('should not destroy twice');
      },
    } as never),
    false,
  );
});
