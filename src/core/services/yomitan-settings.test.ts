import assert from 'node:assert/strict';
import test from 'node:test';

import {
  buildYomitanSettingsUrl,
  configureYomitanSettingsWindowChrome,
  destroyYomitanSettingsWindow,
  showYomitanSettingsWindow,
} from './yomitan-settings';

test('yomitan settings window removes default app menu quit action', () => {
  const calls: string[] = [];

  configureYomitanSettingsWindowChrome({
    setAutoHideMenuBar: (hide: boolean) => calls.push(`auto-hide:${hide}`),
    setMenu: (menu: unknown) => calls.push(`menu:${menu === null ? 'null' : 'custom'}`),
  } as never);

  assert.deepEqual(calls, ['auto-hide:true', 'menu:null']);
});

test('yomitan settings URL disables the embedded popup preview', () => {
  assert.equal(
    buildYomitanSettingsUrl('abc123'),
    'chrome-extension://abc123/settings.html?popup-preview=false&subminer-settings-safe=true',
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
