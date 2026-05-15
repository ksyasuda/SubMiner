import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import test from 'node:test';

import {
  buildYomitanSettingsUrl,
  configureYomitanSettingsWindowChrome,
  destroyYomitanSettingsWindow,
  showYomitanSettingsWindow,
} from './yomitan-settings';

function assertGuardedBySubminerSettingsSafe(source: string, call: string): void {
  const callIndex = source.indexOf(call);
  assert.notEqual(callIndex, -1, `missing call: ${call}`);

  const beforeCall = source.slice(0, callIndex);
  const guardIndex = beforeCall.lastIndexOf('if (!subminerSettingsSafe) {');
  const blockCloseIndex = beforeCall.lastIndexOf('\n    }');
  assert.ok(
    guardIndex > blockCloseIndex,
    `${call} must be inside its own !subminerSettingsSafe startup guard`,
  );
}

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

test('vendored Yomitan settings safe mode skips heavy startup controllers', () => {
  const source = readFileSync(
    'vendor/subminer-yomitan/ext/js/pages/settings/settings-main.js',
    'utf8',
  );

  assert.match(source, /subminer-settings-safe/);
  assertGuardedBySubminerSettingsSafe(source, 'popupPreviewController.prepare()');
  assertGuardedBySubminerSettingsSafe(source, 'persistentStorageController.prepare()');
  assertGuardedBySubminerSettingsSafe(source, 'storageController.prepare()');
  assertGuardedBySubminerSettingsSafe(source, 'dictionaryController.prepare()');
  assertGuardedBySubminerSettingsSafe(source, 'ankiController.prepare()');
  assert.match(source, /if \(!subminerSettingsSafe\)[\s\S]*new AnkiDeckGeneratorController/);
  assert.match(source, /if \(!subminerSettingsSafe\)[\s\S]*new SecondarySearchDictionaryController/);
  assert.match(source, /if \(!subminerSettingsSafe\)[\s\S]*new SortFrequencyDictionaryController/);
});

test('vendored Yomitan settings caches dictionary metadata requests', () => {
  const source = readFileSync(
    'vendor/subminer-yomitan/ext/js/pages/settings/settings-controller.js',
    'utf8',
  );

  assert.match(source, /_dictionaryInfoPromise/);
  assert.match(source, /_dictionaryInfoCache/);
  assert.match(source, /databaseUpdated/);
  assert.match(
    source,
    /this\._dictionaryInfoPromise = this\._application\.api\.getDictionaryInfo\(\)/,
  );
});

test('vendored Yomitan Anki settings reuses SettingsController dictionary metadata cache', () => {
  const source = readFileSync(
    'vendor/subminer-yomitan/ext/js/pages/settings/anki-controller.js',
    'utf8',
  );

  assert.match(source, /this\._settingsController\.getDictionaryInfo\(\)/);
  assert.doesNotMatch(source, /this\._application\.api\.getDictionaryInfo\(\)/);
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
