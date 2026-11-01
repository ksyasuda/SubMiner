import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import {
  createDefaultSetupState,
  getSetupStatePath,
  writeSetupState,
} from '../../shared/setup-state';
import { createFirstRunSetupAssembly, type FirstRunSetupAssemblyDeps } from './first-run-setup';

// Every callback throws, so a test fails if the assembly calls a dep it should only call later.
function makeDeps(configDir: string): FirstRunSetupAssemblyDeps {
  const lazy = (name: string) => (): never => {
    throw new Error(`${name} called unexpectedly`);
  };
  return {
    appState: {
      firstRunSetupCompleted: false,
      firstRunSetupWindow: null,
      backgroundMode: false,
      initialArgs: null,
    },
    configService: { getConfig: lazy('getConfig'), patchRawConfig: lazy('patchRawConfig') },
    logger: { error: lazy('error'), info: lazy('info') },
    configDir,
    userDataPath: configDir,
    dictionaryBackend: 'yomitan',
    ensureYomitanExtensionLoaded: lazy('ensureYomitanExtensionLoaded'),
    getYomitanParserRuntimeDeps: lazy('getYomitanParserRuntimeDeps'),
    createCommandLineLauncherRuntimeOptions: lazy('createCommandLineLauncherRuntimeOptions'),
    getAppPath: lazy('getAppPath'),
    trashItem: lazy('trashItem'),
    writeShortcutLink: lazy('writeShortcutLink'),
    createBrowserWindow: lazy('createBrowserWindow'),
    openDictionarySettings: lazy('openDictionarySettings'),
    openConfigSettingsWindow: lazy('openConfigSettingsWindow'),
    onSetupStateChanged: lazy('onSetupStateChanged'),
    requestAppQuit: lazy('requestAppQuit'),
  };
}

function makeConfigDir(t: { after: (fn: () => void) => void }): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-first-run-assembly-test-'));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  return dir;
}

// main.ts builds the assembly before most siblings exist, so construction must stay lazy.
test('createFirstRunSetupAssembly does not call its deps while constructing', (t) => {
  const firstRunSetup = createFirstRunSetupAssembly(makeDeps(makeConfigDir(t)));
  assert.equal(typeof firstRunSetup.service.isSetupCompleted, 'function');
  assert.equal(typeof firstRunSetup.openWindow, 'function');
});

test('openWindow skips a completed setup unless forced', (t) => {
  const configDir = makeConfigDir(t);
  writeSetupState(getSetupStatePath(configDir), {
    ...createDefaultSetupState(),
    status: 'completed',
  });
  const firstRunSetup = createFirstRunSetupAssembly(makeDeps(configDir));

  firstRunSetup.openWindow();
  assert.throws(() => firstRunSetup.openWindow(true), /createBrowserWindow called unexpectedly/);
});
