import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

import { createFirstRunRuntime } from './first-run-runtime';

function withTempDir(fn: (dir: string) => Promise<void> | void): Promise<void> | void {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-first-run-runtime-test-'));
  const result = fn(dir);
  if (result instanceof Promise) {
    return result.finally(() => {
      fs.rmSync(dir, { recursive: true, force: true });
    });
  }
  fs.rmSync(dir, { recursive: true, force: true });
}

function createMockSetupWindow() {
  const calls: string[] = [];
  let closedHandler: (() => void) | null = null;
  let navigateHandler: ((event: { preventDefault: () => void }, url: string) => void) | null = null;

  return {
    calls,
    window: {
      webContents: {
        on: (event: 'will-navigate', handler: (event: unknown, url: string) => void) => {
          if (event === 'will-navigate') {
            navigateHandler = handler;
          }
        },
      },
      loadURL: async (url: string) => {
        calls.push(`load:${url.slice(0, 24)}`);
      },
      on: (event: 'closed', handler: () => void) => {
        if (event === 'closed') {
          closedHandler = handler;
        }
      },
      isDestroyed: () => false,
      close: () => {
        calls.push('close');
        closedHandler?.();
      },
      focus: () => {
        calls.push('focus');
      },
      triggerNavigate: (url: string) => {
        navigateHandler?.(
          {
            preventDefault: () => {
              calls.push('prevent-default');
            },
          },
          url,
        );
      },
    },
  };
}

test('first-run runtime focuses an existing window instead of creating a new one', async () => {
  await withTempDir(async (root) => {
    const configDir = path.join(root, 'SubMiner');
    fs.mkdirSync(configDir, { recursive: true });
    fs.writeFileSync(path.join(configDir, 'config.jsonc'), '{}');

    let createCount = 0;
    const mock = createMockSetupWindow();
    const runtime = createFirstRunRuntime({
      platform: 'darwin',
      configDir,
      homeDir: os.homedir(),
      binaryPath: process.execPath,
      appPath: '/app',
      resourcesPath: '/resources',
      appDataDir: path.join(root, 'appData'),
      desktopDir: path.join(root, 'desktop'),
      getYomitanDictionaryCount: async () => 1,
      isExternalYomitanConfigured: () => false,
      createBrowserWindow: () => {
        createCount += 1;
        return mock.window;
      },
      writeShortcutLink: () => true,
      openYomitanSettings: () => false,
      shouldQuitWhenClosedIncomplete: () => true,
      quitApp: () => {
        throw new Error('quit should not be called');
      },
      logError: () => {
        throw new Error('logError should not be called');
      },
    });

    runtime.openFirstRunSetupWindow();
    runtime.openFirstRunSetupWindow();

    await new Promise((resolve) => setTimeout(resolve, 0));

    assert.equal(createCount, 1);
    assert.equal(mock.calls.filter((call) => call === 'focus').length, 1);
  });
});

test('first-run runtime closes the setup window after completion', async () => {
  await withTempDir(async (root) => {
    const configDir = path.join(root, 'SubMiner');
    fs.mkdirSync(configDir, { recursive: true });
    fs.writeFileSync(path.join(configDir, 'config.jsonc'), '{}');

    const events: string[] = [];
    const mock = createMockSetupWindow();
    const runtime = createFirstRunRuntime({
      platform: 'linux',
      configDir,
      homeDir: os.homedir(),
      binaryPath: process.execPath,
      appPath: '/app',
      resourcesPath: '/resources',
      appDataDir: path.join(root, 'appData'),
      desktopDir: path.join(root, 'desktop'),
      getYomitanDictionaryCount: async () => 1,
      isExternalYomitanConfigured: () => false,
      createBrowserWindow: () => mock.window,
      writeShortcutLink: () => true,
      openYomitanSettings: () => false,
      shouldQuitWhenClosedIncomplete: () => true,
      quitApp: () => {
        events.push('quit');
      },
      logError: (message, error) => {
        events.push(`${message}:${String(error)}`);
      },
      onStateChanged: (state) => {
        events.push(state.status);
      },
    });

    runtime.openFirstRunSetupWindow();
    await new Promise((resolve) => setTimeout(resolve, 0));

    mock.window.triggerNavigate('subminer://first-run-setup?action=finish');
    await new Promise((resolve) => setTimeout(resolve, 0));

    assert.equal(runtime.isSetupCompleted(), true);
    assert.equal(events[0], 'in_progress');
    assert.equal(events.at(-1), 'completed');
    assert.equal(mock.calls.includes('close'), true);
    assert.equal(events.includes('quit'), false);
  });
});
