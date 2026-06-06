import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { DEFAULT_CONFIG, deepCloneConfig } from '../../config';
import { resolveConfig } from '../../config/resolve';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import { createConfigSettingsRuntime } from './config-settings-runtime';

test('config settings runtime exposes inferred Yomitan Anki deck lookup', async () => {
  const handlers = new Map<string, (event: unknown, ...args: unknown[]) => unknown>();
  const runtime = createConfigSettingsRuntime({
    fields: [],
    getConfigPath: () => '/tmp/config.jsonc',
    getRawConfig: () => ({}),
    getConfig: () => ({
      ...deepCloneConfig(DEFAULT_CONFIG),
      ankiConnect: {
        ...deepCloneConfig(DEFAULT_CONFIG).ankiConnect,
        deck: 'Configured',
      },
    }),
    getWarnings: () => [],
    reloadConfigStrict: () =>
      ({
        ok: true,
        config: deepCloneConfig(DEFAULT_CONFIG),
        warnings: [],
        path: '/tmp/config.jsonc',
      }) as never,
    getSettingsWindow: () => null,
    setSettingsWindow: () => undefined,
    createSettingsWindow: () => ({}) as never,
    settingsHtmlPath: '/tmp/settings.html',
    openPath: async () => '',
    defaultAnkiConnectUrl: DEFAULT_CONFIG.ankiConnect.url,
    createAnkiClient: () =>
      ({
        deckNames: async () => [],
        fieldNamesForDeck: async () => [],
        modelNamesForDeck: async () => [],
        modelNames: async () => [],
        modelFieldNames: async () => [],
      }) as never,
    getYomitanAnkiDeckName: async () => 'Mining',
    ipcMain: {
      handle: (channel, listener) => {
        handlers.set(channel, listener);
      },
    },
    ipcChannels: IPC_CHANNELS.request,
  });

  runtime.registerHandlers();

  const handler = handlers.get(IPC_CHANNELS.request.getConfigSettingsYomitanAnkiDeckName);
  assert.ok(handler);
  assert.deepEqual(await handler({}, undefined), { ok: true, value: 'Mining' });
});

test('config settings runtime persists inferred Yomitan Anki deck when config deck is empty', async () => {
  const handlers = new Map<string, (event: unknown, ...args: unknown[]) => unknown>();
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-settings-'));
  const configPath = path.join(dir, 'config.jsonc');
  fs.writeFileSync(configPath, '{"ankiConnect":{"deck":""}}\n', 'utf-8');

  try {
    let rawConfig = { ankiConnect: { deck: '' } };
    let resolvedConfig = resolveConfig(rawConfig).resolved;
    const runtime = createConfigSettingsRuntime({
      fields: [],
      getConfigPath: () => configPath,
      getRawConfig: () => rawConfig,
      getConfig: () => resolvedConfig,
      getWarnings: () => [],
      reloadConfigStrict: () => {
        rawConfig = JSON.parse(fs.readFileSync(configPath, 'utf-8'));
        resolvedConfig = resolveConfig(rawConfig).resolved;
        return {
          ok: true,
          config: resolvedConfig,
          warnings: [],
          path: configPath,
        };
      },
      getSettingsWindow: () => null,
      setSettingsWindow: () => undefined,
      createSettingsWindow: () => ({}) as never,
      settingsHtmlPath: '/tmp/settings.html',
      openPath: async () => '',
      defaultAnkiConnectUrl: DEFAULT_CONFIG.ankiConnect.url,
      createAnkiClient: () =>
        ({
          deckNames: async () => [],
          fieldNamesForDeck: async () => [],
          modelNamesForDeck: async () => [],
          modelNames: async () => [],
          modelFieldNames: async () => [],
        }) as never,
      getYomitanAnkiDeckName: async () => 'Minecraft',
      ipcMain: {
        handle: (channel, listener) => {
          handlers.set(channel, listener);
        },
      },
      ipcChannels: IPC_CHANNELS.request,
    });

    runtime.registerHandlers();

    const handler = handlers.get(IPC_CHANNELS.request.getConfigSettingsYomitanAnkiDeckName);
    assert.ok(handler);
    assert.deepEqual(await handler({}, undefined), { ok: true, value: 'Minecraft' });
    assert.equal(JSON.parse(fs.readFileSync(configPath, 'utf-8')).ankiConnect.deck, 'Minecraft');
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
