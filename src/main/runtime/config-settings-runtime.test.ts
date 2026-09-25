import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { DEFAULT_CONFIG, deepCloneConfig } from '../../config';
import { resolveConfig } from '../../config/resolve';
import { buildConfigSettingsRegistry } from '../../config/settings/registry';
import type { RawConfig } from '../../types/config';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import { createConfigSettingsRuntime } from './config-settings-runtime';

test('settings saves report live changes and only the sections that actually need restart', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-settings-live-'));
  const configPath = path.join(dir, 'config.jsonc');
  let rawConfig: RawConfig = {};
  let resolvedConfig = resolveConfig(rawConfig).resolved;
  const applied: string[][] = [];
  const runtime = createConfigSettingsRuntime({
    fields: buildConfigSettingsRegistry(DEFAULT_CONFIG),
    getConfigPath: () => configPath,
    getRawConfig: () => rawConfig,
    getConfig: () => resolvedConfig,
    getWarnings: () => [],
    reloadConfigStrict: () => {
      rawConfig = JSON.parse(fs.readFileSync(configPath, 'utf-8'));
      const result = resolveConfig(rawConfig);
      resolvedConfig = result.resolved;
      return { ok: true, config: resolvedConfig, warnings: result.warnings, path: configPath };
    },
    onHotReloadApplied: (diff) => {
      applied.push(diff.hotReloadFields);
    },
    getSettingsWindow: () => null,
    setSettingsWindow: () => {},
    createSettingsWindow: () => {
      throw new Error('Save must not open a window');
    },
    settingsHtmlPath: '/tmp/settings.html',
    openPath: async () => '',
    defaultAnkiConnectUrl: DEFAULT_CONFIG.ankiConnect.url,
    createAnkiClient: () => {
      throw new Error('Save must not query Anki');
    },
    ipcMain: { handle: () => {} },
    ipcChannels: IPC_CHANNELS.request,
  });

  try {
    const live = runtime.savePatch({
      operations: [
        { op: 'set', path: 'notifications.overlayPosition', value: 'top' },
        {
          op: 'set',
          path: 'subtitleGeneration.threads',
          value: DEFAULT_CONFIG.subtitleGeneration.threads + 1,
        },
      ],
    });
    assert.equal(live.ok, true);
    assert.deepEqual(live.restartRequiredFields, []);
    assert.deepEqual(live.restartRequiredSections, []);
    assert.deepEqual(
      new Set(live.hotReloadFields),
      new Set(['notifications.overlayPosition', 'subtitleGeneration.threads']),
    );
    assert.deepEqual(applied, [live.hotReloadFields]);

    const mixed = runtime.savePatch({
      operations: [
        { op: 'set', path: 'ankiConnect.deck', value: 'Mining' },
        { op: 'set', path: 'ankiConnect.url', value: 'http://127.0.0.1:9999' },
      ],
    });
    assert.equal(mixed.ok, true);
    assert.deepEqual(mixed.hotReloadFields, ['ankiConnect.deck']);
    assert.deepEqual(mixed.restartRequiredSections, ['AnkiConnect']);

    const reset = runtime.savePatch({
      operations: [
        { op: 'reset', path: 'notifications.overlayPosition' },
        { op: 'reset', path: 'subtitleGeneration.threads' },
      ],
    });
    assert.equal(reset.ok, true);
    assert.deepEqual(reset.restartRequiredSections, []);
    assert.deepEqual(new Set(reset.hotReloadFields), new Set(live.hotReloadFields));
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

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
