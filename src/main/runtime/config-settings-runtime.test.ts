import test from 'node:test';
import assert from 'node:assert/strict';
import { DEFAULT_CONFIG, deepCloneConfig } from '../../config';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import { createConfigSettingsRuntime } from './config-settings-runtime';

test('config settings runtime exposes inferred Yomitan Anki deck lookup', async () => {
  const handlers = new Map<string, (event: unknown, ...args: unknown[]) => unknown>();
  const runtime = createConfigSettingsRuntime({
    fields: [],
    getConfigPath: () => '/tmp/config.jsonc',
    getRawConfig: () => ({}),
    getConfig: () => deepCloneConfig(DEFAULT_CONFIG),
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
