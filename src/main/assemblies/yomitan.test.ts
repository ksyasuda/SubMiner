import assert from 'node:assert/strict';
import test from 'node:test';
import type { BrowserWindow, IpcMainInvokeEvent } from 'electron';
import { DEFAULT_CONFIG } from '../../config/definitions';
import type { ResolvedConfig } from '../../types';
import type { DictionaryBackend } from '../../types/config';
import { createYomitanAssembly, type YomitanAssemblyDeps } from './yomitan';

type InvokeHandler = (event: IpcMainInvokeEvent, value: unknown) => Promise<void>;

function configWithExternalProfile(externalProfilePath: string): ResolvedConfig {
  return {
    ...DEFAULT_CONFIG,
    yomitan: { ...DEFAULT_CONFIG.yomitan, externalProfilePath },
  };
}

// Every callback except getConfig throws, so a test fails if the assembly calls a dep it
// should only call later. getConfig is read once during construction for the profile policy.
function makeDeps(
  overrides: Partial<YomitanAssemblyDeps> & { backend?: DictionaryBackend } = {},
): YomitanAssemblyDeps {
  const lazy = (name: string) => (): never => {
    throw new Error(`${name} called unexpectedly`);
  };
  const { backend = 'yomitan', ...rest } = overrides;
  return {
    appState: {
      yomitanExt: null,
      yomitanSession: null,
      yomitanParserWindow: null,
      yomitanParserReadyPromise: null,
      yomitanParserInitPromise: null,
      yomitanSettingsWindow: null,
    },
    getConfig: () => DEFAULT_CONFIG,
    logger: { info: lazy('info'), warn: lazy('warn'), error: lazy('error'), debug: lazy('debug') },
    userDataPath: '/tmp/subminer-yomitan-assembly-test',
    dictionaryBackend: backend,
    hachidori: { ensureLoaded: lazy('hachidori.ensureLoaded'), getSession: lazy('getSession') },
    getOverlayWindows: lazy('getOverlayWindows'),
    getOverlayNotifications: lazy('getOverlayNotifications'),
    ipcMain: { handle: lazy('ipcMain.handle') },
    openExternal: lazy('openExternal'),
    ...rest,
  };
}

test('createYomitanAssembly only reads config while constructing', () => {
  const yomitan = createYomitanAssembly(makeDeps());
  assert.equal(typeof yomitan.ensureExtensionLoaded, 'function');
  assert.equal(yomitan.profilePolicy.isExternalReadOnlyMode(), false);
});

test('an external Yomitan profile is read-only for the Yomitan backend only', () => {
  const getConfig = () => configWithExternalProfile('/profiles/main');
  assert.equal(
    createYomitanAssembly(makeDeps({ getConfig })).profilePolicy.isExternalReadOnlyMode(),
    true,
  );
  assert.equal(
    createYomitanAssembly(
      makeDeps({ getConfig, backend: 'hachidori' }),
    ).profilePolicy.isExternalReadOnlyMode(),
    false,
  );
});

test('Yomitan settings refuse to open in read-only external-profile mode', () => {
  const notifications: string[] = [];
  const yomitan = createYomitanAssembly(
    makeDeps({
      getConfig: () => configWithExternalProfile('/profiles/main'),
      logger: { info: () => {}, warn: () => {}, error: () => {}, debug: () => {} },
      getOverlayNotifications: () => ({
        showConfiguredStatusNotification: (message: string) => {
          notifications.push(message);
        },
      }),
    }),
  );

  assert.equal(yomitan.openDictionarySettings(), false);
  assert.deepEqual(notifications, [
    'Yomitan settings unavailable while using read-only external-profile mode.',
  ]);
});

test('dictionary links open only from the main frame of an overlay window', async () => {
  const overlaySender = { mainFrame: 'overlay-main-frame' };
  const overlayWindow = { webContents: overlaySender } as unknown as BrowserWindow;
  let handler: InvokeHandler | null = null;
  const opened: string[] = [];
  const yomitan = createYomitanAssembly(
    makeDeps({
      backend: 'hachidori',
      getOverlayWindows: () => [overlayWindow],
      ipcMain: {
        handle: (_channel: string, listener: InvokeHandler) => {
          handler = listener;
        },
      } as YomitanAssemblyDeps['ipcMain'],
      openExternal: async (url) => {
        opened.push(url);
      },
    }),
  );
  yomitan.registerDictionaryExternalLinkHandler();
  assert.ok(handler);
  const invoke = handler as InvokeHandler;
  const event = (sender: unknown, senderFrame: unknown) =>
    ({ sender, senderFrame }) as unknown as IpcMainInvokeEvent;

  await invoke(event(overlaySender, 'overlay-main-frame'), 'https://example.com/entry');
  await assert.rejects(invoke(event(overlaySender, 'child-frame'), 'https://example.com/x'));
  await assert.rejects(
    invoke(event({ mainFrame: 'other' }, 'other'), 'https://example.com/x'),
    /only available from the active overlay/,
  );
  assert.deepEqual(opened, ['https://example.com/entry']);
});
