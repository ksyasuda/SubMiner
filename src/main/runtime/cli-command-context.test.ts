import test from 'node:test';
import assert from 'node:assert/strict';
import { createCliCommandContext } from './cli-command-context';

function createDeps() {
  let socketPath = '/tmp/mpv.sock';
  const logs: string[] = [];
  const browserErrors: string[] = [];

  return {
    deps: {
      getSocketPath: () => socketPath,
      setSocketPath: (value: string) => {
        socketPath = value;
      },
      getMpvClient: () => null,
      showOsd: () => {},
      texthookerService: {} as never,
      getTexthookerPort: () => 6677,
      setTexthookerPort: () => {},
      getTexthookerWebsocketUrl: () => 'ws://127.0.0.1:6678',
      shouldOpenBrowser: () => true,
      openExternal: async () => {},
      logBrowserOpenError: (url: string) => browserErrors.push(url),
      isOverlayInitialized: () => true,
      initializeOverlay: () => {},
      toggleVisibleOverlay: () => {},
      togglePrimarySubtitleBar: () => {},
      openFirstRunSetup: () => {},
      setVisibleOverlay: () => {},
      copyCurrentSubtitle: () => {},
      startPendingMultiCopy: () => {},
      mineSentenceCard: async () => {},
      startPendingMineSentenceMultiple: () => {},
      updateLastCardFromClipboard: async () => {},
      refreshKnownWordCache: async () => {},
      triggerFieldGrouping: async () => {},
      triggerSubsyncFromConfig: async () => {},
      markLastCardAsAudioCard: async () => {},
      dispatchSessionAction: async () => {},
      getAnilistStatus: () => ({}) as never,
      clearAnilistToken: () => {},
      openAnilistSetup: () => {},
      openJellyfinSetup: () => {},
      getAnilistQueueStatus: () => ({}) as never,
      retryAnilistQueueNow: async () => ({ ok: true, message: 'ok' }),
      generateCharacterDictionary: async () => ({
        zipPath: '/tmp/anilist-1.zip',
        fromCache: false,
        mediaId: 1,
        mediaTitle: 'Test',
        entryCount: 1,
      }),
      runStatsCommand: async () => {},
      runJellyfinCommand: async () => {},
      runYoutubePlaybackFlow: async () => {},
      openYomitanSettings: () => {},
      cycleSecondarySubMode: () => {},
      openRuntimeOptionsPalette: () => {},
      printHelp: () => {},
      stopApp: () => {},
      hasMainWindow: () => true,
      getMultiCopyTimeoutMs: () => 1000,
      schedule: (fn: () => void) => setTimeout(fn, 0),
      logInfo: (message: string) => {
        logs.push(`i:${message}`);
      },
      logWarn: (message: string) => {
        logs.push(`w:${message}`);
      },
      logError: (message: string) => {
        logs.push(`e:${message}`);
      },
    },
    getLogs: () => logs,
    getBrowserErrors: () => browserErrors,
  };
}

test('cli command context proxies socket path getters/setters', () => {
  const { deps } = createDeps();
  const context = createCliCommandContext(deps);
  assert.equal(context.getSocketPath(), '/tmp/mpv.sock');
  context.setSocketPath('/tmp/next.sock');
  assert.equal(context.getSocketPath(), '/tmp/next.sock');
});

test('cli command context openInBrowser reports failures', async () => {
  const { deps, getBrowserErrors } = createDeps();
  deps.openExternal = async () => {
    throw new Error('no browser');
  };
  const context = createCliCommandContext(deps);
  context.openInBrowser('https://example.com');
  await Promise.resolve();
  await Promise.resolve();
  assert.deepEqual(getBrowserErrors(), ['https://example.com']);
});

test('cli command context log methods map to deps loggers', () => {
  const { deps, getLogs } = createDeps();
  const context = createCliCommandContext(deps);
  context.log('info');
  context.warn('warn');
  context.error('error', new Error('x'));
  assert.deepEqual(getLogs(), ['i:info', 'w:warn', 'e:error']);
});
