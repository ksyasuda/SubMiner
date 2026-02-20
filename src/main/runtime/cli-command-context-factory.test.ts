import assert from 'node:assert/strict';
import test from 'node:test';
import { createCliCommandContextFactory } from './cli-command-context-factory';

test('cli command context factory composes main deps and context handlers', () => {
  const calls: string[] = [];
  const appState = {
    mpvSocketPath: '/tmp/mpv.sock',
    mpvClient: null as unknown,
    texthookerPort: 5174,
    overlayRuntimeInitialized: false,
  };

  const createContext = createCliCommandContextFactory({
    appState,
    texthookerService: { start: () => null },
    getResolvedConfig: () => ({ texthooker: { openBrowser: true } }),
    openExternal: async () => {},
    logBrowserOpenError: () => {},
    showMpvOsd: (text) => calls.push(`osd:${text}`),
    initializeOverlayRuntime: () => calls.push('init-overlay'),
    toggleVisibleOverlay: () => calls.push('toggle-visible'),
    toggleInvisibleOverlay: () => calls.push('toggle-invisible'),
    setVisibleOverlayVisible: (visible) => calls.push(`set-visible:${visible}`),
    setInvisibleOverlayVisible: (visible) => calls.push(`set-invisible:${visible}`),
    copyCurrentSubtitle: () => calls.push('copy-sub'),
    startPendingMultiCopy: (timeoutMs) => calls.push(`multi:${timeoutMs}`),
    mineSentenceCard: async () => {},
    startPendingMineSentenceMultiple: () => {},
    updateLastCardFromClipboard: async () => {},
    refreshKnownWordCache: async () => {},
    triggerFieldGrouping: async () => {},
    triggerSubsyncFromConfig: async () => {},
    markLastCardAsAudioCard: async () => {},
    getAnilistStatus: () => ({ status: 'ok' }),
    clearAnilistToken: () => {},
    openAnilistSetupWindow: () => {},
    openJellyfinSetupWindow: () => {},
    getAnilistQueueStatus: () => ({ queued: 0 }),
    processNextAnilistRetryUpdate: async () => ({ ok: true, message: 'ok' }),
    runJellyfinCommand: async () => {},
    openYomitanSettings: () => {},
    cycleSecondarySubMode: () => {},
    openRuntimeOptionsPalette: () => {},
    printHelp: () => {},
    stopApp: () => {},
    hasMainWindow: () => true,
    getMultiCopyTimeoutMs: () => 5000,
    schedule: (fn) => setTimeout(fn, 0),
    logInfo: () => {},
    logWarn: () => {},
    logError: () => {},
  });

  const context = createContext();
  context.setSocketPath('/tmp/new.sock');
  context.showOsd('hello');
  context.setVisibleOverlay(true);
  context.setInvisibleOverlay(false);
  context.toggleVisibleOverlay();
  context.toggleInvisibleOverlay();

  assert.equal(appState.mpvSocketPath, '/tmp/new.sock');
  assert.deepEqual(calls, [
    'osd:hello',
    'set-visible:true',
    'set-invisible:false',
    'toggle-visible',
    'toggle-invisible',
  ]);
});
