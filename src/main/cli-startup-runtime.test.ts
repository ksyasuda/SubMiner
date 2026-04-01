import assert from 'node:assert/strict';
import test from 'node:test';
import type { CliArgs } from '../cli/args';

import { createCliStartupRuntime } from './cli-startup-runtime';

test('cli startup runtime returns callable CLI handlers', () => {
  const calls: string[] = [];

  const runtime = createCliStartupRuntime({
    appState: {
      appState: {} as never,
      getInitialArgs: () => null,
      isBackgroundMode: () => false,
      isTexthookerOnlyMode: () => false,
      setTexthookerOnlyMode: () => {},
      hasImmersionTracker: () => false,
      getMpvClient: () => null,
      isOverlayRuntimeInitialized: () => false,
    },
    config: {
      defaultConfig: { websocket: { port: 6677 }, annotationWebsocket: { port: 6678 } } as never,
      getResolvedConfig: () => ({}) as never,
      setCliLogLevel: () => {},
      hasMpvWebsocketPlugin: () => false,
    },
    io: {
      texthookerService: {} as never,
      openExternal: async () => {},
      logBrowserOpenError: () => {},
      showMpvOsd: () => {},
      schedule: () => 0 as never,
      logInfo: () => {},
      logWarn: () => {},
      logError: () => {},
    },
    commands: {
      initializeOverlayRuntime: () => {},
      toggleVisibleOverlay: () => {},
      openFirstRunSetupWindow: () => {},
      setVisibleOverlayVisible: () => {},
      copyCurrentSubtitle: () => {},
      startPendingMultiCopy: () => {},
      mineSentenceCard: async () => {},
      startPendingMineSentenceMultiple: () => {},
      updateLastCardFromClipboard: async () => {},
      refreshKnownWordCache: async () => {},
      triggerFieldGrouping: async () => {},
      triggerSubsyncFromConfig: async () => {},
      markLastCardAsAudioCard: async () => {},
      getAnilistStatus: () => ({}) as never,
      clearAnilistToken: () => {},
      openAnilistSetupWindow: () => {},
      openJellyfinSetupWindow: () => {},
      getAnilistQueueStatus: () => ({}) as never,
      processNextAnilistRetryUpdate: async () => ({ ok: true, message: 'done' }),
      generateCharacterDictionary: async () => ({
        zipPath: '/tmp/test.zip',
        fromCache: false,
        mediaId: 1,
        mediaTitle: 'Test',
        entryCount: 1,
      }),
      runJellyfinCommand: async () => {},
      runStatsCommand: async () => {},
      runYoutubePlaybackFlow: async () => {},
      openYomitanSettings: () => {},
      cycleSecondarySubMode: () => {},
      openRuntimeOptionsPalette: () => {},
      printHelp: () => {},
      stopApp: () => {},
      hasMainWindow: () => false,
      getMultiCopyTimeoutMs: () => 0,
    },
    startup: {
      shouldEnsureTrayOnStartup: () => false,
      shouldRunHeadlessInitialCommand: () => false,
      ensureTray: () => {},
      commandNeedsOverlayStartupPrereqs: () => false,
      commandNeedsOverlayRuntime: () => false,
      ensureOverlayStartupPrereqs: () => {},
      startBackgroundWarmups: () => {},
    },
    handleCliCommandRuntimeServiceWithContext: (args) => {
      calls.push(`handle:${(args as { command?: string }).command ?? 'unknown'}`);
    },
  });

  assert.equal(typeof runtime.handleCliCommand, 'function');
  assert.equal(typeof runtime.handleInitialArgs, 'function');

  runtime.handleCliCommand({ command: 'start' } as unknown as CliArgs);
  assert.deepEqual(calls, ['handle:start']);
});
