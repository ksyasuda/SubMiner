import assert from 'node:assert/strict';
import test from 'node:test';
import type { CliArgs } from '../../../cli/args';
import { composeCliStartupHandlers } from './cli-startup-composer';

test('composeCliStartupHandlers returns callable CLI startup handlers', () => {
  const calls: string[] = [];
  const handlers = composeCliStartupHandlers({
    cliCommandContextMainDeps: {
      appState: {} as never,
      setLogLevel: () => {},
      texthookerService: {} as never,
      getResolvedConfig: () => ({}) as never,
      defaultWebsocketPort: 6677,
      defaultAnnotationWebsocketPort: 6678,
      hasMpvWebsocketPlugin: () => false,
      openExternal: async () => {},
      logBrowserOpenError: () => {},
      showMpvOsd: () => {},
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
      dispatchSessionAction: async () => {},
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
      schedule: () => 0 as never,
      logInfo: () => {},
      logWarn: () => {},
      logError: () => {},
    },
    cliCommandRuntimeHandlerMainDeps: {
      handleTexthookerOnlyModeTransitionMainDeps: {
        isTexthookerOnlyMode: () => false,
        ensureOverlayStartupPrereqs: () => {},
        setTexthookerOnlyMode: () => {},
        commandNeedsOverlayStartupPrereqs: () => false,
        startBackgroundWarmups: () => {},
        logInfo: () => {},
      },
      handleCliCommandRuntimeServiceWithContext: (args, _source, _ctx) => {
        calls.push(`handleCommand:${(args as { command?: string }).command ?? 'unknown'}`);
      },
    },
    initialArgsRuntimeHandlerMainDeps: {
      getInitialArgs: () => null,
      isBackgroundMode: () => false,
      shouldEnsureTrayOnStartup: () => false,
      shouldRunHeadlessInitialCommand: () => false,
      ensureTray: () => {},
      isTexthookerOnlyMode: () => false,
      hasImmersionTracker: () => false,
      getMpvClient: () => null,
      commandNeedsOverlayStartupPrereqs: () => false,
      commandNeedsOverlayRuntime: () => false,
      ensureOverlayStartupPrereqs: () => {},
      isOverlayRuntimeInitialized: () => false,
      initializeOverlayRuntime: () => {},
      logInfo: () => {},
    },
  });

  assert.equal(typeof handlers.createCliCommandContext, 'function');
  assert.equal(typeof handlers.handleCliCommand, 'function');
  assert.equal(typeof handlers.handleInitialArgs, 'function');

  // handleCliCommand routes to the injected handleCliCommandRuntimeServiceWithContext dep
  handlers.handleCliCommand({ command: 'start' } as unknown as CliArgs);
  assert.deepEqual(calls, ['handleCommand:start']);
});
