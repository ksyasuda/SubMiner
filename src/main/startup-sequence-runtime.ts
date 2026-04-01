import type { CliArgs } from '../cli/args';
import type { ResolvedConfig } from '../types';
import { isAnilistTrackingEnabled } from './runtime/domains/anilist';
import { getStartupModeFlags } from './startup-flags';
import { runHeadlessKnownWordRefresh } from './headless-known-word-refresh';

export interface StartupSequenceRuntimeInput {
  appState: {
    initialArgs: CliArgs | null | undefined;
    runtimeOptionsManager: Parameters<
      typeof runHeadlessKnownWordRefresh
    >[0]['runtimeOptionsManager'];
  };
  userDataPath: string;
  getResolvedConfig: () => ResolvedConfig;
  anilist: {
    refreshAnilistClientSecretStateIfEnabled: (options: {
      force: boolean;
      allowSetupPrompt?: boolean;
    }) => Promise<unknown>;
    refreshRetryQueueState: () => void;
  };
  actions: {
    initializeDiscordPresenceService: () => Promise<void>;
    requestAppQuit: () => void;
  };
  logger: {
    error: (message: string, error?: unknown) => void;
  };
  runHeadlessKnownWordRefresh?: typeof runHeadlessKnownWordRefresh;
  getStartupModeFlags?: typeof getStartupModeFlags;
  isAnilistTrackingEnabled?: typeof isAnilistTrackingEnabled;
}

export interface StartupSequenceRuntime {
  runHeadlessInitialCommand: (input: { handleInitialArgs: () => void }) => Promise<void>;
  runPostStartupInitialization: () => void;
}

export function createStartupSequenceRuntime(
  input: StartupSequenceRuntimeInput,
): StartupSequenceRuntime {
  const runKnownWordRefresh = input.runHeadlessKnownWordRefresh ?? runHeadlessKnownWordRefresh;
  const resolveStartupModeFlags = input.getStartupModeFlags ?? getStartupModeFlags;
  const isTrackingEnabled = input.isAnilistTrackingEnabled ?? isAnilistTrackingEnabled;

  const shouldSkipDeferredStartup = (): boolean => {
    if (!input.appState.initialArgs) {
      return false;
    }

    const startupModeFlags = resolveStartupModeFlags(input.appState.initialArgs);
    return startupModeFlags.shouldUseMinimalStartup || startupModeFlags.shouldSkipHeavyStartup;
  };

  return {
    runHeadlessInitialCommand: async ({ handleInitialArgs }): Promise<void> => {
      if (!input.appState.initialArgs?.refreshKnownWords) {
        handleInitialArgs();
        return;
      }

      await runKnownWordRefresh({
        resolvedConfig: input.getResolvedConfig(),
        runtimeOptionsManager: input.appState.runtimeOptionsManager,
        userDataPath: input.userDataPath,
        logger: input.logger,
        requestAppQuit: input.actions.requestAppQuit,
      });
    },
    runPostStartupInitialization: (): void => {
      if (shouldSkipDeferredStartup()) {
        return;
      }

      if (isTrackingEnabled(input.getResolvedConfig())) {
        void input.anilist
          .refreshAnilistClientSecretStateIfEnabled({
            force: true,
            allowSetupPrompt: false,
          })
          .catch((error) => {
            input.logger.error(
              'Failed to refresh AniList client secret state during startup',
              error,
            );
          });
        input.anilist.refreshRetryQueueState();
      }

      void input.actions.initializeDiscordPresenceService().catch((error) => {
        input.logger.error('Failed to initialize Discord presence service during startup', error);
      });
    },
  };
}
