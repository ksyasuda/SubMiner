import fs from 'node:fs';
import {
  createDefaultSetupState,
  getDefaultConfigFilePaths,
  getSetupStatePath,
  isSetupCompleted,
  readSetupState,
  writeSetupState,
  type SetupPluginInstallStatus,
  type SetupWindowsMpvShortcutInstallStatus,
  type SetupState,
} from '../../shared/setup-state';
import type { CliArgs } from '../../cli/args';

export interface SetupWindowsMpvShortcutSnapshot {
  supported: boolean;
  startMenuEnabled: boolean;
  desktopEnabled: boolean;
  startMenuInstalled: boolean;
  desktopInstalled: boolean;
  status: 'installed' | 'optional' | 'skipped' | 'failed';
  message: string | null;
}

export interface SetupStatusSnapshot {
  configReady: boolean;
  dictionaryCount: number;
  canFinish: boolean;
  externalYomitanConfigured: boolean;
  pluginStatus: 'installed' | 'required' | 'failed';
  pluginInstallPathSummary: string | null;
  windowsMpvShortcuts: SetupWindowsMpvShortcutSnapshot;
  message: string | null;
  state: SetupState;
}

export interface PluginInstallResult {
  ok: boolean;
  pluginInstallStatus: SetupPluginInstallStatus;
  pluginInstallPathSummary: string | null;
  message: string;
}

export interface FirstRunSetupService {
  ensureSetupStateInitialized: () => Promise<SetupStatusSnapshot>;
  getSetupStatus: () => Promise<SetupStatusSnapshot>;
  refreshStatus: (message?: string | null) => Promise<SetupStatusSnapshot>;
  markSetupInProgress: () => Promise<SetupStatusSnapshot>;
  markSetupCancelled: () => Promise<SetupStatusSnapshot>;
  markSetupCompleted: () => Promise<SetupStatusSnapshot>;
  installMpvPlugin: () => Promise<SetupStatusSnapshot>;
  configureWindowsMpvShortcuts: (preferences: {
    startMenuEnabled: boolean;
    desktopEnabled: boolean;
  }) => Promise<SetupStatusSnapshot>;
  isSetupCompleted: () => boolean;
}

function hasAnyStartupCommandBeyondSetup(args: CliArgs): boolean {
  return Boolean(
    args.toggle ||
    args.toggleVisibleOverlay ||
    args.launchMpv ||
    args.settings ||
    args.show ||
    args.hide ||
    args.showVisibleOverlay ||
    args.hideVisibleOverlay ||
    args.copySubtitle ||
    args.copySubtitleMultiple ||
    args.copySubtitleCount !== undefined ||
    args.mineSentence ||
    args.mineSentenceMultiple ||
    args.mineSentenceCount !== undefined ||
    args.updateLastCardFromClipboard ||
    args.refreshKnownWords ||
    args.toggleSecondarySub ||
    args.triggerFieldGrouping ||
    args.triggerSubsync ||
    args.markAudioCard ||
    args.toggleStatsOverlay ||
    args.openRuntimeOptions ||
    args.openJimaku ||
    args.openYoutubePicker ||
    args.openPlaylistBrowser ||
    args.replayCurrentSubtitle ||
    args.playNextSubtitle ||
    args.shiftSubDelayPrevLine ||
    args.shiftSubDelayNextLine ||
    args.cycleRuntimeOptionId !== undefined ||
    args.anilistStatus ||
    args.anilistLogout ||
    args.anilistSetup ||
    args.anilistRetryQueue ||
    args.dictionary ||
    args.jellyfin ||
    args.jellyfinLogin ||
    args.jellyfinLogout ||
    args.jellyfinLibraries ||
    args.jellyfinItems ||
    args.jellyfinSubtitles ||
    args.jellyfinPlay ||
    args.jellyfinRemoteAnnounce ||
    args.jellyfinPreviewAuth ||
    args.texthooker ||
    args.help,
  );
}

export function shouldAutoOpenFirstRunSetup(args: CliArgs): boolean {
  if (args.setup) return true;
  if (!args.start && !args.background) return false;
  return !hasAnyStartupCommandBeyondSetup(args);
}

function getPluginStatus(
  state: SetupState,
  pluginInstalled: boolean,
): SetupStatusSnapshot['pluginStatus'] {
  if (pluginInstalled) return 'installed';
  if (state.pluginInstallStatus === 'failed') return 'failed';
  return 'required';
}

function getWindowsMpvShortcutStatus(
  state: SetupState,
  installed: { startMenuInstalled: boolean; desktopInstalled: boolean },
): SetupWindowsMpvShortcutSnapshot['status'] {
  if (installed.startMenuInstalled || installed.desktopInstalled) return 'installed';
  if (state.windowsMpvShortcutLastStatus === 'skipped') return 'skipped';
  if (state.windowsMpvShortcutLastStatus === 'failed') return 'failed';
  return 'optional';
}

function getEffectiveWindowsMpvShortcutPreferences(
  state: SetupState,
  installed: { startMenuInstalled: boolean; desktopInstalled: boolean },
): { startMenuEnabled: boolean; desktopEnabled: boolean } {
  if (state.windowsMpvShortcutLastStatus === 'unknown') {
    return {
      startMenuEnabled: installed.startMenuInstalled,
      desktopEnabled: installed.desktopInstalled,
    };
  }

  return {
    startMenuEnabled: state.windowsMpvShortcutPreferences.startMenuEnabled,
    desktopEnabled: state.windowsMpvShortcutPreferences.desktopEnabled,
  };
}

function isYomitanSetupSatisfied(options: {
  configReady: boolean;
  dictionaryCount: number;
  externalYomitanConfigured: boolean;
}): boolean {
  if (!options.configReady) {
    return false;
  }
  return options.externalYomitanConfigured || options.dictionaryCount >= 1;
}

export function getFirstRunSetupCompletionMessage(snapshot: {
  configReady: boolean;
  dictionaryCount: number;
  externalYomitanConfigured: boolean;
  pluginStatus: SetupStatusSnapshot['pluginStatus'];
}): string | null {
  if (!snapshot.configReady) {
    return 'Create or provide the config file before finishing setup.';
  }
  if (snapshot.pluginStatus !== 'installed') {
    return 'Install the mpv plugin before finishing setup.';
  }
  if (!snapshot.externalYomitanConfigured && snapshot.dictionaryCount < 1) {
    return 'Install at least one Yomitan dictionary before finishing setup.';
  }
  return null;
}

async function resolveYomitanSetupStatus(deps: {
  configFilePaths: { jsoncPath: string; jsonPath: string };
  getYomitanDictionaryCount: () => Promise<number>;
  isExternalYomitanConfigured?: () => boolean;
}): Promise<{
  configReady: boolean;
  dictionaryCount: number;
  externalYomitanConfigured: boolean;
}> {
  const configReady =
    fs.existsSync(deps.configFilePaths.jsoncPath) || fs.existsSync(deps.configFilePaths.jsonPath);
  const externalYomitanConfigured = deps.isExternalYomitanConfigured?.() ?? false;

  if (configReady && externalYomitanConfigured) {
    return {
      configReady,
      dictionaryCount: 0,
      externalYomitanConfigured,
    };
  }

  return {
    configReady,
    dictionaryCount: await deps.getYomitanDictionaryCount(),
    externalYomitanConfigured,
  };
}

export function createFirstRunSetupService(deps: {
  platform?: NodeJS.Platform;
  configDir: string;
  getYomitanDictionaryCount: () => Promise<number>;
  isExternalYomitanConfigured?: () => boolean;
  detectPluginInstalled: () => boolean | Promise<boolean>;
  installPlugin: () => Promise<PluginInstallResult>;
  detectWindowsMpvShortcuts?: () =>
    | { startMenuInstalled: boolean; desktopInstalled: boolean }
    | Promise<{ startMenuInstalled: boolean; desktopInstalled: boolean }>;
  applyWindowsMpvShortcuts?: (preferences: {
    startMenuEnabled: boolean;
    desktopEnabled: boolean;
  }) => Promise<{ ok: boolean; status: SetupWindowsMpvShortcutInstallStatus; message: string }>;
  onStateChanged?: (state: SetupState) => void;
}): FirstRunSetupService {
  const setupStatePath = getSetupStatePath(deps.configDir);
  const configFilePaths = getDefaultConfigFilePaths(deps.configDir);
  const isWindows = (deps.platform ?? process.platform) === 'win32';
  let completed = false;

  const readState = (): SetupState => readSetupState(setupStatePath) ?? createDefaultSetupState();
  const writeState = (state: SetupState): SetupState => {
    writeSetupState(setupStatePath, state);
    completed = state.status === 'completed';
    deps.onStateChanged?.(state);
    return state;
  };

  const buildSnapshot = async (state: SetupState, message: string | null = null) => {
    const { configReady, dictionaryCount, externalYomitanConfigured } =
      await resolveYomitanSetupStatus({
        configFilePaths,
        getYomitanDictionaryCount: deps.getYomitanDictionaryCount,
        isExternalYomitanConfigured: deps.isExternalYomitanConfigured,
      });
    const pluginInstalled = await deps.detectPluginInstalled();
    const detectedWindowsMpvShortcuts = isWindows
      ? await deps.detectWindowsMpvShortcuts?.()
      : undefined;
    const installedWindowsMpvShortcuts = {
      startMenuInstalled: detectedWindowsMpvShortcuts?.startMenuInstalled ?? false,
      desktopInstalled: detectedWindowsMpvShortcuts?.desktopInstalled ?? false,
    };
    const effectiveWindowsMpvShortcutPreferences = getEffectiveWindowsMpvShortcutPreferences(
      state,
      installedWindowsMpvShortcuts,
    );
    return {
      configReady,
      dictionaryCount,
      canFinish:
        pluginInstalled &&
        isYomitanSetupSatisfied({
          configReady,
          dictionaryCount,
          externalYomitanConfigured,
        }),
      externalYomitanConfigured,
      pluginStatus: getPluginStatus(state, pluginInstalled),
      pluginInstallPathSummary: state.pluginInstallPathSummary,
      windowsMpvShortcuts: {
        supported: isWindows,
        startMenuEnabled: effectiveWindowsMpvShortcutPreferences.startMenuEnabled,
        desktopEnabled: effectiveWindowsMpvShortcutPreferences.desktopEnabled,
        startMenuInstalled: installedWindowsMpvShortcuts.startMenuInstalled,
        desktopInstalled: installedWindowsMpvShortcuts.desktopInstalled,
        status: getWindowsMpvShortcutStatus(state, installedWindowsMpvShortcuts),
        message: null,
      },
      message,
      state,
    } satisfies SetupStatusSnapshot;
  };

  const refreshWithState = async (state: SetupState, message: string | null = null) => {
    const snapshot = await buildSnapshot(state, message);
    if (snapshot.state.lastSeenYomitanDictionaryCount !== snapshot.dictionaryCount) {
      snapshot.state = writeState({
        ...snapshot.state,
        lastSeenYomitanDictionaryCount: snapshot.dictionaryCount,
      });
    }
    return snapshot;
  };

  return {
    ensureSetupStateInitialized: async () => {
      const state = readState();
      const { configReady, dictionaryCount, externalYomitanConfigured } =
        await resolveYomitanSetupStatus({
          configFilePaths,
          getYomitanDictionaryCount: deps.getYomitanDictionaryCount,
          isExternalYomitanConfigured: deps.isExternalYomitanConfigured,
        });
      const pluginInstalled = await deps.detectPluginInstalled();
      const canFinish =
        pluginInstalled &&
        isYomitanSetupSatisfied({
          configReady,
          dictionaryCount,
          externalYomitanConfigured,
        });
      if (isSetupCompleted(state) && canFinish) {
        completed = true;
        return refreshWithState(state);
      }

      if (canFinish) {
        const completedState = writeState({
          ...state,
          status: 'completed',
          completedAt: new Date().toISOString(),
          completionSource: 'legacy_auto_detected',
          yomitanSetupMode: externalYomitanConfigured ? 'external' : 'internal',
          lastSeenYomitanDictionaryCount: dictionaryCount,
        });
        return buildSnapshot(completedState);
      }

      return refreshWithState(
        writeState({
          ...state,
          status: state.status === 'cancelled' ? 'cancelled' : 'incomplete',
          completedAt: null,
          completionSource: null,
          yomitanSetupMode: null,
          lastSeenYomitanDictionaryCount: dictionaryCount,
        }),
      );
    },
    getSetupStatus: async () => refreshWithState(readState()),
    refreshStatus: async (message = null) => refreshWithState(readState(), message),
    markSetupInProgress: async () => {
      const state = readState();
      if (state.status === 'completed') {
        completed = true;
        return refreshWithState(state);
      }
      return refreshWithState(writeState({ ...state, status: 'in_progress' }));
    },
    markSetupCancelled: async () => {
      const state = readState();
      if (state.status === 'completed') {
        completed = true;
        return refreshWithState(state);
      }
      return refreshWithState(writeState({ ...state, status: 'cancelled' }));
    },
    markSetupCompleted: async () => {
      const state = readState();
      const snapshot = await buildSnapshot(state);
      if (!snapshot.canFinish) {
        return snapshot;
      }
      return refreshWithState(
        writeState({
          ...state,
          status: 'completed',
          completedAt: new Date().toISOString(),
          completionSource: 'user',
          yomitanSetupMode: snapshot.externalYomitanConfigured ? 'external' : 'internal',
          lastSeenYomitanDictionaryCount: snapshot.dictionaryCount,
        }),
      );
    },
    installMpvPlugin: async () => {
      const result = await deps.installPlugin();
      return refreshWithState(
        writeState({
          ...readState(),
          pluginInstallStatus: result.pluginInstallStatus,
          pluginInstallPathSummary: result.pluginInstallPathSummary,
        }),
        result.message,
      );
    },
    configureWindowsMpvShortcuts: async (preferences) => {
      if (!isWindows || !deps.applyWindowsMpvShortcuts) {
        return refreshWithState(
          writeState({
            ...readState(),
            windowsMpvShortcutPreferences: {
              startMenuEnabled: preferences.startMenuEnabled,
              desktopEnabled: preferences.desktopEnabled,
            },
          }),
          null,
        );
      }
      const result = await deps.applyWindowsMpvShortcuts(preferences);
      const latestState = readState();
      return refreshWithState(
        writeState({
          ...latestState,
          windowsMpvShortcutPreferences: {
            startMenuEnabled: preferences.startMenuEnabled,
            desktopEnabled: preferences.desktopEnabled,
          },
          windowsMpvShortcutLastStatus: result.status,
        }),
        result.message,
      );
    },
    isSetupCompleted: () => completed || isSetupCompleted(readState()),
  };
}
