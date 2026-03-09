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
  pluginStatus: 'installed' | 'optional' | 'skipped' | 'failed';
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
  skipPluginInstall: () => Promise<SetupStatusSnapshot>;
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
    args.mineSentence ||
    args.mineSentenceMultiple ||
    args.updateLastCardFromClipboard ||
    args.refreshKnownWords ||
    args.toggleSecondarySub ||
    args.triggerFieldGrouping ||
    args.triggerSubsync ||
    args.markAudioCard ||
    args.openRuntimeOptions ||
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
  if (state.pluginInstallStatus === 'skipped') return 'skipped';
  if (state.pluginInstallStatus === 'failed') return 'failed';
  return 'optional';
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

export function createFirstRunSetupService(deps: {
  platform?: NodeJS.Platform;
  configDir: string;
  getYomitanDictionaryCount: () => Promise<number>;
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
    const dictionaryCount = await deps.getYomitanDictionaryCount();
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
    const configReady =
      fs.existsSync(configFilePaths.jsoncPath) || fs.existsSync(configFilePaths.jsonPath);
    return {
      configReady,
      dictionaryCount,
      canFinish: dictionaryCount >= 1,
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
      if (isSetupCompleted(state)) {
        completed = true;
        return refreshWithState(state);
      }

      const dictionaryCount = await deps.getYomitanDictionaryCount();
      const configReady =
        fs.existsSync(configFilePaths.jsoncPath) || fs.existsSync(configFilePaths.jsonPath);
      if (configReady && dictionaryCount >= 1) {
        const completedState = writeState({
          ...state,
          status: 'completed',
          completedAt: new Date().toISOString(),
          completionSource: 'legacy_auto_detected',
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
          lastSeenYomitanDictionaryCount: snapshot.dictionaryCount,
        }),
      );
    },
    skipPluginInstall: async () =>
      refreshWithState(writeState({ ...readState(), pluginInstallStatus: 'skipped' })),
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
