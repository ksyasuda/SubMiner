import type { OverlayHostedModal } from '../shared/ipc/contracts';
import type { FrequencyDictionaryLookup, ResolvedConfig } from '../types';
import type { JlptLookup } from './jlpt-runtime';
import type { DictionarySupportRuntimeInput } from './dictionary-support-runtime';
import { notifyCharacterDictionaryAutoSyncStatus } from './runtime/character-dictionary-auto-sync-notifications';
import type { StartupOsdSequencerCharacterDictionaryEvent } from './runtime/startup-osd-sequencer';

type BrowserWindowLike = {
  isDestroyed: () => boolean;
  webContents: {
    send: (channel: string, payload?: unknown) => void;
  };
};

type ImmersionTrackerLike = {
  handleMediaChange: (path: string, title: string | null) => void;
};

type MpvClientLike = {
  currentVideoPath?: string | null;
  connected?: boolean;
  requestProperty?: (name: string) => Promise<unknown>;
};

type StartupOsdSequencerLike = {
  notifyCharacterDictionaryStatus: (event: StartupOsdSequencerCharacterDictionaryEvent) => void;
};

export interface DictionarySupportRuntimeInputBuilderInput {
  env: {
    platform: NodeJS.Platform;
    dirname: string;
    appPath: string;
    resourcesPath: string;
    userDataPath: string;
    appUserDataPath: string;
    homeDir: string;
    appDataDir?: string;
    cwd: string;
    subtitlePositionsDir: string;
    defaultImmersionDbPath: string;
  };
  config: {
    getResolvedConfig: () => ResolvedConfig;
  };
  dictionaryState: {
    setJlptLevelLookup: (lookup: JlptLookup) => void;
    setFrequencyRankLookup: (lookup: FrequencyDictionaryLookup) => void;
  };
  logger: {
    info: (message: string) => void;
    debug: (message: string) => void;
    warn: (message: string) => void;
    error: (message: string, ...args: unknown[]) => void;
  };
  media: {
    isRemoteMediaPath: (mediaPath: string) => boolean;
    getCurrentMediaPath: () => string | null;
    setCurrentMediaPath: (mediaPath: string | null) => void;
    getCurrentMediaTitle: () => string | null;
    setCurrentMediaTitle: (title: string | null) => void;
    getPendingSubtitlePosition: () => DictionarySupportRuntimeInput['getPendingSubtitlePosition'] extends () => infer T
      ? T
      : never;
    clearPendingSubtitlePosition: () => void;
    setSubtitlePosition: DictionarySupportRuntimeInput['setSubtitlePosition'];
  };
  subtitle: {
    loadSubtitlePosition: DictionarySupportRuntimeInput['loadSubtitlePosition'];
    invalidateTokenizationCache: () => void;
    refreshSubtitlePrefetchFromActiveTrack: () => void;
    refreshCurrentSubtitle: (text: string) => void;
    getCurrentSubtitleText: () => string;
  };
  overlay: {
    broadcastSubtitlePosition: DictionarySupportRuntimeInput<OverlayHostedModal>['broadcastSubtitlePosition'];
    broadcastToOverlayWindows: DictionarySupportRuntimeInput<OverlayHostedModal>['broadcastToOverlayWindows'];
    getMainWindow: () => BrowserWindowLike | null;
    getVisibleOverlayVisible: () => boolean;
    setVisibleOverlayVisible: (visible: boolean) => void;
    getRestoreVisibleOverlayOnModalClose: () => Set<OverlayHostedModal>;
    sendToActiveOverlayWindow: DictionarySupportRuntimeInput<OverlayHostedModal>['sendToActiveOverlayWindow'];
  };
  tracker: {
    getTracker: () => ImmersionTrackerLike | null;
    getMpvClient: () => MpvClientLike | null;
  };
  anilist: {
    guessAnilistMediaInfo: DictionarySupportRuntimeInput['guessAnilistMediaInfo'];
  };
  yomitan: {
    isCharacterDictionaryEnabled: () => boolean;
    getYomitanDictionaryInfo: () => Promise<Array<{ title: string; revision?: string | number }>>;
    importYomitanDictionary: (zipPath: string) => Promise<boolean>;
    deleteYomitanDictionary: (dictionaryTitle: string) => Promise<boolean>;
    upsertYomitanDictionarySettings: (
      dictionaryTitle: string,
      profileScope: ResolvedConfig['anilist']['characterDictionary']['profileScope'],
    ) => Promise<boolean>;
    hasParserWindow: () => boolean;
    clearParserCaches: () => void;
  };
  startup: {
    getNotificationType: () => 'osd' | 'system' | 'both' | 'none' | undefined;
    showMpvOsd: (message: string) => void;
    showDesktopNotification: (title: string, options: { body?: string }) => void;
    startupOsdSequencer?: StartupOsdSequencerLike;
  };
  playback: {
    isYoutubePlaybackActiveNow: () => boolean;
    waitForYomitanMutationReady: () => Promise<void>;
  };
}

export function createDictionarySupportRuntimeInput(
  input: DictionarySupportRuntimeInputBuilderInput,
): DictionarySupportRuntimeInput<OverlayHostedModal> {
  return {
    platform: input.env.platform,
    dirname: input.env.dirname,
    appPath: input.env.appPath,
    resourcesPath: input.env.resourcesPath,
    userDataPath: input.env.userDataPath,
    appUserDataPath: input.env.appUserDataPath,
    homeDir: input.env.homeDir,
    appDataDir: input.env.appDataDir,
    cwd: input.env.cwd,
    subtitlePositionsDir: input.env.subtitlePositionsDir,
    getResolvedConfig: () => input.config.getResolvedConfig(),
    isJlptEnabled: () => input.config.getResolvedConfig().subtitleStyle.enableJlpt,
    isFrequencyDictionaryEnabled: () =>
      input.config.getResolvedConfig().subtitleStyle.frequencyDictionary.enabled,
    getFrequencyDictionarySourcePath: () =>
      input.config.getResolvedConfig().subtitleStyle.frequencyDictionary.sourcePath,
    setJlptLevelLookup: (lookup) => input.dictionaryState.setJlptLevelLookup(lookup),
    setFrequencyRankLookup: (lookup) => input.dictionaryState.setFrequencyRankLookup(lookup),
    logInfo: (message) => input.logger.info(message),
    logDebug: (message) => input.logger.debug(message),
    logWarn: (message) => input.logger.warn(message),
    isRemoteMediaPath: (mediaPath) => input.media.isRemoteMediaPath(mediaPath),
    getCurrentMediaPath: () => input.media.getCurrentMediaPath(),
    setCurrentMediaPath: (mediaPath) => input.media.setCurrentMediaPath(mediaPath),
    getCurrentMediaTitle: () => input.media.getCurrentMediaTitle(),
    setCurrentMediaTitle: (title) => input.media.setCurrentMediaTitle(title),
    getPendingSubtitlePosition: () => input.media.getPendingSubtitlePosition(),
    loadSubtitlePosition: () => input.subtitle.loadSubtitlePosition(),
    clearPendingSubtitlePosition: () => input.media.clearPendingSubtitlePosition(),
    setSubtitlePosition: (position) => input.media.setSubtitlePosition(position),
    broadcastSubtitlePosition: (position) => input.overlay.broadcastSubtitlePosition(position),
    broadcastToOverlayWindows: (channel, payload) =>
      input.overlay.broadcastToOverlayWindows(channel, payload),
    getTracker: () => input.tracker.getTracker(),
    getMpvClient: () => input.tracker.getMpvClient(),
    defaultImmersionDbPath: input.env.defaultImmersionDbPath,
    guessAnilistMediaInfo: (mediaPath, mediaTitle) =>
      input.anilist.guessAnilistMediaInfo(mediaPath, mediaTitle),
    getCollapsibleSectionOpenState: (section) =>
      input.config.getResolvedConfig().anilist.characterDictionary.collapsibleSections[section],
    isCharacterDictionaryEnabled: () => input.yomitan.isCharacterDictionaryEnabled(),
    isYoutubePlaybackActiveNow: () => input.playback.isYoutubePlaybackActiveNow(),
    waitForYomitanMutationReady: () => input.playback.waitForYomitanMutationReady(),
    getYomitanDictionaryInfo: () => input.yomitan.getYomitanDictionaryInfo(),
    importYomitanDictionary: (zipPath) => input.yomitan.importYomitanDictionary(zipPath),
    deleteYomitanDictionary: (dictionaryTitle) =>
      input.yomitan.deleteYomitanDictionary(dictionaryTitle),
    upsertYomitanDictionarySettings: (dictionaryTitle, profileScope) =>
      input.yomitan.upsertYomitanDictionarySettings(dictionaryTitle, profileScope),
    getCharacterDictionaryConfig: () => {
      const config = input.config.getResolvedConfig().anilist.characterDictionary;
      return {
        enabled:
          config.enabled &&
          input.yomitan.isCharacterDictionaryEnabled() &&
          !input.playback.isYoutubePlaybackActiveNow(),
        maxLoaded: config.maxLoaded,
        profileScope: config.profileScope,
      };
    },
    notifyCharacterDictionaryAutoSyncStatus: (event) => {
      notifyCharacterDictionaryAutoSyncStatus(event, {
        getNotificationType: () => input.startup.getNotificationType(),
        showOsd: (message) => input.startup.showMpvOsd(message),
        showDesktopNotification: (title, options) =>
          input.startup.showDesktopNotification(title, options),
        startupOsdSequencer: input.startup.startupOsdSequencer,
      });
    },
    characterDictionaryAutoSyncCompleteDeps: {
      hasParserWindow: () => input.yomitan.hasParserWindow(),
      clearParserCaches: () => input.yomitan.clearParserCaches(),
      invalidateTokenizationCache: () => input.subtitle.invalidateTokenizationCache(),
      refreshSubtitlePrefetch: () => input.subtitle.refreshSubtitlePrefetchFromActiveTrack(),
      refreshCurrentSubtitle: () =>
        input.subtitle.refreshCurrentSubtitle(input.subtitle.getCurrentSubtitleText()),
      logInfo: (message) => input.logger.info(message),
    },
    getMainWindow: () => input.overlay.getMainWindow(),
    getVisibleOverlayVisible: () => input.overlay.getVisibleOverlayVisible(),
    setVisibleOverlayVisible: (visible) => input.overlay.setVisibleOverlayVisible(visible),
    getRestoreVisibleOverlayOnModalClose: () =>
      input.overlay.getRestoreVisibleOverlayOnModalClose(),
    sendToActiveOverlayWindow: (channel, payload, runtimeOptions) =>
      input.overlay.sendToActiveOverlayWindow(channel, payload, runtimeOptions),
  };
}
