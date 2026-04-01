import * as path from 'node:path';

import type { SubtitleCue } from '../core/services/subtitle-cue-parser';
import type { AnilistMediaGuess } from '../core/services/anilist/anilist-updater';
import type { OverlayHostedModal } from '../shared/ipc/contracts';
import type {
  FrequencyDictionaryLookup,
  ResolvedConfig,
  SubtitleData,
  SubtitlePosition,
} from '../types';
import {
  deleteYomitanDictionaryByTitle,
  getYomitanDictionaryInfo,
  importYomitanDictionaryFromZip,
  upsertYomitanDictionarySettings,
  clearYomitanParserCachesForWindow,
} from '../core/services';
import type { YomitanParserRuntimeDeps } from './yomitan-runtime';
import {
  createDictionarySupportRuntime,
  type DictionarySupportRuntime,
} from './dictionary-support-runtime';
import {
  createDictionarySupportRuntimeInput,
  type DictionarySupportRuntimeInputBuilderInput,
} from './dictionary-support-runtime-input';
import {
  createSubtitleRuntime,
  type SubtitleRuntime,
  type SubtitleRuntimeInput,
} from './subtitle-runtime';
import type { JlptLookup } from './jlpt-runtime';
import { formatSkippedYomitanWriteAction } from './runtime/yomitan-read-only-log';

type BrowserWindowLike = {
  isDestroyed: () => boolean;
  webContents: {
    send: (channel: string, payload?: unknown) => void;
  };
};

type ImmersionTrackerLike = {
  handleMediaChange: (path: string, title: string | null) => void;
} | null;

type MpvClientLike = {
  connected?: boolean;
  currentSubStart?: number | null;
  currentSubEnd?: number | null;
  currentTimePos?: number | null;
  currentVideoPath?: string | null;
  requestProperty: (name: string) => Promise<unknown>;
} | null;

type OverlayUiLike = {
  setVisibleOverlayVisible: (visible: boolean) => void;
  getRestoreVisibleOverlayOnModalClose: () => Set<OverlayHostedModal>;
  sendToActiveOverlayWindow: (
    channel: string,
    payload?: unknown,
    runtimeOptions?: {
      restoreOnModalClose?: OverlayHostedModal;
      preferModalWindow?: boolean;
    },
  ) => boolean;
} | null;

type OverlayManagerLike = {
  broadcastToOverlayWindows: (channel: string, payload?: unknown) => void;
  getMainWindow: () => BrowserWindowLike | null;
  getVisibleOverlayVisible: () => boolean;
};

type StartupOsdSequencerLike = NonNullable<
  DictionarySupportRuntimeInputBuilderInput['startup']['startupOsdSequencer']
>;

export interface SubtitleDictionaryRuntimeInput {
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
    configDir: string;
    defaultImmersionDbPath: string;
  };
  appState: {
    currentMediaPath: string | null;
    currentMediaTitle: string | null;
    currentSubText: string;
    currentSubAssText: string;
    mpvClient: MpvClientLike;
    subtitlePosition: SubtitlePosition | null;
    pendingSubtitlePosition: SubtitlePosition | null;
    currentSubtitleData: SubtitleData | null;
    activeParsedSubtitleCues: SubtitleCue[];
    activeParsedSubtitleSource: string | null;
    immersionTracker: ImmersionTrackerLike;
    jlptLevelLookup: JlptLookup;
    frequencyRankLookup: FrequencyDictionaryLookup;
    yomitanParserWindow: BrowserWindowLike | null;
  };
  config: {
    getResolvedConfig: () => ResolvedConfig;
  };
  services: {
    subtitleWsService: SubtitleRuntimeInput['subtitleWsService'];
    annotationSubtitleWsService: SubtitleRuntimeInput['annotationSubtitleWsService'];
    overlayManager: OverlayManagerLike;
    startupOsdSequencer: StartupOsdSequencerLike;
  };
  logging: {
    debug: (message: string) => void;
    info: (message: string) => void;
    warn: (message: string) => void;
    error: (message: string, ...args: unknown[]) => void;
  };
  subtitle: {
    parseSubtitleCues: (content: string, filename: string) => SubtitleCue[];
    createSubtitlePrefetchService: SubtitleRuntimeInput['createSubtitlePrefetchService'];
    schedule: (callback: () => void, delayMs: number) => ReturnType<typeof setTimeout>;
    clearSchedule: (timer: ReturnType<typeof setTimeout>) => void;
  };
  overlay: {
    getOverlayUi: () => OverlayUiLike;
    showMpvOsd: (message: string) => void;
    showDesktopNotification: (title: string, options: { body?: string }) => void;
  };
  playback: {
    isRemoteMediaPath: (mediaPath: string) => boolean;
    isYoutubePlaybackActive: (mediaPath: string | null, videoPath: string | null) => boolean;
    waitForYomitanMutationReady: (mediaKey: string | null) => Promise<void>;
  };
  anilist: {
    guessAnilistMediaInfo: (
      mediaPath: string | null,
      mediaTitle: string | null,
    ) => Promise<AnilistMediaGuess | null>;
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
}

export interface SubtitleDictionaryRuntime {
  subtitle: SubtitleRuntime;
  dictionarySupport: DictionarySupportRuntime;
}

export interface SubtitleDictionaryRuntimeCoordinatorInput {
  env: SubtitleDictionaryRuntimeInput['env'];
  appState: SubtitleDictionaryRuntimeInput['appState'];
  getResolvedConfig: () => ResolvedConfig;
  services: SubtitleDictionaryRuntimeInput['services'];
  logging: {
    debug: (message: string, ...args: unknown[]) => void;
    info: (message: string, ...args: unknown[]) => void;
    warn: (message: string, ...args: unknown[]) => void;
    error: (message: string, ...args: unknown[]) => void;
  };
  overlay: SubtitleDictionaryRuntimeInput['overlay'];
  playback: SubtitleDictionaryRuntimeInput['playback'];
  anilist: SubtitleDictionaryRuntimeInput['anilist'];
  subtitle: {
    parseSubtitleCues: (content: string, filename: string) => SubtitleCue[];
    createSubtitlePrefetchService: SubtitleRuntimeInput['createSubtitlePrefetchService'];
  };
  yomitan: {
    isCharacterDictionaryEnabled: () => boolean;
    isExternalReadOnlyMode: () => boolean;
    logSkippedWrite: (message: string) => void;
    ensureYomitanExtensionLoaded: () => Promise<unknown>;
    getParserRuntimeDeps: () => YomitanParserRuntimeDeps;
  };
}

export function createSubtitleDictionaryRuntime(
  input: SubtitleDictionaryRuntimeInput,
): SubtitleDictionaryRuntime {
  const subtitlePositionsDir = path.join(input.env.configDir, 'subtitle-positions');

  const subtitle = createSubtitleRuntime({
    getResolvedConfig: () => input.config.getResolvedConfig(),
    getCurrentMediaPath: () => input.appState.currentMediaPath,
    getCurrentMediaTitle: () => input.appState.currentMediaTitle,
    getCurrentSubText: () => input.appState.currentSubText,
    getCurrentSubAssText: () => input.appState.currentSubAssText,
    getMpvClient: () => input.appState.mpvClient,
    subtitleWsService: input.services.subtitleWsService,
    annotationSubtitleWsService: input.services.annotationSubtitleWsService,
    broadcastToOverlayWindows: (channel, payload) =>
      input.services.overlayManager.broadcastToOverlayWindows(channel, payload),
    subtitlePositionsDir,
    setSubtitlePosition: (position) => {
      input.appState.subtitlePosition = position;
    },
    setPendingSubtitlePosition: (position) => {
      input.appState.pendingSubtitlePosition = position;
    },
    clearPendingSubtitlePosition: () => {
      input.appState.pendingSubtitlePosition = null;
    },
    setCurrentSubtitleData: (payload) => {
      input.appState.currentSubtitleData = payload;
    },
    setActiveParsedSubtitleState: (cues, sourceKey) => {
      input.appState.activeParsedSubtitleCues = cues;
      input.appState.activeParsedSubtitleSource = sourceKey;
    },
    parseSubtitleCues: (content, filename) => input.subtitle.parseSubtitleCues(content, filename),
    createSubtitlePrefetchService: (deps) => input.subtitle.createSubtitlePrefetchService(deps),
    schedule: (callback, delayMs) => input.subtitle.schedule(callback, delayMs),
    clearSchedule: (timer) => input.subtitle.clearSchedule(timer),
    logDebug: (message) => input.logging.debug(message),
    logInfo: (message) => input.logging.info(message),
    logWarn: (message) => input.logging.warn(message),
  });

  const dictionarySupport = createDictionarySupportRuntime<OverlayHostedModal>(
    createDictionarySupportRuntimeInput({
      env: {
        platform: input.env.platform,
        dirname: input.env.dirname,
        appPath: input.env.appPath,
        resourcesPath: input.env.resourcesPath,
        userDataPath: input.env.userDataPath,
        appUserDataPath: input.env.appUserDataPath,
        homeDir: input.env.homeDir,
        appDataDir: input.env.appDataDir,
        cwd: input.env.cwd,
        subtitlePositionsDir,
        defaultImmersionDbPath: input.env.defaultImmersionDbPath,
      },
      config: {
        getResolvedConfig: () => input.config.getResolvedConfig(),
      },
      dictionaryState: {
        setJlptLevelLookup: (lookup) => {
          input.appState.jlptLevelLookup = lookup;
        },
        setFrequencyRankLookup: (lookup) => {
          input.appState.frequencyRankLookup = lookup;
        },
      },
      logger: {
        info: (message) => input.logging.info(message),
        debug: (message) => input.logging.debug(message),
        warn: (message) => input.logging.warn(message),
        error: (message, ...args) => input.logging.error(message, ...args),
      },
      media: {
        isRemoteMediaPath: (mediaPath) => input.playback.isRemoteMediaPath(mediaPath),
        getCurrentMediaPath: () => input.appState.currentMediaPath,
        setCurrentMediaPath: (mediaPath) => {
          input.appState.currentMediaPath = mediaPath;
        },
        getCurrentMediaTitle: () => input.appState.currentMediaTitle,
        setCurrentMediaTitle: (title) => {
          input.appState.currentMediaTitle = title;
        },
        getPendingSubtitlePosition: () => input.appState.pendingSubtitlePosition,
        clearPendingSubtitlePosition: () => {
          input.appState.pendingSubtitlePosition = null;
        },
        setSubtitlePosition: (position) => {
          input.appState.subtitlePosition = position;
        },
      },
      subtitle: {
        loadSubtitlePosition: () => subtitle.loadSubtitlePosition(),
        invalidateTokenizationCache: () => {
          subtitle.invalidateTokenizationCache();
        },
        refreshSubtitlePrefetchFromActiveTrack: () => {
          subtitle.refreshSubtitlePrefetchFromActiveTrack();
        },
        refreshCurrentSubtitle: (text) => {
          subtitle.refreshCurrentSubtitle(text);
        },
        getCurrentSubtitleText: () => input.appState.currentSubText,
      },
      overlay: {
        broadcastSubtitlePosition: (position) => {
          input.services.overlayManager.broadcastToOverlayWindows('subtitle:position', position);
        },
        broadcastToOverlayWindows: (channel, payload) => {
          input.services.overlayManager.broadcastToOverlayWindows(channel, payload);
        },
        getMainWindow: () => input.services.overlayManager.getMainWindow(),
        getVisibleOverlayVisible: () => input.services.overlayManager.getVisibleOverlayVisible(),
        setVisibleOverlayVisible: (visible) => {
          input.overlay.getOverlayUi()?.setVisibleOverlayVisible(visible);
        },
        getRestoreVisibleOverlayOnModalClose: () =>
          input.overlay.getOverlayUi()?.getRestoreVisibleOverlayOnModalClose() ??
          new Set<OverlayHostedModal>(),
        sendToActiveOverlayWindow: (channel, payload, runtimeOptions) =>
          input.overlay
            .getOverlayUi()
            ?.sendToActiveOverlayWindow(channel, payload, runtimeOptions) ?? false,
      },
      tracker: {
        getTracker: () => input.appState.immersionTracker,
        getMpvClient: () => input.appState.mpvClient,
      },
      anilist: {
        guessAnilistMediaInfo: (mediaPath, mediaTitle) =>
          input.anilist.guessAnilistMediaInfo(mediaPath, mediaTitle),
      },
      yomitan: {
        isCharacterDictionaryEnabled: () => input.yomitan.isCharacterDictionaryEnabled(),
        getYomitanDictionaryInfo: () => input.yomitan.getYomitanDictionaryInfo(),
        importYomitanDictionary: (zipPath) => input.yomitan.importYomitanDictionary(zipPath),
        deleteYomitanDictionary: (dictionaryTitle) =>
          input.yomitan.deleteYomitanDictionary(dictionaryTitle),
        upsertYomitanDictionarySettings: (dictionaryTitle, profileScope) =>
          input.yomitan.upsertYomitanDictionarySettings(dictionaryTitle, profileScope),
        hasParserWindow: () => input.yomitan.hasParserWindow(),
        clearParserCaches: () => input.yomitan.clearParserCaches(),
      },
      startup: {
        getNotificationType: () =>
          input.config.getResolvedConfig().ankiConnect.behavior.notificationType,
        showMpvOsd: (message) => input.overlay.showMpvOsd(message),
        showDesktopNotification: (title, options) =>
          input.overlay.showDesktopNotification(title, options),
        startupOsdSequencer: input.services.startupOsdSequencer,
      },
      playback: {
        isYoutubePlaybackActiveNow: () =>
          input.playback.isYoutubePlaybackActive(
            input.appState.currentMediaPath,
            input.appState.mpvClient?.currentVideoPath ?? null,
          ),
        waitForYomitanMutationReady: () =>
          input.playback.waitForYomitanMutationReady(
            input.appState.currentMediaPath?.trim() ||
              input.appState.mpvClient?.currentVideoPath?.trim() ||
              null,
          ),
      },
    }),
  );

  return {
    subtitle,
    dictionarySupport,
  };
}

export function createSubtitleDictionaryRuntimeCoordinator(
  input: SubtitleDictionaryRuntimeCoordinatorInput,
): SubtitleDictionaryRuntime {
  return createSubtitleDictionaryRuntime({
    env: input.env,
    appState: input.appState,
    config: {
      getResolvedConfig: () => input.getResolvedConfig(),
    },
    services: input.services,
    logging: input.logging,
    subtitle: {
      parseSubtitleCues: (content, filename) => input.subtitle.parseSubtitleCues(content, filename),
      createSubtitlePrefetchService: (deps) => input.subtitle.createSubtitlePrefetchService(deps),
      schedule: (callback, delayMs) => setTimeout(callback, delayMs),
      clearSchedule: (timer) => clearTimeout(timer),
    },
    overlay: input.overlay,
    playback: input.playback,
    anilist: input.anilist,
    yomitan: {
      isCharacterDictionaryEnabled: () => input.yomitan.isCharacterDictionaryEnabled(),
      getYomitanDictionaryInfo: async () => {
        await input.yomitan.ensureYomitanExtensionLoaded();
        return await getYomitanDictionaryInfo(input.yomitan.getParserRuntimeDeps(), {
          error: (message, ...args) => input.logging.error(message, ...args),
          info: (message, ...args) => input.logging.info(message, ...args),
        });
      },
      importYomitanDictionary: async (zipPath) => {
        if (input.yomitan.isExternalReadOnlyMode()) {
          input.yomitan.logSkippedWrite(
            formatSkippedYomitanWriteAction('importYomitanDictionary', zipPath),
          );
          return false;
        }
        await input.yomitan.ensureYomitanExtensionLoaded();
        return await importYomitanDictionaryFromZip(zipPath, input.yomitan.getParserRuntimeDeps(), {
          error: (message, ...args) => input.logging.error(message, ...args),
          info: (message, ...args) => input.logging.info(message, ...args),
        });
      },
      deleteYomitanDictionary: async (dictionaryTitle) => {
        if (input.yomitan.isExternalReadOnlyMode()) {
          input.yomitan.logSkippedWrite(
            formatSkippedYomitanWriteAction('deleteYomitanDictionary', dictionaryTitle),
          );
          return false;
        }
        await input.yomitan.ensureYomitanExtensionLoaded();
        return await deleteYomitanDictionaryByTitle(
          dictionaryTitle,
          input.yomitan.getParserRuntimeDeps(),
          {
            error: (message, ...args) => input.logging.error(message, ...args),
            info: (message, ...args) => input.logging.info(message, ...args),
          },
        );
      },
      upsertYomitanDictionarySettings: async (dictionaryTitle, profileScope) => {
        if (input.yomitan.isExternalReadOnlyMode()) {
          input.yomitan.logSkippedWrite(
            formatSkippedYomitanWriteAction('upsertYomitanDictionarySettings', dictionaryTitle),
          );
          return false;
        }
        await input.yomitan.ensureYomitanExtensionLoaded();
        return await upsertYomitanDictionarySettings(
          dictionaryTitle,
          profileScope,
          input.yomitan.getParserRuntimeDeps(),
          {
            error: (message, ...args) => input.logging.error(message, ...args),
            info: (message, ...args) => input.logging.info(message, ...args),
          },
        );
      },
      hasParserWindow: () => Boolean(input.appState.yomitanParserWindow),
      clearParserCaches: () => {
        if (input.appState.yomitanParserWindow) {
          clearYomitanParserCachesForWindow(input.appState.yomitanParserWindow as never);
        }
      },
    },
  });
}
