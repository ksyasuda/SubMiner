import * as path from 'path';

import type {
  AnilistCharacterDictionaryProfileScope,
  FrequencyDictionaryLookup,
  KikuFieldGroupingChoice,
  KikuFieldGroupingRequestData,
  SubtitlePosition,
  ResolvedConfig,
} from '../types';
import {
  createBuildDictionaryRootsMainHandler,
  createBuildFrequencyDictionaryRuntimeMainDepsHandler,
  createBuildJlptDictionaryRuntimeMainDepsHandler,
} from './runtime/dictionary-runtime-main-deps';
import { createImmersionMediaRuntime } from './runtime/immersion-media';
import {
  createFrequencyDictionaryRuntimeService,
  getFrequencyDictionarySearchPaths,
} from './frequency-dictionary-runtime';
import {
  createJlptDictionaryRuntimeService,
  getJlptDictionarySearchPaths,
  type JlptLookup,
} from './jlpt-runtime';
import { createMediaRuntimeService } from './media-runtime';
import {
  createCharacterDictionaryRuntimeService,
  type CharacterDictionaryBuildResult,
} from './character-dictionary-runtime';
import type { AnilistMediaGuess } from '../core/services/anilist/anilist-updater';
import {
  createCharacterDictionaryAutoSyncRuntimeService,
  type CharacterDictionaryAutoSyncConfig,
  type CharacterDictionaryAutoSyncStatusEvent,
} from './runtime/character-dictionary-auto-sync';
import { handleCharacterDictionaryAutoSyncComplete } from './runtime/character-dictionary-auto-sync-completion';
import { notifyCharacterDictionaryAutoSyncStatus } from './runtime/character-dictionary-auto-sync-notifications';
import { createFieldGroupingOverlayRuntime } from '../core/services/field-grouping-overlay';

type FieldGroupingResolver = ((choice: KikuFieldGroupingChoice) => void) | null;

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

type CharacterDictionaryAutoSyncCompleteDeps = {
  hasParserWindow: () => boolean;
  clearParserCaches: () => void;
  invalidateTokenizationCache: () => void;
  refreshSubtitlePrefetch: () => void;
  refreshCurrentSubtitle: () => void;
  logInfo: (message: string) => void;
};

export interface DictionarySupportRuntimeInput<TModal extends string = string> {
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
  getResolvedConfig: () => ResolvedConfig;
  isJlptEnabled: () => boolean;
  isFrequencyDictionaryEnabled: () => boolean;
  getFrequencyDictionarySourcePath: () => string | undefined;
  setJlptLevelLookup: (lookup: JlptLookup) => void;
  setFrequencyRankLookup: (lookup: FrequencyDictionaryLookup) => void;
  logInfo: (message: string) => void;
  logDebug?: (message: string) => void;
  logWarn: (message: string) => void;
  isRemoteMediaPath: (mediaPath: string) => boolean;
  getCurrentMediaPath: () => string | null;
  setCurrentMediaPath: (mediaPath: string | null) => void;
  getCurrentMediaTitle: () => string | null;
  setCurrentMediaTitle: (title: string | null) => void;
  getPendingSubtitlePosition: () => SubtitlePosition | null;
  loadSubtitlePosition: () => SubtitlePosition | null;
  clearPendingSubtitlePosition: () => void;
  setSubtitlePosition: (position: SubtitlePosition | null) => void;
  broadcastSubtitlePosition: (position: SubtitlePosition | null) => void;
  broadcastToOverlayWindows: (channel: string, payload?: unknown) => void;
  getTracker: () => ImmersionTrackerLike | null;
  getMpvClient: () => MpvClientLike | null;
  defaultImmersionDbPath: string;
  guessAnilistMediaInfo: (
    mediaPath: string | null,
    mediaTitle: string | null,
  ) => Promise<AnilistMediaGuess | null>;
  getCollapsibleSectionOpenState: (
    section: keyof ResolvedConfig['anilist']['characterDictionary']['collapsibleSections'],
  ) => boolean;
  isCharacterDictionaryEnabled: () => boolean;
  isYoutubePlaybackActiveNow: () => boolean;
  waitForYomitanMutationReady: () => Promise<void>;
  getYomitanDictionaryInfo: () => Promise<Array<{ title: string; revision?: string | number }>>;
  importYomitanDictionary: (zipPath: string) => Promise<boolean>;
  deleteYomitanDictionary: (dictionaryTitle: string) => Promise<boolean>;
  upsertYomitanDictionarySettings: (
    dictionaryTitle: string,
    profileScope: AnilistCharacterDictionaryProfileScope,
  ) => Promise<boolean>;
  getCharacterDictionaryConfig: () => CharacterDictionaryAutoSyncConfig;
  notifyCharacterDictionaryAutoSyncStatus: (event: CharacterDictionaryAutoSyncStatusEvent) => void;
  characterDictionaryAutoSyncCompleteDeps: CharacterDictionaryAutoSyncCompleteDeps;
  getMainWindow: () => BrowserWindowLike | null;
  getVisibleOverlayVisible: () => boolean;
  setVisibleOverlayVisible: (visible: boolean) => void;
  getRestoreVisibleOverlayOnModalClose: () => Set<TModal>;
  sendToActiveOverlayWindow: (
    channel: string,
    payload?: unknown,
    runtimeOptions?: { restoreOnModalClose?: TModal },
  ) => boolean;
}

export interface DictionarySupportRuntime {
  ensureJlptDictionaryLookup: () => Promise<void>;
  ensureFrequencyDictionaryLookup: () => Promise<void>;
  getFieldGroupingResolver: () => FieldGroupingResolver;
  setFieldGroupingResolver: (resolver: FieldGroupingResolver) => void;
  createFieldGroupingCallback: () => (
    data: KikuFieldGroupingRequestData,
  ) => Promise<KikuFieldGroupingChoice>;
  getConfiguredDbPath: () => string;
  seedImmersionMediaFromCurrentMedia: () => Promise<void>;
  syncImmersionMediaState: () => void;
  resolveMediaPathForJimaku: (mediaPath: string | null) => string | null;
  updateCurrentMediaPath: (mediaPath: unknown) => void;
  updateCurrentMediaTitle: (mediaTitle: unknown) => void;
  scheduleCharacterDictionarySync: () => void;
  generateCharacterDictionaryForCurrentMedia: (
    targetPath?: string,
  ) => Promise<CharacterDictionaryBuildResult>;
}

export function createDictionarySupportRuntime<TModal extends string>(
  input: DictionarySupportRuntimeInput<TModal>,
): DictionarySupportRuntime {
  const dictionaryRoots = createBuildDictionaryRootsMainHandler({
    platform: input.platform,
    dirname: input.dirname,
    appPath: input.appPath,
    resourcesPath: input.resourcesPath,
    userDataPath: input.userDataPath,
    appUserDataPath: input.appUserDataPath,
    homeDir: input.homeDir,
    appDataDir: input.appDataDir,
    cwd: input.cwd,
    joinPath: (...parts: string[]) => path.join(...parts),
  });

  const jlptRuntime = createJlptDictionaryRuntimeService(
    createBuildJlptDictionaryRuntimeMainDepsHandler({
      isJlptEnabled: () => input.isJlptEnabled(),
      getDictionaryRoots: () => dictionaryRoots(),
      getJlptDictionarySearchPaths,
      setJlptLevelLookup: (lookup) => input.setJlptLevelLookup(lookup),
      logInfo: (message) => input.logInfo(message),
    })(),
  );

  const frequencyRuntime = createFrequencyDictionaryRuntimeService(
    createBuildFrequencyDictionaryRuntimeMainDepsHandler({
      isFrequencyDictionaryEnabled: () => input.isFrequencyDictionaryEnabled(),
      getDictionaryRoots: () => dictionaryRoots(),
      getFrequencyDictionarySearchPaths,
      getSourcePath: () => input.getFrequencyDictionarySourcePath(),
      setFrequencyRankLookup: (lookup) => input.setFrequencyRankLookup(lookup),
      logInfo: (message) => input.logInfo(message),
    })(),
  );

  let fieldGroupingResolver: FieldGroupingResolver = null;
  let fieldGroupingResolverSequence = 0;

  const getFieldGroupingResolver = (): FieldGroupingResolver => fieldGroupingResolver;
  const setFieldGroupingResolver = (resolver: FieldGroupingResolver): void => {
    if (!resolver) {
      fieldGroupingResolver = null;
      return;
    }
    const sequence = ++fieldGroupingResolverSequence;
    fieldGroupingResolver = (choice) => {
      if (sequence !== fieldGroupingResolverSequence) {
        return;
      }
      resolver(choice);
    };
  };

  const fieldGroupingOverlayRuntime = createFieldGroupingOverlayRuntime<TModal>({
    getMainWindow: () => input.getMainWindow(),
    getVisibleOverlayVisible: () => input.getVisibleOverlayVisible(),
    setVisibleOverlayVisible: (visible) => input.setVisibleOverlayVisible(visible),
    getResolver: () => getFieldGroupingResolver(),
    setResolver: (resolver) => setFieldGroupingResolver(resolver),
    getRestoreVisibleOverlayOnModalClose: () => input.getRestoreVisibleOverlayOnModalClose(),
    sendToVisibleOverlay: (channel, payload, runtimeOptions) =>
      input.sendToActiveOverlayWindow(channel, payload, runtimeOptions),
  });

  const immersionMediaRuntime = createImmersionMediaRuntime({
    getResolvedConfig: () => input.getResolvedConfig(),
    defaultImmersionDbPath: input.defaultImmersionDbPath,
    getTracker: () => input.getTracker(),
    getMpvClient: () => input.getMpvClient(),
    getCurrentMediaPath: () => input.getCurrentMediaPath(),
    getCurrentMediaTitle: () => input.getCurrentMediaTitle(),
    logDebug: (message) => (input.logDebug ?? input.logInfo)(message),
    logInfo: (message) => input.logInfo(message),
  });

  const mediaRuntime = createMediaRuntimeService({
    isRemoteMediaPath: (mediaPath) => input.isRemoteMediaPath(mediaPath),
    loadSubtitlePosition: () => input.loadSubtitlePosition(),
    getCurrentMediaPath: () => input.getCurrentMediaPath(),
    getPendingSubtitlePosition: () => input.getPendingSubtitlePosition(),
    getSubtitlePositionsDir: () => input.subtitlePositionsDir,
    setCurrentMediaPath: (mediaPath) => input.setCurrentMediaPath(mediaPath),
    clearPendingSubtitlePosition: () => input.clearPendingSubtitlePosition(),
    setSubtitlePosition: (position) => input.setSubtitlePosition(position),
    broadcastSubtitlePosition: (position) => input.broadcastSubtitlePosition(position),
    getCurrentMediaTitle: () => input.getCurrentMediaTitle(),
    setCurrentMediaTitle: (title) => input.setCurrentMediaTitle(title),
  });

  const characterDictionaryRuntime = createCharacterDictionaryRuntimeService({
    userDataPath: input.userDataPath,
    getCurrentMediaPath: () => input.getCurrentMediaPath(),
    getCurrentMediaTitle: () => input.getCurrentMediaTitle(),
    resolveMediaPathForJimaku: (mediaPath) => mediaRuntime.resolveMediaPathForJimaku(mediaPath),
    guessAnilistMediaInfo: (mediaPath, mediaTitle) =>
      input.guessAnilistMediaInfo(mediaPath, mediaTitle),
    getCollapsibleSectionOpenState: (section) => input.getCollapsibleSectionOpenState(section),
    now: () => Date.now(),
    logInfo: (message) => input.logInfo(message),
    logWarn: (message) => input.logWarn(message),
  });

  const characterDictionaryAutoSyncRuntime = createCharacterDictionaryAutoSyncRuntimeService({
    userDataPath: input.userDataPath,
    getConfig: () => input.getCharacterDictionaryConfig(),
    getOrCreateCurrentSnapshot: (targetPath, progress) =>
      characterDictionaryRuntime.getOrCreateCurrentSnapshot(targetPath, progress),
    buildMergedDictionary: (mediaIds) => characterDictionaryRuntime.buildMergedDictionary(mediaIds),
    waitForYomitanMutationReady: () => input.waitForYomitanMutationReady(),
    getYomitanDictionaryInfo: () => input.getYomitanDictionaryInfo(),
    importYomitanDictionary: (zipPath) => input.importYomitanDictionary(zipPath),
    deleteYomitanDictionary: (dictionaryTitle) => input.deleteYomitanDictionary(dictionaryTitle),
    upsertYomitanDictionarySettings: (dictionaryTitle, profileScope) =>
      input.upsertYomitanDictionarySettings(dictionaryTitle, profileScope),
    now: () => Date.now(),
    schedule: (fn, delayMs) => setTimeout(fn, delayMs),
    clearSchedule: (timer) => clearTimeout(timer),
    logInfo: (message) => input.logInfo(message),
    logWarn: (message) => input.logWarn(message),
    onSyncStatus: (event) => input.notifyCharacterDictionaryAutoSyncStatus(event),
    onSyncComplete: (result) =>
      handleCharacterDictionaryAutoSyncComplete(
        result,
        input.characterDictionaryAutoSyncCompleteDeps,
      ),
  });

  const scheduleCharacterDictionarySync = (): void => {
    if (!input.isCharacterDictionaryEnabled() || input.isYoutubePlaybackActiveNow()) {
      return;
    }
    characterDictionaryAutoSyncRuntime.scheduleSync();
  };

  return {
    ensureJlptDictionaryLookup: () => jlptRuntime.ensureJlptDictionaryLookup(),
    ensureFrequencyDictionaryLookup: () => frequencyRuntime.ensureFrequencyDictionaryLookup(),
    getFieldGroupingResolver,
    setFieldGroupingResolver,
    createFieldGroupingCallback: () => fieldGroupingOverlayRuntime.createFieldGroupingCallback(),
    getConfiguredDbPath: () => immersionMediaRuntime.getConfiguredDbPath(),
    seedImmersionMediaFromCurrentMedia: () => immersionMediaRuntime.seedFromCurrentMedia(),
    syncImmersionMediaState: () => immersionMediaRuntime.syncFromCurrentMediaState(),
    resolveMediaPathForJimaku: (mediaPath) => mediaRuntime.resolveMediaPathForJimaku(mediaPath),
    updateCurrentMediaPath: (mediaPath) => mediaRuntime.updateCurrentMediaPath(mediaPath),
    updateCurrentMediaTitle: (mediaTitle) => mediaRuntime.updateCurrentMediaTitle(mediaTitle),
    scheduleCharacterDictionarySync,
    generateCharacterDictionaryForCurrentMedia: (targetPath?: string) =>
      characterDictionaryRuntime.generateForCurrentMedia(targetPath),
  };
}
