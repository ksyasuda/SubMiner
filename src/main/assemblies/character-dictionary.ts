import {
  clearYomitanParserCachesForWindow,
  deleteYomitanDictionaryByTitle,
  getYomitanDictionaryInfo,
  importYomitanDictionaryFromZip,
  upsertYomitanDictionarySettings,
} from '../../core/services';
import { guessAnilistMediaInfo } from '../../core/services/anilist/anilist-updater';
import type { Logger } from '../../logger';
import type { ResolvedConfig } from '../../types';
import { createCharacterDictionaryRuntimeService } from '../character-dictionary-runtime';
import { createCharacterDictionaryImageLookup } from '../character-dictionary-runtime/image-lookup';
import { createCharacterNameCandidateLookup } from '../character-dictionary-runtime/name-candidates';
import { applyCharacterDictionarySelection } from '../character-dictionary-selection';
import {
  createCharacterDictionaryAutoSyncRuntimeService,
  getCharacterDictionaryManagerSnapshot,
  moveCharacterDictionaryManagedEntry,
  removeCharacterDictionaryManagedEntry,
  replaceCharacterDictionaryManagedEntry,
} from '../runtime/character-dictionary-auto-sync';
import { handleCharacterDictionaryAutoSyncComplete } from '../runtime/character-dictionary-auto-sync-completion';
import {
  notifyCharacterDictionaryAutoSyncStatus,
  type CharacterDictionaryAutoSyncNotificationDeps,
} from '../runtime/character-dictionary-auto-sync-notifications';
import type { createOverlayNotificationsRuntime } from '../runtime/overlay-notifications-runtime';
import { assertYomitanDictionaryMutationSafe } from '../runtime/yomitan-dictionary-integrity';
import type { createYomitanProfilePolicy } from '../runtime/yomitan-profile-policy';
import { formatSkippedYomitanWriteAction } from '../runtime/yomitan-read-only-log';
import type { AppState } from '../state';
import type { CliSlice, IpcSlice } from './slices';

type YomitanParserRuntimeDeps = Parameters<typeof getYomitanDictionaryInfo>[0];

/**
 * What the character dictionary needs from the rest of the app. Siblings built after this
 * assembly in `main.ts` are reached through getters or callbacks, which run only once the app does.
 */
export interface CharacterDictionaryAssemblyDeps {
  userDataPath: string;
  appState: Pick<
    AppState,
    | 'currentMediaPath'
    | 'mpvClient'
    | 'currentMediaTitle'
    | 'mecabTokenizer'
    | 'yomitanParserWindow'
    | 'currentSubText'
  >;
  getConfig: () => ResolvedConfig;
  logger: Pick<Logger, 'info' | 'warn' | 'error'>;
  startupOsdSequencer: CharacterDictionaryAutoSyncNotificationDeps['startupOsdSequencer'];
  resolveMediaPathForJimaku: (mediaPath: string | null) => string | null;
  getYomitanProfilePolicy: () => Pick<
    ReturnType<typeof createYomitanProfilePolicy>,
    | 'isCharacterDictionaryEnabled'
    | 'getCharacterDictionaryDisabledReason'
    | 'isExternalReadOnlyMode'
    | 'logSkippedWrite'
  >;
  isYoutubePlaybackActive: () => boolean;
  waitForMediaTokenizationReady: (mediaPath: string | null) => Promise<void>;
  ensureYomitanExtensionLoaded: () => Promise<unknown>;
  getYomitanParserRuntimeDeps: () => YomitanParserRuntimeDeps;
  getOverlayNotifications: () => Pick<
    ReturnType<typeof createOverlayNotificationsRuntime>,
    'getConfiguredStatusNotificationType' | 'showOverlayNotification'
  >;
  showMpvOsd: (message: string) => void;
  showDesktopNotification: CharacterDictionaryAutoSyncNotificationDeps['showDesktopNotification'];
  invalidateTokenizationCache: () => void;
  refreshCurrentSubtitle: (text: string) => void;
  /** Re-runs subtitle prefetch from the current playback position. */
  refreshSubtitlePrefetch: () => void;
  refreshCurrentSubtitleAnnotations: () => void;
  /** A manual selection invalidates AniList's guess for the current media. */
  resetAnilistMediaGuessState: () => void;
}

export type CharacterDictionaryAssembly = ReturnType<typeof createCharacterDictionaryAssembly>;

/**
 * Owns the AniList character dictionary: snapshot generation, auto-sync into Yomitan, and the
 * portrait and name-candidate lookups the overlay reads. `main.ts` builds it once.
 */
export function createCharacterDictionaryAssembly(deps: CharacterDictionaryAssemblyDeps) {
  const { appState, logger } = deps;

  const runtime = createCharacterDictionaryRuntimeService({
    userDataPath: deps.userDataPath,
    getCurrentMediaPath: () => appState.currentMediaPath,
    getCurrentVideoPath: () => appState.mpvClient?.currentVideoPath,
    getCurrentMediaTitle: () => appState.currentMediaTitle,
    resolveMediaPathForJimaku: (mediaPath) => deps.resolveMediaPathForJimaku(mediaPath),
    guessAnilistMediaInfo: (mediaPath, mediaTitle) => guessAnilistMediaInfo(mediaPath, mediaTitle),
    getNameMatchImagesEnabled: () => deps.getConfig().subtitleStyle.nameMatchImagesEnabled,
    getCollapsibleSectionOpenState: (section) =>
      deps.getConfig().anilist.characterDictionary.collapsibleSections[section],
    tokenizeJapaneseName: async (text) => (await appState.mecabTokenizer?.tokenize(text)) ?? null,
    getJapaneseNameTokenizerAvailable: () => {
      const status = appState.mecabTokenizer?.getStatus();
      return status?.available === true && status.enabled === true;
    },
    now: () => Date.now(),
    logInfo: (message) => logger.info(message),
    logWarn: (message) => logger.warn(message),
  });

  const autoSync = createCharacterDictionaryAutoSyncRuntimeService({
    userDataPath: deps.userDataPath,
    getConfig: () => {
      const config = deps.getConfig().anilist.characterDictionary;
      return {
        enabled:
          deps.getConfig().subtitleStyle.nameMatchEnabled &&
          deps.getYomitanProfilePolicy().isCharacterDictionaryEnabled() &&
          !deps.isYoutubePlaybackActive(),
        maxLoaded: config.maxLoaded,
        profileScope: config.profileScope,
      };
    },
    getOrCreateCurrentSnapshot: (targetPath, progress) =>
      runtime.getOrCreateCurrentSnapshot(targetPath, progress),
    buildMergedDictionary: (mediaIds) => runtime.buildMergedDictionary(mediaIds),
    waitForYomitanMutationReady: () =>
      deps.waitForMediaTokenizationReady(
        appState.currentMediaPath?.trim() || appState.mpvClient?.currentVideoPath?.trim() || null,
      ),
    getYomitanDictionaryInfo: async () => {
      await deps.ensureYomitanExtensionLoaded();
      const dictionaries = await getYomitanDictionaryInfo(deps.getYomitanParserRuntimeDeps(), {
        error: (message, ...args) => logger.error(message, ...args),
        info: (message, ...args) => logger.info(message, ...args),
      });
      assertYomitanDictionaryMutationSafe(deps.userDataPath, dictionaries.length);
      return dictionaries;
    },
    importYomitanDictionary: async (zipPath) => {
      const yomitanProfilePolicy = deps.getYomitanProfilePolicy();
      if (yomitanProfilePolicy.isExternalReadOnlyMode()) {
        yomitanProfilePolicy.logSkippedWrite(
          formatSkippedYomitanWriteAction('importYomitanDictionary', zipPath),
        );
        return false;
      }
      await deps.ensureYomitanExtensionLoaded();
      return await importYomitanDictionaryFromZip(zipPath, deps.getYomitanParserRuntimeDeps(), {
        error: (message, ...args) => logger.error(message, ...args),
        info: (message, ...args) => logger.info(message, ...args),
        warn: (message, ...args) => logger.warn(message, ...args),
      });
    },
    dictionaryImportReplacesExisting: () =>
      deps.getYomitanParserRuntimeDeps().getYomitanExt()?.name === 'Hachidori',
    deleteYomitanDictionary: async (dictionaryTitle) => {
      const yomitanProfilePolicy = deps.getYomitanProfilePolicy();
      if (yomitanProfilePolicy.isExternalReadOnlyMode()) {
        yomitanProfilePolicy.logSkippedWrite(
          formatSkippedYomitanWriteAction('deleteYomitanDictionary', dictionaryTitle),
        );
        return false;
      }
      await deps.ensureYomitanExtensionLoaded();
      return await deleteYomitanDictionaryByTitle(
        dictionaryTitle,
        deps.getYomitanParserRuntimeDeps(),
        {
          error: (message, ...args) => logger.error(message, ...args),
          info: (message, ...args) => logger.info(message, ...args),
        },
      );
    },
    upsertYomitanDictionarySettings: async (dictionaryTitle, profileScope) => {
      const yomitanProfilePolicy = deps.getYomitanProfilePolicy();
      if (yomitanProfilePolicy.isExternalReadOnlyMode()) {
        yomitanProfilePolicy.logSkippedWrite(
          formatSkippedYomitanWriteAction('upsertYomitanDictionarySettings', dictionaryTitle),
        );
        return false;
      }
      await deps.ensureYomitanExtensionLoaded();
      return await upsertYomitanDictionarySettings(
        dictionaryTitle,
        profileScope,
        deps.getYomitanParserRuntimeDeps(),
        {
          error: (message, ...args) => logger.error(message, ...args),
          info: (message, ...args) => logger.info(message, ...args),
        },
      );
    },
    now: () => Date.now(),
    schedule: (fn, delayMs) => setTimeout(fn, delayMs),
    clearSchedule: (timer) => clearTimeout(timer),
    logInfo: (message) => logger.info(message),
    logWarn: (message) => logger.warn(message),
    onSyncStatus: (event) => {
      const notifications = deps.getOverlayNotifications();
      notifyCharacterDictionaryAutoSyncStatus(event, {
        getNotificationType: () => notifications.getConfiguredStatusNotificationType(),
        showOsd: (message) => deps.showMpvOsd(message),
        showOverlayNotification: (payload) => notifications.showOverlayNotification(payload),
        showDesktopNotification: (title, options) => deps.showDesktopNotification(title, options),
        startupOsdSequencer: deps.startupOsdSequencer,
      });
    },
    onSyncComplete: ({ mediaId, mediaTitle, changed }) => {
      handleCharacterDictionaryAutoSyncComplete(
        {
          mediaId,
          mediaTitle,
          changed,
        },
        {
          hasParserWindow: () => Boolean(appState.yomitanParserWindow),
          invalidateCharacterDictionaryLookups: () => {
            imageLookup.invalidate();
            nameCandidates.invalidate();
          },
          clearParserCaches: () => {
            if (appState.yomitanParserWindow) {
              clearYomitanParserCachesForWindow(appState.yomitanParserWindow);
            }
          },
          invalidateTokenizationCache: () => {
            deps.invalidateTokenizationCache();
          },
          refreshSubtitlePrefetch: () => {
            deps.refreshSubtitlePrefetch();
          },
          refreshCurrentSubtitle: () => {
            deps.refreshCurrentSubtitle(appState.currentSubText);
          },
          logInfo: (message) => logger.info(message),
        },
      );
    },
  });

  const imageLookup = createCharacterDictionaryImageLookup({
    userDataPath: deps.userDataPath,
    getCurrentMediaId: () => autoSync.getCurrentMediaId(),
    onIndexReady: () => deps.refreshCurrentSubtitleAnnotations(),
    onIndexReadyError: (error) =>
      logger.warn(
        'Failed to refresh subtitle annotations after character portrait index became ready.',
        error,
      ),
  });

  // Lets the Yomitan scan runtime skip name lookups at positions where no
  // character name can start; absent candidates just mean the exhaustive scan.
  const nameCandidates = createCharacterNameCandidateLookup({
    userDataPath: deps.userDataPath,
    getCurrentMediaId: () => autoSync.getCurrentMediaId(),
  });

  const selectionDeps = {
    setManualSelection: (request: Parameters<typeof runtime.setManualSelection>[0]) =>
      runtime.setManualSelection(request),
    resetAnilistMediaGuessState: () => deps.resetAnilistMediaGuessState(),
    runSyncNow: () => autoSync.runSyncNow(),
    warn: (message: string, error?: unknown) => logger.warn(message, error),
  };

  /** Manager edits that change the merged dictionary rebuild it; a failed rebuild is logged. */
  async function rebuildAfterManagerChange<T extends { ok: boolean; rebuildRequired?: boolean }>(
    result: T,
    change: 'override' | 'removal' | 'reorder',
  ): Promise<T> {
    if (result.ok && result.rebuildRequired) {
      try {
        await autoSync.runSyncNow();
      } catch (error) {
        logger.warn(`Failed to rebuild character dictionary after manager ${change}:`, error);
      }
    }
    return result;
  }

  const ipc: IpcSlice<
    | 'getCharacterDictionarySelection'
    | 'setCharacterDictionarySelection'
    | 'getCharacterDictionaryManagerSnapshot'
    | 'removeCharacterDictionaryManagedEntry'
    | 'moveCharacterDictionaryManagedEntry'
  > = {
    getCharacterDictionarySelection: (searchTitle?: string) =>
      runtime.getManualSelectionSnapshot(undefined, searchTitle),
    setCharacterDictionarySelection: async (
      mediaId: number,
      replaceManagedMediaId?: number,
      mediaTitle?: string,
    ) => {
      if (replaceManagedMediaId !== undefined && mediaTitle) {
        return rebuildAfterManagerChange(
          replaceCharacterDictionaryManagedEntry(deps.userDataPath, replaceManagedMediaId, {
            mediaId,
            mediaTitle,
          }),
          'override',
        );
      }
      return await applyCharacterDictionarySelection({ mediaId }, selectionDeps);
    },
    getCharacterDictionaryManagerSnapshot: async () =>
      getCharacterDictionaryManagerSnapshot(deps.userDataPath, autoSync.getCurrentMediaId()),
    removeCharacterDictionaryManagedEntry: async (mediaId: number) =>
      rebuildAfterManagerChange(
        removeCharacterDictionaryManagedEntry(
          deps.userDataPath,
          mediaId,
          autoSync.getCurrentMediaId(),
        ),
        'removal',
      ),
    moveCharacterDictionaryManagedEntry: async (mediaId: number, direction: 1 | -1) =>
      rebuildAfterManagerChange(
        moveCharacterDictionaryManagedEntry(
          deps.userDataPath,
          mediaId,
          direction,
          autoSync.getCurrentMediaId(),
        ),
        'reorder',
      ),
  };

  const cli: CliSlice<
    | 'generateCharacterDictionary'
    | 'getCharacterDictionarySelection'
    | 'setCharacterDictionarySelection'
  > = {
    generateCharacterDictionary: async (targetPath?: string) => {
      const disabledReason = deps.getYomitanProfilePolicy().getCharacterDictionaryDisabledReason();
      if (disabledReason) {
        throw new Error(disabledReason);
      }
      return await runtime.generateForCurrentMedia(targetPath);
    },
    getCharacterDictionarySelection: async (targetPath?: string) =>
      runtime.getManualSelectionSnapshot(targetPath),
    setCharacterDictionarySelection: async (request) =>
      applyCharacterDictionarySelection(request, selectionDeps),
  };

  return { runtime, autoSync, imageLookup, nameCandidates, ipc, cli };
}
