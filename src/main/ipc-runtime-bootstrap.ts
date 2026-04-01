import * as path from 'node:path';

import type { BrowserWindow } from 'electron';

import type { AnkiIntegration } from '../anki-integration';
import type {
  JimakuApiResponse,
  JimakuLanguagePreference,
  KikuFieldGroupingChoice,
  ResolvedConfig,
  SubsyncManualRunRequest,
  SubsyncResult,
} from '../types';
import { downloadToFile, isRemoteMediaPath, parseMediaInfo } from '../jimaku/utils';
import { applyRuntimeOptionResultRuntime } from '../core/services/runtime-options-ipc';
import {
  playNextSubtitleRuntime,
  replayCurrentSubtitleRuntime,
  sendMpvCommandRuntime,
} from '../core/services';
import type { ConfigService } from '../config';
import { applyControllerConfigUpdate } from './controller-config-update.js';
import type { AnilistRuntime } from './anilist-runtime';
import type { DictionarySupportRuntime } from './dictionary-support-runtime';
import { createIpcRuntimeFromMainState, type IpcRuntime } from './ipc-runtime';
import type { MiningRuntime } from './mining-runtime';
import type { MpvRuntime } from './mpv-runtime';
import type { OverlayModalRuntime } from './overlay-runtime';
import type { OverlayUiRuntime } from './overlay-ui-runtime';
import type { AppState } from './state';
import type { SubtitleRuntime } from './subtitle-runtime';
import type { PlaylistBrowserIpcRuntime } from './runtime/playlist-browser-ipc';
import type { YoutubeRuntime } from './youtube-runtime';
import { resolveSubtitleStyleForRenderer } from './runtime/domains/overlay';
import type { ShortcutsRuntime } from './shortcuts-runtime';

type OverlayManagerLike = {
  getMainWindow: () => BrowserWindow | null;
  getVisibleOverlayVisible: () => boolean;
};

type OverlayUiLike = Pick<
  OverlayUiRuntime<BrowserWindow>,
  | 'broadcastRuntimeOptionsChanged'
  | 'handleOverlayModalClosed'
  | 'openRuntimeOptionsPalette'
  | 'toggleVisibleOverlay'
>;

type OverlayContentMeasurementStoreLike = {
  report: (payload: unknown) => void;
};

type ConfigDerivedRuntimeLike = {
  jimakuFetchJson: <T>(
    endpoint: string,
    query?: Record<string, string | number | boolean | null | undefined>,
  ) => Promise<JimakuApiResponse<T>>;
  getJimakuMaxEntryResults: () => number;
  getJimakuLanguagePreference: () => JimakuLanguagePreference;
  resolveJimakuApiKey: () => Promise<string | null>;
};

type SubsyncRuntimeLike = {
  triggerFromConfig: () => Promise<void>;
  runManualFromIpc: (request: SubsyncManualRunRequest) => Promise<SubsyncResult>;
};

export interface IpcRuntimeBootstrapInput {
  appState: AppState;
  userDataPath: string;
  getResolvedConfig: () => ResolvedConfig;
  configService: Pick<ConfigService, 'getRawConfig' | 'patchRawConfig'>;
  overlay: {
    manager: OverlayManagerLike;
    getOverlayUi: () => OverlayUiLike | undefined;
    modalRuntime: Pick<OverlayModalRuntime, 'notifyOverlayModalOpened'>;
    contentMeasurementStore: OverlayContentMeasurementStoreLike;
  };
  subtitle: SubtitleRuntime;
  mpvRuntime: Pick<MpvRuntime, 'shiftSubtitleDelayToAdjacentCue' | 'showMpvOsd'>;
  shortcuts: Pick<ShortcutsRuntime, 'getConfiguredShortcuts'>;
  actions: {
    requestAppQuit: () => void;
    openYomitanSettings: () => boolean;
    openPlaylistBrowser: () => void | Promise<void>;
    showDesktopNotification: (title: string, options: { body?: string }) => void;
    setAnkiIntegration: (integration: AnkiIntegration | null) => void;
  };
  runtimes: {
    youtube: Pick<YoutubeRuntime, 'openYoutubeTrackPickerFromPlayback' | 'resolveActivePicker'>;
    anilist: Pick<
      AnilistRuntime,
      | 'getStatusSnapshot'
      | 'clearTokenState'
      | 'openAnilistSetupWindow'
      | 'getQueueStatusSnapshot'
      | 'processNextAnilistRetryUpdate'
    >;
    mining: Pick<MiningRuntime, 'appendClipboardVideoToQueue'>;
    dictionarySupport: Pick<
      DictionarySupportRuntime,
      | 'createFieldGroupingCallback'
      | 'getFieldGroupingResolver'
      | 'setFieldGroupingResolver'
      | 'resolveMediaPathForJimaku'
    >;
    playlistBrowser: Pick<PlaylistBrowserIpcRuntime, 'playlistBrowserMainDeps'>;
    configDerived: ConfigDerivedRuntimeLike;
    subsync: SubsyncRuntimeLike;
  };
}

export function createIpcRuntimeBootstrap(input: IpcRuntimeBootstrapInput): IpcRuntime {
  return createIpcRuntimeFromMainState({
    mpv: {
      mainDeps: {
        triggerSubsyncFromConfig: () => input.runtimes.subsync.triggerFromConfig(),
        openRuntimeOptionsPalette: () => input.overlay.getOverlayUi()?.openRuntimeOptionsPalette(),
        openYoutubeTrackPicker: () => input.runtimes.youtube.openYoutubeTrackPickerFromPlayback(),
        openPlaylistBrowser: () => input.actions.openPlaylistBrowser(),
        cycleRuntimeOption: (id, direction) => {
          if (!input.appState.runtimeOptionsManager) {
            return { ok: false, error: 'Runtime options manager unavailable' };
          }
          return applyRuntimeOptionResultRuntime(
            input.appState.runtimeOptionsManager.cycleOption(id, direction),
            (text) => input.mpvRuntime.showMpvOsd(text),
          );
        },
        showMpvOsd: (text) => input.mpvRuntime.showMpvOsd(text),
        replayCurrentSubtitle: () => replayCurrentSubtitleRuntime(input.appState.mpvClient),
        playNextSubtitle: () => playNextSubtitleRuntime(input.appState.mpvClient),
        shiftSubDelayToAdjacentSubtitle: (direction) =>
          input.mpvRuntime.shiftSubtitleDelayToAdjacentCue(direction),
        sendMpvCommand: (rawCommand) => sendMpvCommandRuntime(input.appState.mpvClient, rawCommand),
        getMpvClient: () => input.appState.mpvClient,
        isMpvConnected: () =>
          Boolean(input.appState.mpvClient && input.appState.mpvClient.connected),
        hasRuntimeOptionsManager: () => input.appState.runtimeOptionsManager !== null,
      },
      runSubsyncManualFromIpc: (request) => input.runtimes.subsync.runManualFromIpc(request),
    },
    runtimeOptions: {
      getRuntimeOptionsManager: () => input.appState.runtimeOptionsManager,
      showMpvOsd: (text) => input.mpvRuntime.showMpvOsd(text),
    },
    main: {
      window: {
        getMainWindow: () => input.overlay.manager.getMainWindow(),
        getVisibleOverlayVisibility: () => input.overlay.manager.getVisibleOverlayVisible(),
        focusMainWindow: () => {
          const mainWindow = input.overlay.manager.getMainWindow();
          if (!mainWindow || mainWindow.isDestroyed()) return;
          if (!mainWindow.isFocused()) {
            mainWindow.focus();
          }
        },
        onOverlayModalClosed: (modal) =>
          input.overlay.getOverlayUi()?.handleOverlayModalClosed(modal),
        onOverlayModalOpened: (modal) => {
          input.overlay.modalRuntime.notifyOverlayModalOpened(modal);
        },
        onYoutubePickerResolve: (request) => input.runtimes.youtube.resolveActivePicker(request),
        openYomitanSettings: () => input.actions.openYomitanSettings(),
        quitApp: () => input.actions.requestAppQuit(),
        toggleVisibleOverlay: () => input.overlay.getOverlayUi()?.toggleVisibleOverlay(),
      },
      subtitle: {
        tokenizeCurrentSubtitle: async () => await input.subtitle.tokenizeCurrentSubtitle(),
        getCurrentSubtitleRaw: () => input.appState.currentSubText,
        getCurrentSubtitleAss: () => input.appState.currentSubAssText,
        getSubtitleSidebarSnapshot: async () => await input.subtitle.getSubtitleSidebarSnapshot(),
        getPlaybackPaused: () => input.appState.playbackPaused,
        getSubtitlePosition: () => input.subtitle.loadSubtitlePosition(),
        getSubtitleStyle: () => resolveSubtitleStyleForRenderer(input.getResolvedConfig()),
        saveSubtitlePosition: (position) => input.subtitle.saveSubtitlePosition(position),
        getMecabTokenizer: () => input.appState.mecabTokenizer,
        getKeybindings: () => input.appState.keybindings,
        getConfiguredShortcuts: () => input.shortcuts.getConfiguredShortcuts(),
        getStatsToggleKey: () => input.getResolvedConfig().stats.toggleKey,
        getMarkWatchedKey: () => input.getResolvedConfig().stats.markWatchedKey,
        getSecondarySubMode: () => input.appState.secondarySubMode,
      },
      controller: {
        getControllerConfig: () => input.getResolvedConfig().controller,
        saveControllerConfig: (update) => {
          const currentRawConfig = input.configService.getRawConfig();
          input.configService.patchRawConfig({
            controller: applyControllerConfigUpdate(currentRawConfig.controller, update),
          });
        },
        saveControllerPreference: ({ preferredGamepadId, preferredGamepadLabel }) => {
          input.configService.patchRawConfig({
            controller: {
              preferredGamepadId,
              preferredGamepadLabel,
            },
          });
        },
      },
      runtime: {
        getMpvClient: () => input.appState.mpvClient,
        getAnkiConnectStatus: () => input.appState.ankiIntegration !== null,
        getRuntimeOptions: () => input.appState.runtimeOptionsManager?.listOptions() ?? [],
        reportOverlayContentBounds: (payload) => {
          input.overlay.contentMeasurementStore.report(payload);
        },
        getImmersionTracker: () => input.appState.immersionTracker,
      },
      playlistBrowser: input.runtimes.playlistBrowser.playlistBrowserMainDeps,
      anilist: {
        getStatus: () => input.runtimes.anilist.getStatusSnapshot(),
        clearToken: () => input.runtimes.anilist.clearTokenState(),
        openSetup: () => input.runtimes.anilist.openAnilistSetupWindow(),
        getQueueStatus: () => input.runtimes.anilist.getQueueStatusSnapshot(),
        retryQueueNow: () => input.runtimes.anilist.processNextAnilistRetryUpdate(),
      },
      mining: {
        appendClipboardVideoToQueue: () => input.runtimes.mining.appendClipboardVideoToQueue(),
      },
    },
    ankiJimaku: {
      patchAnkiConnectEnabled: (enabled) => {
        input.configService.patchRawConfig({ ankiConnect: { enabled } });
      },
      getResolvedConfig: () => input.getResolvedConfig(),
      getRuntimeOptionsManager: () => input.appState.runtimeOptionsManager,
      getSubtitleTimingTracker: () => input.appState.subtitleTimingTracker,
      getMpvClient: () => input.appState.mpvClient,
      getAnkiIntegration: () => input.appState.ankiIntegration,
      setAnkiIntegration: (integration) => input.actions.setAnkiIntegration(integration),
      getKnownWordCacheStatePath: () => path.join(input.userDataPath, 'known-words-cache.json'),
      showDesktopNotification: input.actions.showDesktopNotification,
      createFieldGroupingCallback: () =>
        input.runtimes.dictionarySupport.createFieldGroupingCallback(),
      broadcastRuntimeOptionsChanged: () =>
        input.overlay.getOverlayUi()?.broadcastRuntimeOptionsChanged(),
      getFieldGroupingResolver: () => input.runtimes.dictionarySupport.getFieldGroupingResolver(),
      setFieldGroupingResolver: (resolver: ((choice: KikuFieldGroupingChoice) => void) | null) =>
        input.runtimes.dictionarySupport.setFieldGroupingResolver(resolver),
      parseMediaInfo: (mediaPath: string | null) =>
        parseMediaInfo(input.runtimes.dictionarySupport.resolveMediaPathForJimaku(mediaPath)),
      getCurrentMediaPath: () => input.appState.currentMediaPath,
      jimakuFetchJson: <T>(
        endpoint: string,
        query?: Record<string, string | number | boolean | null | undefined>,
      ): Promise<JimakuApiResponse<T>> =>
        input.runtimes.configDerived.jimakuFetchJson<T>(endpoint, query),
      getJimakuMaxEntryResults: () => input.runtimes.configDerived.getJimakuMaxEntryResults(),
      getJimakuLanguagePreference: () => input.runtimes.configDerived.getJimakuLanguagePreference(),
      resolveJimakuApiKey: () => input.runtimes.configDerived.resolveJimakuApiKey(),
      isRemoteMediaPath: (mediaPath: string) => isRemoteMediaPath(mediaPath),
      downloadToFile: (url: string, destPath: string, headers: Record<string, string>) =>
        downloadToFile(url, destPath, headers),
    },
  });
}
