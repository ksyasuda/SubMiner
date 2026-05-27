import electron from 'electron';
import type { BrowserWindow as ElectronBrowserWindow, IpcMainEvent } from 'electron';
import type {
  CompiledSessionBinding,
  ControllerConfigUpdate,
  PlaylistBrowserMutationResult,
  PlaylistBrowserSnapshot,
  ControllerPreferenceUpdate,
  ResolvedControllerConfig,
  RuntimeOptionId,
  RuntimeOptionValue,
  SubtitleMiningContext,
  SubtitleSidebarSnapshot,
  SubtitlePosition,
  SubsyncManualRunRequest,
  SubsyncResult,
  SessionActionDispatchRequest,
  YoutubePickerResolveRequest,
  YoutubePickerResolveResult,
} from '../../types';
import { IPC_CHANNELS, type OverlayHostedModal } from '../../shared/ipc/contracts';
import {
  parseMpvCommand,
  parseControllerConfigUpdate,
  parseControllerPreferenceUpdate,
  parseOptionalForwardingOptions,
  parseOverlayHostedModal,
  parseRuntimeOptionDirection,
  parseRuntimeOptionId,
  parseRuntimeOptionValue,
  parseSessionActionDispatchRequest,
  parseSubtitlePosition,
  parseSubsyncManualRunRequest,
  parseYoutubePickerResolveRequest,
} from '../../shared/ipc/validators';

const { ipcMain } = electron;

export interface IpcServiceDeps {
  onOverlayModalClosed: (
    modal: OverlayHostedModal,
    senderWindow: ElectronBrowserWindow | null,
  ) => void;
  onOverlayModalOpened?: (
    modal: OverlayHostedModal,
    senderWindow: ElectronBrowserWindow | null,
  ) => void;
  onOverlayMouseInteractionChanged?: (
    active: boolean,
    senderWindow: ElectronBrowserWindow | null,
  ) => void;
  openYomitanSettings: () => void;
  quitApp: () => void;
  toggleDevTools: () => void;
  getVisibleOverlayVisibility: () => boolean;
  toggleVisibleOverlay: () => void;
  tokenizeCurrentSubtitle: () => Promise<unknown>;
  getCurrentSubtitleRaw: () => string;
  getCurrentSubtitleAss: () => string;
  getSubtitleSidebarSnapshot?: () => Promise<SubtitleSidebarSnapshot>;
  getPlaybackPaused: () => boolean | null;
  getSubtitlePosition: () => unknown;
  getSubtitleStyle: () => unknown;
  saveSubtitlePosition: (position: SubtitlePosition) => void;
  getMecabStatus: () => {
    available: boolean;
    enabled: boolean;
    path: string | null;
  };
  setMecabEnabled: (enabled: boolean) => void;
  handleMpvCommand: (command: Array<string | number>) => void;
  getKeybindings: () => unknown;
  getSessionBindings?: () => CompiledSessionBinding[];
  getConfiguredShortcuts: () => unknown;
  dispatchSessionAction?: (request: SessionActionDispatchRequest) => void | Promise<void>;
  getStatsToggleKey: () => string;
  getMarkWatchedKey: () => string;
  getControllerConfig: () => ResolvedControllerConfig;
  saveControllerConfig: (update: ControllerConfigUpdate) => void | Promise<void>;
  saveControllerPreference: (update: ControllerPreferenceUpdate) => void | Promise<void>;
  getSecondarySubMode: () => unknown;
  getCurrentSecondarySub: () => string;
  focusMainWindow: () => void;
  runSubsyncManual: (request: SubsyncManualRunRequest) => Promise<SubsyncResult>;
  onYoutubePickerResolve: (
    request: YoutubePickerResolveRequest,
  ) => Promise<YoutubePickerResolveResult>;
  getAnkiConnectStatus: () => boolean;
  getRuntimeOptions: () => unknown;
  setRuntimeOption: (id: RuntimeOptionId, value: RuntimeOptionValue) => unknown;
  cycleRuntimeOption: (id: RuntimeOptionId, direction: 1 | -1) => unknown;
  reportOverlayContentBounds: (payload: unknown) => void;
  getAnilistStatus: () => unknown;
  clearAnilistToken: () => void;
  openAnilistSetup: () => void;
  getAnilistQueueStatus: () => unknown;
  retryAnilistQueueNow: () => Promise<{ ok: boolean; message: string }>;
  runAnilistPostWatchUpdateOnManualMark?: () => Promise<void>;
  recordSubtitleMiningContext?: (context: SubtitleMiningContext | null) => void;
  getCharacterDictionarySelection?: (searchTitle?: string) => Promise<unknown>;
  setCharacterDictionarySelection?: (
    mediaId: number,
    replaceManagedMediaId?: number,
    mediaTitle?: string,
  ) => Promise<unknown>;
  getCharacterDictionaryManagerSnapshot?: () => Promise<unknown>;
  removeCharacterDictionaryManagedEntry?: (mediaId: number) => Promise<unknown>;
  moveCharacterDictionaryManagedEntry?: (mediaId: number, direction: 1 | -1) => Promise<unknown>;
  appendClipboardVideoToQueue: () => { ok: boolean; message: string };
  getPlaylistBrowserSnapshot: () => Promise<PlaylistBrowserSnapshot>;
  appendPlaylistBrowserFile: (filePath: string) => Promise<PlaylistBrowserMutationResult>;
  playPlaylistBrowserIndex: (index: number) => Promise<PlaylistBrowserMutationResult>;
  removePlaylistBrowserIndex: (index: number) => Promise<PlaylistBrowserMutationResult>;
  movePlaylistBrowserIndex: (
    index: number,
    direction: 1 | -1,
  ) => Promise<PlaylistBrowserMutationResult>;
  immersionTracker?: {
    recordYomitanLookup: () => void;
    getSessionSummaries: (limit?: number) => Promise<unknown>;
    getDailyRollups: (limit?: number) => Promise<unknown>;
    getMonthlyRollups: (limit?: number) => Promise<unknown>;
    getQueryHints: () => Promise<{
      totalSessions: number;
      activeSessions: number;
      episodesToday: number;
      activeAnimeCount: number;
      totalActiveMin: number;
      totalCards: number;
      activeDays: number;
      totalEpisodesWatched: number;
      totalAnimeCompleted: number;
      totalTokensSeen: number;
      totalLookupCount: number;
      totalLookupHits: number;
      totalYomitanLookupCount: number;
      newWordsToday: number;
      newWordsThisWeek: number;
    }>;
    getSessionTimeline: (sessionId: number, limit?: number) => Promise<unknown>;
    getSessionEvents: (sessionId: number, limit?: number) => Promise<unknown>;
    getVocabularyStats: (limit?: number) => Promise<unknown>;
    getKanjiStats: (limit?: number) => Promise<unknown>;
    getMediaLibrary: () => Promise<unknown>;
    getMediaDetail: (videoId: number) => Promise<unknown>;
    getMediaSessions: (videoId: number, limit?: number) => Promise<unknown>;
    getMediaDailyRollups: (videoId: number, limit?: number) => Promise<unknown>;
    getCoverArt: (videoId: number) => Promise<unknown>;
    markActiveVideoWatched: () => Promise<boolean>;
  } | null;
}

interface WindowLike {
  isDestroyed: () => boolean;
  focus: () => void;
  setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => void;
  webContents: {
    toggleDevTools: () => void;
  };
}

interface MecabTokenizerLike {
  getStatus: () => {
    available: boolean;
    enabled: boolean;
    path: string | null;
  };
  setEnabled: (enabled: boolean) => void;
}

interface MpvClientLike {
  currentSecondarySubText?: string;
}

interface IpcMainRegistrar {
  on: (channel: string, listener: (event: unknown, ...args: unknown[]) => void) => void;
  handle: (channel: string, listener: (event: unknown, ...args: unknown[]) => unknown) => void;
}

function parseSubtitleMiningContext(payload: unknown): SubtitleMiningContext | null {
  if (!payload || typeof payload !== 'object') {
    return null;
  }

  const record = payload as Record<string, unknown>;
  const source = record.source;
  const text = record.text;
  const startTime = record.startTime;
  const endTime = record.endTime;
  const capturedAtMs = record.capturedAtMs;

  if (
    source !== 'subtitle-sidebar' ||
    typeof text !== 'string' ||
    text.trim().length === 0 ||
    typeof startTime !== 'number' ||
    typeof endTime !== 'number' ||
    !Number.isFinite(startTime) ||
    !Number.isFinite(endTime) ||
    endTime <= startTime
  ) {
    return null;
  }

  const parsed: SubtitleMiningContext = {
    source: 'subtitle-sidebar',
    text,
    startTime,
    endTime,
  };
  if (typeof capturedAtMs === 'number' && Number.isFinite(capturedAtMs)) {
    parsed.capturedAtMs = capturedAtMs;
  }
  return parsed;
}

export interface IpcDepsRuntimeOptions {
  getMainWindow: () => WindowLike | null;
  getVisibleOverlayVisibility: () => boolean;
  onOverlayModalClosed: (
    modal: OverlayHostedModal,
    senderWindow: ElectronBrowserWindow | null,
  ) => void;
  onOverlayModalOpened?: (
    modal: OverlayHostedModal,
    senderWindow: ElectronBrowserWindow | null,
  ) => void;
  onOverlayMouseInteractionChanged?: (
    active: boolean,
    senderWindow: ElectronBrowserWindow | null,
  ) => void;
  openYomitanSettings: () => void;
  quitApp: () => void;
  toggleVisibleOverlay: () => void;
  tokenizeCurrentSubtitle: () => Promise<unknown>;
  getCurrentSubtitleRaw: () => string;
  getCurrentSubtitleAss: () => string;
  getSubtitleSidebarSnapshot?: () => Promise<SubtitleSidebarSnapshot>;
  getPlaybackPaused: () => boolean | null;
  getSubtitlePosition: () => unknown;
  getSubtitleStyle: () => unknown;
  saveSubtitlePosition: (position: SubtitlePosition) => void;
  getMecabTokenizer: () => MecabTokenizerLike | null;
  handleMpvCommand: (command: Array<string | number>) => void;
  getKeybindings: () => unknown;
  getSessionBindings?: () => CompiledSessionBinding[];
  getConfiguredShortcuts: () => unknown;
  dispatchSessionAction?: (request: SessionActionDispatchRequest) => void | Promise<void>;
  getStatsToggleKey: () => string;
  getMarkWatchedKey: () => string;
  getControllerConfig: () => ResolvedControllerConfig;
  saveControllerConfig: (update: ControllerConfigUpdate) => void | Promise<void>;
  saveControllerPreference: (update: ControllerPreferenceUpdate) => void | Promise<void>;
  getSecondarySubMode: () => unknown;
  getMpvClient: () => MpvClientLike | null;
  focusMainWindow: () => void;
  runSubsyncManual: (request: SubsyncManualRunRequest) => Promise<SubsyncResult>;
  onYoutubePickerResolve: (
    request: YoutubePickerResolveRequest,
  ) => Promise<YoutubePickerResolveResult>;
  getAnkiConnectStatus: () => boolean;
  getRuntimeOptions: () => unknown;
  setRuntimeOption: (id: RuntimeOptionId, value: RuntimeOptionValue) => unknown;
  cycleRuntimeOption: (id: RuntimeOptionId, direction: 1 | -1) => unknown;
  reportOverlayContentBounds: (payload: unknown) => void;
  getAnilistStatus: () => unknown;
  clearAnilistToken: () => void;
  openAnilistSetup: () => void;
  getAnilistQueueStatus: () => unknown;
  retryAnilistQueueNow: () => Promise<{ ok: boolean; message: string }>;
  runAnilistPostWatchUpdateOnManualMark?: () => Promise<void>;
  recordSubtitleMiningContext?: (context: SubtitleMiningContext | null) => void;
  getCharacterDictionarySelection?: (searchTitle?: string) => Promise<unknown>;
  setCharacterDictionarySelection?: (
    mediaId: number,
    replaceManagedMediaId?: number,
    mediaTitle?: string,
  ) => Promise<unknown>;
  getCharacterDictionaryManagerSnapshot?: () => Promise<unknown>;
  removeCharacterDictionaryManagedEntry?: (mediaId: number) => Promise<unknown>;
  moveCharacterDictionaryManagedEntry?: (mediaId: number, direction: 1 | -1) => Promise<unknown>;
  appendClipboardVideoToQueue: () => { ok: boolean; message: string };
  getPlaylistBrowserSnapshot: () => Promise<PlaylistBrowserSnapshot>;
  appendPlaylistBrowserFile: (filePath: string) => Promise<PlaylistBrowserMutationResult>;
  playPlaylistBrowserIndex: (index: number) => Promise<PlaylistBrowserMutationResult>;
  removePlaylistBrowserIndex: (index: number) => Promise<PlaylistBrowserMutationResult>;
  movePlaylistBrowserIndex: (
    index: number,
    direction: 1 | -1,
  ) => Promise<PlaylistBrowserMutationResult>;
  getImmersionTracker?: () => IpcServiceDeps['immersionTracker'];
}

export function createIpcDepsRuntime(options: IpcDepsRuntimeOptions): IpcServiceDeps {
  return {
    onOverlayModalClosed: options.onOverlayModalClosed,
    onOverlayModalOpened: options.onOverlayModalOpened,
    onOverlayMouseInteractionChanged: options.onOverlayMouseInteractionChanged,
    openYomitanSettings: options.openYomitanSettings,
    recordSubtitleMiningContext: options.recordSubtitleMiningContext,
    quitApp: options.quitApp,
    toggleDevTools: () => {
      const mainWindow = options.getMainWindow();
      if (!mainWindow || mainWindow.isDestroyed()) return;
      mainWindow.webContents.toggleDevTools();
    },
    getVisibleOverlayVisibility: options.getVisibleOverlayVisibility,
    toggleVisibleOverlay: options.toggleVisibleOverlay,
    tokenizeCurrentSubtitle: options.tokenizeCurrentSubtitle,
    getCurrentSubtitleRaw: options.getCurrentSubtitleRaw,
    getCurrentSubtitleAss: options.getCurrentSubtitleAss,
    getSubtitleSidebarSnapshot: options.getSubtitleSidebarSnapshot,
    getPlaybackPaused: options.getPlaybackPaused,
    getSubtitlePosition: options.getSubtitlePosition,
    getSubtitleStyle: options.getSubtitleStyle,
    saveSubtitlePosition: options.saveSubtitlePosition,
    getMecabStatus: () => {
      const mecabTokenizer = options.getMecabTokenizer();
      return mecabTokenizer
        ? mecabTokenizer.getStatus()
        : { available: false, enabled: false, path: null };
    },
    setMecabEnabled: (enabled) => {
      const mecabTokenizer = options.getMecabTokenizer();
      if (!mecabTokenizer) return;
      mecabTokenizer.setEnabled(enabled);
    },
    handleMpvCommand: options.handleMpvCommand,
    getKeybindings: options.getKeybindings,
    getSessionBindings: options.getSessionBindings ?? (() => []),
    getConfiguredShortcuts: options.getConfiguredShortcuts,
    dispatchSessionAction: options.dispatchSessionAction ?? (async () => {}),
    getStatsToggleKey: options.getStatsToggleKey,
    getMarkWatchedKey: options.getMarkWatchedKey,
    getControllerConfig: options.getControllerConfig,
    saveControllerConfig: options.saveControllerConfig,
    saveControllerPreference: options.saveControllerPreference,
    getSecondarySubMode: options.getSecondarySubMode,
    getCurrentSecondarySub: () => options.getMpvClient()?.currentSecondarySubText || '',
    focusMainWindow: () => {
      const mainWindow = options.getMainWindow();
      if (!mainWindow || mainWindow.isDestroyed()) return;
      mainWindow.focus();
    },
    runSubsyncManual: options.runSubsyncManual,
    onYoutubePickerResolve: options.onYoutubePickerResolve,
    getAnkiConnectStatus: options.getAnkiConnectStatus,
    getRuntimeOptions: options.getRuntimeOptions,
    setRuntimeOption: options.setRuntimeOption,
    cycleRuntimeOption: options.cycleRuntimeOption,
    reportOverlayContentBounds: options.reportOverlayContentBounds,
    getAnilistStatus: options.getAnilistStatus,
    clearAnilistToken: options.clearAnilistToken,
    openAnilistSetup: options.openAnilistSetup,
    getAnilistQueueStatus: options.getAnilistQueueStatus,
    retryAnilistQueueNow: options.retryAnilistQueueNow,
    runAnilistPostWatchUpdateOnManualMark: options.runAnilistPostWatchUpdateOnManualMark,
    getCharacterDictionarySelection:
      options.getCharacterDictionarySelection ??
      (async () => ({
        seriesKey: '',
        guessTitle: null,
        current: null,
        override: null,
        candidates: [],
      })),
    setCharacterDictionarySelection:
      options.setCharacterDictionarySelection ??
      (async () => ({
        ok: false,
        seriesKey: '',
        selected: { id: 0, title: '', episodes: null },
        staleMediaIds: [],
      })),
    getCharacterDictionaryManagerSnapshot:
      options.getCharacterDictionaryManagerSnapshot ?? (async () => ({ entries: [] })),
    removeCharacterDictionaryManagedEntry:
      options.removeCharacterDictionaryManagedEntry ??
      (async () => ({
        ok: false,
        message: 'Character dictionary manager unavailable.',
        entries: [],
      })),
    moveCharacterDictionaryManagedEntry:
      options.moveCharacterDictionaryManagedEntry ??
      (async () => ({
        ok: false,
        message: 'Character dictionary manager unavailable.',
        entries: [],
      })),
    appendClipboardVideoToQueue: options.appendClipboardVideoToQueue,
    getPlaylistBrowserSnapshot: options.getPlaylistBrowserSnapshot,
    appendPlaylistBrowserFile: options.appendPlaylistBrowserFile,
    playPlaylistBrowserIndex: options.playPlaylistBrowserIndex,
    removePlaylistBrowserIndex: options.removePlaylistBrowserIndex,
    movePlaylistBrowserIndex: options.movePlaylistBrowserIndex,
    get immersionTracker() {
      return options.getImmersionTracker?.() ?? null;
    },
  };
}

export function registerIpcHandlers(deps: IpcServiceDeps, ipc: IpcMainRegistrar = ipcMain): void {
  const parsePositiveIntLimit = (
    value: unknown,
    defaultValue: number,
    maxValue: number,
  ): number => {
    if (!Number.isInteger(value) || (value as number) < 1) {
      return defaultValue;
    }
    return Math.min(value as number, maxValue);
  };

  const parsePositiveInteger = (value: unknown): number | null => {
    if (typeof value !== 'number' || !Number.isInteger(value) || value <= 0) {
      return null;
    }
    return value;
  };

  ipc.on(
    IPC_CHANNELS.command.setIgnoreMouseEvents,
    (event: unknown, ignore: unknown, options: unknown = {}) => {
      if (typeof ignore !== 'boolean') return;
      const parsedOptions = parseOptionalForwardingOptions(options);
      const senderWindow =
        electron.BrowserWindow?.fromWebContents((event as IpcMainEvent).sender) ?? null;
      if (senderWindow && !senderWindow.isDestroyed()) {
        senderWindow.setIgnoreMouseEvents(ignore, parsedOptions);
      }
      deps.onOverlayMouseInteractionChanged?.(!ignore, senderWindow);
    },
  );

  ipc.on(IPC_CHANNELS.command.overlayModalClosed, (event: unknown, modal: unknown) => {
    const parsedModal = parseOverlayHostedModal(modal);
    if (!parsedModal) return;
    const senderWindow =
      electron.BrowserWindow?.fromWebContents((event as IpcMainEvent).sender) ?? null;
    deps.onOverlayModalClosed(parsedModal, senderWindow);
  });
  ipc.on(IPC_CHANNELS.command.overlayModalOpened, (event: unknown, modal: unknown) => {
    const parsedModal = parseOverlayHostedModal(modal);
    if (!parsedModal) return;
    if (!deps.onOverlayModalOpened) return;
    const senderWindow =
      electron.BrowserWindow?.fromWebContents((event as IpcMainEvent).sender) ?? null;
    deps.onOverlayModalOpened(parsedModal, senderWindow);
  });

  ipc.handle(
    IPC_CHANNELS.request.youtubePickerResolve,
    async (_event: unknown, request: unknown) => {
      const parsedRequest = parseYoutubePickerResolveRequest(request);
      if (!parsedRequest) {
        return { ok: false, message: 'Invalid YouTube picker resolve payload' };
      }
      return await deps.onYoutubePickerResolve(parsedRequest);
    },
  );

  ipc.on(IPC_CHANNELS.command.openYomitanSettings, () => {
    deps.openYomitanSettings();
  });

  ipc.on(IPC_CHANNELS.command.recordYomitanLookup, (_event: unknown, payload: unknown) => {
    deps.recordSubtitleMiningContext?.(parseSubtitleMiningContext(payload));
    deps.immersionTracker?.recordYomitanLookup();
  });

  ipc.handle(IPC_CHANNELS.command.markActiveVideoWatched, async () => {
    const marked = (await deps.immersionTracker?.markActiveVideoWatched()) ?? false;
    if (marked) {
      try {
        await deps.runAnilistPostWatchUpdateOnManualMark?.();
      } catch (error) {
        console.warn(
          'Failed to run AniList post-watch update after manual watched mark:',
          (error as Error).message,
        );
      }
    }
    return marked;
  });

  ipc.on(IPC_CHANNELS.command.quitApp, () => {
    deps.quitApp();
  });

  ipc.on(IPC_CHANNELS.command.toggleDevTools, () => {
    deps.toggleDevTools();
  });

  ipc.on(IPC_CHANNELS.command.toggleOverlay, () => {
    deps.toggleVisibleOverlay();
  });

  ipc.handle(IPC_CHANNELS.request.getVisibleOverlayVisibility, () => {
    return deps.getVisibleOverlayVisibility();
  });

  ipc.handle(IPC_CHANNELS.request.getCurrentSubtitle, async () => {
    return await deps.tokenizeCurrentSubtitle();
  });

  ipc.handle(IPC_CHANNELS.request.getCurrentSubtitleRaw, () => {
    return deps.getCurrentSubtitleRaw();
  });

  ipc.handle(IPC_CHANNELS.request.getCurrentSubtitleAss, () => {
    return deps.getCurrentSubtitleAss();
  });

  ipc.handle(IPC_CHANNELS.request.getSubtitleSidebarSnapshot, async () => {
    if (!deps.getSubtitleSidebarSnapshot) {
      throw new Error('Subtitle sidebar snapshot is unavailable.');
    }
    return await deps.getSubtitleSidebarSnapshot();
  });

  ipc.handle(IPC_CHANNELS.request.getPlaybackPaused, () => {
    return deps.getPlaybackPaused();
  });

  ipc.handle(IPC_CHANNELS.request.getSubtitlePosition, () => {
    return deps.getSubtitlePosition();
  });

  ipc.handle(IPC_CHANNELS.request.getSubtitleStyle, () => {
    return deps.getSubtitleStyle();
  });

  ipc.on(IPC_CHANNELS.command.saveSubtitlePosition, (_event: unknown, position: unknown) => {
    const parsedPosition = parseSubtitlePosition(position);
    if (!parsedPosition) return;
    deps.saveSubtitlePosition(parsedPosition);
  });

  ipc.handle(
    IPC_CHANNELS.command.saveControllerPreference,
    async (_event: unknown, update: unknown) => {
      const parsedUpdate = parseControllerPreferenceUpdate(update);
      if (!parsedUpdate) {
        throw new Error('Invalid controller preference payload');
      }
      await deps.saveControllerPreference(parsedUpdate);
    },
  );

  ipc.handle(
    IPC_CHANNELS.command.saveControllerConfig,
    async (_event: unknown, update: unknown) => {
      const parsedUpdate = parseControllerConfigUpdate(update);
      if (!parsedUpdate) {
        throw new Error('Invalid controller config payload');
      }
      await deps.saveControllerConfig(parsedUpdate);
    },
  );

  ipc.handle(IPC_CHANNELS.request.getMecabStatus, () => {
    return deps.getMecabStatus();
  });

  ipc.on(IPC_CHANNELS.command.setMecabEnabled, (_event: unknown, enabled: unknown) => {
    if (typeof enabled !== 'boolean') return;
    deps.setMecabEnabled(enabled);
  });

  ipc.on(IPC_CHANNELS.command.mpvCommand, (_event: unknown, command: unknown) => {
    const parsedCommand = parseMpvCommand(command);
    if (!parsedCommand) return;
    deps.handleMpvCommand(parsedCommand);
  });

  ipc.handle(
    IPC_CHANNELS.command.dispatchSessionAction,
    async (_event: unknown, request: unknown) => {
      const parsedRequest = parseSessionActionDispatchRequest(request);
      if (!parsedRequest) {
        throw new Error('Invalid session action payload');
      }
      await deps.dispatchSessionAction?.(parsedRequest);
    },
  );

  ipc.handle(IPC_CHANNELS.request.getKeybindings, () => {
    return deps.getKeybindings();
  });

  ipc.handle(IPC_CHANNELS.request.getSessionBindings, () => {
    return deps.getSessionBindings?.() ?? [];
  });

  ipc.handle(IPC_CHANNELS.request.getConfigShortcuts, () => {
    return deps.getConfiguredShortcuts();
  });

  ipc.handle(IPC_CHANNELS.request.getStatsToggleKey, () => {
    return deps.getStatsToggleKey();
  });

  ipc.handle(IPC_CHANNELS.request.getMarkWatchedKey, () => {
    return deps.getMarkWatchedKey();
  });

  ipc.handle(IPC_CHANNELS.request.getControllerConfig, () => {
    return deps.getControllerConfig();
  });

  ipc.handle(IPC_CHANNELS.request.getSecondarySubMode, () => {
    return deps.getSecondarySubMode();
  });

  ipc.handle(IPC_CHANNELS.request.getCurrentSecondarySub, () => {
    return deps.getCurrentSecondarySub();
  });

  ipc.handle(IPC_CHANNELS.request.focusMainWindow, () => {
    deps.focusMainWindow();
  });

  ipc.handle(IPC_CHANNELS.request.runSubsyncManual, async (_event, request: unknown) => {
    const parsedRequest = parseSubsyncManualRunRequest(request);
    if (!parsedRequest) {
      return { ok: false, message: 'Invalid subsync manual request payload' };
    }
    return await deps.runSubsyncManual(parsedRequest);
  });

  ipc.handle(IPC_CHANNELS.request.getAnkiConnectStatus, () => {
    return deps.getAnkiConnectStatus();
  });

  ipc.handle(IPC_CHANNELS.request.getRuntimeOptions, () => {
    return deps.getRuntimeOptions();
  });

  ipc.handle(IPC_CHANNELS.request.setRuntimeOption, (_event, id: unknown, value: unknown) => {
    const parsedId = parseRuntimeOptionId(id);
    if (!parsedId) {
      return { ok: false, error: 'Invalid runtime option id' };
    }
    const parsedValue = parseRuntimeOptionValue(value);
    if (parsedValue === null) {
      return { ok: false, error: 'Invalid runtime option value payload' };
    }
    return deps.setRuntimeOption(parsedId, parsedValue);
  });

  ipc.handle(IPC_CHANNELS.request.cycleRuntimeOption, (_event, id: unknown, direction: unknown) => {
    const parsedId = parseRuntimeOptionId(id);
    if (!parsedId) {
      return { ok: false, error: 'Invalid runtime option id' };
    }
    const parsedDirection = parseRuntimeOptionDirection(direction);
    if (!parsedDirection) {
      return { ok: false, error: 'Invalid runtime option cycle direction' };
    }
    return deps.cycleRuntimeOption(parsedId, parsedDirection);
  });

  ipc.on(IPC_CHANNELS.command.reportOverlayContentBounds, (_event: unknown, payload: unknown) => {
    deps.reportOverlayContentBounds(payload);
  });

  ipc.handle(IPC_CHANNELS.request.getAnilistStatus, () => {
    return deps.getAnilistStatus();
  });

  ipc.handle(IPC_CHANNELS.request.clearAnilistToken, () => {
    deps.clearAnilistToken();
    return { ok: true };
  });

  ipc.handle(IPC_CHANNELS.request.openAnilistSetup, () => {
    deps.openAnilistSetup();
    return { ok: true };
  });

  ipc.handle(IPC_CHANNELS.request.getAnilistQueueStatus, () => {
    return deps.getAnilistQueueStatus();
  });

  ipc.handle(IPC_CHANNELS.request.retryAnilistNow, async () => {
    return await deps.retryAnilistQueueNow();
  });

  ipc.handle(IPC_CHANNELS.request.getCharacterDictionarySelection, async (_event, searchTitle) => {
    const normalizedSearchTitle = typeof searchTitle === 'string' ? searchTitle.trim() : undefined;
    return await (deps.getCharacterDictionarySelection?.(normalizedSearchTitle) ??
      Promise.resolve({
        seriesKey: '',
        guessTitle: null,
        current: null,
        override: null,
        candidates: [],
      }));
  });

  ipc.handle(
    IPC_CHANNELS.request.setCharacterDictionarySelection,
    async (_event, mediaId: unknown, replaceManagedMediaId: unknown, mediaTitle: unknown) => {
      if (!Number.isSafeInteger(mediaId) || (mediaId as number) <= 0) {
        return { ok: false, message: 'Invalid AniList media ID.' };
      }
      const normalizedReplaceManagedMediaId =
        Number.isSafeInteger(replaceManagedMediaId) && (replaceManagedMediaId as number) > 0
          ? (replaceManagedMediaId as number)
          : undefined;
      const normalizedMediaTitle =
        typeof mediaTitle === 'string' && mediaTitle.trim() ? mediaTitle.trim() : undefined;
      return await (deps.setCharacterDictionarySelection?.(
        mediaId as number,
        normalizedReplaceManagedMediaId,
        normalizedMediaTitle,
      ) ??
        Promise.resolve({
          ok: false,
          message: 'Character dictionary selection unavailable.',
        }));
    },
  );

  ipc.handle(IPC_CHANNELS.request.getCharacterDictionaryManagerSnapshot, async () => {
    return await (deps.getCharacterDictionaryManagerSnapshot?.() ??
      Promise.resolve({ entries: [] }));
  });

  ipc.handle(
    IPC_CHANNELS.request.removeCharacterDictionaryManagedEntry,
    async (_event, mediaId: unknown) => {
      if (!Number.isSafeInteger(mediaId) || (mediaId as number) <= 0) {
        return { ok: false, message: 'Invalid AniList media ID.', entries: [] };
      }
      return await (deps.removeCharacterDictionaryManagedEntry?.(mediaId as number) ??
        Promise.resolve({
          ok: false,
          message: 'Character dictionary manager unavailable.',
          entries: [],
        }));
    },
  );

  ipc.handle(
    IPC_CHANNELS.request.moveCharacterDictionaryManagedEntry,
    async (_event, mediaId: unknown, direction: unknown) => {
      if (!Number.isSafeInteger(mediaId) || (mediaId as number) <= 0) {
        return { ok: false, message: 'Invalid AniList media ID.', entries: [] };
      }
      if (direction !== 1 && direction !== -1) {
        return { ok: false, message: 'Invalid move direction.', entries: [] };
      }
      return await (deps.moveCharacterDictionaryManagedEntry?.(mediaId as number, direction) ??
        Promise.resolve({
          ok: false,
          message: 'Character dictionary manager unavailable.',
          entries: [],
        }));
    },
  );

  ipc.handle(IPC_CHANNELS.request.appendClipboardVideoToQueue, () => {
    return deps.appendClipboardVideoToQueue();
  });

  ipc.handle(IPC_CHANNELS.request.getPlaylistBrowserSnapshot, async () => {
    return await deps.getPlaylistBrowserSnapshot();
  });

  ipc.handle(IPC_CHANNELS.request.appendPlaylistBrowserFile, async (_event, filePath: unknown) => {
    if (typeof filePath !== 'string' || filePath.trim().length === 0) {
      return { ok: false, message: 'Invalid playlist browser file path.' };
    }
    return await deps.appendPlaylistBrowserFile(filePath);
  });

  ipc.handle(IPC_CHANNELS.request.playPlaylistBrowserIndex, async (_event, index: unknown) => {
    if (!Number.isSafeInteger(index) || (index as number) < 0) {
      return { ok: false, message: 'Invalid playlist browser index.' };
    }
    return await deps.playPlaylistBrowserIndex(index as number);
  });

  ipc.handle(IPC_CHANNELS.request.removePlaylistBrowserIndex, async (_event, index: unknown) => {
    if (!Number.isSafeInteger(index) || (index as number) < 0) {
      return { ok: false, message: 'Invalid playlist browser index.' };
    }
    return await deps.removePlaylistBrowserIndex(index as number);
  });

  ipc.handle(
    IPC_CHANNELS.request.movePlaylistBrowserIndex,
    async (_event, index: unknown, direction: unknown) => {
      if (!Number.isSafeInteger(index) || (index as number) < 0) {
        return { ok: false, message: 'Invalid playlist browser index.' };
      }
      if (direction !== 1 && direction !== -1) {
        return { ok: false, message: 'Invalid playlist browser move direction.' };
      }
      return await deps.movePlaylistBrowserIndex(index as number, direction as 1 | -1);
    },
  );

  // Stats request handlers
  ipc.handle(IPC_CHANNELS.request.statsGetOverview, async () => {
    const tracker = deps.immersionTracker;
    if (!tracker) {
      return {
        sessions: [],
        rollups: [],
        hints: {
          totalSessions: 0,
          activeSessions: 0,
          episodesToday: 0,
          activeAnimeCount: 0,
          totalActiveMin: 0,
          totalCards: 0,
          activeDays: 0,
          totalEpisodesWatched: 0,
          totalAnimeCompleted: 0,
          totalTokensSeen: 0,
          totalLookupCount: 0,
          totalLookupHits: 0,
          totalYomitanLookupCount: 0,
          newWordsToday: 0,
          newWordsThisWeek: 0,
        },
      };
    }
    const [sessions, rollups, hints] = await Promise.all([
      tracker.getSessionSummaries(5),
      tracker.getDailyRollups(14),
      tracker.getQueryHints(),
    ]);
    return { sessions, rollups, hints };
  });

  ipc.handle(IPC_CHANNELS.request.statsGetDailyRollups, async (_event, limit: unknown) => {
    const parsedLimit = parsePositiveIntLimit(limit, 60, 500);
    return deps.immersionTracker?.getDailyRollups(parsedLimit) ?? [];
  });

  ipc.handle(IPC_CHANNELS.request.statsGetMonthlyRollups, async (_event, limit: unknown) => {
    const parsedLimit = parsePositiveIntLimit(limit, 24, 120);
    return deps.immersionTracker?.getMonthlyRollups(parsedLimit) ?? [];
  });

  ipc.handle(IPC_CHANNELS.request.statsGetSessions, async (_event, limit: unknown) => {
    const parsedLimit = parsePositiveIntLimit(limit, 50, 500);
    return deps.immersionTracker?.getSessionSummaries(parsedLimit) ?? [];
  });

  ipc.handle(
    IPC_CHANNELS.request.statsGetSessionTimeline,
    async (_event, sessionId: unknown, limit: unknown) => {
      const parsedSessionId = parsePositiveInteger(sessionId);
      if (parsedSessionId === null) return [];
      const parsedLimit = limit === undefined ? undefined : parsePositiveIntLimit(limit, 200, 1000);
      return deps.immersionTracker?.getSessionTimeline(parsedSessionId, parsedLimit) ?? [];
    },
  );

  ipc.handle(
    IPC_CHANNELS.request.statsGetSessionEvents,
    async (_event, sessionId: unknown, limit: unknown) => {
      const parsedSessionId = parsePositiveInteger(sessionId);
      if (parsedSessionId === null) return [];
      const parsedLimit = parsePositiveIntLimit(limit, 500, 1000);
      return deps.immersionTracker?.getSessionEvents(parsedSessionId, parsedLimit) ?? [];
    },
  );

  ipc.handle(IPC_CHANNELS.request.statsGetVocabulary, async (_event, limit: unknown) => {
    const parsedLimit = parsePositiveIntLimit(limit, 100, 500);
    return deps.immersionTracker?.getVocabularyStats(parsedLimit) ?? [];
  });

  ipc.handle(IPC_CHANNELS.request.statsGetKanji, async (_event, limit: unknown) => {
    const parsedLimit = parsePositiveIntLimit(limit, 100, 500);
    return deps.immersionTracker?.getKanjiStats(parsedLimit) ?? [];
  });

  ipc.handle(IPC_CHANNELS.request.statsGetMediaLibrary, async () => {
    return deps.immersionTracker?.getMediaLibrary() ?? [];
  });

  ipc.handle(IPC_CHANNELS.request.statsGetMediaDetail, async (_event, videoId: unknown) => {
    if (typeof videoId !== 'number') return null;
    return deps.immersionTracker?.getMediaDetail(videoId) ?? null;
  });

  ipc.handle(
    IPC_CHANNELS.request.statsGetMediaSessions,
    async (_event, videoId: unknown, limit: unknown) => {
      if (typeof videoId !== 'number') return [];
      const parsedLimit = parsePositiveIntLimit(limit, 100, 500);
      return deps.immersionTracker?.getMediaSessions(videoId, parsedLimit) ?? [];
    },
  );

  ipc.handle(
    IPC_CHANNELS.request.statsGetMediaDailyRollups,
    async (_event, videoId: unknown, limit: unknown) => {
      if (typeof videoId !== 'number') return [];
      const parsedLimit = parsePositiveIntLimit(limit, 90, 500);
      return deps.immersionTracker?.getMediaDailyRollups(videoId, parsedLimit) ?? [];
    },
  );

  ipc.handle(IPC_CHANNELS.request.statsGetMediaCover, async (_event, videoId: unknown) => {
    if (typeof videoId !== 'number') return null;
    return deps.immersionTracker?.getCoverArt(videoId) ?? null;
  });
}
