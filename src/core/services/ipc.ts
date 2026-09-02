import electron from 'electron';
import type { BrowserWindow as ElectronBrowserWindow, IpcMainEvent } from 'electron';
import type {
  ChangelogSnapshot,
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
import { applyOverlayClickThrough } from './overlay-click-through';

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
  onOverlayInteractiveHint?: (
    interactive: boolean,
    senderWindow: ElectronBrowserWindow | null,
  ) => void;
  handleOverlayNotificationAction?: (
    notificationId: string,
    actionId: string,
    noteId?: number,
  ) => void | Promise<void>;
  openYomitanSettings: () => void;
  quitApp: () => void;
  toggleDevTools: () => void;
  getVisibleOverlayVisibility: () => boolean;
  toggleVisibleOverlay: () => void;
  tokenizeCurrentSubtitle: () => Promise<unknown>;
  getCurrentSubtitleRaw: () => string;
  getCurrentSubtitleAss: () => string;
  getSubtitleSidebarSnapshot?: () => Promise<SubtitleSidebarSnapshot>;
  getSubtitleSidebarOpen?: () => boolean;
  getPlaybackPaused: () => boolean | null | Promise<boolean | null>;
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
  getOverlayNotificationPosition: () => string;
  getControllerConfig: () => ResolvedControllerConfig;
  saveControllerConfig: (update: ControllerConfigUpdate) => void | Promise<void>;
  saveControllerPreference: (update: ControllerPreferenceUpdate) => void | Promise<void>;
  getSecondarySubMode: () => unknown;
  getCurrentSecondarySub: () => string;
  focusMainWindow: () => void;
  activatePlaybackWindowForOverlayInteraction?: () => boolean | Promise<boolean>;
  runSubsyncManual: (request: SubsyncManualRunRequest) => Promise<SubsyncResult>;
  onYoutubePickerResolve: (
    request: YoutubePickerResolveRequest,
  ) => Promise<YoutubePickerResolveResult>;
  getAnkiConnectStatus: () => boolean;
  getRuntimeOptions: () => unknown;
  setRuntimeOption: (id: RuntimeOptionId, value: RuntimeOptionValue) => unknown;
  cycleRuntimeOption: (id: RuntimeOptionId, direction: 1 | -1) => unknown;
  reportOverlayContentBounds: (
    payload: unknown,
    senderWindow: ElectronBrowserWindow | null,
  ) => void;
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
  getChangelogSnapshot?: (options?: { refresh?: boolean }) => Promise<ChangelogSnapshot>;
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

function parseOverlayNotificationActionPayload(
  payload: unknown,
): { notificationId: string; actionId: string; noteId?: number } | null {
  if (!payload || typeof payload !== 'object') return null;
  const record = payload as Record<string, unknown>;
  const notificationId = record.notificationId;
  const actionId = record.actionId;
  const noteId = record.noteId;
  if (typeof notificationId !== 'string' || notificationId.trim().length === 0) return null;
  if (typeof actionId !== 'string' || actionId.trim().length === 0) return null;
  if (
    noteId !== undefined &&
    (typeof noteId !== 'number' || !Number.isInteger(noteId) || noteId <= 0)
  ) {
    return null;
  }
  return { notificationId, actionId, ...(typeof noteId === 'number' ? { noteId } : {}) };
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
  onOverlayInteractiveHint?: (
    interactive: boolean,
    senderWindow: ElectronBrowserWindow | null,
  ) => void;
  handleOverlayNotificationAction?: (
    notificationId: string,
    actionId: string,
    noteId?: number,
  ) => void | Promise<void>;
  openYomitanSettings: () => void;
  quitApp: () => void;
  toggleVisibleOverlay: () => void;
  tokenizeCurrentSubtitle: () => Promise<unknown>;
  getCurrentSubtitleRaw: () => string;
  getCurrentSubtitleAss: () => string;
  getSubtitleSidebarSnapshot?: () => Promise<SubtitleSidebarSnapshot>;
  getSubtitleSidebarOpen?: () => boolean;
  getPlaybackPaused: () => boolean | null | Promise<boolean | null>;
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
  getOverlayNotificationPosition: () => string;
  getControllerConfig: () => ResolvedControllerConfig;
  saveControllerConfig: (update: ControllerConfigUpdate) => void | Promise<void>;
  saveControllerPreference: (update: ControllerPreferenceUpdate) => void | Promise<void>;
  getSecondarySubMode: () => unknown;
  getMpvClient: () => MpvClientLike | null;
  focusMainWindow: () => void;
  activatePlaybackWindowForOverlayInteraction?: () => boolean | Promise<boolean>;
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
  getChangelogSnapshot?: (options?: { refresh?: boolean }) => Promise<ChangelogSnapshot>;
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
    onOverlayInteractiveHint: options.onOverlayInteractiveHint,
    handleOverlayNotificationAction: options.handleOverlayNotificationAction,
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
    getSubtitleSidebarOpen: options.getSubtitleSidebarOpen ?? (() => false),
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
    getOverlayNotificationPosition: options.getOverlayNotificationPosition,
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
    activatePlaybackWindowForOverlayInteraction:
      options.activatePlaybackWindowForOverlayInteraction ?? (() => false),
    runSubsyncManual: options.runSubsyncManual,
    onYoutubePickerResolve: options.onYoutubePickerResolve,
    getAnkiConnectStatus: options.getAnkiConnectStatus,
    getRuntimeOptions: options.getRuntimeOptions,
    setRuntimeOption: options.setRuntimeOption,
    cycleRuntimeOption: options.cycleRuntimeOption,
    reportOverlayContentBounds: (payload, senderWindow) => {
      const mainWindow = options.getMainWindow();
      if (!mainWindow || mainWindow.isDestroyed()) return;
      if (!senderWindow || senderWindow !== (mainWindow as unknown as ElectronBrowserWindow))
        return;
      options.reportOverlayContentBounds(payload);
    },
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
    getChangelogSnapshot: options.getChangelogSnapshot,
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
  ipc.on(
    IPC_CHANNELS.command.setIgnoreMouseEvents,
    (event: unknown, ignore: unknown, options: unknown = {}) => {
      if (typeof ignore !== 'boolean') return;
      const parsedOptions = parseOptionalForwardingOptions(options);
      const senderWindow =
        electron.BrowserWindow?.fromWebContents((event as IpcMainEvent).sender) ?? null;
      if (senderWindow && !senderWindow.isDestroyed()) {
        // Route forwarding requests through the platform-aware helper so Windows never
        // installs Electron's global mouse hook (see overlay-click-through.ts).
        if (ignore && parsedOptions?.forward) {
          applyOverlayClickThrough(senderWindow);
        } else {
          senderWindow.setIgnoreMouseEvents(ignore, parsedOptions);
        }
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
  ipc.on(IPC_CHANNELS.command.overlayNotificationAction, (_event: unknown, payload: unknown) => {
    const parsedPayload = parseOverlayNotificationActionPayload(payload);
    if (!parsedPayload) return;
    void Promise.resolve(
      deps.handleOverlayNotificationAction?.(
        parsedPayload.notificationId,
        parsedPayload.actionId,
        parsedPayload.noteId,
      ),
    ).catch((error) => {
      console.warn(
        'Failed to handle overlay notification action:',
        error instanceof Error ? error.message : String(error),
      );
    });
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
    try {
      deps.recordSubtitleMiningContext?.(parseSubtitleMiningContext(payload));
    } catch (error) {
      console.warn(
        'Failed to record subtitle mining context:',
        error instanceof Error ? error.message : String(error),
      );
    }
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

  ipc.handle(IPC_CHANNELS.request.getSubtitleSidebarOpen, () => {
    return deps.getSubtitleSidebarOpen?.() ?? false;
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

  ipc.handle(IPC_CHANNELS.request.getOverlayNotificationPosition, () => {
    return deps.getOverlayNotificationPosition();
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

  ipc.handle(IPC_CHANNELS.request.activatePlaybackWindowForOverlayInteraction, async () => {
    return (await deps.activatePlaybackWindowForOverlayInteraction?.()) ?? false;
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

  ipc.on(IPC_CHANNELS.command.reportOverlayContentBounds, (event: unknown, payload: unknown) => {
    const senderWindow =
      electron.BrowserWindow?.fromWebContents((event as IpcMainEvent).sender) ?? null;
    deps.reportOverlayContentBounds(payload, senderWindow);
  });

  ipc.on(IPC_CHANNELS.command.reportOverlayInteractive, (event: unknown, interactive: unknown) => {
    if (typeof interactive !== 'boolean') return;
    const senderWindow =
      electron.BrowserWindow?.fromWebContents((event as IpcMainEvent).sender) ?? null;
    deps.onOverlayInteractiveHint?.(interactive, senderWindow);
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

  ipc.handle(IPC_CHANNELS.request.getChangelogSnapshot, async (_event, payload: unknown) => {
    const refresh =
      typeof payload === 'object' && payload !== null && 'refresh' in payload
        ? (payload as { refresh?: unknown }).refresh === true
        : false;
    if (!deps.getChangelogSnapshot) {
      throw new Error('Changelog service is unavailable.');
    }
    return await deps.getChangelogSnapshot({ refresh });
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
}
