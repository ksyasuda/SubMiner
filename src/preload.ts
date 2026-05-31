/*
 * SubMiner - All-in-one sentence mining overlay
 * Copyright (C) 2024 sudacode
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */

import { contextBridge, ipcRenderer, IpcRendererEvent } from 'electron';
import { resolveOverlayLayerFromArgv } from './preload-args';
import type {
  SubtitleData,
  SubtitlePosition,
  MecabStatus,
  Keybinding,
  ElectronAPI,
  SecondarySubMode,
  SubtitleStyleConfig,
  JimakuMediaInfo,
  JimakuSearchQuery,
  JimakuFilesQuery,
  JimakuDownloadQuery,
  JimakuEntry,
  JimakuFileEntry,
  JimakuApiResponse,
  JimakuDownloadResult,
  SubsyncManualPayload,
  SubsyncManualRunRequest,
  SubsyncResult,
  ClipboardAppendResult,
  PlaylistBrowserMutationResult,
  PlaylistBrowserSnapshot,
  KikuFieldGroupingRequestData,
  KikuFieldGroupingChoice,
  KikuMergePreviewRequest,
  KikuMergePreviewResponse,
  RuntimeOptionApplyResult,
  RuntimeOptionId,
  RuntimeOptionState,
  RuntimeOptionValue,
  OverlayContentMeasurement,
  ShortcutsConfig,
  ConfigHotReloadPayload,
  ControllerConfigUpdate,
  ControllerPreferenceUpdate,
  ResolvedControllerConfig,
  SessionNumericSelectionStartPayload,
  SubtitleMiningContext,
  YoutubePickerOpenPayload,
  YoutubePickerResolveRequest,
  YoutubePickerResolveResult,
} from './types';
import { IPC_CHANNELS } from './shared/ipc/contracts';

const overlayLayer = resolveOverlayLayerFromArgv(process.argv);

type EmptyListener = () => void;
type PayloadedListener<T> = (payload: T) => void;

function createQueuedIpcListener(channel: string): (listener: EmptyListener) => void {
  let count = 0;
  const listeners: EmptyListener[] = [];

  const dispatch = (): void => {
    if (listeners.length === 0) {
      count += 1;
      return;
    }
    for (const listener of listeners) {
      listener();
    }
  };

  ipcRenderer.on(channel, () => {
    dispatch();
  });

  return (listener: EmptyListener): void => {
    listeners.push(listener);
    while (count > 0) {
      count -= 1;
      listener();
    }
  };
}

function createQueuedIpcListenerWithPayload<T>(
  channel: string,
  normalize: (payload: unknown) => T,
): (listener: PayloadedListener<T>) => void {
  const pending: T[] = [];
  const listeners: PayloadedListener<T>[] = [];

  const dispatch = (payload: T): void => {
    if (listeners.length === 0) {
      pending.push(payload);
      return;
    }
    for (const listener of listeners) {
      listener(payload);
    }
  };

  ipcRenderer.on(channel, (_event: IpcRendererEvent, payloadArg: unknown) => {
    dispatch(normalize(payloadArg));
  });

  return (listener: PayloadedListener<T>): void => {
    listeners.push(listener);
    while (pending.length > 0) {
      const payload = pending.shift();
      listener(payload as T);
    }
  };
}

function createLatestValueIpcListenerWithPayload<T>(
  channel: string,
  normalize: (payload: unknown) => T,
): (listener: PayloadedListener<T>) => void {
  let pending: T | undefined;
  const listeners: PayloadedListener<T>[] = [];

  const dispatch = (payload: T): void => {
    if (listeners.length === 0) {
      pending = payload;
      return;
    }
    for (const listener of listeners) {
      listener(payload);
    }
  };

  ipcRenderer.on(channel, (_event: IpcRendererEvent, payloadArg: unknown) => {
    dispatch(normalize(payloadArg));
  });

  return (listener: PayloadedListener<T>): void => {
    listeners.push(listener);
    if (pending !== undefined) {
      const payload = pending;
      pending = undefined;
      listener(payload);
    }
  };
}

const onOpenRuntimeOptionsEvent = createQueuedIpcListener(IPC_CHANNELS.event.runtimeOptionsOpen);
const onOpenSessionHelpEvent = createQueuedIpcListener(IPC_CHANNELS.event.sessionHelpOpen);
const onOpenCharacterDictionaryManagerEvent = createQueuedIpcListener(
  IPC_CHANNELS.event.characterDictionaryManagerOpen,
);
const onOpenControllerSelectEvent = createQueuedIpcListener(
  IPC_CHANNELS.event.controllerSelectOpen,
);
const onOpenControllerDebugEvent = createQueuedIpcListener(IPC_CHANNELS.event.controllerDebugOpen);
const onOpenJimakuEvent = createQueuedIpcListener(IPC_CHANNELS.event.jimakuOpen);
const onOpenYoutubeTrackPickerEvent = createQueuedIpcListenerWithPayload<YoutubePickerOpenPayload>(
  IPC_CHANNELS.event.youtubePickerOpen,
  (payload) => payload as YoutubePickerOpenPayload,
);
const onOpenPlaylistBrowserEvent = createQueuedIpcListener(IPC_CHANNELS.event.playlistBrowserOpen);
const onCancelYoutubeTrackPickerEvent = createQueuedIpcListener(
  IPC_CHANNELS.event.youtubePickerCancel,
);
const onSessionNumericSelectionStartEvent =
  createQueuedIpcListenerWithPayload<SessionNumericSelectionStartPayload>(
    IPC_CHANNELS.event.sessionNumericSelectionStart,
    (payload) => payload as SessionNumericSelectionStartPayload,
  );
const onKeyboardModeToggleRequestedEvent = createQueuedIpcListener(
  IPC_CHANNELS.event.keyboardModeToggleRequested,
);
const onLookupWindowToggleRequestedEvent = createQueuedIpcListener(
  IPC_CHANNELS.event.lookupWindowToggleRequested,
);
const onSubsyncManualOpenEvent = createQueuedIpcListenerWithPayload<SubsyncManualPayload>(
  IPC_CHANNELS.event.subsyncOpenManual,
  (payload) => payload as SubsyncManualPayload,
);
const onSubtitleSidebarToggleEvent = createQueuedIpcListener(
  IPC_CHANNELS.event.subtitleSidebarToggle,
);
const onPrimarySubtitleBarToggleEvent = createQueuedIpcListener(
  IPC_CHANNELS.event.primarySubtitleBarToggle,
);
const onKikuFieldGroupingRequestEvent =
  createQueuedIpcListenerWithPayload<KikuFieldGroupingRequestData>(
    IPC_CHANNELS.event.kikuFieldGroupingRequest,
    (payload) => payload as KikuFieldGroupingRequestData,
  );
const onSubtitleSetEvent = createLatestValueIpcListenerWithPayload<SubtitleData>(
  IPC_CHANNELS.event.subtitleSet,
  (payload) => payload as SubtitleData,
);
const onSubtitleVisibilityEvent = createLatestValueIpcListenerWithPayload<boolean>(
  IPC_CHANNELS.event.subtitleVisibility,
  (payload) => payload === true,
);
const onSubtitlePositionSetEvent = createLatestValueIpcListenerWithPayload<SubtitlePosition | null>(
  IPC_CHANNELS.event.subtitlePositionSet,
  (payload) => payload as SubtitlePosition | null,
);
const onSecondarySubtitleSetEvent = createLatestValueIpcListenerWithPayload<string>(
  IPC_CHANNELS.event.secondarySubtitleSet,
  (payload) => (typeof payload === 'string' ? payload : ''),
);
const onSecondarySubtitleModeEvent = createLatestValueIpcListenerWithPayload<SecondarySubMode>(
  IPC_CHANNELS.event.secondarySubtitleMode,
  (payload) => payload as SecondarySubMode,
);

const electronAPI: ElectronAPI = {
  getOverlayLayer: () => overlayLayer,
  onSubtitle: (callback: (data: SubtitleData) => void) => {
    onSubtitleSetEvent(callback);
  },

  onVisibility: (callback: (visible: boolean) => void) => {
    onSubtitleVisibilityEvent(callback);
  },

  onSubtitlePosition: (callback: (position: SubtitlePosition | null) => void) => {
    onSubtitlePositionSetEvent(callback);
  },

  getOverlayVisibility: (): Promise<boolean> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getVisibleOverlayVisibility),
  getCurrentSubtitle: (): Promise<SubtitleData> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getCurrentSubtitle),
  getCurrentSubtitleRaw: (): Promise<string> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getCurrentSubtitleRaw),
  getCurrentSubtitleAss: (): Promise<string> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getCurrentSubtitleAss),
  getSubtitleSidebarOpen: (): Promise<boolean> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getSubtitleSidebarOpen),
  getSubtitleSidebarSnapshot: () =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getSubtitleSidebarSnapshot),
  getPlaybackPaused: (): Promise<boolean | null> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getPlaybackPaused),
  onSubtitleAss: (callback: (assText: string) => void) => {
    ipcRenderer.on(
      IPC_CHANNELS.event.subtitleAssSet,
      (_event: IpcRendererEvent, assText: string) => {
        callback(assText);
      },
    );
  },

  setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
    ipcRenderer.send(IPC_CHANNELS.command.setIgnoreMouseEvents, ignore, options);
  },

  reportOverlayInteractive: (interactive: boolean) => {
    ipcRenderer.send(IPC_CHANNELS.command.reportOverlayInteractive, interactive);
  },

  openYomitanSettings: () => {
    ipcRenderer.send(IPC_CHANNELS.command.openYomitanSettings);
  },

  recordYomitanLookup: (context?: SubtitleMiningContext | null) => {
    ipcRenderer.send(IPC_CHANNELS.command.recordYomitanLookup, context ?? null);
  },

  getSubtitlePosition: (): Promise<SubtitlePosition | null> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getSubtitlePosition),
  saveSubtitlePosition: (position: SubtitlePosition) => {
    ipcRenderer.send(IPC_CHANNELS.command.saveSubtitlePosition, position);
  },

  getMecabStatus: (): Promise<MecabStatus> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getMecabStatus),
  setMecabEnabled: (enabled: boolean) => {
    ipcRenderer.send(IPC_CHANNELS.command.setMecabEnabled, enabled);
  },

  sendMpvCommand: (command: (string | number)[]) => {
    ipcRenderer.send(IPC_CHANNELS.command.mpvCommand, command);
  },

  getKeybindings: (): Promise<Keybinding[]> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getKeybindings),
  getSessionBindings: () => ipcRenderer.invoke(IPC_CHANNELS.request.getSessionBindings),
  getConfiguredShortcuts: (): Promise<Required<ShortcutsConfig>> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getConfigShortcuts),
  dispatchSessionAction: (actionId, payload) =>
    ipcRenderer.invoke(IPC_CHANNELS.command.dispatchSessionAction, { actionId, payload }),
  getStatsToggleKey: (): Promise<string> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getStatsToggleKey),
  getMarkWatchedKey: (): Promise<string> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getMarkWatchedKey),
  markActiveVideoWatched: (): Promise<boolean> =>
    ipcRenderer.invoke(IPC_CHANNELS.command.markActiveVideoWatched),
  getControllerConfig: (): Promise<ResolvedControllerConfig> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getControllerConfig),
  saveControllerConfig: (update: ControllerConfigUpdate): Promise<void> =>
    ipcRenderer.invoke(IPC_CHANNELS.command.saveControllerConfig, update),
  saveControllerPreference: (update: ControllerPreferenceUpdate): Promise<void> =>
    ipcRenderer.invoke(IPC_CHANNELS.command.saveControllerPreference, update),

  getJimakuMediaInfo: (): Promise<JimakuMediaInfo> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.jimakuGetMediaInfo),
  jimakuSearchEntries: (query: JimakuSearchQuery): Promise<JimakuApiResponse<JimakuEntry[]>> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.jimakuSearchEntries, query),
  jimakuListFiles: (query: JimakuFilesQuery): Promise<JimakuApiResponse<JimakuFileEntry[]>> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.jimakuListFiles, query),
  jimakuDownloadFile: (query: JimakuDownloadQuery): Promise<JimakuDownloadResult> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.jimakuDownloadFile, query),

  quitApp: () => {
    ipcRenderer.send(IPC_CHANNELS.command.quitApp);
  },

  toggleDevTools: () => {
    ipcRenderer.send(IPC_CHANNELS.command.toggleDevTools);
  },

  toggleOverlay: () => {
    ipcRenderer.send(IPC_CHANNELS.command.toggleOverlay);
  },

  toggleStatsOverlay: () => {
    ipcRenderer.send(IPC_CHANNELS.command.toggleStatsOverlay);
  },

  getAnkiConnectStatus: (): Promise<boolean> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getAnkiConnectStatus),
  setAnkiConnectEnabled: (enabled: boolean) => {
    ipcRenderer.send(IPC_CHANNELS.command.setAnkiConnectEnabled, enabled);
  },
  clearAnkiConnectHistory: () => {
    ipcRenderer.send(IPC_CHANNELS.command.clearAnkiConnectHistory);
  },

  onSecondarySub: (callback: (text: string) => void) => {
    onSecondarySubtitleSetEvent(callback);
  },

  onSecondarySubMode: (callback: (mode: SecondarySubMode) => void) => {
    onSecondarySubtitleModeEvent(callback);
  },

  getSecondarySubMode: (): Promise<SecondarySubMode> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getSecondarySubMode),
  getCurrentSecondarySub: (): Promise<string> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getCurrentSecondarySub),
  focusMainWindow: () => ipcRenderer.invoke(IPC_CHANNELS.request.focusMainWindow) as Promise<void>,
  getSubtitleStyle: (): Promise<SubtitleStyleConfig | null> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getSubtitleStyle),
  onSubsyncManualOpen: onSubsyncManualOpenEvent,
  runSubsyncManual: (request: SubsyncManualRunRequest): Promise<SubsyncResult> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.runSubsyncManual, request),

  onKikuFieldGroupingRequest: onKikuFieldGroupingRequestEvent,
  kikuBuildMergePreview: (request: KikuMergePreviewRequest): Promise<KikuMergePreviewResponse> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.kikuBuildMergePreview, request),

  kikuFieldGroupingRespond: (choice: KikuFieldGroupingChoice) => {
    ipcRenderer.send(IPC_CHANNELS.command.kikuFieldGroupingRespond, choice);
  },

  getRuntimeOptions: (): Promise<RuntimeOptionState[]> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getRuntimeOptions),
  setRuntimeOptionValue: (
    id: RuntimeOptionId,
    value: RuntimeOptionValue,
  ): Promise<RuntimeOptionApplyResult> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.setRuntimeOption, id, value),
  cycleRuntimeOption: (id: RuntimeOptionId, direction: 1 | -1): Promise<RuntimeOptionApplyResult> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.cycleRuntimeOption, id, direction),
  onRuntimeOptionsChanged: (callback: (options: RuntimeOptionState[]) => void) => {
    ipcRenderer.on(
      IPC_CHANNELS.event.runtimeOptionsChanged,
      (_event: IpcRendererEvent, options: RuntimeOptionState[]) => {
        callback(options);
      },
    );
  },
  onOpenRuntimeOptions: onOpenRuntimeOptionsEvent,
  onOpenSessionHelp: onOpenSessionHelpEvent,
  onOpenControllerSelect: onOpenControllerSelectEvent,
  onOpenControllerDebug: onOpenControllerDebugEvent,
  onOpenJimaku: onOpenJimakuEvent,
  onOpenYoutubeTrackPicker: onOpenYoutubeTrackPickerEvent,
  onOpenPlaylistBrowser: onOpenPlaylistBrowserEvent,
  onOpenCharacterDictionaryManager: onOpenCharacterDictionaryManagerEvent,
  onSubtitleSidebarToggle: onSubtitleSidebarToggleEvent,
  onPrimarySubtitleBarToggle: onPrimarySubtitleBarToggleEvent,
  onCancelYoutubeTrackPicker: onCancelYoutubeTrackPickerEvent,
  onSessionNumericSelectionStart: onSessionNumericSelectionStartEvent,
  onKeyboardModeToggleRequested: onKeyboardModeToggleRequestedEvent,
  onLookupWindowToggleRequested: onLookupWindowToggleRequestedEvent,
  appendClipboardVideoToQueue: (): Promise<ClipboardAppendResult> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.appendClipboardVideoToQueue),
  getPlaylistBrowserSnapshot: (): Promise<PlaylistBrowserSnapshot> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getPlaylistBrowserSnapshot),
  appendPlaylistBrowserFile: (pathValue: string): Promise<PlaylistBrowserMutationResult> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.appendPlaylistBrowserFile, pathValue),
  playPlaylistBrowserIndex: (index: number): Promise<PlaylistBrowserMutationResult> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.playPlaylistBrowserIndex, index),
  removePlaylistBrowserIndex: (index: number): Promise<PlaylistBrowserMutationResult> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.removePlaylistBrowserIndex, index),
  movePlaylistBrowserIndex: (
    index: number,
    direction: 1 | -1,
  ): Promise<PlaylistBrowserMutationResult> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.movePlaylistBrowserIndex, index, direction),
  youtubePickerResolve: (
    request: YoutubePickerResolveRequest,
  ): Promise<YoutubePickerResolveResult> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.youtubePickerResolve, request),
  getCharacterDictionarySelection: (searchTitle?: string) =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getCharacterDictionarySelection, searchTitle),
  setCharacterDictionarySelection: (
    mediaId: number,
    replaceManagedMediaId?: number,
    mediaTitle?: string,
  ) =>
    ipcRenderer.invoke(
      IPC_CHANNELS.request.setCharacterDictionarySelection,
      mediaId,
      replaceManagedMediaId,
      mediaTitle,
    ),
  getCharacterDictionaryManagerSnapshot: () =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getCharacterDictionaryManagerSnapshot),
  removeCharacterDictionaryManagedEntry: (mediaId: number) =>
    ipcRenderer.invoke(IPC_CHANNELS.request.removeCharacterDictionaryManagedEntry, mediaId),
  moveCharacterDictionaryManagedEntry: (mediaId: number, direction: 1 | -1) =>
    ipcRenderer.invoke(
      IPC_CHANNELS.request.moveCharacterDictionaryManagedEntry,
      mediaId,
      direction,
    ),
  notifyOverlayModalClosed: (modal) => {
    ipcRenderer.send(IPC_CHANNELS.command.overlayModalClosed, modal);
  },
  notifyOverlayModalOpened: (modal) => {
    ipcRenderer.send(IPC_CHANNELS.command.overlayModalOpened, modal);
  },
  reportOverlayContentBounds: (measurement: OverlayContentMeasurement) => {
    ipcRenderer.send(IPC_CHANNELS.command.reportOverlayContentBounds, measurement);
  },
  onConfigHotReload: (callback: (payload: ConfigHotReloadPayload) => void) => {
    ipcRenderer.on(
      IPC_CHANNELS.event.configHotReload,
      (_event: IpcRendererEvent, payload: ConfigHotReloadPayload) => {
        callback(payload);
      },
    );
  },
};

contextBridge.exposeInMainWorld('electronAPI', electronAPI);
