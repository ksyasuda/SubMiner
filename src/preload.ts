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
} from './types';
import { IPC_CHANNELS } from './shared/ipc/contracts';

const overlayLayerArg = process.argv.find((arg) => arg.startsWith('--overlay-layer='));
const overlayLayerFromArg = overlayLayerArg?.slice('--overlay-layer='.length);
const overlayLayer =
  overlayLayerFromArg === 'visible' || overlayLayerFromArg === 'modal' ? overlayLayerFromArg : null;

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

const onOpenRuntimeOptionsEvent = createQueuedIpcListener(IPC_CHANNELS.event.runtimeOptionsOpen);
const onOpenJimakuEvent = createQueuedIpcListener(IPC_CHANNELS.event.jimakuOpen);
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
const onKikuFieldGroupingRequestEvent =
  createQueuedIpcListenerWithPayload<KikuFieldGroupingRequestData>(
    IPC_CHANNELS.event.kikuFieldGroupingRequest,
    (payload) => payload as KikuFieldGroupingRequestData,
  );

const electronAPI: ElectronAPI = {
  getOverlayLayer: () => overlayLayer,
  onSubtitle: (callback: (data: SubtitleData) => void) => {
    ipcRenderer.on(IPC_CHANNELS.event.subtitleSet, (_event: IpcRendererEvent, data: SubtitleData) =>
      callback(data),
    );
  },

  onVisibility: (callback: (visible: boolean) => void) => {
    ipcRenderer.on(
      IPC_CHANNELS.event.subtitleVisibility,
      (_event: IpcRendererEvent, visible: boolean) => callback(visible),
    );
  },

  onSubtitlePosition: (callback: (position: SubtitlePosition | null) => void) => {
    ipcRenderer.on(
      IPC_CHANNELS.event.subtitlePositionSet,
      (_event: IpcRendererEvent, position: SubtitlePosition | null) => {
        callback(position);
      },
    );
  },

  getOverlayVisibility: (): Promise<boolean> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getVisibleOverlayVisibility),
  getCurrentSubtitle: (): Promise<SubtitleData> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getCurrentSubtitle),
  getCurrentSubtitleRaw: (): Promise<string> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getCurrentSubtitleRaw),
  getCurrentSubtitleAss: (): Promise<string> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getCurrentSubtitleAss),
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

  openYomitanSettings: () => {
    ipcRenderer.send(IPC_CHANNELS.command.openYomitanSettings);
  },

  recordYomitanLookup: () => {
    ipcRenderer.send(IPC_CHANNELS.command.recordYomitanLookup);
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
  getConfiguredShortcuts: (): Promise<Required<ShortcutsConfig>> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.getConfigShortcuts),
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
    ipcRenderer.on(
      IPC_CHANNELS.event.secondarySubtitleSet,
      (_event: IpcRendererEvent, text: string) => callback(text),
    );
  },

  onSecondarySubMode: (callback: (mode: SecondarySubMode) => void) => {
    ipcRenderer.on(
      IPC_CHANNELS.event.secondarySubtitleMode,
      (_event: IpcRendererEvent, mode: SecondarySubMode) => callback(mode),
    );
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
  onOpenJimaku: onOpenJimakuEvent,
  onKeyboardModeToggleRequested: onKeyboardModeToggleRequestedEvent,
  onLookupWindowToggleRequested: onLookupWindowToggleRequestedEvent,
  appendClipboardVideoToQueue: (): Promise<ClipboardAppendResult> =>
    ipcRenderer.invoke(IPC_CHANNELS.request.appendClipboardVideoToQueue),
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
