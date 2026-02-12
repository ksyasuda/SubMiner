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

import { contextBridge, ipcRenderer, IpcRendererEvent } from "electron";
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
  KikuFieldGroupingRequestData,
  KikuFieldGroupingChoice,
  KikuMergePreviewRequest,
  KikuMergePreviewResponse,
  RuntimeOptionApplyResult,
  RuntimeOptionId,
  RuntimeOptionState,
  RuntimeOptionValue,
  MpvSubtitleRenderMetrics,
  OverlayContentMeasurement,
} from "./types";

const overlayLayerArg = process.argv.find((arg) =>
  arg.startsWith("--overlay-layer="),
);
const overlayLayerFromArg = overlayLayerArg?.slice("--overlay-layer=".length);
const overlayLayer =
  overlayLayerFromArg === "visible" || overlayLayerFromArg === "invisible"
    ? overlayLayerFromArg
    : null;

const electronAPI: ElectronAPI = {
  getOverlayLayer: () => overlayLayer,
  onSubtitle: (callback: (data: SubtitleData) => void) => {
    ipcRenderer.on(
      "subtitle:set",
      (_event: IpcRendererEvent, data: SubtitleData) => callback(data),
    );
  },

  onVisibility: (callback: (visible: boolean) => void) => {
    ipcRenderer.on(
      "mpv:subVisibility",
      (_event: IpcRendererEvent, visible: boolean) => callback(visible),
    );
  },

  onSubtitlePosition: (
    callback: (position: SubtitlePosition | null) => void,
  ) => {
    ipcRenderer.on(
      "subtitle-position:set",
      (_event: IpcRendererEvent, position: SubtitlePosition | null) => {
        callback(position);
      },
    );
  },

  getOverlayVisibility: (): Promise<boolean> =>
    ipcRenderer.invoke("get-overlay-visibility"),
  getCurrentSubtitle: (): Promise<SubtitleData> =>
    ipcRenderer.invoke("get-current-subtitle"),
  getCurrentSubtitleAss: (): Promise<string> =>
    ipcRenderer.invoke("get-current-subtitle-ass"),
  getMpvSubtitleRenderMetrics: () =>
    ipcRenderer.invoke("get-mpv-subtitle-render-metrics"),
  onMpvSubtitleRenderMetrics: (
    callback: (metrics: MpvSubtitleRenderMetrics) => void,
  ) => {
    ipcRenderer.on(
      "mpv-subtitle-render-metrics:set",
      (_event: IpcRendererEvent, metrics: MpvSubtitleRenderMetrics) => {
        callback(metrics);
      },
    );
  },
  onSubtitleAss: (callback: (assText: string) => void) => {
    ipcRenderer.on(
      "subtitle-ass:set",
      (_event: IpcRendererEvent, assText: string) => {
        callback(assText);
      },
    );
  },
  onOverlayDebugVisualization: (callback: (enabled: boolean) => void) => {
    ipcRenderer.on(
      "overlay-debug-visualization:set",
      (_event: IpcRendererEvent, enabled: boolean) => {
        callback(enabled);
      },
    );
  },

  setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
    ipcRenderer.send("set-ignore-mouse-events", ignore, options);
  },

  openYomitanSettings: () => {
    ipcRenderer.send("open-yomitan-settings");
  },

  getSubtitlePosition: (): Promise<SubtitlePosition | null> =>
    ipcRenderer.invoke("get-subtitle-position"),
  saveSubtitlePosition: (position: SubtitlePosition) => {
    ipcRenderer.send("save-subtitle-position", position);
  },

  getMecabStatus: (): Promise<MecabStatus> =>
    ipcRenderer.invoke("get-mecab-status"),
  setMecabEnabled: (enabled: boolean) => {
    ipcRenderer.send("set-mecab-enabled", enabled);
  },

  sendMpvCommand: (command: (string | number)[]) => {
    ipcRenderer.send("mpv-command", command);
  },

  getKeybindings: (): Promise<Keybinding[]> =>
    ipcRenderer.invoke("get-keybindings"),

  getJimakuMediaInfo: (): Promise<JimakuMediaInfo> =>
    ipcRenderer.invoke("jimaku:get-media-info"),
  jimakuSearchEntries: (
    query: JimakuSearchQuery,
  ): Promise<JimakuApiResponse<JimakuEntry[]>> =>
    ipcRenderer.invoke("jimaku:search-entries", query),
  jimakuListFiles: (
    query: JimakuFilesQuery,
  ): Promise<JimakuApiResponse<JimakuFileEntry[]>> =>
    ipcRenderer.invoke("jimaku:list-files", query),
  jimakuDownloadFile: (
    query: JimakuDownloadQuery,
  ): Promise<JimakuDownloadResult> =>
    ipcRenderer.invoke("jimaku:download-file", query),

  quitApp: () => {
    ipcRenderer.send("quit-app");
  },

  toggleDevTools: () => {
    ipcRenderer.send("toggle-dev-tools");
  },

  toggleOverlay: () => {
    ipcRenderer.send("toggle-overlay");
  },

  getAnkiConnectStatus: (): Promise<boolean> =>
    ipcRenderer.invoke("get-anki-connect-status"),
  setAnkiConnectEnabled: (enabled: boolean) => {
    ipcRenderer.send("set-anki-connect-enabled", enabled);
  },
  clearAnkiConnectHistory: () => {
    ipcRenderer.send("clear-anki-connect-history");
  },

  onSecondarySub: (callback: (text: string) => void) => {
    ipcRenderer.on(
      "secondary-subtitle:set",
      (_event: IpcRendererEvent, text: string) => callback(text),
    );
  },

  onSecondarySubMode: (callback: (mode: SecondarySubMode) => void) => {
    ipcRenderer.on(
      "secondary-subtitle:mode",
      (_event: IpcRendererEvent, mode: SecondarySubMode) => callback(mode),
    );
  },

  getSecondarySubMode: (): Promise<SecondarySubMode> =>
    ipcRenderer.invoke("get-secondary-sub-mode"),
  getCurrentSecondarySub: (): Promise<string> =>
    ipcRenderer.invoke("get-current-secondary-sub"),
  getSubtitleStyle: (): Promise<SubtitleStyleConfig | null> =>
    ipcRenderer.invoke("get-subtitle-style"),
  onSubsyncManualOpen: (callback: (payload: SubsyncManualPayload) => void) => {
    ipcRenderer.on(
      "subsync:open-manual",
      (_event: IpcRendererEvent, payload: SubsyncManualPayload) => {
        callback(payload);
      },
    );
  },
  runSubsyncManual: (
    request: SubsyncManualRunRequest,
  ): Promise<SubsyncResult> =>
    ipcRenderer.invoke("subsync:run-manual", request),

  onKikuFieldGroupingRequest: (
    callback: (data: KikuFieldGroupingRequestData) => void,
  ) => {
    ipcRenderer.on(
      "kiku:field-grouping-request",
      (_event: IpcRendererEvent, data: KikuFieldGroupingRequestData) =>
        callback(data),
    );
  },
  kikuBuildMergePreview: (
    request: KikuMergePreviewRequest,
  ): Promise<KikuMergePreviewResponse> =>
    ipcRenderer.invoke("kiku:build-merge-preview", request),

  kikuFieldGroupingRespond: (choice: KikuFieldGroupingChoice) => {
    ipcRenderer.send("kiku:field-grouping-respond", choice);
  },

  getRuntimeOptions: (): Promise<RuntimeOptionState[]> =>
    ipcRenderer.invoke("runtime-options:get"),
  setRuntimeOptionValue: (
    id: RuntimeOptionId,
    value: RuntimeOptionValue,
  ): Promise<RuntimeOptionApplyResult> =>
    ipcRenderer.invoke("runtime-options:set", id, value),
  cycleRuntimeOption: (
    id: RuntimeOptionId,
    direction: 1 | -1,
  ): Promise<RuntimeOptionApplyResult> =>
    ipcRenderer.invoke("runtime-options:cycle", id, direction),
  onRuntimeOptionsChanged: (
    callback: (options: RuntimeOptionState[]) => void,
  ) => {
    ipcRenderer.on(
      "runtime-options:changed",
      (_event: IpcRendererEvent, options: RuntimeOptionState[]) => {
        callback(options);
      },
    );
  },
  onOpenRuntimeOptions: (callback: () => void) => {
    ipcRenderer.on("runtime-options:open", () => {
      callback();
    });
  },
  onOpenJimaku: (callback: () => void) => {
    ipcRenderer.on("jimaku:open", () => {
      callback();
    });
  },
  notifyOverlayModalClosed: (modal: "runtime-options" | "subsync") => {
    ipcRenderer.send("overlay:modal-closed", modal);
  },
  reportOverlayContentBounds: (measurement: OverlayContentMeasurement) => {
    ipcRenderer.send("overlay-content-bounds:report", measurement);
  },
};

contextBridge.exposeInMainWorld("electronAPI", electronAPI);
