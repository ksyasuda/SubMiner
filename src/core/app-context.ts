import {
  AnkiConnectConfig,
  JimakuApiResponse,
  JimakuDownloadQuery,
  JimakuDownloadResult,
  JimakuEntry,
  JimakuFileEntry,
  JimakuFilesQuery,
  JimakuMediaInfo,
  JimakuSearchQuery,
  RuntimeOptionState,
  SubsyncManualRunRequest,
  SubsyncMode,
  SubsyncResult,
} from "../types";

export interface RuntimeOptionsModuleContext {
  getAnkiConfig: () => AnkiConnectConfig;
  applyAnkiPatch: (patch: Partial<AnkiConnectConfig>) => void;
  onOptionsChanged: (options: RuntimeOptionState[]) => void;
}

export interface AppContext {
  runtimeOptions?: RuntimeOptionsModuleContext;
  jimaku?: {
    getMediaInfo: () => JimakuMediaInfo;
    searchEntries: (
      query: JimakuSearchQuery,
    ) => Promise<JimakuApiResponse<JimakuEntry[]>>;
    listFiles: (
      query: JimakuFilesQuery,
    ) => Promise<JimakuApiResponse<JimakuFileEntry[]>>;
    downloadFile: (
      query: JimakuDownloadQuery,
    ) => Promise<JimakuDownloadResult>;
  };
  subsync?: {
    getDefaultMode: () => SubsyncMode;
    openManualPicker: () => Promise<void>;
    runAuto: () => Promise<SubsyncResult>;
    runManual: (request: SubsyncManualRunRequest) => Promise<SubsyncResult>;
    showOsd: (message: string) => void;
    runWithSpinner: <T>(task: () => Promise<T>, label?: string) => Promise<T>;
  };
}
