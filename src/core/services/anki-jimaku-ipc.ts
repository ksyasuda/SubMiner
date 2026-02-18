import { ipcMain, IpcMainEvent } from "electron";
import * as fs from "fs";
import * as path from "path";
import * as os from "os";
import { createLogger } from "../../logger";
import {
  JimakuApiResponse,
  JimakuDownloadQuery,
  JimakuDownloadResult,
  JimakuEntry,
  JimakuFileEntry,
  JimakuFilesQuery,
  JimakuMediaInfo,
  JimakuSearchQuery,
  KikuFieldGroupingChoice,
  KikuMergePreviewRequest,
  KikuMergePreviewResponse,
} from "../../types";

const logger = createLogger("main:anki-jimaku-ipc");

export interface AnkiJimakuIpcDeps {
  setAnkiConnectEnabled: (enabled: boolean) => void;
  clearAnkiHistory: () => void;
  refreshKnownWords: () => Promise<void> | void;
  respondFieldGrouping: (choice: KikuFieldGroupingChoice) => void;
  buildKikuMergePreview: (
    request: KikuMergePreviewRequest,
  ) => Promise<KikuMergePreviewResponse>;
  getJimakuMediaInfo: () => JimakuMediaInfo;
  searchJimakuEntries: (
    query: JimakuSearchQuery,
  ) => Promise<JimakuApiResponse<JimakuEntry[]>>;
  listJimakuFiles: (
    query: JimakuFilesQuery,
  ) => Promise<JimakuApiResponse<JimakuFileEntry[]>>;
  resolveJimakuApiKey: () => Promise<string | null>;
  getCurrentMediaPath: () => string | null;
  isRemoteMediaPath: (mediaPath: string) => boolean;
  downloadToFile: (
    url: string,
    destPath: string,
    headers: Record<string, string>,
  ) => Promise<JimakuDownloadResult>;
  onDownloadedSubtitle: (pathToSubtitle: string) => void;
}

export function registerAnkiJimakuIpcHandlers(deps: AnkiJimakuIpcDeps): void {
  ipcMain.on(
    "set-anki-connect-enabled",
    (_event: IpcMainEvent, enabled: boolean) => {
      deps.setAnkiConnectEnabled(enabled);
    },
  );

  ipcMain.on("clear-anki-connect-history", () => {
    deps.clearAnkiHistory();
  });

  ipcMain.on("anki:refresh-known-words", async () => {
    await deps.refreshKnownWords();
  });

  ipcMain.on(
    "kiku:field-grouping-respond",
    (_event: IpcMainEvent, choice: KikuFieldGroupingChoice) => {
      deps.respondFieldGrouping(choice);
    },
  );

  ipcMain.handle(
    "kiku:build-merge-preview",
    async (
      _event,
      request: KikuMergePreviewRequest,
    ): Promise<KikuMergePreviewResponse> => {
      return deps.buildKikuMergePreview(request);
    },
  );

  ipcMain.handle("jimaku:get-media-info", (): JimakuMediaInfo => {
    return deps.getJimakuMediaInfo();
  });

  ipcMain.handle(
    "jimaku:search-entries",
    async (
      _event,
      query: JimakuSearchQuery,
    ): Promise<JimakuApiResponse<JimakuEntry[]>> => {
      return deps.searchJimakuEntries(query);
    },
  );

  ipcMain.handle(
    "jimaku:list-files",
    async (
      _event,
      query: JimakuFilesQuery,
    ): Promise<JimakuApiResponse<JimakuFileEntry[]>> => {
      return deps.listJimakuFiles(query);
    },
  );

  ipcMain.handle(
    "jimaku:download-file",
    async (
      _event,
      query: JimakuDownloadQuery,
    ): Promise<JimakuDownloadResult> => {
      const apiKey = await deps.resolveJimakuApiKey();
      if (!apiKey) {
        return {
          ok: false,
          error: {
            error:
              "Jimaku API key not set. Configure jimaku.apiKey or jimaku.apiKeyCommand.",
            code: 401,
          },
        };
      }

      const currentMediaPath = deps.getCurrentMediaPath();
      if (!currentMediaPath) {
        return { ok: false, error: { error: "No media file loaded in MPV." } };
      }

      const mediaDir = deps.isRemoteMediaPath(currentMediaPath)
        ? fs.mkdtempSync(path.join(os.tmpdir(), "subminer-jimaku-"))
        : path.dirname(path.resolve(currentMediaPath));
      const safeName = path.basename(query.name);
      if (!safeName) {
        return { ok: false, error: { error: "Invalid subtitle filename." } };
      }

      const ext = path.extname(safeName);
      const baseName = ext ? safeName.slice(0, -ext.length) : safeName;
      let targetPath = path.join(mediaDir, safeName);
      if (fs.existsSync(targetPath)) {
        targetPath = path.join(
          mediaDir,
          `${baseName} (jimaku-${query.entryId})${ext}`,
        );
        let counter = 2;
        while (fs.existsSync(targetPath)) {
          targetPath = path.join(
            mediaDir,
            `${baseName} (jimaku-${query.entryId}-${counter})${ext}`,
          );
          counter += 1;
        }
      }

      logger.info(
        `[jimaku] download-file name="${query.name}" entryId=${query.entryId}`,
      );
      const result = await deps.downloadToFile(query.url, targetPath, {
        Authorization: apiKey,
        "User-Agent": "SubMiner",
      });

      if (result.ok) {
        logger.info(`[jimaku] download-file saved to ${result.path}`);
        deps.onDownloadedSubtitle(result.path);
      } else {
        logger.error(
          `[jimaku] download-file failed: ${result.error?.error ?? "unknown error"}`,
        );
      }

      return result;
    },
  );
}
