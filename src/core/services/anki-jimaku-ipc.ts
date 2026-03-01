import { ipcMain } from 'electron';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import { createLogger } from '../../logger';
import {
  JimakuApiResponse,
  JimakuDownloadResult,
  JimakuEntry,
  JimakuFileEntry,
  JimakuFilesQuery,
  JimakuMediaInfo,
  JimakuSearchQuery,
  KikuFieldGroupingChoice,
  KikuMergePreviewRequest,
  KikuMergePreviewResponse,
} from '../../types';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import {
  parseJimakuDownloadQuery,
  parseJimakuFilesQuery,
  parseJimakuSearchQuery,
  parseKikuFieldGroupingChoice,
  parseKikuMergePreviewRequest,
} from '../../shared/ipc/validators';
import { buildJimakuSubtitleFilenameFromMediaPath } from './jimaku-download-path';

const logger = createLogger('main:anki-jimaku-ipc');

export interface AnkiJimakuIpcDeps {
  setAnkiConnectEnabled: (enabled: boolean) => void;
  clearAnkiHistory: () => void;
  refreshKnownWords: () => Promise<void> | void;
  respondFieldGrouping: (choice: KikuFieldGroupingChoice) => void;
  buildKikuMergePreview: (request: KikuMergePreviewRequest) => Promise<KikuMergePreviewResponse>;
  getJimakuMediaInfo: () => JimakuMediaInfo;
  searchJimakuEntries: (query: JimakuSearchQuery) => Promise<JimakuApiResponse<JimakuEntry[]>>;
  listJimakuFiles: (query: JimakuFilesQuery) => Promise<JimakuApiResponse<JimakuFileEntry[]>>;
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

interface IpcMainRegistrar {
  on: (channel: string, listener: (event: unknown, ...args: unknown[]) => void) => void;
  handle: (channel: string, listener: (event: unknown, ...args: unknown[]) => unknown) => void;
}

export function registerAnkiJimakuIpcHandlers(
  deps: AnkiJimakuIpcDeps,
  ipc: IpcMainRegistrar = ipcMain,
): void {
  ipc.on(IPC_CHANNELS.command.setAnkiConnectEnabled, (_event: unknown, enabled: unknown) => {
    if (typeof enabled !== 'boolean') return;
    deps.setAnkiConnectEnabled(enabled);
  });

  ipc.on(IPC_CHANNELS.command.clearAnkiConnectHistory, () => {
    deps.clearAnkiHistory();
  });

  ipc.on(IPC_CHANNELS.command.refreshKnownWords, async () => {
    await deps.refreshKnownWords();
  });

  ipc.on(IPC_CHANNELS.command.kikuFieldGroupingRespond, (_event: unknown, choice: unknown) => {
    const parsedChoice = parseKikuFieldGroupingChoice(choice);
    if (!parsedChoice) return;
    deps.respondFieldGrouping(parsedChoice);
  });

  ipc.handle(
    IPC_CHANNELS.request.kikuBuildMergePreview,
    async (_event, request: unknown): Promise<KikuMergePreviewResponse> => {
      const parsedRequest = parseKikuMergePreviewRequest(request);
      if (!parsedRequest) {
        return { ok: false, error: 'Invalid merge preview request payload' };
      }
      return deps.buildKikuMergePreview(parsedRequest);
    },
  );

  ipc.handle(IPC_CHANNELS.request.jimakuGetMediaInfo, (): JimakuMediaInfo => {
    return deps.getJimakuMediaInfo();
  });

  ipc.handle(
    IPC_CHANNELS.request.jimakuSearchEntries,
    async (_event, query: unknown): Promise<JimakuApiResponse<JimakuEntry[]>> => {
      const parsedQuery = parseJimakuSearchQuery(query);
      if (!parsedQuery) {
        return { ok: false, error: { error: 'Invalid Jimaku search query payload', code: 400 } };
      }
      return deps.searchJimakuEntries(parsedQuery);
    },
  );

  ipc.handle(
    IPC_CHANNELS.request.jimakuListFiles,
    async (_event, query: unknown): Promise<JimakuApiResponse<JimakuFileEntry[]>> => {
      const parsedQuery = parseJimakuFilesQuery(query);
      if (!parsedQuery) {
        return { ok: false, error: { error: 'Invalid Jimaku files query payload', code: 400 } };
      }
      return deps.listJimakuFiles(parsedQuery);
    },
  );

  ipc.handle(
    IPC_CHANNELS.request.jimakuDownloadFile,
    async (_event, query: unknown): Promise<JimakuDownloadResult> => {
      const parsedQuery = parseJimakuDownloadQuery(query);
      if (!parsedQuery) {
        return {
          ok: false,
          error: {
            error: 'Invalid Jimaku download query payload',
            code: 400,
          },
        };
      }

      const apiKey = await deps.resolveJimakuApiKey();
      if (!apiKey) {
        return {
          ok: false,
          error: {
            error: 'Jimaku API key not set. Configure jimaku.apiKey or jimaku.apiKeyCommand.',
            code: 401,
          },
        };
      }

      const currentMediaPath = deps.getCurrentMediaPath();
      if (!currentMediaPath) {
        return { ok: false, error: { error: 'No media file loaded in MPV.' } };
      }

      const mediaDir = deps.isRemoteMediaPath(currentMediaPath)
        ? fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-jimaku-'))
        : path.dirname(path.resolve(currentMediaPath));
      const safeName = path.basename(parsedQuery.name);
      if (!safeName) {
        return { ok: false, error: { error: 'Invalid subtitle filename.' } };
      }
      const subtitleFilename = buildJimakuSubtitleFilenameFromMediaPath(currentMediaPath, safeName);

      const ext = path.extname(subtitleFilename);
      const baseName = ext ? subtitleFilename.slice(0, -ext.length) : subtitleFilename;
      let targetPath = path.join(mediaDir, subtitleFilename);
      if (fs.existsSync(targetPath)) {
        targetPath = path.join(mediaDir, `${baseName} (jimaku-${parsedQuery.entryId})${ext}`);
        let counter = 2;
        while (fs.existsSync(targetPath)) {
          targetPath = path.join(
            mediaDir,
            `${baseName} (jimaku-${parsedQuery.entryId}-${counter})${ext}`,
          );
          counter += 1;
        }
      }

      logger.info(
        `[jimaku] download-file name="${parsedQuery.name}" entryId=${parsedQuery.entryId}`,
      );
      const result = await deps.downloadToFile(parsedQuery.url, targetPath, {
        Authorization: apiKey,
        'User-Agent': 'SubMiner',
      });

      if (result.ok) {
        logger.info(`[jimaku] download-file saved to ${result.path}`);
        deps.onDownloadedSubtitle(result.path);
      } else {
        logger.error(`[jimaku] download-file failed: ${result.error?.error ?? 'unknown error'}`);
      }

      return result;
    },
  );
}
