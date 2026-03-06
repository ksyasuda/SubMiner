import * as fs from 'fs';
import * as path from 'path';
import type {
  AnilistCharacterDictionaryEvictionPolicy,
  AnilistCharacterDictionaryProfileScope,
} from '../../types';
import type {
  CharacterDictionaryBuildResult,
  CharacterDictionaryGenerateOptions,
} from '../character-dictionary-runtime';

type AutoSyncStateDictionaryEntry = {
  mediaId: number;
  dictionaryTitle: string;
  lastImportedRevision: string | null;
  lastUsedAt: number;
};

type AutoSyncState = {
  activeMediaIds: number[];
  dictionariesByMediaId: Record<string, AutoSyncStateDictionaryEntry>;
};

type AutoSyncDictionaryInfo = {
  title: string;
  revision?: string | number;
};

export interface CharacterDictionaryAutoSyncConfig {
  enabled: boolean;
  refreshTtlHours: number;
  maxLoaded: number;
  evictionPolicy: AnilistCharacterDictionaryEvictionPolicy;
  profileScope: AnilistCharacterDictionaryProfileScope;
}

export interface CharacterDictionaryAutoSyncRuntimeDeps {
  userDataPath: string;
  getConfig: () => CharacterDictionaryAutoSyncConfig;
  generateCharacterDictionary: (
    options?: CharacterDictionaryGenerateOptions,
  ) => Promise<CharacterDictionaryBuildResult>;
  getYomitanDictionaryInfo: () => Promise<AutoSyncDictionaryInfo[]>;
  importYomitanDictionary: (zipPath: string) => Promise<boolean>;
  deleteYomitanDictionary: (dictionaryTitle: string) => Promise<boolean>;
  upsertYomitanDictionarySettings: (
    dictionaryTitle: string,
    profileScope: AnilistCharacterDictionaryProfileScope,
  ) => Promise<boolean>;
  removeYomitanDictionarySettings: (
    dictionaryTitle: string,
    profileScope: AnilistCharacterDictionaryProfileScope,
    mode: 'delete' | 'disable',
  ) => Promise<boolean>;
  now: () => number;
  schedule?: (fn: () => void, delayMs: number) => ReturnType<typeof setTimeout>;
  clearSchedule?: (timer: ReturnType<typeof setTimeout>) => void;
  operationTimeoutMs?: number;
  logInfo?: (message: string) => void;
  logWarn?: (message: string) => void;
}

function ensureDir(dirPath: string): void {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
}

function readAutoSyncState(statePath: string): AutoSyncState {
  try {
    const raw = fs.readFileSync(statePath, 'utf8');
    const parsed = JSON.parse(raw) as Partial<AutoSyncState>;
    if (!parsed || typeof parsed !== 'object') {
      return { activeMediaIds: [], dictionariesByMediaId: {} };
    }
    const dictionariesByMediaId = parsed.dictionariesByMediaId ?? {};
    if (!dictionariesByMediaId || typeof dictionariesByMediaId !== 'object') {
      return { activeMediaIds: [], dictionariesByMediaId: {} };
    }

    const normalizedEntries: Record<string, AutoSyncStateDictionaryEntry> = {};
    for (const [key, value] of Object.entries(dictionariesByMediaId)) {
      if (!value || typeof value !== 'object') {
        continue;
      }
      const mediaId = Number.parseInt(key, 10);
      const dictionaryTitle =
        typeof (value as { dictionaryTitle?: unknown }).dictionaryTitle === 'string'
          ? (value as { dictionaryTitle: string }).dictionaryTitle.trim()
          : '';
      if (!Number.isFinite(mediaId) || mediaId <= 0 || !dictionaryTitle) {
        continue;
      }

      const lastImportedRevisionRaw = (value as { lastImportedRevision?: unknown })
        .lastImportedRevision;
      const lastUsedAtRaw = (value as { lastUsedAt?: unknown }).lastUsedAt;
      normalizedEntries[String(mediaId)] = {
        mediaId,
        dictionaryTitle,
        lastImportedRevision:
          typeof lastImportedRevisionRaw === 'string' && lastImportedRevisionRaw.length > 0
            ? lastImportedRevisionRaw
            : null,
        lastUsedAt:
          typeof lastUsedAtRaw === 'number' && Number.isFinite(lastUsedAtRaw) ? lastUsedAtRaw : 0,
      };
    }

    const activeMediaIdsRaw = Array.isArray(parsed.activeMediaIds) ? parsed.activeMediaIds : [];
    const activeMediaIds = activeMediaIdsRaw
      .filter((value): value is number => typeof value === 'number' && Number.isFinite(value))
      .map((value) => Math.max(1, Math.floor(value)))
      .filter((value, index, all) => all.indexOf(value) === index)
      .filter((value) => normalizedEntries[String(value)] !== undefined);

    return {
      activeMediaIds,
      dictionariesByMediaId: normalizedEntries,
    };
  } catch {
    return { activeMediaIds: [], dictionariesByMediaId: {} };
  }
}

function writeAutoSyncState(statePath: string, state: AutoSyncState): void {
  ensureDir(path.dirname(statePath));
  fs.writeFileSync(statePath, JSON.stringify(state, null, 2), 'utf8');
}

function buildDictionaryTitle(mediaId: number): string {
  return `SubMiner Character Dictionary (AniList ${mediaId})`;
}

export function createCharacterDictionaryAutoSyncRuntimeService(
  deps: CharacterDictionaryAutoSyncRuntimeDeps,
): {
  scheduleSync: () => void;
  runSyncNow: () => Promise<void>;
} {
  const dictionariesDir = path.join(deps.userDataPath, 'character-dictionaries');
  const statePath = path.join(dictionariesDir, 'auto-sync-state.json');
  const schedule = deps.schedule ?? ((fn, delayMs) => setTimeout(fn, delayMs));
  const clearSchedule = deps.clearSchedule ?? ((timer) => clearTimeout(timer));
  const debounceMs = 800;
  const operationTimeoutMs = Math.max(1, Math.floor(deps.operationTimeoutMs ?? 7_000));

  let debounceTimer: ReturnType<typeof setTimeout> | null = null;
  let syncInFlight = false;
  let runQueued = false;

  const withOperationTimeout = async <T>(label: string, promise: Promise<T>): Promise<T> => {
    let timer: ReturnType<typeof setTimeout> | null = null;
    try {
      return await Promise.race([
        promise,
        new Promise<never>((_resolve, reject) => {
          timer = setTimeout(() => {
            reject(new Error(`${label} timed out after ${operationTimeoutMs}ms`));
          }, operationTimeoutMs);
        }),
      ]);
    } finally {
      if (timer !== null) {
        clearTimeout(timer);
      }
    }
  };

  const runSyncOnce = async (): Promise<void> => {
    const config = deps.getConfig();
    if (!config.enabled) {
      return;
    }

    const refreshTtlMs = Math.max(1, Math.floor(config.refreshTtlHours)) * 60 * 60 * 1000;
    const generation = await deps.generateCharacterDictionary({ refreshTtlMs });
    const dictionaryTitle = generation.dictionaryTitle ?? buildDictionaryTitle(generation.mediaId);
    const revision =
      typeof generation.revision === 'string' && generation.revision.length > 0
        ? generation.revision
        : null;

    const state = readAutoSyncState(statePath);
    const dictionaryInfo = await withOperationTimeout(
      'getYomitanDictionaryInfo',
      deps.getYomitanDictionaryInfo(),
    );
    const existing = dictionaryInfo.find((entry) => entry.title === dictionaryTitle) ?? null;
    const existingRevision =
      existing && (typeof existing.revision === 'string' || typeof existing.revision === 'number')
        ? String(existing.revision)
        : null;
    const shouldImport =
      existing === null || (revision !== null && existingRevision !== revision);

    if (shouldImport) {
      if (existing !== null) {
        await withOperationTimeout(
          `deleteYomitanDictionary(${dictionaryTitle})`,
          deps.deleteYomitanDictionary(dictionaryTitle),
        );
      }
      deps.logInfo?.(
        `[dictionary:auto-sync] importing AniList ${generation.mediaId}: ${generation.zipPath}`,
      );
      const imported = await withOperationTimeout(
        `importYomitanDictionary(${path.basename(generation.zipPath)})`,
        deps.importYomitanDictionary(generation.zipPath),
      );
      if (!imported) {
        throw new Error(`Failed to import dictionary ZIP: ${generation.zipPath}`);
      }
    }

    await withOperationTimeout(
      `upsertYomitanDictionarySettings(${dictionaryTitle})`,
      deps.upsertYomitanDictionarySettings(dictionaryTitle, config.profileScope),
    );

    const mediaIdKey = String(generation.mediaId);
    state.dictionariesByMediaId[mediaIdKey] = {
      mediaId: generation.mediaId,
      dictionaryTitle,
      lastImportedRevision: revision,
      lastUsedAt: deps.now(),
    };
    state.activeMediaIds = [
      generation.mediaId,
      ...state.activeMediaIds.filter((value) => value !== generation.mediaId),
    ];

    const maxLoaded = Math.max(1, Math.floor(config.maxLoaded));
    while (state.activeMediaIds.length > maxLoaded) {
      const evictedMediaId = state.activeMediaIds.pop();
      if (evictedMediaId === undefined) {
        break;
      }
      const evicted = state.dictionariesByMediaId[String(evictedMediaId)];
      if (!evicted) {
        continue;
      }

      await withOperationTimeout(
        `removeYomitanDictionarySettings(${evicted.dictionaryTitle})`,
        deps.removeYomitanDictionarySettings(
          evicted.dictionaryTitle,
          config.profileScope,
          config.evictionPolicy,
        ),
      );
      if (config.evictionPolicy === 'delete') {
        await withOperationTimeout(
          `deleteYomitanDictionary(${evicted.dictionaryTitle})`,
          deps.deleteYomitanDictionary(evicted.dictionaryTitle),
        );
        delete state.dictionariesByMediaId[String(evictedMediaId)];
      }
    }

    writeAutoSyncState(statePath, state);
    deps.logInfo?.(
      `[dictionary:auto-sync] synced AniList ${generation.mediaId}: ${dictionaryTitle} (${generation.entryCount} entries)`,
    );
  };

  const enqueueSync = (): void => {
    runQueued = true;
    if (syncInFlight) {
      return;
    }

    syncInFlight = true;
    void (async () => {
      while (runQueued) {
        runQueued = false;
        try {
          await runSyncOnce();
        } catch (error) {
          deps.logWarn?.(
            `[dictionary:auto-sync] sync failed: ${(error as Error)?.message ?? String(error)}`,
          );
        }
      }
    })().finally(() => {
      syncInFlight = false;
    });
  };

  return {
    scheduleSync: () => {
      const config = deps.getConfig();
      if (!config.enabled) {
        return;
      }
      if (debounceTimer !== null) {
        clearSchedule(debounceTimer);
      }
      debounceTimer = schedule(() => {
        debounceTimer = null;
        enqueueSync();
      }, debounceMs);
    },
    runSyncNow: async () => {
      await runSyncOnce();
    },
  };
}
