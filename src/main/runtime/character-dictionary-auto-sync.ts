import * as fs from 'fs';
import * as path from 'path';
import type { AnilistCharacterDictionaryProfileScope } from '../../types';
import type {
  CharacterDictionarySnapshotResult,
  MergedCharacterDictionaryBuildResult,
} from '../character-dictionary-runtime';

type AutoSyncState = {
  activeMediaIds: number[];
  mergedRevision: string | null;
  mergedDictionaryTitle: string | null;
};

type AutoSyncDictionaryInfo = {
  title: string;
  revision?: string | number;
};

export interface CharacterDictionaryAutoSyncConfig {
  enabled: boolean;
  maxLoaded: number;
  profileScope: AnilistCharacterDictionaryProfileScope;
}

export interface CharacterDictionaryAutoSyncRuntimeDeps {
  userDataPath: string;
  getConfig: () => CharacterDictionaryAutoSyncConfig;
  getOrCreateCurrentSnapshot: (targetPath?: string) => Promise<CharacterDictionarySnapshotResult>;
  buildMergedDictionary: (mediaIds: number[]) => Promise<MergedCharacterDictionaryBuildResult>;
  getYomitanDictionaryInfo: () => Promise<AutoSyncDictionaryInfo[]>;
  importYomitanDictionary: (zipPath: string) => Promise<boolean>;
  deleteYomitanDictionary: (dictionaryTitle: string) => Promise<boolean>;
  upsertYomitanDictionarySettings: (
    dictionaryTitle: string,
    profileScope: AnilistCharacterDictionaryProfileScope,
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
    const activeMediaIds = Array.isArray(parsed.activeMediaIds)
      ? parsed.activeMediaIds
          .filter((value): value is number => typeof value === 'number' && Number.isFinite(value))
          .map((value) => Math.max(1, Math.floor(value)))
          .filter((value, index, all) => all.indexOf(value) === index)
      : [];
    return {
      activeMediaIds,
      mergedRevision:
        typeof parsed.mergedRevision === 'string' && parsed.mergedRevision.length > 0
          ? parsed.mergedRevision
          : null,
      mergedDictionaryTitle:
        typeof parsed.mergedDictionaryTitle === 'string' && parsed.mergedDictionaryTitle.length > 0
          ? parsed.mergedDictionaryTitle
          : null,
    };
  } catch {
    return {
      activeMediaIds: [],
      mergedRevision: null,
      mergedDictionaryTitle: null,
    };
  }
}

function writeAutoSyncState(statePath: string, state: AutoSyncState): void {
  ensureDir(path.dirname(statePath));
  fs.writeFileSync(statePath, JSON.stringify(state, null, 2), 'utf8');
}

function arraysEqual(left: number[], right: number[]): boolean {
  if (left.length !== right.length) return false;
  for (let i = 0; i < left.length; i += 1) {
    if (left[i] !== right[i]) return false;
  }
  return true;
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

    deps.logInfo?.('[dictionary:auto-sync] syncing current anime snapshot');
    const snapshot = await deps.getOrCreateCurrentSnapshot();
    const state = readAutoSyncState(statePath);
    const nextActiveMediaIds = [
      snapshot.mediaId,
      ...state.activeMediaIds.filter((mediaId) => mediaId !== snapshot.mediaId),
    ].slice(0, Math.max(1, Math.floor(config.maxLoaded)));
    deps.logInfo?.(
      `[dictionary:auto-sync] active AniList media set: ${nextActiveMediaIds.join(', ')}`,
    );

    const retainedChanged = !arraysEqual(nextActiveMediaIds, state.activeMediaIds);
    let merged: MergedCharacterDictionaryBuildResult | null = null;
    if (
      retainedChanged ||
      !state.mergedRevision ||
      !state.mergedDictionaryTitle ||
      !snapshot.fromCache
    ) {
      deps.logInfo?.('[dictionary:auto-sync] rebuilding merged dictionary for active anime set');
      merged = await deps.buildMergedDictionary(nextActiveMediaIds);
    }

    const dictionaryTitle = merged?.dictionaryTitle ?? state.mergedDictionaryTitle;
    const revision = merged?.revision ?? state.mergedRevision;
    if (!dictionaryTitle || !revision) {
      throw new Error('Merged character dictionary state is incomplete.');
    }

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
      merged !== null || existing === null || existingRevision === null || existingRevision !== revision;

    if (shouldImport) {
      if (existing !== null) {
        await withOperationTimeout(
          `deleteYomitanDictionary(${dictionaryTitle})`,
          deps.deleteYomitanDictionary(dictionaryTitle),
        );
      }
      if (merged === null) {
        merged = await deps.buildMergedDictionary(nextActiveMediaIds);
      }
      deps.logInfo?.(`[dictionary:auto-sync] importing merged dictionary: ${merged.zipPath}`);
      const imported = await withOperationTimeout(
        `importYomitanDictionary(${path.basename(merged.zipPath)})`,
        deps.importYomitanDictionary(merged.zipPath),
      );
      if (!imported) {
        throw new Error(`Failed to import dictionary ZIP: ${merged.zipPath}`);
      }
    }

    deps.logInfo?.(`[dictionary:auto-sync] applying Yomitan settings for ${dictionaryTitle}`);
    await withOperationTimeout(
      `upsertYomitanDictionarySettings(${dictionaryTitle})`,
      deps.upsertYomitanDictionarySettings(dictionaryTitle, config.profileScope),
    );

    writeAutoSyncState(statePath, {
      activeMediaIds: nextActiveMediaIds,
      mergedRevision: merged?.revision ?? revision,
      mergedDictionaryTitle: merged?.dictionaryTitle ?? dictionaryTitle,
    });
    deps.logInfo?.(
      `[dictionary:auto-sync] synced AniList ${snapshot.mediaId}: ${dictionaryTitle} (${snapshot.entryCount} entries)`,
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
