import * as fs from 'fs';
import * as path from 'path';
import type { AnilistCharacterDictionaryProfileScope } from '../../types';
import type {
  CharacterDictionarySnapshotProgressCallbacks,
  CharacterDictionarySnapshotResult,
  MergedCharacterDictionaryBuildResult,
} from '../character-dictionary-runtime';

type AutoSyncMediaEntry = {
  mediaId: number;
  label: string;
};

type AutoSyncState = {
  activeMediaIds: AutoSyncMediaEntry[];
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

export interface CharacterDictionaryAutoSyncStatusEvent {
  phase: 'checking' | 'generating' | 'syncing' | 'building' | 'importing' | 'ready' | 'failed';
  mediaId?: number;
  mediaTitle?: string;
  message: string;
  changed?: boolean;
}

export interface CharacterDictionaryAutoSyncRuntimeDeps {
  userDataPath: string;
  getConfig: () => CharacterDictionaryAutoSyncConfig;
  getOrCreateCurrentSnapshot: (
    targetPath?: string,
    progress?: CharacterDictionarySnapshotProgressCallbacks,
  ) => Promise<CharacterDictionarySnapshotResult>;
  buildMergedDictionary: (mediaIds: number[]) => Promise<MergedCharacterDictionaryBuildResult>;
  waitForYomitanMutationReady?: () => Promise<void>;
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
  onSyncStatus?: (event: CharacterDictionaryAutoSyncStatusEvent) => void;
  onSyncComplete?: (result: { mediaId: number; mediaTitle: string; changed: boolean }) => void;
}

function ensureDir(dirPath: string): void {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
}

function normalizeMediaId(rawMediaId: number): number | null {
  const mediaId = Math.max(1, Math.floor(rawMediaId));
  return Number.isFinite(mediaId) ? mediaId : null;
}

function parseActiveMediaEntry(rawEntry: unknown): AutoSyncMediaEntry | null {
  if (typeof rawEntry === 'number') {
    const mediaId = normalizeMediaId(rawEntry);
    if (mediaId === null) {
      return null;
    }
    return { mediaId, label: String(mediaId) };
  }

  if (typeof rawEntry !== 'string') {
    return null;
  }

  const trimmed = rawEntry.trim();
  if (!trimmed) {
    return null;
  }

  const [rawId, ...rawTitleParts] = trimmed.split(' - ');
  if (!rawId || !/^\d+$/.test(rawId)) {
    return null;
  }
  const mediaId = normalizeMediaId(Number.parseInt(rawId ?? '', 10));
  if (mediaId === null || mediaId <= 0) {
    return null;
  }

  const rawLabel = rawTitleParts.length > 0 ? rawTitleParts.join(' - ').trim() : '';
  return { mediaId, label: rawLabel ? `${mediaId} - ${rawLabel}` : String(mediaId) };
}

function buildActiveMediaLabel(mediaId: number, mediaTitle: string | null | undefined): string {
  const normalizedId = normalizeMediaId(mediaId);
  const trimmedTitle = typeof mediaTitle === 'string' ? mediaTitle.trim() : '';
  if (normalizedId === null) {
    return trimmedTitle;
  }
  return trimmedTitle.length > 0 ? `${normalizedId} - ${trimmedTitle}` : String(normalizedId);
}

function readAutoSyncState(statePath: string): AutoSyncState {
  try {
    const raw = fs.readFileSync(statePath, 'utf8');
    const parsed = JSON.parse(raw) as Partial<AutoSyncState>;
    const activeMediaIds: AutoSyncMediaEntry[] = [];
    const activeMediaIdSet = new Set<number>();
    if (Array.isArray(parsed.activeMediaIds)) {
      for (const value of parsed.activeMediaIds) {
        const entry = parseActiveMediaEntry(value);
        if (entry && !activeMediaIdSet.has(entry.mediaId)) {
          activeMediaIdSet.add(entry.mediaId);
          activeMediaIds.push(entry);
        }
      }
    }
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
  const persistedState = {
    activeMediaIds: state.activeMediaIds.map((entry) => entry.label),
    mergedRevision: state.mergedRevision,
    mergedDictionaryTitle: state.mergedDictionaryTitle,
  };
  fs.writeFileSync(statePath, JSON.stringify(persistedState, null, 2), 'utf8');
}

function arraysEqual(left: number[], right: number[]): boolean {
  if (left.length !== right.length) return false;
  for (let i = 0; i < left.length; i += 1) {
    if (left[i] !== right[i]) return false;
  }
  return true;
}

function sameMembership(left: number[], right: number[]): boolean {
  if (left.length !== right.length) return false;
  const leftSorted = [...left].sort((a, b) => a - b);
  const rightSorted = [...right].sort((a, b) => a - b);
  return arraysEqual(leftSorted, rightSorted);
}

function buildSyncingMessage(mediaTitle: string): string {
  return `Updating character dictionary for ${mediaTitle}...`;
}

function buildCheckingMessage(mediaTitle: string): string {
  return `Checking character dictionary for ${mediaTitle}...`;
}

function buildGeneratingMessage(mediaTitle: string): string {
  return `Generating character dictionary for ${mediaTitle}...`;
}

function buildImportingMessage(mediaTitle: string): string {
  return `Importing character dictionary for ${mediaTitle}...`;
}

function buildBuildingMessage(mediaTitle: string): string {
  return `Building character dictionary for ${mediaTitle}...`;
}

function buildReadyMessage(mediaTitle: string): string {
  return `Character dictionary ready for ${mediaTitle}`;
}

function buildFailedMessage(mediaTitle: string | null, errorMessage: string): string {
  if (mediaTitle) {
    return `Character dictionary sync failed for ${mediaTitle}: ${errorMessage}`;
  }
  return `Character dictionary sync failed: ${errorMessage}`;
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

    let currentMediaId: number | undefined;
    let currentMediaTitle: string | null = null;

    try {
      deps.logInfo?.('[dictionary:auto-sync] syncing current anime snapshot');
      const snapshot = await deps.getOrCreateCurrentSnapshot(undefined, {
        onChecking: ({ mediaId, mediaTitle }) => {
          currentMediaId = mediaId;
          currentMediaTitle = mediaTitle;
          deps.onSyncStatus?.({
            phase: 'checking',
            mediaId,
            mediaTitle,
            message: buildCheckingMessage(mediaTitle),
          });
        },
        onGenerating: ({ mediaId, mediaTitle }) => {
          currentMediaId = mediaId;
          currentMediaTitle = mediaTitle;
          deps.onSyncStatus?.({
            phase: 'generating',
            mediaId,
            mediaTitle,
            message: buildGeneratingMessage(mediaTitle),
          });
        },
      });
      currentMediaId = snapshot.mediaId;
      currentMediaTitle = snapshot.mediaTitle;
      deps.onSyncStatus?.({
        phase: 'syncing',
        mediaId: snapshot.mediaId,
        mediaTitle: snapshot.mediaTitle,
        message: buildSyncingMessage(snapshot.mediaTitle),
      });
      const state = readAutoSyncState(statePath);
      const nextActiveMediaIds = [
        {
          mediaId: snapshot.mediaId,
          label: buildActiveMediaLabel(snapshot.mediaId, snapshot.mediaTitle),
        },
        ...state.activeMediaIds.filter((entry) => entry.mediaId !== snapshot.mediaId),
      ].slice(0, Math.max(1, Math.floor(config.maxLoaded)));
      const nextActiveMediaIdValues = nextActiveMediaIds.map((entry) => entry.mediaId);
      deps.logInfo?.(
        `[dictionary:auto-sync] active AniList media set: ${nextActiveMediaIds
          .map((entry) => entry.label)
          .join(', ')}`,
      );

      const stateMediaIds = state.activeMediaIds.map((entry) => entry.mediaId);
      const retainedOrderChanged = !arraysEqual(nextActiveMediaIdValues, stateMediaIds);
      const retainedMembershipChanged = !sameMembership(nextActiveMediaIdValues, stateMediaIds);
      let merged: MergedCharacterDictionaryBuildResult | null = null;
      if (
        retainedMembershipChanged ||
        !state.mergedRevision ||
        !state.mergedDictionaryTitle ||
        !snapshot.fromCache
      ) {
        deps.onSyncStatus?.({
          phase: 'building',
          mediaId: snapshot.mediaId,
          mediaTitle: snapshot.mediaTitle,
          message: buildBuildingMessage(snapshot.mediaTitle),
        });
        deps.logInfo?.('[dictionary:auto-sync] rebuilding merged dictionary for active anime set');
        merged = await deps.buildMergedDictionary(nextActiveMediaIdValues);
      }

      const dictionaryTitle = merged?.dictionaryTitle ?? state.mergedDictionaryTitle;
      const revision = merged?.revision ?? state.mergedRevision;
      if (!dictionaryTitle || !revision) {
        throw new Error('Merged character dictionary state is incomplete.');
      }

      writeAutoSyncState(statePath, {
        activeMediaIds: nextActiveMediaIds,
        mergedRevision: merged?.revision ?? revision,
        mergedDictionaryTitle: merged?.dictionaryTitle ?? dictionaryTitle,
      });

      await deps.waitForYomitanMutationReady?.();

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
        merged !== null ||
        existing === null ||
        existingRevision === null ||
        existingRevision !== revision;
      let changed = merged !== null || retainedOrderChanged;

      if (shouldImport) {
        deps.onSyncStatus?.({
          phase: 'importing',
          mediaId: snapshot.mediaId,
          mediaTitle: snapshot.mediaTitle,
          message: buildImportingMessage(snapshot.mediaTitle),
        });
        if (existing !== null) {
          await withOperationTimeout(
            `deleteYomitanDictionary(${dictionaryTitle})`,
            deps.deleteYomitanDictionary(dictionaryTitle),
          );
        }
        if (merged === null) {
          merged = await deps.buildMergedDictionary(nextActiveMediaIdValues);
        }
        deps.logInfo?.(`[dictionary:auto-sync] importing merged dictionary: ${merged.zipPath}`);
        const imported = await withOperationTimeout(
          `importYomitanDictionary(${path.basename(merged.zipPath)})`,
          deps.importYomitanDictionary(merged.zipPath),
        );
        if (!imported) {
          throw new Error(`Failed to import dictionary ZIP: ${merged.zipPath}`);
        }
        changed = true;
      }

      deps.logInfo?.(`[dictionary:auto-sync] applying Yomitan settings for ${dictionaryTitle}`);
      const settingsUpdated = await withOperationTimeout(
        `upsertYomitanDictionarySettings(${dictionaryTitle})`,
        deps.upsertYomitanDictionarySettings(dictionaryTitle, config.profileScope),
      );
      changed = changed || settingsUpdated === true;

      deps.logInfo?.(
        `[dictionary:auto-sync] synced AniList ${snapshot.mediaId}: ${dictionaryTitle} (${snapshot.entryCount} entries)`,
      );
      deps.onSyncStatus?.({
        phase: 'ready',
        mediaId: snapshot.mediaId,
        mediaTitle: snapshot.mediaTitle,
        message: buildReadyMessage(snapshot.mediaTitle),
        changed,
      });
      deps.onSyncComplete?.({
        mediaId: snapshot.mediaId,
        mediaTitle: snapshot.mediaTitle,
        changed,
      });
    } catch (error) {
      const errorMessage = (error as Error)?.message ?? String(error);
      deps.onSyncStatus?.({
        phase: 'failed',
        mediaId: currentMediaId,
        mediaTitle: currentMediaTitle ?? undefined,
        message: buildFailedMessage(currentMediaTitle, errorMessage),
      });
      throw error;
    }
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
