import * as fs from 'fs';
import * as path from 'path';
import { ensureDir } from '../../shared/fs-utils';
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

export interface CharacterDictionaryManagerEntry {
  mediaId: number;
  label: string;
  title: string;
  current: boolean;
}

export interface CharacterDictionaryManagerSnapshot {
  entries: CharacterDictionaryManagerEntry[];
}

export type CharacterDictionaryManagerMutationResult =
  | (CharacterDictionaryManagerSnapshot & { ok: true; rebuildRequired?: boolean })
  | { ok: false; message: string; entries: CharacterDictionaryManagerEntry[] };

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

function getAutoSyncStatePath(userDataPath: string): string {
  return path.join(userDataPath, 'character-dictionaries', 'auto-sync-state.json');
}

function parseActiveMediaTitle(entry: AutoSyncMediaEntry): string {
  const prefix = `${entry.mediaId} - `;
  if (entry.label.startsWith(prefix)) {
    return entry.label.slice(prefix.length).trim();
  }
  return entry.label === String(entry.mediaId) ? '' : entry.label.trim();
}

function resolveCurrentManagerMediaId(
  state: AutoSyncState,
  currentMediaId?: number | null,
): number | null {
  const normalizedCurrentMediaId =
    typeof currentMediaId === 'number' ? normalizeMediaId(currentMediaId) : null;
  if (normalizedCurrentMediaId !== null) return normalizedCurrentMediaId;
  return state.activeMediaIds[0]?.mediaId ?? null;
}

function toManagerEntries(
  state: AutoSyncState,
  currentMediaId?: number | null,
): CharacterDictionaryManagerEntry[] {
  const resolvedCurrentMediaId = resolveCurrentManagerMediaId(state, currentMediaId);
  return state.activeMediaIds.map((entry, index) => ({
    mediaId: entry.mediaId,
    label: entry.label,
    title: parseActiveMediaTitle(entry),
    current:
      resolvedCurrentMediaId !== null ? entry.mediaId === resolvedCurrentMediaId : index === 0,
  }));
}

export function getCharacterDictionaryManagerSnapshot(
  userDataPath: string,
  currentMediaId?: number | null,
): CharacterDictionaryManagerSnapshot {
  return {
    entries: toManagerEntries(
      readAutoSyncState(getAutoSyncStatePath(userDataPath)),
      currentMediaId,
    ),
  };
}

export function moveCharacterDictionaryManagedEntry(
  userDataPath: string,
  mediaId: number,
  direction: 1 | -1,
  currentMediaId?: number | null,
): CharacterDictionaryManagerMutationResult {
  const statePath = getAutoSyncStatePath(userDataPath);
  const state = readAutoSyncState(statePath);
  const managerEntries = toManagerEntries(state, currentMediaId);
  const index = state.activeMediaIds.findIndex((entry) => entry.mediaId === mediaId);
  if (index < 0) {
    return {
      ok: false,
      message: 'Character dictionary entry not found.',
      entries: managerEntries,
    };
  }
  if (managerEntries[index]?.current) {
    return {
      ok: false,
      message: 'The current anime stays anchored while you are watching it.',
      entries: managerEntries,
    };
  }
  const targetIndex = Math.min(state.activeMediaIds.length - 1, Math.max(0, index + direction));
  if (targetIndex === index) {
    return { ok: true, entries: managerEntries };
  }
  const nextActiveMediaIds = [...state.activeMediaIds];
  const [entry] = nextActiveMediaIds.splice(index, 1);
  if (entry) {
    nextActiveMediaIds.splice(targetIndex, 0, entry);
  }
  const nextState = { ...state, activeMediaIds: nextActiveMediaIds };
  writeAutoSyncState(statePath, nextState);
  return { ok: true, entries: toManagerEntries(nextState, currentMediaId) };
}

export function removeCharacterDictionaryManagedEntry(
  userDataPath: string,
  mediaId: number,
  currentMediaId?: number | null,
): CharacterDictionaryManagerMutationResult {
  const statePath = getAutoSyncStatePath(userDataPath);
  const state = readAutoSyncState(statePath);
  const managerEntries = toManagerEntries(state, currentMediaId);
  const index = state.activeMediaIds.findIndex((entry) => entry.mediaId === mediaId);
  if (index < 0) {
    return {
      ok: false,
      message: 'Character dictionary entry not found.',
      entries: managerEntries,
    };
  }
  if (managerEntries[index]?.current) {
    return {
      ok: false,
      message: 'The current anime stays loaded while you are watching it.',
      entries: managerEntries,
    };
  }
  const nextState = {
    ...state,
    activeMediaIds: state.activeMediaIds.filter((entry) => entry.mediaId !== mediaId),
    mergedRevision: null,
  };
  writeAutoSyncState(statePath, nextState);
  return { ok: true, entries: toManagerEntries(nextState, currentMediaId), rebuildRequired: true };
}

export function replaceCharacterDictionaryManagedEntry(
  userDataPath: string,
  mediaId: number,
  replacement: { mediaId: number; mediaTitle: string },
): CharacterDictionaryManagerMutationResult {
  const statePath = getAutoSyncStatePath(userDataPath);
  const state = readAutoSyncState(statePath);
  const index = state.activeMediaIds.findIndex((entry) => entry.mediaId === mediaId);
  if (index < 0) {
    return {
      ok: false,
      message: 'Character dictionary entry not found.',
      entries: toManagerEntries(state),
    };
  }
  const normalizedReplacementMediaId = normalizeMediaId(replacement.mediaId);
  const mediaTitle = replacement.mediaTitle.trim();
  if (normalizedReplacementMediaId === null || !mediaTitle) {
    return {
      ok: false,
      message: 'Invalid replacement AniList media.',
      entries: toManagerEntries(state),
    };
  }
  const replacementEntry = {
    mediaId: normalizedReplacementMediaId,
    label: buildActiveMediaLabel(normalizedReplacementMediaId, mediaTitle),
  };
  const nextActiveMediaIds = state.activeMediaIds
    .map((entry, entryIndex) => (entryIndex === index ? replacementEntry : entry))
    .filter(
      (entry, entryIndex, entries) =>
        entries.findIndex((candidate) => candidate.mediaId === entry.mediaId) === entryIndex,
    );
  const nextState = {
    ...state,
    activeMediaIds: nextActiveMediaIds,
    mergedRevision: null,
  };
  writeAutoSyncState(statePath, nextState);
  return { ok: true, entries: toManagerEntries(nextState), rebuildRequired: true };
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
  getCurrentMediaId: () => number | null;
} {
  const dictionariesDir = path.join(deps.userDataPath, 'character-dictionaries');
  const statePath = getAutoSyncStatePath(deps.userDataPath);
  const schedule = deps.schedule ?? ((fn, delayMs) => setTimeout(fn, delayMs));
  const clearSchedule = deps.clearSchedule ?? ((timer) => clearTimeout(timer));
  const debounceMs = 800;
  const operationTimeoutMs = Math.max(1, Math.floor(deps.operationTimeoutMs ?? 7_000));

  let debounceTimer: ReturnType<typeof setTimeout> | null = null;
  let syncInFlight = false;
  let runQueued = false;
  let activeCurrentMediaId: number | null = null;

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
      activeCurrentMediaId = null;
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
          activeCurrentMediaId = mediaId;
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
          activeCurrentMediaId = mediaId;
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
      activeCurrentMediaId = snapshot.mediaId;
      const state = readAutoSyncState(statePath);
      const staleMediaIds = new Set(
        (snapshot.staleMediaIds ?? [])
          .map((mediaId) => normalizeMediaId(mediaId))
          .filter((mediaId): mediaId is number => mediaId !== null),
      );
      const nextActiveMediaIds = [
        {
          mediaId: snapshot.mediaId,
          label: buildActiveMediaLabel(snapshot.mediaId, snapshot.mediaTitle),
        },
        ...state.activeMediaIds.filter(
          (entry) => entry.mediaId !== snapshot.mediaId && !staleMediaIds.has(entry.mediaId),
        ),
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
          const existingMergedZipPath = path.join(dictionariesDir, 'merged.zip');
          if (fs.existsSync(existingMergedZipPath)) {
            merged = {
              zipPath: existingMergedZipPath,
              revision,
              dictionaryTitle,
              entryCount: snapshot.entryCount,
            };
          } else {
            merged = await deps.buildMergedDictionary(nextActiveMediaIdValues);
          }
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
    getCurrentMediaId: () => activeCurrentMediaId,
  };
}
