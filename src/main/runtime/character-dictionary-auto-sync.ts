import * as fs from 'fs';
import * as path from 'path';
import { ensureDir } from '../../shared/fs-utils';
import type { AnilistCharacterDictionaryProfileScope } from '../../types';
import type {
  CharacterDictionarySnapshotProgressCallbacks,
  CharacterDictionarySnapshotResult,
  CharacterDictionarySnapshotStageProgress,
  MergedCharacterDictionaryBuildResult,
} from '../character-dictionary-runtime';

const DEFAULT_IMPORT_TIMEOUT_BASE_MS = 120_000;
const IMPORT_TIMEOUT_MS_PER_MB = 6_000;
const IMPORT_TIMEOUT_MAX_MS = 1_800_000;

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
  /** Budget for the quick Yomitan queries (dictionary info, settings upsert). */
  operationTimeoutMs?: number;
  /**
   * Base budget for the slow Yomitan mutations (delete + import). The effective budget grows with
   * the merged ZIP size, because importing a large dictionary can take several minutes.
   */
  dictionaryImportTimeoutBaseMs?: number;
  heartbeatMs?: number;
  progressThrottleMs?: number;
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
  const nextState = { ...state, activeMediaIds: nextActiveMediaIds, mergedRevision: null };
  writeAutoSyncState(statePath, nextState);
  return { ok: true, entries: toManagerEntries(nextState, currentMediaId), rebuildRequired: true };
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

function buildGeneratingMessage(mediaTitle: string, detail?: string): string {
  return detail
    ? `Generating character dictionary for ${mediaTitle} (${detail})...`
    : `Generating character dictionary for ${mediaTitle}...`;
}

function formatCharacterDictionaryProgressDetail(
  progress: CharacterDictionarySnapshotStageProgress,
  remainingMs: number | null,
): string {
  if (progress.stage === 'saving') {
    return 'saving snapshot';
  }
  if (progress.stage === 'names') {
    return progress.total !== null && progress.total > 0
      ? `name ${progress.completed}/${progress.total}`
      : `${progress.completed} names`;
  }
  if (progress.stage === 'images') {
    const counted =
      progress.total !== null && progress.total > 0
        ? `image ${progress.completed}/${progress.total}`
        : `${progress.completed} images`;
    return remainingMs !== null
      ? `${counted}, ~${formatRemainingDuration(remainingMs)} left`
      : counted;
  }
  const page = typeof progress.page === 'number' ? `page ${progress.page}, ` : '';
  return `${page}${progress.completed} characters`;
}

function formatElapsedDuration(elapsedMs: number): string {
  const totalSeconds = Math.max(0, Math.floor(elapsedMs / 1000));
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;
  return minutes > 0 ? `${minutes}m ${String(seconds).padStart(2, '0')}s` : `${seconds}s`;
}

/** Coarser than the elapsed clock: an estimate that ticks every second reads as precision it lacks. */
function formatRemainingDuration(remainingMs: number): string {
  const totalSeconds = Math.max(1, Math.round(remainingMs / 1000));
  if (totalSeconds >= 60) {
    return `${Math.max(1, Math.round(totalSeconds / 60))}m`;
  }
  return `${Math.max(5, Math.ceil(totalSeconds / 5) * 5)}s`;
}

/**
 * The elapsed clock is the part that proves the app is alive: a stalled network fetch freezes the
 * counts, but the clock keeps moving.
 */
function joinGeneratingDetail(detail: string | null, elapsedMs: number): string | undefined {
  const parts = [detail, elapsedMs >= 5_000 ? formatElapsedDuration(elapsedMs) : null].filter(
    (part): part is string => typeof part === 'string' && part.length > 0,
  );
  return parts.length > 0 ? parts.join(' · ') : undefined;
}

function buildImportingMessage(mediaTitle: string, elapsedMs?: number): string {
  const elapsed =
    typeof elapsedMs === 'number' && elapsedMs >= 1000
      ? ` (${formatElapsedDuration(elapsedMs)})`
      : '';
  return `Importing character dictionary for ${mediaTitle}${elapsed}...`;
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
  const dictionaryImportTimeoutBaseMs = Math.max(
    1,
    Math.floor(deps.dictionaryImportTimeoutBaseMs ?? DEFAULT_IMPORT_TIMEOUT_BASE_MS),
  );
  const heartbeatMs = Math.max(1, Math.floor(deps.heartbeatMs ?? 5_000));
  const progressThrottleMs = Math.max(0, Math.floor(deps.progressThrottleMs ?? 1_000));

  let debounceTimer: ReturnType<typeof setTimeout> | null = null;
  let syncInFlight = false;
  let runQueued = false;
  let activeCurrentMediaId: number | null = null;

  const withTimeout = async <T>(
    label: string,
    promise: Promise<T>,
    timeoutMs: number,
  ): Promise<T> => {
    let timer: ReturnType<typeof setTimeout> | null = null;
    try {
      return await Promise.race([
        promise,
        new Promise<never>((_resolve, reject) => {
          timer = setTimeout(() => {
            reject(new Error(`${label} timed out after ${timeoutMs}ms`));
          }, timeoutMs);
        }),
      ]);
    } finally {
      if (timer !== null) {
        clearTimeout(timer);
      }
    }
  };

  const withOperationTimeout = <T>(label: string, promise: Promise<T>): Promise<T> =>
    withTimeout(label, promise, operationTimeoutMs);

  /**
   * Importing a merged dictionary means Yomitan writing every term and image into IndexedDB, which
   * scales with the ZIP: a single-season dictionary lands in seconds, One Piece takes minutes. Size
   * the budget off the archive instead of failing a healthy import on a flat deadline.
   */
  const resolveImportTimeoutMs = (zipPath: string | null): number => {
    let bytes = 0;
    if (zipPath) {
      try {
        bytes = fs.statSync(zipPath).size;
      } catch {
        bytes = 0;
      }
    }
    const sizeAllowanceMs = (bytes / (1024 * 1024)) * IMPORT_TIMEOUT_MS_PER_MB;
    return Math.min(
      IMPORT_TIMEOUT_MAX_MS,
      Math.max(
        dictionaryImportTimeoutBaseMs,
        Math.round(dictionaryImportTimeoutBaseMs + sizeAllowanceMs),
      ),
    );
  };

  /** Keeps the persistent notification ticking so a multi-minute import never looks hung. */
  const withHeartbeat = async <T>(
    run: () => Promise<T>,
    onTick: (elapsedMs: number) => void,
  ): Promise<T> => {
    const startedAt = deps.now();
    let stopped = false;
    let timer: ReturnType<typeof setTimeout> | null = null;
    const tick = (): void => {
      if (stopped) return;
      onTick(deps.now() - startedAt);
      timer = schedule(tick, heartbeatMs);
    };
    timer = schedule(tick, heartbeatMs);
    try {
      return await run();
    } finally {
      stopped = true;
      if (timer !== null) {
        clearSchedule(timer);
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
    let lastProgressAt = Number.NEGATIVE_INFINITY;
    let lastProgressStage: CharacterDictionarySnapshotStageProgress['stage'] | null = null;
    let generating: {
      mediaId: number;
      mediaTitle: string;
      startedAt: number;
      detail: string | null;
    } | null = null;
    let imageRate: { startedAt: number; startCompleted: number } | null = null;

    const emitGeneratingStatus = (): void => {
      if (!generating) {
        return;
      }
      deps.onSyncStatus?.({
        phase: 'generating',
        mediaId: generating.mediaId,
        mediaTitle: generating.mediaTitle,
        message: buildGeneratingMessage(
          generating.mediaTitle,
          joinGeneratingDetail(generating.detail, deps.now() - generating.startedAt),
        ),
      });
    };

    // Image downloads are serial and evenly paced, so the observed rate predicts the tail well.
    const estimateRemainingImageMs = (
      progress: CharacterDictionarySnapshotStageProgress,
      nowMs: number,
    ): number | null => {
      if (progress.stage !== 'images' || progress.total === null || imageRate === null) {
        return null;
      }
      const done = progress.completed - imageRate.startCompleted;
      const elapsedMs = nowMs - imageRate.startedAt;
      const remaining = progress.total - progress.completed;
      if (done <= 0 || elapsedMs <= 0 || remaining <= 0) {
        return null;
      }
      return Math.round((elapsedMs / done) * remaining);
    };

    try {
      deps.logInfo?.('[dictionary:auto-sync] syncing current anime snapshot');
      const snapshot = await withHeartbeat(
        () =>
          deps.getOrCreateCurrentSnapshot(undefined, {
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
              lastProgressAt = Number.NEGATIVE_INFINITY;
              lastProgressStage = null;
              imageRate = null;
              generating = { mediaId, mediaTitle, startedAt: deps.now(), detail: null };
              emitGeneratingStatus();
            },
            // Long-running work (AniList character pages, then one image download per character
            // and voice actor, then MeCab name splits) reports counts so the notification shows
            // movement instead of a frozen "Generating..." for the minutes a large series takes.
            onGenerateProgress: (progress) => {
              const nowMs = deps.now();
              if (!generating) {
                generating = {
                  mediaId: progress.mediaId,
                  mediaTitle: progress.mediaTitle,
                  startedAt: nowMs,
                  detail: null,
                };
              }
              if (progress.stage === 'images' && imageRate === null) {
                imageRate = { startedAt: nowMs, startCompleted: progress.completed };
              }
              // Stage changes and the last item of a stage always report; the throttle only thins
              // out the run of identical-looking updates in between.
              const isFinal = progress.total !== null && progress.completed >= progress.total;
              const isStageChange = progress.stage !== lastProgressStage;
              if (!isFinal && !isStageChange && nowMs - lastProgressAt < progressThrottleMs) {
                return;
              }
              lastProgressAt = nowMs;
              lastProgressStage = progress.stage;
              generating.mediaId = progress.mediaId;
              generating.mediaTitle = progress.mediaTitle;
              generating.detail = formatCharacterDictionaryProgressDetail(
                progress,
                estimateRemainingImageMs(progress, nowMs),
              );
              emitGeneratingStatus();
            },
          }),
        // Ticks even when a step stalls, so the message keeps moving while the counts do not.
        () => emitGeneratingStatus(),
      );
      generating = null;
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
        await withHeartbeat(
          async () => {
            const importTimeoutMs = resolveImportTimeoutMs(
              merged?.zipPath ?? path.join(dictionariesDir, 'merged.zip'),
            );
            if (existing !== null) {
              await withTimeout(
                `deleteYomitanDictionary(${dictionaryTitle})`,
                deps.deleteYomitanDictionary(dictionaryTitle),
                importTimeoutMs,
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
            const mergedZipPath = merged.zipPath;
            const mergedImportTimeoutMs = resolveImportTimeoutMs(mergedZipPath);
            deps.logInfo?.(
              `[dictionary:auto-sync] importing merged dictionary: ${mergedZipPath} (timeout ${mergedImportTimeoutMs}ms)`,
            );
            const imported = await withTimeout(
              `importYomitanDictionary(${path.basename(mergedZipPath)})`,
              deps.importYomitanDictionary(mergedZipPath),
              mergedImportTimeoutMs,
            );
            if (!imported) {
              throw new Error(`Failed to import dictionary ZIP: ${merged.zipPath}`);
            }
          },
          (elapsedMs) => {
            deps.onSyncStatus?.({
              phase: 'importing',
              mediaId: snapshot.mediaId,
              mediaTitle: snapshot.mediaTitle,
              message: buildImportingMessage(snapshot.mediaTitle, elapsedMs),
            });
          },
        );
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
