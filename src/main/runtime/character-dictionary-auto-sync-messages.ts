import type { CharacterDictionarySnapshotStageProgress } from '../character-dictionary-runtime';

export function buildSyncingMessage(mediaTitle: string): string {
  return `Updating character dictionary for ${mediaTitle}...`;
}

export function buildCheckingMessage(mediaTitle: string): string {
  return `Checking character dictionary for ${mediaTitle}...`;
}

export function buildGeneratingMessage(mediaTitle: string, detail?: string): string {
  return detail
    ? `Generating character dictionary for ${mediaTitle} (${detail})...`
    : `Generating character dictionary for ${mediaTitle}...`;
}

export function formatCharacterDictionaryProgressDetail(
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

export function formatElapsedDuration(elapsedMs: number): string {
  const totalSeconds = Math.max(0, Math.floor(elapsedMs / 1000));
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;
  return minutes > 0 ? `${minutes}m ${String(seconds).padStart(2, '0')}s` : `${seconds}s`;
}

/** Coarser than the elapsed clock: an estimate that ticks every second reads as precision it lacks. */
export function formatRemainingDuration(remainingMs: number): string {
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
export function joinGeneratingDetail(detail: string | null, elapsedMs: number): string | undefined {
  const parts = [detail, elapsedMs >= 5_000 ? formatElapsedDuration(elapsedMs) : null].filter(
    (part): part is string => typeof part === 'string' && part.length > 0,
  );
  return parts.length > 0 ? parts.join(' · ') : undefined;
}

export function buildImportingMessage(mediaTitle: string, elapsedMs?: number): string {
  const elapsed =
    typeof elapsedMs === 'number' && elapsedMs >= 1000
      ? ` (${formatElapsedDuration(elapsedMs)})`
      : '';
  return `Importing character dictionary for ${mediaTitle}${elapsed}...`;
}

export function buildBuildingMessage(mediaTitle: string): string {
  return `Building character dictionary for ${mediaTitle}...`;
}

export function buildReadyMessage(mediaTitle: string): string {
  return `Character dictionary ready for ${mediaTitle}`;
}

export function buildFailedMessage(mediaTitle: string | null, errorMessage: string): string {
  if (mediaTitle) {
    return `Character dictionary sync failed for ${mediaTitle}: ${errorMessage}`;
  }
  return `Character dictionary sync failed: ${errorMessage}`;
}
