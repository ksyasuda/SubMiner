// NDJSON progress protocol emitted by `subminer sync --json` and consumed by
// the sync UI's launcher client. Keep this file free of Electron/bun imports —
// it is shared between the launcher (bun) and the app main process (node).

export interface SyncMergeSummary {
  sessionsMerged: number;
  sessionsAlreadyPresent: number;
  activeSessionsSkipped: number;
  animeAdded: number;
  videosAdded: number;
  wordsAdded: number;
  kanjiAdded: number;
  subtitleLinesAdded: number;
  telemetryRowsAdded: number;
  eventsAdded: number;
  excludedWordsAdded: number;
  dailyRollupsCopied: number;
  monthlyRollupsCopied: number;
  rollupGroupsRecomputed: number;
}

export type SyncStage =
  | 'snapshot-local'
  | 'snapshot-remote'
  | 'download'
  | 'upload'
  | 'merge-local'
  | 'merge-remote';

export type SyncProgressEvent =
  | { type: 'stage'; stage: SyncStage; message: string }
  | { type: 'merge-summary'; target: 'local' | 'remote'; summary: SyncMergeSummary }
  | { type: 'remote-output'; text: string }
  | { type: 'snapshot-created'; path: string }
  | {
      type: 'check-result';
      host: string;
      sshOk: boolean;
      remoteCommand: string | null;
      remoteVersion: string | null;
      ok: boolean;
      error: string | null;
    }
  | { type: 'result'; ok: boolean; error: string | null };

const SUMMARY_KEYS: readonly (keyof SyncMergeSummary)[] = [
  'sessionsMerged',
  'sessionsAlreadyPresent',
  'activeSessionsSkipped',
  'animeAdded',
  'videosAdded',
  'wordsAdded',
  'kanjiAdded',
  'subtitleLinesAdded',
  'telemetryRowsAdded',
  'eventsAdded',
  'excludedWordsAdded',
  'dailyRollupsCopied',
  'monthlyRollupsCopied',
  'rollupGroupsRecomputed',
];

const STAGES: readonly SyncStage[] = [
  'snapshot-local',
  'snapshot-remote',
  'download',
  'upload',
  'merge-local',
  'merge-remote',
];

function isString(value: unknown): value is string {
  return typeof value === 'string';
}

function isNullableString(value: unknown): value is string | null {
  return value === null || typeof value === 'string';
}

function isMergeSummary(value: unknown): value is SyncMergeSummary {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const record = value as Record<string, unknown>;
  return SUMMARY_KEYS.every((key) => typeof record[key] === 'number');
}

/**
 * The events cross a process boundary (the sync child's stdout), so each variant
 * is validated field-by-field: a half-formed `result` or `merge-summary` would
 * otherwise be cast straight into the UI, which reads `.ok`/`.summary` directly.
 */
export function parseSyncProgressLine(line: string): SyncProgressEvent | null {
  const trimmed = line.trim();
  if (!trimmed.startsWith('{')) return null;
  let parsed: unknown;
  try {
    parsed = JSON.parse(trimmed);
  } catch {
    return null;
  }
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) return null;
  const event = parsed as Record<string, unknown>;

  switch (event.type) {
    case 'stage':
      return STAGES.includes(event.stage as SyncStage) && isString(event.message)
        ? (parsed as SyncProgressEvent)
        : null;
    case 'merge-summary':
      return (event.target === 'local' || event.target === 'remote') && isMergeSummary(event.summary)
        ? (parsed as SyncProgressEvent)
        : null;
    case 'remote-output':
      return isString(event.text) ? (parsed as SyncProgressEvent) : null;
    case 'snapshot-created':
      return isString(event.path) ? (parsed as SyncProgressEvent) : null;
    case 'check-result':
      return isString(event.host) &&
        typeof event.sshOk === 'boolean' &&
        typeof event.ok === 'boolean' &&
        isNullableString(event.remoteCommand) &&
        isNullableString(event.remoteVersion) &&
        isNullableString(event.error)
        ? (parsed as SyncProgressEvent)
        : null;
    case 'result':
      return typeof event.ok === 'boolean' && isNullableString(event.error)
        ? (parsed as SyncProgressEvent)
        : null;
    default:
      return null;
  }
}
