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

const EVENT_TYPES = new Set([
  'stage',
  'merge-summary',
  'remote-output',
  'snapshot-created',
  'check-result',
  'result',
]);

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
  const type = (parsed as { type?: unknown }).type;
  if (typeof type !== 'string' || !EVENT_TYPES.has(type)) return null;
  return parsed as SyncProgressEvent;
}
