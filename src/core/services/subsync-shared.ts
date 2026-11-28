import { redactUrlCredentials } from './subtitle-generation-process';
import { SubsyncManualPayload } from '../../types';
import {
  CommandResult,
  fileExists,
  findSubsyncExecutable,
  hasPathSeparators,
  SubsyncResolvedConfig,
} from '../../subsync/utils';
import { type AlignmentResult } from '../../subsync/alass-align';
import { type EvaluatedSubtitleReference } from '../../subsync/reference-discovery';
import { type SubtitleTimeSpan } from '../../subsync/subtitle-lines';

export interface FileExtractionResult {
  path: string;
  temporary: boolean;
}

export type SubtitleSlot = 'primary' | 'secondary';

export const SYNCED_TRACK_LOOKUP_ATTEMPTS = 5;

export const SYNCED_TRACK_LOOKUP_RETRY_MS = 100;

const ALREADY_IN_SYNC_MS = 80;

export const MIN_REFERENCE_DIALOGUE_LINES = 15;

export const AUTO_SUBTITLE_CONFIDENCE_THRESHOLD = 0.65;

export const AUTO_REMOTE_CONFIDENCE_THRESHOLD = 0.7;

export const AUTO_AUDIO_CONFIDENCE_THRESHOLD = 0.65;

export const AUDIO_SPEECH_LEAD_SEC = 0.1;

export const JELLYFIN_SUBTITLE_FETCH_TIMEOUT_MS = 15_000;

const COMMAND_OUTPUT_SUMMARY_CHARS = 300;

// Keep the tail of tool output (errors come last; alass prints progress first) and strip
// stream tokens: ffmpeg echoes input URLs, and Jellyfin URLs carry ApiKey.
function summarizeCommandOutput(text: string): string {
  const trimmed = redactUrlCredentials(text.trim());
  return trimmed.length > COMMAND_OUTPUT_SUMMARY_CHARS
    ? `…${trimmed.slice(-COMMAND_OUTPUT_SUMMARY_CHARS)}`
    : trimmed;
}

export function summarizeCommandFailure(command: string, result: CommandResult): string {
  const parts = [
    `code=${result.code ?? 'n/a'}`,
    result.stderr ? `stderr: ${summarizeCommandOutput(result.stderr)}` : '',
    result.stdout ? `stdout: ${summarizeCommandOutput(result.stdout)}` : '',
    result.error ? `error: ${summarizeCommandOutput(result.error)}` : '',
  ]
    .map((value) => value.trim())
    .filter(Boolean);

  if (parts.length === 0) return `command failed (${command})`;
  return `command failed (${command}) ${parts.join(' | ')}`;
}

export interface MpvClientLike {
  connected: boolean;
  currentAudioStreamIndex: number | null;
  send: (payload: { command: (string | number)[] }) => void;
  requestProperty: (name: string) => Promise<unknown>;
}

export interface SubsyncVadPaths {
  vadExecutablePath: string | null;
  vadModelPath: string | null;
}

export interface SubsyncCoreDeps {
  getMpvClient: () => MpvClientLike | null;
  getResolvedConfig: () => SubsyncResolvedConfig;
  getVadPaths?: () => SubsyncVadPaths;
  canPickReferenceFile?: () => boolean;
  fetchImpl?: typeof fetch;
  /** Aborts the run (media changed); kills spawned tools and stops between stages. */
  signal?: AbortSignal;
}

/** Network and cancellation handles threaded into extraction helpers. */
export type SubsyncIo = Pick<SubsyncCoreDeps, 'fetchImpl' | 'signal'>;

export interface TriggerSubsyncFromConfigDeps extends SubsyncCoreDeps {
  isSubsyncInProgress: () => boolean;
  setSubsyncInProgress: (inProgress: boolean) => void;
  showMpvOsd: (text: string) => void;
  runWithSubsyncSpinner: <T>(task: () => Promise<T>) => Promise<T>;
  openManualPicker: (payload: SubsyncManualPayload) => void;
}

export function getMpvClientForSubsync(deps: SubsyncCoreDeps): MpvClientLike {
  const client = deps.getMpvClient();
  if (!client || !client.connected) {
    throw new Error('MPV not connected');
  }
  return client;
}

export function ensureExecutablePath(pathOrName: string, name: string): string {
  if (!pathOrName) {
    throw new Error(`Missing ${name} path in config`);
  }

  if (hasPathSeparators(pathOrName) && !fileExists(pathOrName)) {
    throw new Error(`Configured ${name} executable not found: ${pathOrName}`);
  }
  return pathOrName;
}

export function isExecutableAvailable(configuredPath: string, names: readonly string[]): boolean {
  if (!configuredPath) return findSubsyncExecutable(undefined, names) !== null;
  if (hasPathSeparators(configuredPath)) return fileExists(configuredPath);
  return findSubsyncExecutable(configuredPath, names) !== null;
}

export function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export function formatOffsetSummary(offsetMs: number): string {
  const sec = offsetMs / 1000;
  const formatted = Math.abs(sec) >= 10 ? sec.toFixed(1) : sec.toFixed(2);
  return offsetMs >= 0 ? `+${formatted}s` : `${formatted}s`;
}

export function formatAlignmentSummary(sourceLabel: string, alignment: AlignmentResult): string {
  const parts = [`Synced via ${sourceLabel}`, formatOffsetSummary(alignment.offsetMs)];
  if (Math.abs(alignment.ratio - 1.0) > 1e-4) {
    parts.push(`speed ${alignment.ratio.toFixed(4)}x`);
  }
  if (alignment.segments > 1) {
    parts.push(`${alignment.segments} segments`);
  }
  const pct = Math.round(alignment.confidence * 100);
  if (pct > 0) {
    parts.push(`${pct}%`);
  }
  return parts.join(' · ');
}

export function isAlreadyInSync(alignment: AlignmentResult): boolean {
  return (
    alignment.confident &&
    !alignment.cut &&
    alignment.segments === 1 &&
    Math.abs(alignment.ratio - 1.0) < 1e-4 &&
    Math.abs(alignment.offsetMs) <= ALREADY_IN_SYNC_MS
  );
}

export interface CandidateWithSpans {
  evaluated: EvaluatedSubtitleReference;
  dialogueSpans: SubtitleTimeSpan[];
  allSpans: SubtitleTimeSpan[];
}

export const PLAYBACK_CHANGED_RESULT_MESSAGE =
  'Playback changed while syncing; discarded stale result.';

export const STALE_AUTO_RESULT_MESSAGE =
  'Subsync: playback changed while syncing; discarded result';
