import type { MediaGenerator } from '../../../media-generator.js';
import type { AnkiConnectConfig } from '../../../types.js';
import { createLogger } from '../../../logger.js';
import type { RetimedSecondarySubtitleInput } from '../secondary-subtitle-sidecar.js';

export type StatsServerNoteInfo = {
  noteId: number;
  fields: Record<string, { value: string }>;
};

export type StatsServerMediaGenerator = {
  generateAudio: (...args: Parameters<MediaGenerator['generateAudio']>) => Promise<Buffer | null>;
  generateScreenshot: (
    ...args: Parameters<MediaGenerator['generateScreenshot']>
  ) => Promise<Buffer | null>;
  generateAnimatedImage: (
    ...args: Parameters<MediaGenerator['generateAnimatedImage']>
  ) => Promise<Buffer | null>;
};

export type StatsMiningTimingEvent = {
  mode: 'word' | 'sentence' | 'audio';
  phase: string;
  elapsedMs: number;
  noteId?: number;
};

export type StatsMiningRouteOptions = {
  ankiConnectConfig?: AnkiConnectConfig;
  getAnkiConnectConfig?: () => AnkiConnectConfig | undefined;
  getYomitanAnkiDeckName?: () => Promise<string | null | undefined> | string | null | undefined;
  secondarySubtitleLanguages?: string[];
  getSecondarySubtitleLanguages?: () => string[] | undefined;
  statsMiningAlassPath?: string;
  getStatsMiningAlassPath?: () => string | null | undefined;
  resolveRetimedSecondarySubtitleText?: (
    input: RetimedSecondarySubtitleInput,
  ) => Promise<string> | string;
  addYomitanNote?: (word: string) => Promise<number | null>;
  createMediaGenerator?: () => StatsServerMediaGenerator;
  onMiningTiming?: (event: StatsMiningTimingEvent) => void;
  nowMs?: () => number;
};

type StatsMiningLogger = {
  debug: (message: string, ...meta: unknown[]) => void;
  warn: (message: string, ...meta: unknown[]) => void;
};

export const statsMiningLogger: StatsMiningLogger = createLogger('stats:mining');

export function resolveStatsNoteFieldName(
  noteInfo: StatsServerNoteInfo,
  ...preferredNames: (string | undefined)[]
): string | null {
  for (const preferredName of preferredNames) {
    if (!preferredName) continue;
    const resolved = Object.keys(noteInfo.fields).find(
      (fieldName) => fieldName.toLowerCase() === preferredName.toLowerCase(),
    );
    if (resolved) return resolved;
  }
  return null;
}

function uniqueFieldNames(...fieldNames: (string | null | undefined)[]): string[] {
  const seen = new Set<string>();
  const result: string[] = [];
  for (const fieldName of fieldNames) {
    const normalized = fieldName?.trim();
    if (!normalized) continue;
    const key = normalized.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    result.push(normalized);
  }
  return result;
}

export function getStatsWordMiningAudioFieldName(
  ankiConfig: AnkiConnectConfig,
  noteInfo: StatsServerNoteInfo | null,
): string {
  return (
    (noteInfo
      ? resolveStatsNoteFieldName(noteInfo, 'SentenceAudio', ankiConfig.fields?.audio)
      : null) ??
    ankiConfig.fields?.audio ??
    'ExpressionAudio'
  );
}

export function shouldUseStatsLapisKikuCardFields(ankiConfig: AnkiConnectConfig): boolean {
  return ankiConfig.isLapis?.enabled === true || ankiConfig.isKiku?.enabled === true;
}

export function applyStatsWordAndSentenceCardFields(
  fields: Record<string, string>,
  noteInfo: StatsServerNoteInfo | null,
  ankiConfig: AnkiConnectConfig,
): void {
  if (!shouldUseStatsLapisKikuCardFields(ankiConfig) || !noteInfo) return;
  const wordAndSentenceFlag = resolveStatsNoteFieldName(noteInfo, 'IsWordAndSentenceCard');
  if (!wordAndSentenceFlag) return;

  fields[wordAndSentenceFlag] = 'x';
  for (const flagName of ['IsSentenceCard', 'IsAudioCard']) {
    const resolved = resolveStatsNoteFieldName(noteInfo, flagName);
    if (resolved && resolved !== wordAndSentenceFlag) fields[resolved] = '';
  }
}

export function getStatsDirectMiningAudioFieldNames(
  ankiConfig: AnkiConnectConfig,
  noteInfo: StatsServerNoteInfo | null,
  mode: 'sentence' | 'audio',
): string[] {
  const configuredAudioField = ankiConfig.fields?.audio ?? 'ExpressionAudio';
  if (!ankiConfig.isLapis?.enabled && !ankiConfig.isKiku?.enabled) {
    return [configuredAudioField];
  }
  const sentenceAudioField = noteInfo
    ? resolveStatsNoteFieldName(noteInfo, 'SentenceAudio', configuredAudioField)
    : 'SentenceAudio';
  const expressionAudioField = noteInfo
    ? resolveStatsNoteFieldName(noteInfo, configuredAudioField)
    : configuredAudioField;
  return mode === 'sentence'
    ? uniqueFieldNames(sentenceAudioField)
    : uniqueFieldNames(sentenceAudioField, expressionAudioField);
}

export function createStatsMiningContext(options?: StatsMiningRouteOptions) {
  const nowMs = options?.nowMs ?? (() => Date.now());
  const getAnkiConnectConfig = (): AnkiConnectConfig | undefined =>
    options?.getAnkiConnectConfig?.() ?? options?.ankiConnectConfig;
  const getSecondarySubtitleLanguages = (): string[] =>
    options?.getSecondarySubtitleLanguages?.() ?? options?.secondarySubtitleLanguages ?? [];
  const getStatsMiningAlassPath = (): string | null | undefined =>
    options?.getStatsMiningAlassPath?.() ?? options?.statsMiningAlassPath;
  const getEffectiveMiningDeckName = async (ankiConfig: AnkiConnectConfig): Promise<string> => {
    const configuredDeckName = ankiConfig.deck?.trim() ?? '';
    if (configuredDeckName) return configuredDeckName;
    try {
      const yomitanDeckName = await options?.getYomitanAnkiDeckName?.();
      return typeof yomitanDeckName === 'string' ? yomitanDeckName.trim() : '';
    } catch (error) {
      statsMiningLogger.warn(
        'Failed to resolve Yomitan Anki deck for stats mining:',
        error instanceof Error ? error.message : String(error),
      );
      return '';
    }
  };
  const recordMiningTiming = (event: StatsMiningTimingEvent): void => {
    try {
      options?.onMiningTiming?.(event);
    } catch {
      // Timing observers must not affect mining execution.
    }
    statsMiningLogger.debug(
      `[stats:mining] ${event.mode} ${event.phase} ${Math.round(event.elapsedMs)}ms`,
      event,
    );
  };
  const timeMiningPhase = async <T>(
    mode: StatsMiningTimingEvent['mode'],
    phase: string,
    fn: () => Promise<T>,
    details?: (value: T) => Partial<StatsMiningTimingEvent>,
  ): Promise<T> => {
    const startedAtMs = nowMs();
    try {
      const value = await fn();
      recordMiningTiming({
        mode,
        phase,
        elapsedMs: nowMs() - startedAtMs,
        ...details?.(value),
      });
      return value;
    } catch (error) {
      recordMiningTiming({ mode, phase, elapsedMs: nowMs() - startedAtMs });
      throw error;
    }
  };

  return {
    getAnkiConnectConfig,
    getEffectiveMiningDeckName,
    getSecondarySubtitleLanguages,
    getStatsMiningAlassPath,
    timeMiningPhase,
  };
}
