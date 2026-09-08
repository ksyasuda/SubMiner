import {
  isSubtitleGenerationModelId,
  RECOMMENDED_SUBTITLE_GENERATION_MODEL,
  type SubtitleGenerationModelId,
} from './subtitle-generation-model-catalog';

export interface SubtitleGenerationConfig {
  whisperPath: string;
  modelPath: string;
  managedModel: SubtitleGenerationModelId;
  threads: number;
  ffmpegPath: string;
  ffprobePath: string;
  vadModelPath: string;
  vadPath: string;
}

export const DEFAULT_SUBTITLE_GENERATION_CONFIG: SubtitleGenerationConfig = {
  whisperPath: '',
  modelPath: '',
  managedModel: RECOMMENDED_SUBTITLE_GENERATION_MODEL,
  threads: 4,
  ffmpegPath: '',
  ffprobePath: '',
  vadModelPath: '',
  vadPath: '',
};

export interface SubtitleGenerationProgress {
  stage: 'download' | 'extract' | 'transcribe' | 'write';
  percent?: number;
  message: string;
}

export type SubtitleGenerationModelStatus =
  | { kind: 'external' | 'managed'; path: string }
  | { kind: 'missing'; path: string }
  | { kind: 'invalid'; path: string; message: string };

export function resolveSubtitleGenerationConfig(
  value: unknown,
  onWarning?: (key: string, value: unknown, message: string) => void,
): SubtitleGenerationConfig {
  const result = { ...DEFAULT_SUBTITLE_GENERATION_CONFIG };
  if (value === undefined) return result;
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    onWarning?.('subtitleGeneration', value, 'Expected an object.');
    return result;
  }
  for (const [key, rawValue] of Object.entries(value)) {
    if (
      key !== 'whisperPath' &&
      key !== 'modelPath' &&
      key !== 'ffmpegPath' &&
      key !== 'ffprobePath' &&
      key !== 'vadModelPath' &&
      key !== 'vadPath'
    )
      continue;
    const candidate: unknown = rawValue;
    if (typeof candidate === 'string') {
      result[key] = candidate.trim();
    } else onWarning?.(key, candidate, 'Expected a string.');
  }
  if ('managedModel' in value) {
    if (isSubtitleGenerationModelId(value.managedModel)) {
      result.managedModel = value.managedModel;
    } else
      onWarning?.(
        'managedModel',
        value.managedModel,
        'Expected a supported multilingual Whisper model.',
      );
  }
  if ('threads' in value) {
    if (
      typeof value.threads === 'number' &&
      Number.isInteger(value.threads) &&
      value.threads >= 1 &&
      value.threads <= 256
    ) {
      result.threads = value.threads;
    } else onWarning?.('threads', value.threads, 'Expected an integer between 1 and 256.');
  }
  return result;
}
