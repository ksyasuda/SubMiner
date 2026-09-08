import type {
  SubtitleGenerationConfig,
  SubtitleGenerationModelStatus,
  SubtitleGenerationProgress,
} from './subtitle-generation';

export type SubtitleGenerationResult =
  | { ok: true; message: string; outputPath?: string }
  | { ok: false; message: string };

export interface SubtitleGenerationStatus {
  model: SubtitleGenerationModelStatus;
  vad: { enabled: boolean; model: SubtitleGenerationModelStatus };
  managedModel: SubtitleGenerationConfig['managedModel'];
  externalModelPath: string | null;
  mediaPath: string | null;
  running: boolean;
  progress: SubtitleGenerationProgress | null;
  lastResult: SubtitleGenerationResult | null;
}
