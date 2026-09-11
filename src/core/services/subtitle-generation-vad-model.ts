import { access, stat } from 'node:fs/promises';
import { constants } from 'node:fs';
import path from 'node:path';
import type {
  SubtitleGenerationConfig,
  SubtitleGenerationModelStatus,
  SubtitleGenerationProgress,
} from '../../shared/subtitle-generation';
import { SUBTITLE_GENERATION_VAD_MODEL } from '../../shared/subtitle-generation-vad-model';
import { expandSubtitleGenerationPath } from './subtitle-generation-files';
import { isMissingFile } from './subtitle-generation-models';
import { downloadSubtitleGenerationArtifact } from './subtitle-generation-download';

export async function resolveSubtitleGenerationVadModel(
  config: SubtitleGenerationConfig,
  modelDirectory: string,
): Promise<SubtitleGenerationModelStatus> {
  const external = config.vadModelPath.trim();
  const modelPath = external
    ? path.resolve(expandSubtitleGenerationPath(external))
    : path.resolve(modelDirectory, SUBTITLE_GENERATION_VAD_MODEL.filename);
  try {
    const info = await stat(modelPath);
    if (
      !info.isFile() ||
      info.size === 0 ||
      (!external && info.size !== SUBTITLE_GENERATION_VAD_MODEL.size)
    )
      return {
        kind: 'invalid',
        path: modelPath,
        message: 'Speech detection model has an invalid size.',
      };
    await access(modelPath, constants.R_OK);
    return { kind: external ? 'external' : 'managed', path: modelPath };
  } catch (error) {
    if (!external && isMissingFile(error)) return { kind: 'missing', path: modelPath };
    return {
      kind: 'invalid',
      path: modelPath,
      message: `Cannot read speech detection model: ${error instanceof Error ? error.message : String(error)}`,
    };
  }
}

export async function downloadSubtitleGenerationVadModel(input: {
  config: SubtitleGenerationConfig;
  modelDirectory: string;
  onProgress?: (progress: SubtitleGenerationProgress) => void;
  signal?: AbortSignal;
}): Promise<string> {
  input.signal?.throwIfAborted();
  const current = await resolveSubtitleGenerationVadModel(input.config, input.modelDirectory);
  if (current.kind === 'invalid') throw new Error(current.message);
  if (current.kind !== 'missing') return current.path;
  return downloadSubtitleGenerationArtifact({
    ...SUBTITLE_GENERATION_VAD_MODEL,
    destination: current.path,
    label: 'Silero speech detection model',
    onProgress: input.onProgress,
    signal: input.signal,
  });
}
