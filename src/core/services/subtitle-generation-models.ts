import { access, open, stat } from 'node:fs/promises';
import { constants } from 'node:fs';
import path from 'node:path';
import { getSubtitleGenerationModel } from '../../shared/subtitle-generation-model-catalog';
import { expandSubtitleGenerationPath } from './subtitle-generation-files';
import { downloadSubtitleGenerationArtifact } from './subtitle-generation-download';
import type {
  SubtitleGenerationConfig,
  SubtitleGenerationModelStatus,
  SubtitleGenerationProgress,
} from '../../shared/subtitle-generation';

export function isMissingFile(error: unknown): boolean {
  return error instanceof Error && 'code' in error && error.code === 'ENOENT';
}

async function modelCompatibilityError(modelPath: string): Promise<string | undefined> {
  const file = await open(modelPath, 'r');
  try {
    // whisper_model_load reads GGML magic, then n_vocab. is_multilingual uses n_vocab >= 51865.
    const header = Buffer.alloc(8);
    const { bytesRead } = await file.read(header, 0, header.length, 0);
    if (
      bytesRead !== header.length ||
      header.readUInt32LE(0) !== 0x67676d6c ||
      header.readInt32LE(4) <= 0
    ) {
      return 'Unsupported model format. Choose a whisper.cpp GGML .bin model.';
    }
    if (header.readInt32LE(4) < 51865) {
      return 'This Whisper model is English-only. Japanese subtitle generation requires a multilingual model.';
    }
    return undefined;
  } finally {
    await file.close();
  }
}

export async function resolveSubtitleGenerationModel(
  config: SubtitleGenerationConfig,
  modelDirectory: string,
): Promise<SubtitleGenerationModelStatus> {
  const external = config.modelPath.trim();
  const modelPath = external
    ? path.resolve(expandSubtitleGenerationPath(external))
    : path.resolve(modelDirectory, `ggml-${config.managedModel}.bin`);
  try {
    const info = await stat(modelPath);
    if (!info.isFile() || info.size === 0) {
      return { kind: 'invalid', path: modelPath, message: 'Model must be a nonempty file.' };
    }
    await access(modelPath, constants.R_OK);
    if (!external && info.size !== getSubtitleGenerationModel(config.managedModel).size) {
      return { kind: 'invalid', path: modelPath, message: 'Managed model has an unexpected size.' };
    }
    const compatibilityError = await modelCompatibilityError(modelPath);
    if (compatibilityError)
      return { kind: 'invalid', path: modelPath, message: compatibilityError };
    return { kind: external ? 'external' : 'managed', path: modelPath };
  } catch (error) {
    if (!external && isMissingFile(error)) return { kind: 'missing', path: modelPath };
    return {
      kind: 'invalid',
      path: modelPath,
      message: `Cannot read model: ${error instanceof Error ? error.message : String(error)}`,
    };
  }
}

export async function downloadSubtitleGenerationModel(input: {
  config: SubtitleGenerationConfig;
  modelDirectory: string;
  onProgress?: (progress: SubtitleGenerationProgress) => void;
  signal?: AbortSignal;
}): Promise<string> {
  input.signal?.throwIfAborted();
  const current = await resolveSubtitleGenerationModel(input.config, input.modelDirectory);
  if (current.kind === 'external' || current.kind === 'managed') return current.path;
  if (current.kind === 'invalid') throw new Error(current.message);
  const model = getSubtitleGenerationModel(input.config.managedModel);
  return downloadSubtitleGenerationArtifact({
    url: `https://huggingface.co/ggerganov/whisper.cpp/resolve/5359861c739e955e79d9a303bcbc70fb988958b1/ggml-${input.config.managedModel}.bin`,
    size: model.size,
    sha256: model.sha256,
    destination: current.path,
    label: 'Whisper model',
    onProgress: input.onProgress,
    signal: input.signal,
  });
}
