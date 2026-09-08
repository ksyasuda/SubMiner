import type {
  SubtitleGenerationModelStatus,
  SubtitleGenerationProgress,
} from '../../shared/subtitle-generation';
import type { SubtitleGenerationStatus } from '../../shared/subtitle-generation-ipc';

export function describeGenerationVad(vad: SubtitleGenerationStatus['vad']) {
  const model = describeGenerationModel(vad.model);
  return {
    ready: !vad.enabled || model.ready,
    download: vad.enabled && model.download,
    text: !vad.enabled
      ? 'Optional. Generate from the full audio when unchecked.'
      : vad.model.kind === 'missing'
        ? 'Download Silero to prioritize spoken dialogue.'
        : vad.model.kind === 'invalid'
          ? vad.model.message
          : vad.model.kind === 'external'
            ? `Your speech detection model: ${vad.model.path}`
            : 'Silero speech detection model installed.',
  };
}

export function describeGenerationModel(model: SubtitleGenerationModelStatus) {
  switch (model.kind) {
    case 'external':
      return { ready: true, download: false, text: `Your model: ${model.path}` };
    case 'managed':
      return { ready: true, download: false, text: `SubMiner model: ${model.path}` };
    case 'missing':
      return { ready: false, download: true, text: 'Download a speech model to get started.' };
    case 'invalid':
      return { ready: false, download: false, text: model.message };
    default: {
      const exhaustive: never = model;
      return exhaustive;
    }
  }
}

const STAGE_LABELS = {
  download: 'Downloading speech model',
  extract: 'Preparing audio',
  transcribe: 'Recognizing Japanese speech',
  write: 'Saving subtitles',
} satisfies Record<SubtitleGenerationProgress['stage'], string>;

export function describeGenerationProgress(progress: SubtitleGenerationProgress | null) {
  const raw = progress?.percent;
  const percent =
    raw !== undefined && Number.isFinite(raw) ? Math.max(0, Math.min(100, raw)) : null;
  return {
    stage: progress ? STAGE_LABELS[progress.stage] : 'Preparing',
    percent,
    label: percent === null ? 'Working...' : `${Math.floor(percent)}%`,
  };
}
