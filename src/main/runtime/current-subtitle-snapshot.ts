import type { SubtitleData } from '../../types';

export async function resolveCurrentSubtitleForRenderer(deps: {
  currentSubText: string;
  currentSubtitleData: SubtitleData | null;
  withCurrentSubtitleTiming: (payload: SubtitleData) => SubtitleData;
}): Promise<SubtitleData> {
  if (deps.currentSubtitleData?.text === deps.currentSubText) {
    return deps.withCurrentSubtitleTiming(deps.currentSubtitleData);
  }

  if (!deps.currentSubText.trim()) {
    return deps.withCurrentSubtitleTiming({
      text: deps.currentSubText,
      tokens: null,
    });
  }

  return deps.withCurrentSubtitleTiming({
    text: deps.currentSubText,
    tokens: null,
  });
}
