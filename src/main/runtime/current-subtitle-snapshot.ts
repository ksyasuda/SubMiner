import type { SubtitleData } from '../../types';

export async function resolveCurrentSubtitleForRenderer(deps: {
  currentSubText: string;
  currentSubtitleData: SubtitleData | null;
  withCurrentSubtitleTiming: (payload: SubtitleData) => SubtitleData;
  tokenizeSubtitle?: (text: string) => Promise<SubtitleData | null>;
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

  const tokenized = await deps.tokenizeSubtitle?.(deps.currentSubText);
  if (tokenized) {
    return deps.withCurrentSubtitleTiming(tokenized);
  }

  return deps.withCurrentSubtitleTiming({
    text: deps.currentSubText,
    tokens: null,
  });
}
