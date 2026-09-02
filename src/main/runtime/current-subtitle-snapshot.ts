import type { SubtitleData } from '../../types';

type CurrentSubtitleMpvClient = {
  connected?: boolean;
  requestProperty: (name: string) => Promise<unknown>;
};

export async function resolveCurrentSubtitleForRenderer(deps: {
  currentSubText: string;
  currentSubtitleData: SubtitleData | null;
  withCurrentSubtitleTiming: (payload: SubtitleData) => SubtitleData;
  tokenizeSubtitle?: (text: string) => Promise<SubtitleData | null>;
  tokenizeUncached?: boolean;
  onResolvedSubtitle?: (payload: SubtitleData) => void;
}): Promise<SubtitleData> {
  const resolve = (payload: SubtitleData): SubtitleData => {
    const timedPayload = deps.withCurrentSubtitleTiming(payload);
    deps.onResolvedSubtitle?.(timedPayload);
    return timedPayload;
  };

  if (deps.currentSubtitleData?.text === deps.currentSubText) {
    return resolve(deps.currentSubtitleData);
  }

  if (!deps.currentSubText.trim()) {
    return resolve({
      text: deps.currentSubText,
      tokens: null,
    });
  }

  if (deps.tokenizeUncached !== false) {
    const tokenized = await deps.tokenizeSubtitle?.(deps.currentSubText);
    if (tokenized) {
      return resolve(tokenized);
    }
  }

  return resolve({
    text: deps.currentSubText,
    tokens: null,
  });
}

export async function primeVisibleOverlaySubtitleFromMpv(deps: {
  getMpvClient: () => CurrentSubtitleMpvClient | null;
  setCurrentSubText: (text: string) => void;
  resolvePrimarySubtitleText?: (text: string) => string;
  getCurrentSubtitleData: () => SubtitleData | null;
  consumeCachedSubtitle: (text: string) => SubtitleData | null;
  onSubtitleChange: (text: string) => void;
  refreshCurrentSubtitle: (text: string) => void;
  emitSubtitle: (payload: SubtitleData) => void;
  deferUncachedRefresh?: boolean;
  setCurrentSecondarySubText?: (text: string) => void;
  emitSecondarySubtitle?: (text: string) => void;
  logDebug?: (message: string) => void;
}): Promise<void> {
  const client = deps.getMpvClient();
  if (!client?.connected) {
    return;
  }

  let subTextRaw: unknown;
  try {
    subTextRaw = await client.requestProperty('sub-text');
  } catch (error) {
    deps.logDebug?.(
      `[visible-overlay-subtitle-prime] failed to read sub-text: ${
        error instanceof Error ? error.message : String(error)
      }`,
    );
    return;
  }

  const liveText = typeof subTextRaw === 'string' ? subTextRaw : '';
  const text = deps.resolvePrimarySubtitleText?.(liveText) ?? liveText;
  deps.setCurrentSubText(text);

  const primeSecondarySubtitle = async (): Promise<void> => {
    if (!deps.setCurrentSecondarySubText && !deps.emitSecondarySubtitle) {
      return;
    }

    try {
      const secondarySubTextRaw = await client.requestProperty('secondary-sub-text');
      const secondaryText = typeof secondarySubTextRaw === 'string' ? secondarySubTextRaw : '';
      deps.setCurrentSecondarySubText?.(secondaryText);
      deps.emitSecondarySubtitle?.(secondaryText);
    } catch (error) {
      deps.logDebug?.(
        `[visible-overlay-subtitle-prime] failed to read secondary-sub-text: ${
          error instanceof Error ? error.message : String(error)
        }`,
      );
    }
  };

  if (!text.trim()) {
    deps.onSubtitleChange(text);
    deps.emitSubtitle({ text, tokens: null });
    await primeSecondarySubtitle();
    return;
  }

  const currentPayload = deps.getCurrentSubtitleData();
  if (currentPayload?.text === text) {
    deps.emitSubtitle(currentPayload);
    deps.refreshCurrentSubtitle(text);
    await primeSecondarySubtitle();
    return;
  }

  const cachedPayload = deps.consumeCachedSubtitle(text);
  if (cachedPayload) {
    deps.onSubtitleChange(text);
    deps.emitSubtitle(cachedPayload);
    await primeSecondarySubtitle();
    return;
  }

  if (deps.deferUncachedRefresh === true) {
    await primeSecondarySubtitle();
    return;
  }

  deps.refreshCurrentSubtitle(text);
  await primeSecondarySubtitle();
}
