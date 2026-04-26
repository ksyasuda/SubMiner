import type { CharacterDictionaryManualSelectionResult } from './character-dictionary-runtime/types';

export type CharacterDictionarySelectionRequest = {
  targetPath?: string;
  mediaId: number;
};

export type CharacterDictionarySelectionDeps = {
  setManualSelection: (
    request: CharacterDictionarySelectionRequest,
  ) => Promise<CharacterDictionaryManualSelectionResult>;
  resetAnilistMediaGuessState: () => void;
  runSyncNow: () => Promise<void>;
  warn: (message: string, error?: unknown) => void;
};

export async function applyCharacterDictionarySelection(
  request: CharacterDictionarySelectionRequest,
  deps: CharacterDictionarySelectionDeps,
): Promise<CharacterDictionaryManualSelectionResult> {
  const result = await deps.setManualSelection(request);
  deps.resetAnilistMediaGuessState();
  try {
    await deps.runSyncNow();
  } catch (error) {
    deps.warn('Character dictionary auto-sync failed after manual selection', error);
  }
  return result;
}
