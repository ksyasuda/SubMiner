export function handleCharacterDictionaryAutoSyncComplete(
  completion: {
    mediaId: number;
    mediaTitle: string;
    changed: boolean;
  },
  deps: {
    hasParserWindow: () => boolean;
    clearParserCaches: () => void;
    invalidateTokenizationCache: () => void;
    refreshSubtitlePrefetch: () => void;
    refreshCurrentSubtitle: () => void;
    logInfo: (message: string) => void;
  },
): void {
  if (completion.changed) {
    if (deps.hasParserWindow()) {
      deps.clearParserCaches();
    }
    deps.invalidateTokenizationCache();
    deps.refreshSubtitlePrefetch();
    deps.refreshCurrentSubtitle();
  }
  deps.logInfo(
    `[dictionary:auto-sync] refreshed current subtitle after sync (AniList ${completion.mediaId}, changed=${completion.changed ? 'yes' : 'no'}, title=${completion.mediaTitle})`,
  );
}
