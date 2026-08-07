export function handleCharacterDictionaryAutoSyncComplete(
  completion: {
    mediaId: number;
    mediaTitle: string;
    changed: boolean;
  },
  deps: {
    hasParserWindow: () => boolean;
    clearParserCaches: () => void;
    /**
     * Drops cached reads of the generated dictionary (character images, and the
     * name candidates the scanner uses to skip lookups). Runs before the
     * refreshes below so they re-tokenize against the new dictionary content.
     */
    invalidateCharacterDictionaryLookups?: () => void;
    invalidateTokenizationCache: () => void;
    refreshSubtitlePrefetch: () => void;
    refreshCurrentSubtitle: () => void;
    logInfo: (message: string) => void;
  },
): void {
  if (completion.changed) {
    deps.invalidateCharacterDictionaryLookups?.();
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
