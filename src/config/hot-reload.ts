function pathStartsWith(path: string, prefix: string): boolean {
  return path === prefix || path.startsWith(`${prefix}.`);
}

const HOT_RELOAD_ROOTS = ['subtitleStyle', 'keybindings', 'shortcuts', 'subtitleSidebar'] as const;

const HOT_RELOAD_EXACT_OR_PREFIX_PATHS = [
  'secondarySub.defaultMode',
  'mpv.aniskipEnabled',
  'mpv.aniskipButtonKey',
  'ankiConnect.ai.enabled',
  'stats.toggleKey',
  'stats.markWatchedKey',
  'logging.level',
  'logging.rotation',
  'logging.files',
  'youtube.primarySubLanguages',
  'ankiConnect.deck',
  'ankiConnect.media.normalizeAudio',
  'ankiConnect.media.mirrorMpvVolume',
  'ankiConnect.media.reviewTiming',
  'ankiConnect.behavior.autoUpdateNewCards',
  'ankiConnect.knownWords.highlightEnabled',
  'ankiConnect.knownWords.refreshMinutes',
  'ankiConnect.knownWords.addMinedWordsImmediately',
  'ankiConnect.knownWords.matchMode',
  'ankiConnect.knownWords.decks',
  'ankiConnect.nPlusOne.enabled',
  'ankiConnect.nPlusOne.minSentenceWords',
  'ankiConnect.fields.word',
  'ankiConnect.fields.audio',
  'ankiConnect.fields.image',
  'ankiConnect.fields.sentence',
  'ankiConnect.fields.miscInfo',
  'ankiConnect.isLapis.sentenceCardModel',
  'ankiConnect.isKiku.fieldGrouping',
  'ankiConnect.isSenren.fieldGrouping',
  'ankiConnect.lapisKiku.wordCardKind',
] as const;

export function getConfigHotReloadField(path: string): string | null {
  for (const root of HOT_RELOAD_ROOTS) {
    if (pathStartsWith(path, root)) {
      return root;
    }
  }

  for (const hotPath of HOT_RELOAD_EXACT_OR_PREFIX_PATHS) {
    if (pathStartsWith(path, hotPath)) {
      return hotPath;
    }
  }

  // These consumers read the current config when the next operation starts.
  if (
    ['jimaku', 'subsync', 'notifications', 'subtitleGeneration'].some((root) =>
      pathStartsWith(path, root),
    )
  ) {
    return path;
  }

  return null;
}
