export interface StartupOsdSequencerCharacterDictionaryEvent {
  phase: 'checking' | 'generating' | 'syncing' | 'building' | 'importing' | 'ready' | 'failed';
  message: string;
}

export function createStartupOsdSequencer(deps: { showOsd: (message: string) => boolean | void }): {
  reset: () => void;
  markTokenizationReady: () => void;
  showAnnotationLoading: (message: string) => void;
  markAnnotationLoadingComplete: (message: string) => void;
  notifyCharacterDictionaryStatus: (event: StartupOsdSequencerCharacterDictionaryEvent) => boolean;
} {
  let tokenizationReady = false;
  let tokenizationWarmupCompleted = false;
  let annotationLoadingMessage: string | null = null;
  let pendingDictionaryProgress: StartupOsdSequencerCharacterDictionaryEvent | null = null;
  let pendingDictionaryFailure: StartupOsdSequencerCharacterDictionaryEvent | null = null;
  let pendingDictionaryReady: StartupOsdSequencerCharacterDictionaryEvent | null = null;
  let dictionaryProgressShown = false;

  const canShowDictionaryStatus = (): boolean =>
    tokenizationReady && annotationLoadingMessage === null;
  const showOsd = (message: string): boolean => deps.showOsd(message) !== false;

  const flushBufferedDictionaryStatus = (): boolean => {
    if (!canShowDictionaryStatus()) {
      return false;
    }
    if (pendingDictionaryProgress) {
      if (dictionaryProgressShown) {
        return true;
      }
      dictionaryProgressShown = showOsd(pendingDictionaryProgress.message);
      return dictionaryProgressShown;
    }
    if (pendingDictionaryReady) {
      const shown = showOsd(pendingDictionaryReady.message);
      if (shown) {
        pendingDictionaryReady = null;
        dictionaryProgressShown = false;
      }
      return shown;
    }
    if (pendingDictionaryFailure) {
      const shown = showOsd(pendingDictionaryFailure.message);
      if (shown) {
        pendingDictionaryFailure = null;
        dictionaryProgressShown = false;
      }
      return shown;
    }
    return false;
  };

  return {
    reset: () => {
      tokenizationReady = tokenizationWarmupCompleted;
      annotationLoadingMessage = null;
      pendingDictionaryProgress = null;
      pendingDictionaryFailure = null;
      pendingDictionaryReady = null;
      dictionaryProgressShown = false;
    },
    markTokenizationReady: () => {
      tokenizationWarmupCompleted = true;
      tokenizationReady = true;
      if (annotationLoadingMessage !== null) {
        showOsd(annotationLoadingMessage);
        return;
      }
      flushBufferedDictionaryStatus();
    },
    showAnnotationLoading: (message) => {
      annotationLoadingMessage = message;
      if (tokenizationReady) {
        showOsd(message);
      }
    },
    markAnnotationLoadingComplete: (message) => {
      annotationLoadingMessage = null;
      if (!tokenizationReady) {
        return;
      }
      if (flushBufferedDictionaryStatus()) {
        return;
      }
      showOsd(message);
    },
    notifyCharacterDictionaryStatus: (event) => {
      if (
        event.phase === 'checking' ||
        event.phase === 'generating' ||
        event.phase === 'syncing' ||
        event.phase === 'building' ||
        event.phase === 'importing'
      ) {
        pendingDictionaryProgress = event;
        pendingDictionaryFailure = null;
        pendingDictionaryReady = null;
        if (canShowDictionaryStatus()) {
          dictionaryProgressShown = showOsd(event.message);
        } else if (tokenizationReady) {
          dictionaryProgressShown = showOsd(event.message);
        }
        return dictionaryProgressShown;
      }

      pendingDictionaryProgress = null;
      if (event.phase === 'failed') {
        pendingDictionaryReady = null;
        if (canShowDictionaryStatus()) {
          if (!showOsd(event.message)) {
            pendingDictionaryFailure = event;
            return false;
          }
          dictionaryProgressShown = false;
          return true;
        } else {
          pendingDictionaryFailure = event;
        }
        dictionaryProgressShown = false;
        return false;
      }

      pendingDictionaryFailure = null;
      if (canShowDictionaryStatus()) {
        if (!showOsd(event.message)) {
          pendingDictionaryReady = event;
          dictionaryProgressShown = false;
          return false;
        }
        pendingDictionaryReady = null;
        dictionaryProgressShown = false;
        return true;
      } else {
        pendingDictionaryReady = event;
      }
      dictionaryProgressShown = false;
      return false;
    },
  };
}
