export interface StartupOsdSequencerCharacterDictionaryEvent {
  phase: 'checking' | 'generating' | 'syncing' | 'building' | 'importing' | 'ready' | 'failed';
  message: string;
}

export function createStartupOsdSequencer(deps: { showOsd: (message: string) => void }): {
  reset: () => void;
  markTokenizationReady: () => void;
  showAnnotationLoading: (message: string) => void;
  markAnnotationLoadingComplete: (message: string) => void;
  notifyCharacterDictionaryStatus: (event: StartupOsdSequencerCharacterDictionaryEvent) => void;
} {
  let tokenizationReady = false;
  let annotationLoadingMessage: string | null = null;
  let pendingDictionaryProgress: StartupOsdSequencerCharacterDictionaryEvent | null = null;
  let pendingDictionaryFailure: StartupOsdSequencerCharacterDictionaryEvent | null = null;
  let dictionaryProgressShown = false;

  const canShowDictionaryStatus = (): boolean =>
    tokenizationReady && annotationLoadingMessage === null;

  const flushBufferedDictionaryStatus = (): boolean => {
    if (!canShowDictionaryStatus()) {
      return false;
    }
    if (pendingDictionaryProgress) {
      deps.showOsd(pendingDictionaryProgress.message);
      dictionaryProgressShown = true;
      return true;
    }
    if (pendingDictionaryFailure) {
      deps.showOsd(pendingDictionaryFailure.message);
      pendingDictionaryFailure = null;
      dictionaryProgressShown = false;
      return true;
    }
    return false;
  };

  return {
    reset: () => {
      tokenizationReady = false;
      annotationLoadingMessage = null;
      pendingDictionaryProgress = null;
      pendingDictionaryFailure = null;
      dictionaryProgressShown = false;
    },
    markTokenizationReady: () => {
      tokenizationReady = true;
      if (annotationLoadingMessage !== null) {
        deps.showOsd(annotationLoadingMessage);
        return;
      }
      flushBufferedDictionaryStatus();
    },
    showAnnotationLoading: (message) => {
      annotationLoadingMessage = message;
      if (tokenizationReady) {
        deps.showOsd(message);
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
      deps.showOsd(message);
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
        if (canShowDictionaryStatus()) {
          deps.showOsd(event.message);
          dictionaryProgressShown = true;
        }
        return;
      }

      pendingDictionaryProgress = null;
      if (event.phase === 'failed') {
        if (canShowDictionaryStatus()) {
          deps.showOsd(event.message);
        } else {
          pendingDictionaryFailure = event;
        }
        dictionaryProgressShown = false;
        return;
      }

      pendingDictionaryFailure = null;
      if (canShowDictionaryStatus() && dictionaryProgressShown) {
        deps.showOsd(event.message);
      }
      dictionaryProgressShown = false;
    },
  };
}
