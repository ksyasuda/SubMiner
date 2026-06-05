import type { NotificationType, OverlayNotificationPayload } from '../../types/notification';
import { shouldShowDesktop, shouldShowOverlay, shouldShowOsd } from './notification-routing';

export interface StartupOsdSequencerCharacterDictionaryEvent {
  phase: 'checking' | 'generating' | 'syncing' | 'building' | 'importing' | 'ready' | 'failed';
  message: string;
}

export interface StartupOsdSequencerDeps {
  getNotificationType?: () => NotificationType | undefined;
  showOsd: (message: string) => boolean | void;
  showOverlayNotification?: (payload: OverlayNotificationPayload) => void;
  showDesktopNotification?: (title: string, options: { body?: string }) => void;
}

interface StartupStatusNotificationOptions {
  id: string;
  title: string;
  message: string;
  variant: OverlayNotificationPayload['variant'];
  persistent: boolean;
  desktop?: boolean;
}

export function createStartupOsdSequencer(deps: StartupOsdSequencerDeps): {
  reset: () => void;
  showTokenizationLoading: (message: string) => void;
  markTokenizationReady: () => void;
  showAnnotationLoading: (message: string) => void;
  markAnnotationLoadingComplete: (message: string) => void;
  notifyCharacterDictionaryStatus: (event: StartupOsdSequencerCharacterDictionaryEvent) => boolean;
} {
  let tokenizationReady = false;
  let tokenizationWarmupCompleted = false;
  let tokenizationLoadingShown = false;
  let annotationLoadingMessage: string | null = null;
  let pendingDictionaryProgress: StartupOsdSequencerCharacterDictionaryEvent | null = null;
  let pendingDictionaryFailure: StartupOsdSequencerCharacterDictionaryEvent | null = null;
  let pendingDictionaryReady: StartupOsdSequencerCharacterDictionaryEvent | null = null;
  let dictionaryProgressShown = false;

  const canShowDictionaryStatus = (): boolean =>
    tokenizationReady && annotationLoadingMessage === null;
  const getNotificationType = (): NotificationType => deps.getNotificationType?.() ?? 'osd';
  const notifyStartupStatus = (options: StartupStatusNotificationOptions): boolean => {
    const type = getNotificationType();
    if (type === 'none') {
      return false;
    }
    let shown = false;
    if (shouldShowOverlay(type)) {
      deps.showOverlayNotification?.({
        id: options.id,
        title: options.title,
        body: options.message,
        variant: options.variant,
        persistent: options.persistent,
      });
      shown = true;
    }
    if (shouldShowOsd(type)) {
      shown = deps.showOsd(options.message) !== false || shown;
    }
    if (options.desktop !== false && shouldShowDesktop(type)) {
      deps.showDesktopNotification?.('SubMiner', { body: options.message });
      shown = true;
    }
    return shown;
  };
  const showOsd = (message: string): boolean =>
    notifyStartupStatus({
      id: 'startup-status',
      title: 'SubMiner',
      message,
      variant: 'info',
      persistent: false,
    });
  const notifyTokenization = (
    message: string,
    variant: OverlayNotificationPayload['variant'],
    persistent: boolean,
  ): boolean =>
    notifyStartupStatus({
      id: 'startup-tokenization',
      title: 'Subtitle tokenization',
      message,
      variant,
      persistent,
      desktop: !persistent,
    });
  const notifyAnnotation = (
    message: string,
    variant: OverlayNotificationPayload['variant'],
    persistent: boolean,
  ): boolean =>
    notifyStartupStatus({
      id: 'startup-subtitle-annotations',
      title: 'Subtitle annotations',
      message,
      variant,
      persistent,
      desktop: !persistent,
    });

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
      tokenizationLoadingShown = false;
      annotationLoadingMessage = null;
      pendingDictionaryProgress = null;
      pendingDictionaryFailure = null;
      pendingDictionaryReady = null;
      dictionaryProgressShown = false;
    },
    showTokenizationLoading: (message) => {
      if (tokenizationReady) {
        return;
      }
      tokenizationLoadingShown = true;
      notifyTokenization(message, 'progress', true);
    },
    markTokenizationReady: () => {
      tokenizationWarmupCompleted = true;
      tokenizationReady = true;
      if (tokenizationLoadingShown) {
        notifyTokenization('Subtitle tokenization ready', 'success', false);
        tokenizationLoadingShown = false;
      }
      if (annotationLoadingMessage !== null) {
        notifyAnnotation(annotationLoadingMessage, 'progress', true);
        return;
      }
      flushBufferedDictionaryStatus();
    },
    showAnnotationLoading: (message) => {
      annotationLoadingMessage = message;
      if (tokenizationReady) {
        notifyAnnotation(message, 'progress', true);
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
      notifyAnnotation(message, 'success', false);
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
