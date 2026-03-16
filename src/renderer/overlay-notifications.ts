import type { OverlayNotificationPayload } from '../types';

type OverlayNotificationsDom = {
  overlayNotificationToast: {
    classList: {
      add: (...tokens: string[]) => void;
      remove: (...tokens: string[]) => void;
      contains?: (token: string) => boolean;
    };
    dataset: {
      kind?: string;
    };
  };
  overlayNotificationTitle: {
    textContent: string;
  };
  overlayNotificationMessage: {
    textContent: string;
  };
  overlayNotificationSpinner: {
    classList: {
      add: (...tokens: string[]) => void;
      remove: (...tokens: string[]) => void;
    };
  };
};

type OverlayNotificationsTimerDeps = {
  durationMs?: number;
  setTimeout?: (callback: () => void, delayMs: number) => number | ReturnType<typeof setTimeout>;
  clearTimeout?: (timeout: number | ReturnType<typeof setTimeout>) => void;
};

export function createOverlayNotificationsController(
  dom: OverlayNotificationsDom,
  deps: OverlayNotificationsTimerDeps = {},
) {
  let hideTimer: number | ReturnType<typeof setTimeout> | null = null;
  const durationMs = deps.durationMs ?? 2200;
  const setTimeoutHandler = deps.setTimeout ?? setTimeout;
  const clearTimeoutHandler = deps.clearTimeout ?? clearTimeout;

  const clearHideTimer = (): void => {
    if (hideTimer === null) {
      return;
    }
    clearTimeoutHandler(hideTimer);
    hideTimer = null;
  };

  const hide = (): void => {
    clearHideTimer();
    dom.overlayNotificationToast.classList.add('hidden');
    dom.overlayNotificationSpinner.classList.add('hidden');
    dom.overlayNotificationToast.dataset.kind = '';
    dom.overlayNotificationTitle.textContent = '';
    dom.overlayNotificationMessage.textContent = '';
  };

  const show = (payload: OverlayNotificationPayload): void => {
    clearHideTimer();
    dom.overlayNotificationToast.classList.remove('hidden');
    dom.overlayNotificationToast.dataset.kind = payload.kind;
    dom.overlayNotificationTitle.textContent = payload.title ?? '';
    dom.overlayNotificationMessage.textContent = payload.message;

    if (payload.kind === 'loading') {
      dom.overlayNotificationSpinner.classList.remove('hidden');
      return;
    }

    dom.overlayNotificationSpinner.classList.add('hidden');
    hideTimer = setTimeoutHandler(() => {
      hide();
    }, payload.durationMs ?? durationMs);
  };

  return {
    show,
    hide,
  };
}
