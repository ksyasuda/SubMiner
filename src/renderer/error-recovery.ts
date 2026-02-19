export type RendererErrorSource =
  | 'callback'
  | 'window.onerror'
  | 'window.unhandledrejection'
  | 'bootstrap';

export type RendererRecoveryContext = {
  source: RendererErrorSource;
  action: string;
  details?: Record<string, unknown>;
};

export type RendererRecoverySnapshot = {
  activeModal: string | null;
  subtitlePreview: string;
  secondarySubtitlePreview: string;
  isOverlayInteractive: boolean;
  isOverSubtitle: boolean;
  invisiblePositionEditMode: boolean;
  overlayLayer: 'visible' | 'invisible';
};

type NormalizedRendererError = {
  message: string;
  stack: string | null;
};

export type RendererRecoveryLogPayload = {
  kind: 'renderer-overlay-recovery';
  context: RendererRecoveryContext;
  error: NormalizedRendererError;
  snapshot: RendererRecoverySnapshot;
  timestamp: string;
};

type RendererRecoveryDeps = {
  dismissActiveUi: () => void;
  restoreOverlayInteraction: () => void;
  showToast: (message: string) => void;
  getSnapshot: () => RendererRecoverySnapshot;
  logError: (payload: RendererRecoveryLogPayload) => void;
  toastMessage?: string;
};

type RendererRecoveryController = {
  handleError: (error: unknown, context: RendererRecoveryContext) => void;
};

type RendererRecoveryWindow = Pick<Window, 'addEventListener' | 'removeEventListener'>;

const DEFAULT_TOAST_MESSAGE = 'Renderer error recovered. Overlay is still running.';

function normalizeRendererError(error: unknown): NormalizedRendererError {
  if (error instanceof Error) {
    return {
      message: error.message || 'Unknown renderer error',
      stack: typeof error.stack === 'string' ? error.stack : null,
    };
  }

  if (typeof error === 'string') {
    return {
      message: error,
      stack: null,
    };
  }

  if (typeof error === 'object' && error !== null) {
    try {
      return {
        message: JSON.stringify(error),
        stack: null,
      };
    } catch {
      return {
        message: '[unserializable error object]',
        stack: null,
      };
    }
  }

  return {
    message: String(error),
    stack: null,
  };
}

export function createRendererRecoveryController(
  deps: RendererRecoveryDeps,
): RendererRecoveryController {
  let inRecovery = false;

  const toastMessage = deps.toastMessage ?? DEFAULT_TOAST_MESSAGE;

  const invokeRecoveryStep = (
    step: 'dismissActiveUi' | 'restoreOverlayInteraction' | 'showToast',
    fn: () => void,
  ): void => {
    try {
      fn();
    } catch (error) {
      try {
        deps.logError({
          kind: 'renderer-overlay-recovery',
          context: {
            source: 'callback',
            action: `recovery-step:${step}`,
          },
          error: normalizeRendererError(error),
          snapshot: deps.getSnapshot(),
          timestamp: new Date().toISOString(),
        });
      } catch {
        // Avoid recursive failures from logging inside the recovery path.
      }
    }
  };

  const handleError = (error: unknown, context: RendererRecoveryContext): void => {
    if (inRecovery) {
      return;
    }

    inRecovery = true;
    try {
      deps.logError({
        kind: 'renderer-overlay-recovery',
        context,
        error: normalizeRendererError(error),
        snapshot: deps.getSnapshot(),
        timestamp: new Date().toISOString(),
      });

      invokeRecoveryStep('dismissActiveUi', deps.dismissActiveUi);
      invokeRecoveryStep('restoreOverlayInteraction', deps.restoreOverlayInteraction);
      invokeRecoveryStep('showToast', () => deps.showToast(toastMessage));
    } finally {
      inRecovery = false;
    }
  };

  return { handleError };
}

export function registerRendererGlobalErrorHandlers(
  recoveryWindow: RendererRecoveryWindow,
  controller: RendererRecoveryController,
): () => void {
  const onError = (event: Event): void => {
    const errorEvent = event as ErrorEvent;
    controller.handleError(errorEvent.error ?? errorEvent.message, {
      source: 'window.onerror',
      action: 'global-error',
      details: {
        filename: errorEvent.filename,
        lineno: errorEvent.lineno,
        colno: errorEvent.colno,
      },
    });
  };

  const onUnhandledRejection = (event: Event): void => {
    const rejectionEvent = event as PromiseRejectionEvent;
    controller.handleError(rejectionEvent.reason, {
      source: 'window.unhandledrejection',
      action: 'global-unhandledrejection',
    });
  };

  recoveryWindow.addEventListener('error', onError);
  recoveryWindow.addEventListener('unhandledrejection', onUnhandledRejection);

  return () => {
    recoveryWindow.removeEventListener('error', onError);
    recoveryWindow.removeEventListener('unhandledrejection', onUnhandledRejection);
  };
}
