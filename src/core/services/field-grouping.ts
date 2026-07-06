import { KikuFieldGroupingChoice, KikuFieldGroupingRequestData } from '../../types';

const DEFAULT_FIELD_GROUPING_RESPONSE_TIMEOUT_MS = 90000;

export function createFieldGroupingCallback(options: {
  getVisibleOverlayVisible: () => boolean;
  setVisibleOverlayVisible: (visible: boolean) => void;
  getResolver: () => ((choice: KikuFieldGroupingChoice) => void) | null;
  setResolver: (resolver: ((choice: KikuFieldGroupingChoice) => void) | null) => void;
  sendRequestToVisibleOverlay: (data: KikuFieldGroupingRequestData) => boolean | Promise<boolean>;
  /**
   * Tears down the modal UI when the request is abandoned without a renderer response
   * (send failure or response timeout). Without this the dedicated modal window, its
   * restore-set entry, and the forced main-overlay passthrough stay orphaned — which on
   * Wayland leaves an invisible modal covering mpv and the overlay stuck unresponsive.
   */
  dismissModalUi?: () => void;
  responseTimeoutMs?: number;
}): (data: KikuFieldGroupingRequestData) => Promise<KikuFieldGroupingChoice> {
  const cancelledChoice = (): KikuFieldGroupingChoice => ({
    keepNoteId: 0,
    deleteNoteId: 0,
    deleteDuplicate: true,
    cancelled: true,
  });

  const dismissModalUi = (): void => {
    try {
      options.dismissModalUi?.();
    } catch (error) {
      console.error('Failed to dismiss Kiku field grouping modal UI:', error);
    }
  };

  return async (data: KikuFieldGroupingRequestData): Promise<KikuFieldGroupingChoice> => {
    return new Promise((resolve) => {
      if (options.getResolver()) {
        resolve(cancelledChoice());
        return;
      }

      const previousVisibleOverlay = options.getVisibleOverlayVisible();
      let settled = false;
      let timeout: ReturnType<typeof setTimeout> | null = null;

      const finish = (choice: KikuFieldGroupingChoice, abandoned = false): void => {
        if (settled) return;
        settled = true;
        if (timeout !== null) {
          clearTimeout(timeout);
          timeout = null;
        }
        // Always release the resolver. Callers (main) wrap this in a sequence-guarded
        // resolver, so an identity check against `finish` never matches and would leak
        // the resolver — blocking every later grouping attempt with an instant cancel.
        options.setResolver(null);
        resolve(choice);

        // When abandoned without a renderer response, tear down the modal window/state
        // that the request path spun up. A normal response already routes through the
        // renderer's close handler, so only the abandon paths need this.
        if (abandoned) {
          dismissModalUi();
        }

        if (!previousVisibleOverlay && options.getVisibleOverlayVisible()) {
          options.setVisibleOverlayVisible(false);
        }
      };

      options.setResolver(finish);
      void Promise.resolve(options.sendRequestToVisibleOverlay(data)).then(
        (sent) => {
          if (settled) return;
          if (!sent) {
            finish(cancelledChoice(), true);
            return;
          }
          timeout = setTimeout(() => {
            if (!settled) {
              finish(cancelledChoice(), true);
            }
          }, options.responseTimeoutMs ?? DEFAULT_FIELD_GROUPING_RESPONSE_TIMEOUT_MS);
        },
        () => {
          finish(cancelledChoice(), true);
        },
      );
    });
  };
}
