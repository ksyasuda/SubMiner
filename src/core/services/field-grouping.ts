import { KikuFieldGroupingChoice, KikuFieldGroupingRequestData } from '../../types';

export function createFieldGroupingCallback(options: {
  getVisibleOverlayVisible: () => boolean;
  getInvisibleOverlayVisible: () => boolean;
  setVisibleOverlayVisible: (visible: boolean) => void;
  setInvisibleOverlayVisible: (visible: boolean) => void;
  getResolver: () => ((choice: KikuFieldGroupingChoice) => void) | null;
  setResolver: (resolver: ((choice: KikuFieldGroupingChoice) => void) | null) => void;
  sendRequestToVisibleOverlay: (data: KikuFieldGroupingRequestData) => boolean;
}): (data: KikuFieldGroupingRequestData) => Promise<KikuFieldGroupingChoice> {
  return async (data: KikuFieldGroupingRequestData): Promise<KikuFieldGroupingChoice> => {
    return new Promise((resolve) => {
      if (options.getResolver()) {
        resolve({
          keepNoteId: 0,
          deleteNoteId: 0,
          deleteDuplicate: true,
          cancelled: true,
        });
        return;
      }

      const previousVisibleOverlay = options.getVisibleOverlayVisible();
      const previousInvisibleOverlay = options.getInvisibleOverlayVisible();
      let settled = false;

      const finish = (choice: KikuFieldGroupingChoice): void => {
        if (settled) return;
        settled = true;
        if (options.getResolver() === finish) {
          options.setResolver(null);
        }
        resolve(choice);

        if (!previousVisibleOverlay && options.getVisibleOverlayVisible()) {
          options.setVisibleOverlayVisible(false);
        }
        if (options.getInvisibleOverlayVisible() !== previousInvisibleOverlay) {
          options.setInvisibleOverlayVisible(previousInvisibleOverlay);
        }
      };

      options.setResolver(finish);
      if (!options.sendRequestToVisibleOverlay(data)) {
        finish({
          keepNoteId: 0,
          deleteNoteId: 0,
          deleteDuplicate: true,
          cancelled: true,
        });
        return;
      }
      setTimeout(() => {
        if (!settled) {
          finish({
            keepNoteId: 0,
            deleteNoteId: 0,
            deleteDuplicate: true,
            cancelled: true,
          });
        }
      }, 90000);
    });
  };
}
