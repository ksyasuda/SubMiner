import { KikuFieldGroupingChoice, KikuFieldGroupingRequestData } from '../../types';

export function createFieldGroupingCallback(options: {
  getVisibleOverlayVisible: () => boolean;
  setVisibleOverlayVisible: (visible: boolean) => void;
  getResolver: () => ((choice: KikuFieldGroupingChoice) => void) | null;
  setResolver: (resolver: ((choice: KikuFieldGroupingChoice) => void) | null) => void;
  sendRequestToVisibleOverlay: (data: KikuFieldGroupingRequestData) => boolean | Promise<boolean>;
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
      let settled = false;
      let timeout: ReturnType<typeof setTimeout> | null = null;

      const finish = (choice: KikuFieldGroupingChoice): void => {
        if (settled) return;
        settled = true;
        if (timeout !== null) {
          clearTimeout(timeout);
          timeout = null;
        }
        if (options.getResolver() === finish) {
          options.setResolver(null);
        }
        resolve(choice);

        if (!previousVisibleOverlay && options.getVisibleOverlayVisible()) {
          options.setVisibleOverlayVisible(false);
        }
      };

      options.setResolver(finish);
      void Promise.resolve(options.sendRequestToVisibleOverlay(data)).then(
        (sent) => {
          if (settled) return;
          if (!sent) {
            finish({
              keepNoteId: 0,
              deleteNoteId: 0,
              deleteDuplicate: true,
              cancelled: true,
            });
            return;
          }
          timeout = setTimeout(() => {
            if (!settled) {
              finish({
                keepNoteId: 0,
                deleteNoteId: 0,
                deleteDuplicate: true,
                cancelled: true,
              });
            }
          }, 90000);
        },
        () => {
          finish({
            keepNoteId: 0,
            deleteNoteId: 0,
            deleteDuplicate: true,
            cancelled: true,
          });
        },
      );
    });
  };
}
