import type { SubtitleData } from '../../types';

export interface SubtitleProcessingControllerDeps {
  tokenizeSubtitle: (text: string) => Promise<SubtitleData | null>;
  emitSubtitle: (payload: SubtitleData) => void;
  logDebug?: (message: string) => void;
  now?: () => number;
}

export interface SubtitleProcessingController {
  onSubtitleChange: (text: string) => void;
}

export function createSubtitleProcessingController(
  deps: SubtitleProcessingControllerDeps,
): SubtitleProcessingController {
  let latestText = '';
  let lastPlainText = '';
  let processing = false;
  let staleDropCount = 0;
  const now = deps.now ?? (() => Date.now());

  const emitPlainSubtitle = (text: string): void => {
    if (text === lastPlainText) {
      return;
    }
    lastPlainText = text;
    deps.emitSubtitle({ text, tokens: null });
  };

  const processLatest = (): void => {
    if (processing) {
      return;
    }

    processing = true;

    void (async () => {
      while (true) {
        const text = latestText;
        if (!text.trim()) {
          break;
        }

        const startedAtMs = now();
        let tokenized: SubtitleData | null = null;
        try {
          tokenized = await deps.tokenizeSubtitle(text);
        } catch (error) {
          deps.logDebug?.(`Subtitle tokenization failed: ${(error as Error).message}`);
        }

        if (latestText !== text) {
          staleDropCount += 1;
          deps.logDebug?.(
            `Dropped stale subtitle tokenization result; dropped=${staleDropCount}, elapsed=${now() - startedAtMs}ms`,
          );
          continue;
        }

        if (tokenized) {
          deps.emitSubtitle(tokenized);
          deps.logDebug?.(
            `Subtitle tokenization delivered; elapsed=${now() - startedAtMs}ms, staleDrops=${staleDropCount}`,
          );
        }
        break;
      }
    })()
      .catch((error) => {
        deps.logDebug?.(`Subtitle processing loop failed: ${(error as Error).message}`);
      })
      .finally(() => {
        processing = false;
        if (latestText !== lastPlainText) {
          processLatest();
        }
      });
  };

  return {
    onSubtitleChange: (text: string) => {
      if (text === latestText) {
        return;
      }
      const plainStartedAtMs = now();
      latestText = text;
      emitPlainSubtitle(text);
      deps.logDebug?.(`Subtitle plain emit completed in ${now() - plainStartedAtMs}ms`);
      if (!text.trim()) {
        return;
      }
      processLatest();
    },
  };
}
