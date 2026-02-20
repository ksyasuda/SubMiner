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
  let lastEmittedText = '';
  let processing = false;
  let staleDropCount = 0;
  const now = deps.now ?? (() => Date.now());

  const processLatest = (): void => {
    if (processing) {
      return;
    }

    processing = true;

    void (async () => {
      while (true) {
        const text = latestText;
        const startedAtMs = now();

        if (!text.trim()) {
          deps.emitSubtitle({ text, tokens: null });
          lastEmittedText = text;
          break;
        }

        let output: SubtitleData = { text, tokens: null };
        try {
          const tokenized = await deps.tokenizeSubtitle(text);
          if (tokenized) {
            output = tokenized;
          }
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

        deps.emitSubtitle(output);
        lastEmittedText = text;
        deps.logDebug?.(
          `Subtitle tokenization delivered; elapsed=${now() - startedAtMs}ms, staleDrops=${staleDropCount}`,
        );
        break;
      }
    })()
      .catch((error) => {
        deps.logDebug?.(`Subtitle processing loop failed: ${(error as Error).message}`);
      })
      .finally(() => {
        processing = false;
        if (latestText !== lastEmittedText) {
          processLatest();
        }
      });
  };

  return {
    onSubtitleChange: (text: string) => {
      if (text === latestText) {
        return;
      }
      latestText = text;
      processLatest();
    },
  };
}
