import {
  createCancelNumericShortcutSessionHandler,
  createStartNumericShortcutSessionHandler,
} from './numeric-shortcut-session-handlers';

type CancelNumericShortcutSessionMainDeps = Parameters<
  typeof createCancelNumericShortcutSessionHandler
>[0];

export function createNumericShortcutSessionRuntimeHandlers(deps: {
  multiCopySession: CancelNumericShortcutSessionMainDeps['session'];
  mineSentenceSession: CancelNumericShortcutSessionMainDeps['session'];
  onMultiCopyDigit: (count: number) => void;
  onMineSentenceDigit: (count: number) => void;
  tryBeginMultiCopyOverlaySelection?: (timeoutMs: number) => boolean;
  tryBeginMineSentenceOverlaySelection?: (timeoutMs: number) => boolean;
}) {
  const cancelPendingMultiCopyHandler = createCancelNumericShortcutSessionHandler({
    session: deps.multiCopySession,
  });

  const startPendingMultiCopyHandler = createStartNumericShortcutSessionHandler({
    session: deps.multiCopySession,
    onDigit: deps.onMultiCopyDigit,
    messages: {
      prompt: 'Copy how many lines? Press 1-9 (Esc to cancel)',
      timeout: 'Copy timeout',
      cancelled: 'Cancelled',
    },
  });

  const cancelPendingMineSentenceMultipleHandler = createCancelNumericShortcutSessionHandler({
    session: deps.mineSentenceSession,
  });

  const startPendingMineSentenceMultipleHandler = createStartNumericShortcutSessionHandler({
    session: deps.mineSentenceSession,
    onDigit: deps.onMineSentenceDigit,
    messages: {
      prompt: 'Mine how many lines? Press 1-9 (Esc to cancel)',
      timeout: 'Mine sentence timeout',
      cancelled: 'Cancelled',
    },
  });

  return {
    cancelPendingMultiCopy: () => cancelPendingMultiCopyHandler(),
    startPendingMultiCopy: (timeoutMs: number) => {
      if (deps.tryBeginMultiCopyOverlaySelection?.(timeoutMs)) return;
      startPendingMultiCopyHandler(timeoutMs);
    },
    cancelPendingMineSentenceMultiple: () => cancelPendingMineSentenceMultipleHandler(),
    startPendingMineSentenceMultiple: (timeoutMs: number) => {
      if (deps.tryBeginMineSentenceOverlaySelection?.(timeoutMs)) return;
      startPendingMineSentenceMultipleHandler(timeoutMs);
    },
  };
}
