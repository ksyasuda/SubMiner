import {
  createCancelNumericShortcutSessionHandler,
  createStartNumericShortcutSessionHandler,
} from './numeric-shortcut-session-handlers';
import {
  createBuildCancelNumericShortcutSessionMainDepsHandler,
  createBuildStartNumericShortcutSessionMainDepsHandler,
} from './numeric-shortcut-session-main-deps';

type CancelNumericShortcutSessionMainDeps = Parameters<
  typeof createBuildCancelNumericShortcutSessionMainDepsHandler
>[0];

export function createNumericShortcutSessionRuntimeHandlers(deps: {
  multiCopySession: CancelNumericShortcutSessionMainDeps['session'];
  mineSentenceSession: CancelNumericShortcutSessionMainDeps['session'];
  onMultiCopyDigit: (count: number) => void;
  onMineSentenceDigit: (count: number) => void;
  tryBeginMultiCopyOverlaySelection?: (timeoutMs: number) => boolean;
  tryBeginMineSentenceOverlaySelection?: (timeoutMs: number) => boolean;
}) {
  const cancelPendingMultiCopyMainDeps = createBuildCancelNumericShortcutSessionMainDepsHandler({
    session: deps.multiCopySession,
  })();
  const cancelPendingMultiCopyHandler = createCancelNumericShortcutSessionHandler(
    cancelPendingMultiCopyMainDeps,
  );

  const startPendingMultiCopyMainDeps = createBuildStartNumericShortcutSessionMainDepsHandler({
    session: deps.multiCopySession,
    onDigit: deps.onMultiCopyDigit,
    messages: {
      prompt: 'Copy how many lines? Press 1-9 (Esc to cancel)',
      timeout: 'Copy timeout',
      cancelled: 'Cancelled',
    },
  })();
  const startPendingMultiCopyHandler = createStartNumericShortcutSessionHandler(
    startPendingMultiCopyMainDeps,
  );

  const cancelPendingMineSentenceMultipleMainDeps =
    createBuildCancelNumericShortcutSessionMainDepsHandler({
      session: deps.mineSentenceSession,
    })();
  const cancelPendingMineSentenceMultipleHandler = createCancelNumericShortcutSessionHandler(
    cancelPendingMineSentenceMultipleMainDeps,
  );

  const startPendingMineSentenceMultipleMainDeps =
    createBuildStartNumericShortcutSessionMainDepsHandler({
      session: deps.mineSentenceSession,
      onDigit: deps.onMineSentenceDigit,
      messages: {
        prompt: 'Mine how many lines? Press 1-9 (Esc to cancel)',
        timeout: 'Mine sentence timeout',
        cancelled: 'Cancelled',
      },
    })();
  const startPendingMineSentenceMultipleHandler = createStartNumericShortcutSessionHandler(
    startPendingMineSentenceMultipleMainDeps,
  );

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
