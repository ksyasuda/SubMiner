import type {
  createCancelNumericShortcutSessionHandler,
  createStartNumericShortcutSessionHandler,
} from './numeric-shortcut-session-handlers';

type CancelNumericShortcutSessionMainDeps = Parameters<typeof createCancelNumericShortcutSessionHandler>[0];
type StartNumericShortcutSessionMainDeps = Parameters<typeof createStartNumericShortcutSessionHandler>[0];

export function createBuildCancelNumericShortcutSessionMainDepsHandler(
  deps: CancelNumericShortcutSessionMainDeps,
) {
  return (): CancelNumericShortcutSessionMainDeps => ({
    session: deps.session,
  });
}

export function createBuildStartNumericShortcutSessionMainDepsHandler(
  deps: StartNumericShortcutSessionMainDeps,
) {
  return (): StartNumericShortcutSessionMainDeps => ({
    session: deps.session,
    onDigit: (digit: number) => deps.onDigit(digit),
    messages: deps.messages,
  });
}
