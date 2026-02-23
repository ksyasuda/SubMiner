import type {
  NumericShortcutSessionMessages,
  NumericShortcutSessionStartParams,
} from '../../core/services/numeric-shortcut';

type NumericShortcutSessionLike = {
  start: (params: NumericShortcutSessionStartParams) => void;
  cancel: () => void;
};

export function createCancelNumericShortcutSessionHandler(deps: {
  session: NumericShortcutSessionLike;
}) {
  return (): void => {
    deps.session.cancel();
  };
}

export function createStartNumericShortcutSessionHandler(deps: {
  session: NumericShortcutSessionLike;
  onDigit: (digit: number) => void;
  messages: NumericShortcutSessionMessages;
}) {
  return (timeoutMs: number): void => {
    deps.session.start({
      timeoutMs,
      onDigit: deps.onDigit,
      messages: deps.messages,
    });
  };
}
