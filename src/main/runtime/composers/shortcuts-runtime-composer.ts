import { createNumericShortcutRuntime } from '../../../core/services/numeric-shortcut';
import {
  createBuildNumericShortcutRuntimeMainDepsHandler,
  createGlobalShortcutsRuntimeHandlers,
  createNumericShortcutSessionRuntimeHandlers,
  createOverlayShortcutsRuntimeHandlers,
} from '../domains/shortcuts';
import type { ComposerInputs, ComposerOutputs } from './contracts';

type GlobalShortcutsOptions = Parameters<typeof createGlobalShortcutsRuntimeHandlers>[0];
type NumericShortcutRuntimeMainDeps = Parameters<
  typeof createBuildNumericShortcutRuntimeMainDepsHandler
>[0];
type NumericSessionOptions = Omit<
  Parameters<typeof createNumericShortcutSessionRuntimeHandlers>[0],
  'multiCopySession' | 'mineSentenceSession'
>;
type OverlayShortcutsMainDeps = Parameters<
  typeof createOverlayShortcutsRuntimeHandlers
>[0]['overlayShortcutsRuntimeMainDeps'];

export type ShortcutsRuntimeComposerOptions = ComposerInputs<{
  globalShortcuts: GlobalShortcutsOptions;
  numericShortcutRuntimeMainDeps: NumericShortcutRuntimeMainDeps;
  numericSessions: NumericSessionOptions;
  overlayShortcutsRuntimeMainDeps: OverlayShortcutsMainDeps;
}>;

export type ShortcutsRuntimeComposerResult = ComposerOutputs<
  ReturnType<typeof createGlobalShortcutsRuntimeHandlers> &
    ReturnType<typeof createNumericShortcutSessionRuntimeHandlers> &
    ReturnType<typeof createOverlayShortcutsRuntimeHandlers>
>;

export function composeShortcutRuntimes(
  options: ShortcutsRuntimeComposerOptions,
): ShortcutsRuntimeComposerResult {
  const globalShortcuts = createGlobalShortcutsRuntimeHandlers(options.globalShortcuts);

  const numericShortcutRuntimeMainDeps = createBuildNumericShortcutRuntimeMainDepsHandler(
    options.numericShortcutRuntimeMainDeps,
  )();
  const numericShortcutRuntime = createNumericShortcutRuntime(numericShortcutRuntimeMainDeps);
  const numericSessions = createNumericShortcutSessionRuntimeHandlers({
    multiCopySession: numericShortcutRuntime.createSession(),
    mineSentenceSession: numericShortcutRuntime.createSession(),
    onMultiCopyDigit: options.numericSessions.onMultiCopyDigit,
    onMineSentenceDigit: options.numericSessions.onMineSentenceDigit,
  });

  const overlayShortcuts = createOverlayShortcutsRuntimeHandlers({
    overlayShortcutsRuntimeMainDeps: options.overlayShortcutsRuntimeMainDeps,
  });

  return {
    ...globalShortcuts,
    ...numericSessions,
    ...overlayShortcuts,
  };
}
