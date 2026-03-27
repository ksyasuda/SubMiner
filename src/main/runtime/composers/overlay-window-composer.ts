import { createOverlayWindowRuntimeHandlers } from '../overlay-window-runtime-handlers';
import type { ComposerInputs, ComposerOutputs } from './contracts';

type OverlayWindowRuntimeDeps<TWindow> =
  Parameters<typeof createOverlayWindowRuntimeHandlers<TWindow>>[0];
type OverlayWindowRuntimeHandlers<TWindow> = ReturnType<
  typeof createOverlayWindowRuntimeHandlers<TWindow>
>;

export type OverlayWindowComposerOptions<TWindow> = ComposerInputs<OverlayWindowRuntimeDeps<TWindow>>;
export type OverlayWindowComposerResult<TWindow> =
  ComposerOutputs<OverlayWindowRuntimeHandlers<TWindow>>;

export function composeOverlayWindowHandlers<TWindow>(
  options: OverlayWindowComposerOptions<TWindow>,
): OverlayWindowComposerResult<TWindow> {
  return createOverlayWindowRuntimeHandlers<TWindow>(options);
}
