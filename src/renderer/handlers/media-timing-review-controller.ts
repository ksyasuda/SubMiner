import {
  CONTROLLER_REVIEW_SCREEN,
  type ControllerReviewStep,
} from '../../shared/controller-review';
import type { ControllerReviewAction, ResolvedControllerConfig } from '../../types';
import {
  createControllerComboScreenLayout,
  formatControllerScreenHints,
} from './controller-screen-layout.js';
import type { ControllerScreenLayout } from './gamepad-controller.js';

export { getControllerReviewInputContext } from './controller-screen-layout.js';

/** Every button on the media timing review screen, named by what it does. */
export type MediaTimingReviewControllerAction = ControllerReviewAction;

/** The review shows either the timing editor or the "Stop reviewing?" step. */
export type MediaTimingReviewStep = ControllerReviewStep;

/** Hint text per action for the review buttons, e.g. `A` or `L1 + ▶`; null when unbound. */
export function formatMediaTimingReviewControllerHints(
  config: ResolvedControllerConfig,
  mapping: string | null | undefined,
): Record<MediaTimingReviewControllerAction, string | null> {
  return formatControllerScreenHints(
    CONTROLLER_REVIEW_SCREEN,
    config.reviewBindings,
    config,
    mapping,
  );
}

export function createMediaTimingReviewControllerLayout(options: {
  getStep: () => MediaTimingReviewStep;
  dispatch: (action: MediaTimingReviewControllerAction) => void;
}): ControllerScreenLayout {
  return createControllerComboScreenLayout({
    screen: CONTROLLER_REVIEW_SCREEN,
    getBindings: (config) => config.reviewBindings,
    getStep: options.getStep,
    dispatch: options.dispatch,
  });
}
