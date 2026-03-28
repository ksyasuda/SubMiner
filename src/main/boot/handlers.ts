import { composeOverlayWindowHandlers } from '../runtime/composers/overlay-window-composer';
import {
  composeCliStartupHandlers,
  composeHeadlessStartupHandlers,
  composeIpcRuntimeHandlers,
  composeStartupLifecycleHandlers,
} from '../runtime/composers';

export const composeBootStartupLifecycleHandlers = composeStartupLifecycleHandlers;
export const composeBootIpcRuntimeHandlers = composeIpcRuntimeHandlers;
export const composeBootCliStartupHandlers = composeCliStartupHandlers;
export const composeBootHeadlessStartupHandlers = composeHeadlessStartupHandlers;
export const composeBootOverlayWindowHandlers = composeOverlayWindowHandlers;
