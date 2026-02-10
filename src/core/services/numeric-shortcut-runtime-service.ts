import {
  createNumericShortcutSessionService,
} from "./numeric-shortcut-session-service";

interface GlobalShortcutLike {
  register: (accelerator: string, callback: () => void) => boolean;
  unregister: (accelerator: string) => void;
}

export interface NumericShortcutRuntimeOptions {
  globalShortcut: GlobalShortcutLike;
  showMpvOsd: (text: string) => void;
  setTimer: (
    handler: () => void,
    timeoutMs: number,
  ) => ReturnType<typeof setTimeout>;
  clearTimer: (timer: ReturnType<typeof setTimeout>) => void;
}

export function createNumericShortcutRuntimeService(
  options: NumericShortcutRuntimeOptions,
) {
  const createSession = () =>
    createNumericShortcutSessionService({
      registerShortcut: (accelerator, handler) =>
        options.globalShortcut.register(accelerator, handler),
      unregisterShortcut: (accelerator) =>
        options.globalShortcut.unregister(accelerator),
      setTimer: options.setTimer,
      clearTimer: options.clearTimer,
      showMpvOsd: options.showMpvOsd,
    });

  return {
    createSession,
  };
}
