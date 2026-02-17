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

export function createNumericShortcutRuntime(
  options: NumericShortcutRuntimeOptions,
) {
  const createSession = () =>
    createNumericShortcutSession({
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

export interface NumericShortcutSessionMessages {
  prompt: string;
  timeout: string;
  cancelled?: string;
}

export interface NumericShortcutSessionDeps {
  registerShortcut: (accelerator: string, handler: () => void) => boolean;
  unregisterShortcut: (accelerator: string) => void;
  setTimer: (handler: () => void, timeoutMs: number) => ReturnType<typeof setTimeout>;
  clearTimer: (timer: ReturnType<typeof setTimeout>) => void;
  showMpvOsd: (text: string) => void;
}

export interface NumericShortcutSessionStartParams {
  timeoutMs: number;
  onDigit: (digit: number) => void;
  messages: NumericShortcutSessionMessages;
}

export function createNumericShortcutSession(
  deps: NumericShortcutSessionDeps,
) {
  let active = false;
  let timeout: ReturnType<typeof setTimeout> | null = null;
  let digitShortcuts: string[] = [];
  let escapeShortcut: string | null = null;

  let cancelledMessage = "Cancelled";

  const cancel = (showCancelled = false): void => {
    if (!active) return;
    active = false;

    if (timeout) {
      deps.clearTimer(timeout);
      timeout = null;
    }

    for (const shortcut of digitShortcuts) {
      deps.unregisterShortcut(shortcut);
    }
    digitShortcuts = [];

    if (escapeShortcut) {
      deps.unregisterShortcut(escapeShortcut);
      escapeShortcut = null;
    }

    if (showCancelled) {
      deps.showMpvOsd(cancelledMessage);
    }
  };

  const start = ({
    timeoutMs,
    onDigit,
    messages,
  }: NumericShortcutSessionStartParams): void => {
    cancel();
    cancelledMessage = messages.cancelled ?? "Cancelled";
    active = true;

    for (let i = 1; i <= 9; i++) {
      const shortcut = i.toString();
      if (
        deps.registerShortcut(shortcut, () => {
          if (!active) return;
          cancel();
          onDigit(i);
        })
      ) {
        digitShortcuts.push(shortcut);
      }
    }

    if (
      deps.registerShortcut("Escape", () => {
        cancel(true);
      })
    ) {
      escapeShortcut = "Escape";
    }

    timeout = deps.setTimer(() => {
      if (!active) return;
      cancel();
      deps.showMpvOsd(messages.timeout);
    }, timeoutMs);

    deps.showMpvOsd(messages.prompt);
  };

  return {
    start,
    cancel,
    isActive: (): boolean => active,
  };
}
