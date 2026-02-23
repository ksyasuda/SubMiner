import type { NumericShortcutRuntimeOptions } from '../../core/services/numeric-shortcut';

export function createBuildNumericShortcutRuntimeMainDepsHandler(deps: NumericShortcutRuntimeOptions) {
  return (): NumericShortcutRuntimeOptions => ({
    globalShortcut: deps.globalShortcut,
    showMpvOsd: (text: string) => deps.showMpvOsd(text),
    setTimer: (handler: () => void, timeoutMs: number) => deps.setTimer(handler, timeoutMs),
    clearTimer: (timer: ReturnType<typeof setTimeout>) => deps.clearTimer(timer),
  });
}
