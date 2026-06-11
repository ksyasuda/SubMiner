const DEFAULT_OVERLAY_LOADING_OSD_TICK_MS = 180;
const OVERLAY_LOADING_OSD_FRAMES = ['|', '/', '-', '\\'] as const;

export function createOverlayLoadingOsdController(deps: {
  showOsd: (message: string) => void;
  clearOsd: () => void;
  setInterval?: (callback: () => void, delayMs: number) => unknown;
  clearInterval?: (timer: unknown) => void;
}) {
  const setIntervalHandler =
    deps.setInterval ??
    ((callback: () => void, delayMs: number): unknown => setInterval(callback, delayMs));
  const clearIntervalHandler =
    deps.clearInterval ??
    ((timer: unknown): void => clearInterval(timer as ReturnType<typeof setInterval>));
  let active = false;
  let frame = 0;
  let timer: unknown = null;

  const showNextFrame = (): void => {
    deps.showOsd(
      `Overlay loading ${OVERLAY_LOADING_OSD_FRAMES[frame % OVERLAY_LOADING_OSD_FRAMES.length]}`,
    );
    frame += 1;
  };

  return {
    start(): void {
      if (active) {
        return;
      }
      active = true;
      frame = 0;
      showNextFrame();
      timer = setIntervalHandler(showNextFrame, DEFAULT_OVERLAY_LOADING_OSD_TICK_MS);
    },
    stop(): void {
      if (!active) {
        return;
      }
      active = false;
      if (timer !== null) {
        clearIntervalHandler(timer);
        timer = null;
      }
      deps.clearOsd();
    },
  };
}
