import type { SecondarySubMode } from '../../types';

export function createBuildCycleSecondarySubModeMainDepsHandler(deps: {
  getSecondarySubMode: () => SecondarySubMode;
  setSecondarySubMode: (mode: SecondarySubMode) => void;
  getLastSecondarySubToggleAtMs: () => number;
  setLastSecondarySubToggleAtMs: (timestampMs: number) => void;
  broadcastToOverlayWindows: (channel: string, mode: SecondarySubMode) => void;
  showMpvOsd: (text: string) => void;
}) {
  return () => ({
    getSecondarySubMode: () => deps.getSecondarySubMode(),
    setSecondarySubMode: (mode: SecondarySubMode) => deps.setSecondarySubMode(mode),
    getLastSecondarySubToggleAtMs: () => deps.getLastSecondarySubToggleAtMs(),
    setLastSecondarySubToggleAtMs: (timestampMs: number) =>
      deps.setLastSecondarySubToggleAtMs(timestampMs),
    broadcastSecondarySubMode: (mode: SecondarySubMode) =>
      deps.broadcastToOverlayWindows('secondary-subtitle:mode', mode),
    showMpvOsd: (text: string) => deps.showMpvOsd(text),
  });
}
