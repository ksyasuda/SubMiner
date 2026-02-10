import { SecondarySubMode } from "../../types";

export interface CycleSecondarySubModeDeps {
  getSecondarySubMode: () => SecondarySubMode;
  setSecondarySubMode: (mode: SecondarySubMode) => void;
  getLastSecondarySubToggleAtMs: () => number;
  setLastSecondarySubToggleAtMs: (timestampMs: number) => void;
  broadcastSecondarySubMode: (mode: SecondarySubMode) => void;
  showMpvOsd: (text: string) => void;
  now?: () => number;
}

const SECONDARY_SUB_CYCLE: SecondarySubMode[] = ["hidden", "visible", "hover"];
const SECONDARY_SUB_TOGGLE_DEBOUNCE_MS = 120;

export function cycleSecondarySubModeService(
  deps: CycleSecondarySubModeDeps,
): void {
  const now = deps.now ? deps.now() : Date.now();
  if (now - deps.getLastSecondarySubToggleAtMs() < SECONDARY_SUB_TOGGLE_DEBOUNCE_MS) {
    return;
  }
  deps.setLastSecondarySubToggleAtMs(now);

  const currentMode = deps.getSecondarySubMode();
  const currentIndex = SECONDARY_SUB_CYCLE.indexOf(currentMode);
  const nextMode =
    SECONDARY_SUB_CYCLE[(currentIndex + 1) % SECONDARY_SUB_CYCLE.length];
  deps.setSecondarySubMode(nextMode);
  deps.broadcastSecondarySubMode(nextMode);
  deps.showMpvOsd(`Secondary subtitle: ${nextMode}`);
}
