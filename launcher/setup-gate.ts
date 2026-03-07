import { isSetupCompleted, type SetupState } from '../src/shared/setup-state.js';

export async function waitForSetupCompletion(deps: {
  readSetupState: () => SetupState | null;
  sleep: (ms: number) => Promise<void>;
  now: () => number;
  timeoutMs: number;
  pollIntervalMs: number;
}): Promise<'completed' | 'cancelled' | 'timeout'> {
  const deadline = deps.now() + deps.timeoutMs;

  while (deps.now() <= deadline) {
    const state = deps.readSetupState();
    if (isSetupCompleted(state)) {
      return 'completed';
    }
    if (state?.status === 'cancelled') {
      return 'cancelled';
    }
    await deps.sleep(deps.pollIntervalMs);
  }

  return 'timeout';
}

export async function ensureLauncherSetupReady(deps: {
  readSetupState: () => SetupState | null;
  launchSetupApp: () => void;
  sleep: (ms: number) => Promise<void>;
  now: () => number;
  timeoutMs: number;
  pollIntervalMs: number;
}): Promise<boolean> {
  if (isSetupCompleted(deps.readSetupState())) {
    return true;
  }

  deps.launchSetupApp();
  const result = await waitForSetupCompletion(deps);
  return result === 'completed';
}
