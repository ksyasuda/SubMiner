import { isSetupCompleted, type SetupState } from '../src/shared/setup-state.js';

export async function waitForSetupCompletion(deps: {
  readSetupState: () => SetupState | null;
  sleep: (ms: number) => Promise<void>;
  now: () => number;
  timeoutMs: number;
  pollIntervalMs: number;
  ignoreInitialCancelledState?: boolean;
}): Promise<'completed' | 'cancelled' | 'timeout'> {
  const deadline = deps.now() + deps.timeoutMs;
  let ignoringCancelled = deps.ignoreInitialCancelledState === true;

  while (deps.now() <= deadline) {
    const state = deps.readSetupState();
    if (isSetupCompleted(state)) {
      return 'completed';
    }
    if (ignoringCancelled && state?.status !== 'cancelled') {
      ignoringCancelled = false;
    }
    if (state?.status === 'cancelled') {
      if (ignoringCancelled) {
        await deps.sleep(deps.pollIntervalMs);
        continue;
      }
      return 'cancelled';
    }
    await deps.sleep(deps.pollIntervalMs);
  }

  return 'timeout';
}

export async function ensureLauncherSetupReady(deps: {
  readSetupState: () => SetupState | null;
  isExternalYomitanConfigured?: () => boolean;
  isPluginInstalled?: () => boolean;
  launchSetupApp: () => void;
  sleep: (ms: number) => Promise<void>;
  now: () => number;
  timeoutMs: number;
  pollIntervalMs: number;
}): Promise<boolean> {
  if (deps.isExternalYomitanConfigured?.()) {
    return true;
  }
  if (deps.isPluginInstalled?.()) {
    return true;
  }
  const initialState = deps.readSetupState();
  if (isSetupCompleted(initialState)) {
    return true;
  }

  deps.launchSetupApp();
  const result = await waitForSetupCompletion({
    ...deps,
    ignoreInitialCancelledState: initialState?.status === 'cancelled',
  });
  return result === 'completed';
}
