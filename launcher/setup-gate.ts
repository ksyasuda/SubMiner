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
    if (ignoringCancelled && state != null && state.status !== 'cancelled') {
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

export async function waitForLegacyMpvPluginPromptResolution(deps: {
  readSetupState: () => SetupState | null;
  sleep: (ms: number) => Promise<void>;
  now: () => number;
  timeoutMs: number;
  pollIntervalMs: number;
  initialState?: SetupState | null;
}): Promise<'acknowledged' | 'cancelled' | 'timeout'> {
  const deadline = deps.now() + deps.timeoutMs;
  const initialCompleted = isSetupCompleted(deps.initialState);
  const initialCompletedAt = deps.initialState?.completedAt ?? null;

  while (deps.now() <= deadline) {
    const state = deps.readSetupState();
    if (
      isSetupCompleted(state) &&
      (!initialCompleted || state?.completedAt !== initialCompletedAt)
    ) {
      return 'acknowledged';
    }
    if (!initialCompleted && state?.status === 'cancelled') {
      return 'cancelled';
    }

    await deps.sleep(deps.pollIntervalMs);
  }

  return 'timeout';
}

export async function ensureLauncherSetupReady(deps: {
  readSetupState: () => SetupState | null;
  isExternalYomitanConfigured?: () => boolean;
  hasLegacyMpvPlugin?: () => boolean;
  launchSetupApp: () => void;
  sleep: (ms: number) => Promise<void>;
  now: () => number;
  timeoutMs: number;
  pollIntervalMs: number;
}): Promise<boolean> {
  const initialState = deps.readSetupState();
  let setupLaunched = false;
  const launchSetupApp = () => {
    if (setupLaunched) return;
    setupLaunched = true;
    deps.launchSetupApp();
  };

  if (deps.hasLegacyMpvPlugin?.()) {
    launchSetupApp();
    const result = await waitForLegacyMpvPluginPromptResolution({
      readSetupState: deps.readSetupState,
      sleep: deps.sleep,
      now: deps.now,
      timeoutMs: deps.timeoutMs,
      pollIntervalMs: deps.pollIntervalMs,
      initialState,
    });
    if (result === 'cancelled' || result === 'timeout') {
      return false;
    }
  }

  if (deps.isExternalYomitanConfigured?.()) {
    return true;
  }
  const stateAfterLegacyPrompt = deps.readSetupState();
  if (isSetupCompleted(stateAfterLegacyPrompt)) {
    return true;
  }

  launchSetupApp();
  const result = await waitForSetupCompletion({
    ...deps,
    ignoreInitialCancelledState: stateAfterLegacyPrompt?.status === 'cancelled',
  });
  return result === 'completed';
}
