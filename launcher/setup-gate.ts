import type { DictionaryBackend } from '../src/types/config.js';
import {
  getSetupStateDictionaryBackend,
  isSetupCompleted,
  type SetupState,
} from '../src/shared/setup-state.js';

export async function waitForSetupCompletion(deps: {
  readSetupState: () => SetupState | null;
  dictionaryBackend?: DictionaryBackend;
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
    if (isSetupCompleted(state, deps.dictionaryBackend)) {
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
  dictionaryBackend?: DictionaryBackend;
  sleep: (ms: number) => Promise<void>;
  now: () => number;
  timeoutMs: number;
  pollIntervalMs: number;
  initialState?: SetupState | null;
}): Promise<'acknowledged' | 'cancelled' | 'timeout'> {
  const deadline = deps.now() + deps.timeoutMs;
  const initialCompleted = isSetupCompleted(deps.initialState, deps.dictionaryBackend);
  const initialCompletedAt = deps.initialState?.completedAt ?? null;

  while (deps.now() <= deadline) {
    const state = deps.readSetupState();
    if (
      isSetupCompleted(state, deps.dictionaryBackend) &&
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

/**
 * The app pins its dictionary backend at startup while the config file can change
 * underneath it. When an app is already running, gate on the backend it recorded
 * in the setup state rather than the config value it has not restarted into.
 */
export async function resolveLauncherGateBackend(deps: {
  configuredBackend: DictionaryBackend;
  state: SetupState | null;
  isAppRunning?: () => Promise<boolean>;
  warn?: (message: string) => void;
}): Promise<DictionaryBackend> {
  const runningBackend = deps.state
    ? getSetupStateDictionaryBackend(deps.state)
    : deps.configuredBackend;
  if (runningBackend === deps.configuredBackend || !(await deps.isAppRunning?.())) {
    return deps.configuredBackend;
  }
  deps.warn?.(
    `SubMiner is running with the ${runningBackend} dictionary backend; restart it to switch to ${deps.configuredBackend}.`,
  );
  return runningBackend;
}

export async function ensureLauncherSetupReady(deps: {
  readSetupState: () => SetupState | null;
  dictionaryBackend?: DictionaryBackend;
  isAppRunning?: () => Promise<boolean>;
  warn?: (message: string) => void;
  isExternalYomitanConfigured?: () => boolean;
  hasLegacyMpvPlugin?: () => boolean;
  launchSetupApp: () => void;
  sleep: (ms: number) => Promise<void>;
  now: () => number;
  timeoutMs: number;
  pollIntervalMs: number;
}): Promise<boolean> {
  const initialState = deps.readSetupState();
  const dictionaryBackend = await resolveLauncherGateBackend({
    configuredBackend: deps.dictionaryBackend ?? 'yomitan',
    state: initialState,
    isAppRunning: deps.isAppRunning,
    warn: deps.warn,
  });
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
      dictionaryBackend,
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

  if (dictionaryBackend !== 'hachidori' && deps.isExternalYomitanConfigured?.()) {
    return true;
  }
  const stateAfterLegacyPrompt = deps.readSetupState();
  if (isSetupCompleted(stateAfterLegacyPrompt, dictionaryBackend)) {
    return true;
  }

  launchSetupApp();
  const result = await waitForSetupCompletion({
    ...deps,
    dictionaryBackend,
    ignoreInitialCancelledState: stateAfterLegacyPrompt?.status === 'cancelled',
  });
  return result === 'completed';
}
