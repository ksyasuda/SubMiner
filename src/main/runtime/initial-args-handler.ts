import type { CliArgs } from '../../cli/args';

type MpvClientLike = {
  connected: boolean;
  connect: () => void;
};

export function createHandleInitialArgsHandler(deps: {
  getInitialArgs: () => CliArgs | null;
  isBackgroundMode: () => boolean;
  shouldEnsureTrayOnStartup: () => boolean;
  shouldRunHeadlessInitialCommand: (args: CliArgs) => boolean;
  ensureTray: () => void;
  isTexthookerOnlyMode: () => boolean;
  hasImmersionTracker: () => boolean;
  getMpvClient: () => MpvClientLike | null;
  commandNeedsOverlayRuntime: (args: CliArgs) => boolean;
  ensureOverlayStartupPrereqs: () => void;
  isOverlayRuntimeInitialized: () => boolean;
  initializeOverlayRuntime: () => void;
  logInfo: (message: string) => void;
  handleCliCommand: (args: CliArgs, source: 'initial') => void;
}) {
  return (): void => {
    const initialArgs = deps.getInitialArgs();
    if (!initialArgs) return;
    const runHeadless = deps.shouldRunHeadlessInitialCommand(initialArgs);

    if (!runHeadless && (deps.isBackgroundMode() || deps.shouldEnsureTrayOnStartup())) {
      deps.ensureTray();
    }

    const mpvClient = deps.getMpvClient();
    if (
      !runHeadless &&
      !deps.isTexthookerOnlyMode() &&
      !initialArgs.stats &&
      deps.hasImmersionTracker() &&
      mpvClient &&
      !mpvClient.connected
    ) {
      deps.logInfo('Auto-connecting MPV client for immersion tracking');
      mpvClient.connect();
    }

    if (!runHeadless && deps.commandNeedsOverlayRuntime(initialArgs)) {
      deps.ensureOverlayStartupPrereqs();
      if (!deps.isOverlayRuntimeInitialized()) {
        deps.initializeOverlayRuntime();
      }
    }

    deps.handleCliCommand(initialArgs, 'initial');
  };
}
