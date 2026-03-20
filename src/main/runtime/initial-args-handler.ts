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

    deps.handleCliCommand(initialArgs, 'initial');
  };
}
