import type { CliArgs } from '../../cli/args';

type MpvClientLike = {
  connected: boolean;
  connect: () => void;
};

export function createHandleInitialArgsHandler(deps: {
  getInitialArgs: () => CliArgs | null;
  isBackgroundMode: () => boolean;
  shouldEnsureTrayOnStartup: () => boolean;
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

    if (deps.isBackgroundMode() || deps.shouldEnsureTrayOnStartup()) {
      deps.ensureTray();
    }

    const mpvClient = deps.getMpvClient();
    if (
      !deps.isTexthookerOnlyMode() &&
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
