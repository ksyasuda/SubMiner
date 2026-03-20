import type { CliArgs } from '../../cli/args';

export function createBuildHandleInitialArgsMainDepsHandler(deps: {
  getInitialArgs: () => CliArgs | null;
  isBackgroundMode: () => boolean;
  shouldEnsureTrayOnStartup: () => boolean;
  shouldRunHeadlessInitialCommand: (args: CliArgs) => boolean;
  ensureTray: () => void;
  isTexthookerOnlyMode: () => boolean;
  hasImmersionTracker: () => boolean;
  getMpvClient: () => { connected: boolean; connect: () => void } | null;
  logInfo: (message: string) => void;
  handleCliCommand: (args: CliArgs, source: 'initial') => void;
}) {
  return () => ({
    getInitialArgs: () => deps.getInitialArgs(),
    isBackgroundMode: () => deps.isBackgroundMode(),
    shouldEnsureTrayOnStartup: () => deps.shouldEnsureTrayOnStartup(),
    shouldRunHeadlessInitialCommand: (args: CliArgs) => deps.shouldRunHeadlessInitialCommand(args),
    ensureTray: () => deps.ensureTray(),
    isTexthookerOnlyMode: () => deps.isTexthookerOnlyMode(),
    hasImmersionTracker: () => deps.hasImmersionTracker(),
    getMpvClient: () => deps.getMpvClient(),
    logInfo: (message: string) => deps.logInfo(message),
    handleCliCommand: (args: CliArgs, source: 'initial') => deps.handleCliCommand(args, source),
  });
}
