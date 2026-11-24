import {
  createHandleJellyfinRemoteGeneralCommand,
  createHandleJellyfinRemotePlay,
  createHandleJellyfinRemotePlaystate,
  createJellyfinRemoteReportTracker,
  createReportJellyfinRemoteProgressHandler,
  createReportJellyfinRemoteStoppedHandler,
  createReportSupersededJellyfinPlaybackHandler,
  type JellyfinRemoteGeneralCommandHandlerDeps,
  type JellyfinRemotePlayHandlerDeps,
  type JellyfinRemotePlaystateHandlerDeps,
  type JellyfinRemoteProgressReporterDeps,
} from '../domains/jellyfin';
import type { ComposerInputs, ComposerOutputs } from './contracts';

type RemotePlayPayload = Parameters<ReturnType<typeof createHandleJellyfinRemotePlay>>[0];
type RemotePlaystatePayload = Parameters<ReturnType<typeof createHandleJellyfinRemotePlaystate>>[0];
type RemoteGeneralPayload = Parameters<
  ReturnType<typeof createHandleJellyfinRemoteGeneralCommand>
>[0];

export type JellyfinRemoteComposerOptions = ComposerInputs<{
  getConfiguredSession: JellyfinRemotePlayHandlerDeps['getConfiguredSession'];
  getClientInfo: JellyfinRemotePlayHandlerDeps['getClientInfo'];
  getJellyfinConfig: JellyfinRemotePlayHandlerDeps['getJellyfinConfig'];
  playJellyfinItem: JellyfinRemotePlayHandlerDeps['playJellyfinItem'];
  logWarn: JellyfinRemotePlayHandlerDeps['logWarn'];
  getMpvClient: JellyfinRemoteProgressReporterDeps['getMpvClient'];
  sendMpvCommand: JellyfinRemotePlaystateHandlerDeps['sendMpvCommand'];
  jellyfinTicksToSeconds: JellyfinRemotePlaystateHandlerDeps['jellyfinTicksToSeconds'];
  getActivePlayback: JellyfinRemoteGeneralCommandHandlerDeps['getActivePlayback'];
  clearActivePlayback: JellyfinRemoteProgressReporterDeps['clearActivePlayback'];
  getNow: JellyfinRemoteProgressReporterDeps['getNow'];
  getLastProgressAtMs: JellyfinRemoteProgressReporterDeps['getLastProgressAtMs'];
  setLastProgressAtMs: JellyfinRemoteProgressReporterDeps['setLastProgressAtMs'];
  progressIntervalMs: number;
  ticksPerSecond: number;
  logDebug: JellyfinRemoteProgressReporterDeps['logDebug'];
}>;

export type JellyfinRemoteComposerResult = ComposerOutputs<{
  reportJellyfinRemoteProgress: ReturnType<typeof createReportJellyfinRemoteProgressHandler>;
  reportJellyfinRemoteStopped: ReturnType<typeof createReportJellyfinRemoteStoppedHandler>;
  reportJellyfinPlaybackSuperseded: ReturnType<
    typeof createReportSupersededJellyfinPlaybackHandler
  >;
  /** Resolves once no episode start or stop report is in flight. */
  waitForJellyfinTransitionReports: () => Promise<void>;
  /** Tracks an episode start report, so its own stop and later starts wait for it. */
  trackJellyfinPlaybackStart: (
    report: Promise<void>,
    playback: Parameters<ReturnType<typeof createReportSupersededJellyfinPlaybackHandler>>[0],
  ) => void;
  handleJellyfinRemotePlay: (payload: RemotePlayPayload) => Promise<void>;
  handleJellyfinRemotePlaystate: (payload: RemotePlaystatePayload) => Promise<void>;
  handleJellyfinRemoteGeneralCommand: (payload: RemoteGeneralPayload) => Promise<void>;
}>;

export function composeJellyfinRemoteHandlers(
  options: JellyfinRemoteComposerOptions,
): JellyfinRemoteComposerResult {
  const reportTracker = createJellyfinRemoteReportTracker();
  const reportJellyfinRemoteProgress = createReportJellyfinRemoteProgressHandler({
    getActivePlayback: options.getActivePlayback,
    clearActivePlayback: options.clearActivePlayback,
    getMpvClient: options.getMpvClient,
    getNow: options.getNow,
    getLastProgressAtMs: options.getLastProgressAtMs,
    setLastProgressAtMs: options.setLastProgressAtMs,
    progressIntervalMs: options.progressIntervalMs,
    ticksPerSecond: options.ticksPerSecond,
    logDebug: options.logDebug,
    reportTracker,
  });
  // Every episode start and stop report in flight. Jellyfin keeps whichever start arrives last
  // as the now-playing item and clears it on any stop, so a start waits for all of these.
  const transitionReportTracker = createJellyfinRemoteReportTracker();
  const reportStopped = createReportJellyfinRemoteStoppedHandler({
    getActivePlayback: options.getActivePlayback,
    clearActivePlayback: options.clearActivePlayback,
    getMpvClient: options.getMpvClient,
    getNow: options.getNow,
    ticksPerSecond: options.ticksPerSecond,
    logDebug: options.logDebug,
    logWarn: options.logWarn,
    reportTracker,
  });
  const reportSuperseded = createReportSupersededJellyfinPlaybackHandler({
    ticksPerSecond: options.ticksPerSecond,
    logDebug: options.logDebug,
    logWarn: options.logWarn,
    reportTracker,
  });
  const reportJellyfinRemoteStopped = (): Promise<void> => {
    const report = reportStopped();
    transitionReportTracker.track(report, null);
    return report;
  };
  const reportJellyfinPlaybackSuperseded = (
    playback: Parameters<typeof reportSuperseded>[0],
  ): Promise<void> => {
    const report = reportSuperseded(playback);
    transitionReportTracker.track(report, playback);
    return report;
  };

  return {
    reportJellyfinRemoteProgress,
    reportJellyfinRemoteStopped,
    reportJellyfinPlaybackSuperseded,
    waitForJellyfinTransitionReports: () => transitionReportTracker.settled(),
    trackJellyfinPlaybackStart: (report, playback) => {
      // Tagged with its playback in the shared tracker, so that playback's stop waits for it.
      reportTracker.track(report, playback);
      transitionReportTracker.track(report, playback);
    },
    handleJellyfinRemotePlay: createHandleJellyfinRemotePlay({
      getConfiguredSession: options.getConfiguredSession,
      getClientInfo: options.getClientInfo,
      getJellyfinConfig: options.getJellyfinConfig,
      getActivePlayback: options.getActivePlayback,
      playJellyfinItem: options.playJellyfinItem,
      logWarn: options.logWarn,
    }),
    handleJellyfinRemotePlaystate: createHandleJellyfinRemotePlaystate({
      getMpvClient: options.getMpvClient,
      sendMpvCommand: options.sendMpvCommand,
      reportJellyfinRemoteProgress: (force) => reportJellyfinRemoteProgress(force),
      reportJellyfinRemoteStopped: () => reportJellyfinRemoteStopped(),
      jellyfinTicksToSeconds: options.jellyfinTicksToSeconds,
    }),
    handleJellyfinRemoteGeneralCommand: createHandleJellyfinRemoteGeneralCommand({
      getMpvClient: options.getMpvClient,
      sendMpvCommand: options.sendMpvCommand,
      getActivePlayback: options.getActivePlayback,
      reportJellyfinRemoteProgress: (force) => reportJellyfinRemoteProgress(force),
      logDebug: (message) => options.logDebug(message, undefined),
    }),
  };
}
