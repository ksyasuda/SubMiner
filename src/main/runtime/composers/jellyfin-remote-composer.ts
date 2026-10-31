import {
  createHandleJellyfinRemoteGeneralCommand,
  createHandleJellyfinRemotePlay,
  createHandleJellyfinRemotePlaystate,
  createJellyfinRemoteReportTracker,
  createReportJellyfinRemoteProgressHandler,
  createReportJellyfinRemoteStoppedHandler,
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
  const reportJellyfinRemoteStopped = createReportJellyfinRemoteStoppedHandler({
    getActivePlayback: options.getActivePlayback,
    clearActivePlayback: options.clearActivePlayback,
    getMpvClient: options.getMpvClient,
    getNow: options.getNow,
    ticksPerSecond: options.ticksPerSecond,
    logDebug: options.logDebug,
    logWarn: options.logWarn,
    reportTracker,
  });

  return {
    reportJellyfinRemoteProgress,
    reportJellyfinRemoteStopped,
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
