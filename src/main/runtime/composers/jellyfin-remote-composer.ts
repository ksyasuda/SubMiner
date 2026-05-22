import {
  createBuildHandleJellyfinRemoteGeneralCommandMainDepsHandler,
  createBuildHandleJellyfinRemotePlayMainDepsHandler,
  createBuildHandleJellyfinRemotePlaystateMainDepsHandler,
  createBuildReportJellyfinRemoteProgressMainDepsHandler,
  createBuildReportJellyfinRemoteStoppedMainDepsHandler,
  createHandleJellyfinRemoteGeneralCommand,
  createHandleJellyfinRemotePlay,
  createHandleJellyfinRemotePlaystate,
  createReportJellyfinRemoteProgressHandler,
  createReportJellyfinRemoteStoppedHandler,
} from '../domains/jellyfin';
import type { ComposerInputs, ComposerOutputs } from './contracts';

type RemotePlayPayload = Parameters<ReturnType<typeof createHandleJellyfinRemotePlay>>[0];
type RemotePlaystatePayload = Parameters<ReturnType<typeof createHandleJellyfinRemotePlaystate>>[0];
type RemoteGeneralPayload = Parameters<
  ReturnType<typeof createHandleJellyfinRemoteGeneralCommand>
>[0];
type JellyfinRemotePlayMainDeps = Parameters<
  typeof createBuildHandleJellyfinRemotePlayMainDepsHandler
>[0];
type JellyfinRemotePlaystateMainDeps = Parameters<
  typeof createBuildHandleJellyfinRemotePlaystateMainDepsHandler
>[0];
type JellyfinRemoteGeneralMainDeps = Parameters<
  typeof createBuildHandleJellyfinRemoteGeneralCommandMainDepsHandler
>[0];
type JellyfinRemoteProgressMainDeps = Parameters<
  typeof createBuildReportJellyfinRemoteProgressMainDepsHandler
>[0];

export type JellyfinRemoteComposerOptions = ComposerInputs<{
  getConfiguredSession: JellyfinRemotePlayMainDeps['getConfiguredSession'];
  getClientInfo: JellyfinRemotePlayMainDeps['getClientInfo'];
  getJellyfinConfig: JellyfinRemotePlayMainDeps['getJellyfinConfig'];
  playJellyfinItem: JellyfinRemotePlayMainDeps['playJellyfinItem'];
  logWarn: JellyfinRemotePlayMainDeps['logWarn'];
  getMpvClient: JellyfinRemoteProgressMainDeps['getMpvClient'];
  sendMpvCommand: JellyfinRemotePlaystateMainDeps['sendMpvCommand'];
  jellyfinTicksToSeconds: Parameters<
    typeof createBuildHandleJellyfinRemotePlaystateMainDepsHandler
  >[0]['jellyfinTicksToSeconds'];
  getActivePlayback: JellyfinRemoteGeneralMainDeps['getActivePlayback'];
  clearActivePlayback: JellyfinRemoteProgressMainDeps['clearActivePlayback'];
  getSession: JellyfinRemoteProgressMainDeps['getSession'];
  getNow: JellyfinRemoteProgressMainDeps['getNow'];
  getLastProgressAtMs: Parameters<
    typeof createBuildReportJellyfinRemoteProgressMainDepsHandler
  >[0]['getLastProgressAtMs'];
  setLastProgressAtMs: Parameters<
    typeof createBuildReportJellyfinRemoteProgressMainDepsHandler
  >[0]['setLastProgressAtMs'];
  progressIntervalMs: number;
  ticksPerSecond: number;
  logDebug: Parameters<
    typeof createBuildReportJellyfinRemoteProgressMainDepsHandler
  >[0]['logDebug'];
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
  const buildReportJellyfinRemoteProgressMainDepsHandler =
    createBuildReportJellyfinRemoteProgressMainDepsHandler({
      getActivePlayback: options.getActivePlayback,
      clearActivePlayback: options.clearActivePlayback,
      getSession: options.getSession,
      getMpvClient: options.getMpvClient,
      getNow: options.getNow,
      getLastProgressAtMs: options.getLastProgressAtMs,
      setLastProgressAtMs: options.setLastProgressAtMs,
      progressIntervalMs: options.progressIntervalMs,
      ticksPerSecond: options.ticksPerSecond,
      logDebug: options.logDebug,
    });
  const buildReportJellyfinRemoteStoppedMainDepsHandler =
    createBuildReportJellyfinRemoteStoppedMainDepsHandler({
      getActivePlayback: options.getActivePlayback,
      clearActivePlayback: options.clearActivePlayback,
      getSession: options.getSession,
      logDebug: options.logDebug,
    });
  const reportJellyfinRemoteProgress = createReportJellyfinRemoteProgressHandler(
    buildReportJellyfinRemoteProgressMainDepsHandler(),
  );
  const reportJellyfinRemoteStopped = createReportJellyfinRemoteStoppedHandler(
    buildReportJellyfinRemoteStoppedMainDepsHandler(),
  );

  const buildHandleJellyfinRemotePlayMainDepsHandler =
    createBuildHandleJellyfinRemotePlayMainDepsHandler({
      getConfiguredSession: options.getConfiguredSession,
      getClientInfo: options.getClientInfo,
      getJellyfinConfig: options.getJellyfinConfig,
      getActivePlayback: options.getActivePlayback,
      playJellyfinItem: options.playJellyfinItem,
      logWarn: options.logWarn,
    });
  const buildHandleJellyfinRemotePlaystateMainDepsHandler =
    createBuildHandleJellyfinRemotePlaystateMainDepsHandler({
      getMpvClient: options.getMpvClient,
      sendMpvCommand: options.sendMpvCommand,
      reportJellyfinRemoteProgress: (force) => reportJellyfinRemoteProgress(force),
      reportJellyfinRemoteStopped: () => reportJellyfinRemoteStopped(),
      jellyfinTicksToSeconds: options.jellyfinTicksToSeconds,
    });
  const buildHandleJellyfinRemoteGeneralCommandMainDepsHandler =
    createBuildHandleJellyfinRemoteGeneralCommandMainDepsHandler({
      getMpvClient: options.getMpvClient,
      sendMpvCommand: options.sendMpvCommand,
      getActivePlayback: options.getActivePlayback,
      reportJellyfinRemoteProgress: (force) => reportJellyfinRemoteProgress(force),
      logDebug: (message) => options.logDebug(message, undefined),
    });

  return {
    reportJellyfinRemoteProgress,
    reportJellyfinRemoteStopped,
    handleJellyfinRemotePlay: createHandleJellyfinRemotePlay(
      buildHandleJellyfinRemotePlayMainDepsHandler(),
    ),
    handleJellyfinRemotePlaystate: createHandleJellyfinRemotePlaystate(
      buildHandleJellyfinRemotePlaystateMainDepsHandler(),
    ),
    handleJellyfinRemoteGeneralCommand: createHandleJellyfinRemoteGeneralCommand(
      buildHandleJellyfinRemoteGeneralCommandMainDepsHandler(),
    ),
  };
}
