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

type RemotePlayPayload = Parameters<ReturnType<typeof createHandleJellyfinRemotePlay>>[0];
type RemotePlaystatePayload = Parameters<ReturnType<typeof createHandleJellyfinRemotePlaystate>>[0];
type RemoteGeneralPayload = Parameters<ReturnType<typeof createHandleJellyfinRemoteGeneralCommand>>[0];

export type JellyfinRemoteComposerOptions = {
  getConfiguredSession: Parameters<typeof createBuildHandleJellyfinRemotePlayMainDepsHandler>[0]['getConfiguredSession'];
  getClientInfo: Parameters<typeof createBuildHandleJellyfinRemotePlayMainDepsHandler>[0]['getClientInfo'];
  getJellyfinConfig: Parameters<typeof createBuildHandleJellyfinRemotePlayMainDepsHandler>[0]['getJellyfinConfig'];
  playJellyfinItem: Parameters<typeof createBuildHandleJellyfinRemotePlayMainDepsHandler>[0]['playJellyfinItem'];
  logWarn: Parameters<typeof createBuildHandleJellyfinRemotePlayMainDepsHandler>[0]['logWarn'];
  getMpvClient: Parameters<typeof createBuildReportJellyfinRemoteProgressMainDepsHandler>[0]['getMpvClient'];
  sendMpvCommand: Parameters<typeof createBuildHandleJellyfinRemotePlaystateMainDepsHandler>[0]['sendMpvCommand'];
  jellyfinTicksToSeconds: Parameters<
    typeof createBuildHandleJellyfinRemotePlaystateMainDepsHandler
  >[0]['jellyfinTicksToSeconds'];
  getActivePlayback: Parameters<typeof createBuildHandleJellyfinRemoteGeneralCommandMainDepsHandler>[0]['getActivePlayback'];
  clearActivePlayback: Parameters<typeof createBuildReportJellyfinRemoteProgressMainDepsHandler>[0]['clearActivePlayback'];
  getSession: Parameters<typeof createBuildReportJellyfinRemoteProgressMainDepsHandler>[0]['getSession'];
  getNow: Parameters<typeof createBuildReportJellyfinRemoteProgressMainDepsHandler>[0]['getNow'];
  getLastProgressAtMs: Parameters<
    typeof createBuildReportJellyfinRemoteProgressMainDepsHandler
  >[0]['getLastProgressAtMs'];
  setLastProgressAtMs: Parameters<
    typeof createBuildReportJellyfinRemoteProgressMainDepsHandler
  >[0]['setLastProgressAtMs'];
  progressIntervalMs: number;
  ticksPerSecond: number;
  logDebug: Parameters<typeof createBuildReportJellyfinRemoteProgressMainDepsHandler>[0]['logDebug'];
};

export type JellyfinRemoteComposerResult = {
  reportJellyfinRemoteProgress: ReturnType<typeof createReportJellyfinRemoteProgressHandler>;
  reportJellyfinRemoteStopped: ReturnType<typeof createReportJellyfinRemoteStoppedHandler>;
  handleJellyfinRemotePlay: (payload: RemotePlayPayload) => Promise<void>;
  handleJellyfinRemotePlaystate: (payload: RemotePlaystatePayload) => Promise<void>;
  handleJellyfinRemoteGeneralCommand: (payload: RemoteGeneralPayload) => Promise<void>;
};

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
      playJellyfinItem: options.playJellyfinItem,
      logWarn: options.logWarn,
    });
  const buildHandleJellyfinRemotePlaystateMainDepsHandler =
    createBuildHandleJellyfinRemotePlaystateMainDepsHandler({
      getMpvClient: options.getMpvClient as Parameters<
        typeof createBuildHandleJellyfinRemotePlaystateMainDepsHandler
      >[0]['getMpvClient'],
      sendMpvCommand: options.sendMpvCommand,
      reportJellyfinRemoteProgress: (force) => reportJellyfinRemoteProgress(force),
      reportJellyfinRemoteStopped: () => reportJellyfinRemoteStopped(),
      jellyfinTicksToSeconds: options.jellyfinTicksToSeconds,
    });
  const buildHandleJellyfinRemoteGeneralCommandMainDepsHandler =
    createBuildHandleJellyfinRemoteGeneralCommandMainDepsHandler({
      getMpvClient: options.getMpvClient as Parameters<
        typeof createBuildHandleJellyfinRemoteGeneralCommandMainDepsHandler
      >[0]['getMpvClient'],
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
