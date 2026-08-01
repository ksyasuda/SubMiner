import {
  buildJellyfinSetupFormHtml,
  createEnsureMpvConnectedForJellyfinPlaybackHandler,
  createBuildEnsureMpvConnectedForJellyfinPlaybackMainDepsHandler,
  createBuildGetJellyfinClientInfoMainDepsHandler,
  createBuildGetResolvedJellyfinConfigMainDepsHandler,
  createBuildHandleJellyfinAuthCommandsMainDepsHandler,
  createBuildHandleJellyfinListCommandsMainDepsHandler,
  createBuildHandleJellyfinPlayCommandMainDepsHandler,
  createBuildHandleJellyfinRemoteAnnounceCommandMainDepsHandler,
  createBuildLaunchMpvIdleForJellyfinPlaybackMainDepsHandler,
  createBuildOpenJellyfinSetupWindowMainDepsHandler,
  createBuildPlayJellyfinItemInMpvMainDepsHandler,
  createBuildPreloadJellyfinExternalSubtitlesMainDepsHandler,
  createBuildRunJellyfinCommandMainDepsHandler,
  createBuildStartJellyfinRemoteSessionMainDepsHandler,
  createBuildStopJellyfinRemoteSessionMainDepsHandler,
  createBuildWaitForMpvConnectedMainDepsHandler,
  createGetJellyfinClientInfoHandler,
  createGetResolvedJellyfinConfigHandler,
  createHandleJellyfinAuthCommands,
  createHandleJellyfinListCommands,
  createHandleJellyfinPlayCommand,
  createHandleJellyfinRemoteAnnounceCommand,
  createLaunchMpvIdleForJellyfinPlaybackHandler,
  createOpenJellyfinSetupWindowHandler,
  createPlayJellyfinItemInMpvHandler,
  createPreloadJellyfinExternalSubtitlesHandler,
  createRunJellyfinCommandHandler,
  createStartJellyfinRemoteSessionHandler,
  createStopJellyfinRemoteSessionHandler,
  createWaitForMpvConnectedHandler,
  createMaybeFocusExistingJellyfinSetupWindowHandler,
  parseJellyfinSetupSubmissionUrl,
} from '../domains/jellyfin';
import {
  composeJellyfinRemoteHandlers,
  type JellyfinRemoteComposerOptions,
} from './jellyfin-remote-composer';
import type { ComposerInputs, ComposerOutputs } from './contracts';

type EnsureMpvConnectedMainDeps = Parameters<
  typeof createBuildEnsureMpvConnectedForJellyfinPlaybackMainDepsHandler
>[0];
type PlayJellyfinItemMainDeps = Parameters<
  typeof createBuildPlayJellyfinItemInMpvMainDepsHandler
>[0];
type HandlePlayCommandMainDeps = Parameters<
  typeof createBuildHandleJellyfinPlayCommandMainDepsHandler
>[0];
type HandleRemoteAnnounceMainDeps = Parameters<
  typeof createBuildHandleJellyfinRemoteAnnounceCommandMainDepsHandler
>[0];
type StartRemoteSessionMainDeps = Parameters<
  typeof createBuildStartJellyfinRemoteSessionMainDepsHandler
>[0];
type RunJellyfinCommandMainDeps = Parameters<
  typeof createBuildRunJellyfinCommandMainDepsHandler
>[0];
type OpenJellyfinSetupWindowMainDeps = Parameters<
  typeof createBuildOpenJellyfinSetupWindowMainDepsHandler
>[0];

export type JellyfinRuntimeComposerOptions = ComposerInputs<{
  getResolvedJellyfinConfigMainDeps: Parameters<
    typeof createBuildGetResolvedJellyfinConfigMainDepsHandler
  >[0];
  getJellyfinClientInfoMainDeps: Parameters<
    typeof createBuildGetJellyfinClientInfoMainDepsHandler
  >[0];
  waitForMpvConnectedMainDeps: Parameters<typeof createBuildWaitForMpvConnectedMainDepsHandler>[0];
  launchMpvIdleForJellyfinPlaybackMainDeps: Parameters<
    typeof createBuildLaunchMpvIdleForJellyfinPlaybackMainDepsHandler
  >[0];
  ensureMpvConnectedForJellyfinPlaybackMainDeps: Omit<
    EnsureMpvConnectedMainDeps,
    'waitForMpvConnected' | 'launchMpvIdleForJellyfinPlayback'
  >;
  preloadJellyfinExternalSubtitlesMainDeps: Parameters<
    typeof createBuildPreloadJellyfinExternalSubtitlesMainDepsHandler
  >[0];
  playJellyfinItemInMpvMainDeps: Omit<
    PlayJellyfinItemMainDeps,
    'ensureMpvConnectedForPlayback' | 'preloadExternalSubtitles'
  >;
  remoteComposerOptions: Omit<
    JellyfinRemoteComposerOptions,
    'getClientInfo' | 'getJellyfinConfig' | 'playJellyfinItem'
  >;
  handleJellyfinAuthCommandsMainDeps: Parameters<
    typeof createBuildHandleJellyfinAuthCommandsMainDepsHandler
  >[0];
  handleJellyfinListCommandsMainDeps: Parameters<
    typeof createBuildHandleJellyfinListCommandsMainDepsHandler
  >[0];
  handleJellyfinPlayCommandMainDeps: Omit<HandlePlayCommandMainDeps, 'playJellyfinItemInMpv'>;
  handleJellyfinRemoteAnnounceCommandMainDeps: Omit<
    HandleRemoteAnnounceMainDeps,
    'startJellyfinRemoteSession'
  >;
  startJellyfinRemoteSessionMainDeps: Omit<
    StartRemoteSessionMainDeps,
    | 'getJellyfinConfig'
    | 'getClientInfo'
    | 'handlePlay'
    | 'handlePlaystate'
    | 'handleGeneralCommand'
  >;
  stopJellyfinRemoteSessionMainDeps: Parameters<
    typeof createBuildStopJellyfinRemoteSessionMainDepsHandler
  >[0];
  runJellyfinCommandMainDeps: Omit<
    RunJellyfinCommandMainDeps,
    | 'getJellyfinConfig'
    | 'getJellyfinClientInfo'
    | 'handleAuthCommands'
    | 'handleRemoteAnnounceCommand'
    | 'handleListCommands'
    | 'handlePlayCommand'
  >;
  maybeFocusExistingJellyfinSetupWindowMainDeps: Parameters<
    typeof createMaybeFocusExistingJellyfinSetupWindowHandler
  >[0];
  openJellyfinSetupWindowMainDeps: Omit<
    OpenJellyfinSetupWindowMainDeps,
    'maybeFocusExistingSetupWindow' | 'getResolvedJellyfinConfig' | 'getJellyfinClientInfo'
  >;
}>;

export type JellyfinRuntimeComposerResult = ComposerOutputs<{
  getResolvedJellyfinConfig: ReturnType<typeof createGetResolvedJellyfinConfigHandler>;
  getJellyfinClientInfo: ReturnType<typeof createGetJellyfinClientInfoHandler>;
  ensureMpvConnectedForPlayback: ReturnType<
    typeof createEnsureMpvConnectedForJellyfinPlaybackHandler
  >;
  reportJellyfinRemoteProgress: ReturnType<
    typeof composeJellyfinRemoteHandlers
  >['reportJellyfinRemoteProgress'];
  reportJellyfinRemoteStopped: ReturnType<
    typeof composeJellyfinRemoteHandlers
  >['reportJellyfinRemoteStopped'];
  handleJellyfinRemotePlay: ReturnType<
    typeof composeJellyfinRemoteHandlers
  >['handleJellyfinRemotePlay'];
  handleJellyfinRemotePlaystate: ReturnType<
    typeof composeJellyfinRemoteHandlers
  >['handleJellyfinRemotePlaystate'];
  handleJellyfinRemoteGeneralCommand: ReturnType<
    typeof composeJellyfinRemoteHandlers
  >['handleJellyfinRemoteGeneralCommand'];
  playJellyfinItemInMpv: ReturnType<typeof createPlayJellyfinItemInMpvHandler>;
  cleanupJellyfinSubtitleCache: () => void;
  startJellyfinRemoteSession: ReturnType<typeof createStartJellyfinRemoteSessionHandler>;
  stopJellyfinRemoteSession: ReturnType<typeof createStopJellyfinRemoteSessionHandler>;
  runJellyfinCommand: ReturnType<typeof createRunJellyfinCommandHandler>;
  openJellyfinSetupWindow: ReturnType<typeof createOpenJellyfinSetupWindowHandler>;
}>;

export function createRestartJellyfinRemoteSessionAfterSetupLoginHandler(deps: {
  getCurrentSession: () => unknown | null;
  startJellyfinRemoteSession: (options?: { explicit?: boolean }) => Promise<void>;
}) {
  return async (): Promise<void> => {
    const hasActiveSession = deps.getCurrentSession() !== null;
    await deps.startJellyfinRemoteSession(hasActiveSession ? { explicit: true } : undefined);
  };
}

export function composeJellyfinRuntimeHandlers(
  options: JellyfinRuntimeComposerOptions,
): JellyfinRuntimeComposerResult {
  const getResolvedJellyfinConfig = createGetResolvedJellyfinConfigHandler(
    createBuildGetResolvedJellyfinConfigMainDepsHandler(
      options.getResolvedJellyfinConfigMainDeps,
    )(),
  );
  const getJellyfinClientInfo = createGetJellyfinClientInfoHandler(
    createBuildGetJellyfinClientInfoMainDepsHandler(options.getJellyfinClientInfoMainDeps)(),
  );

  const waitForMpvConnected = createWaitForMpvConnectedHandler(
    createBuildWaitForMpvConnectedMainDepsHandler(options.waitForMpvConnectedMainDeps)(),
  );
  const launchMpvIdleForJellyfinPlayback = createLaunchMpvIdleForJellyfinPlaybackHandler(
    createBuildLaunchMpvIdleForJellyfinPlaybackMainDepsHandler(
      options.launchMpvIdleForJellyfinPlaybackMainDeps,
    )(),
  );
  const ensureMpvConnectedForJellyfinPlayback = createEnsureMpvConnectedForJellyfinPlaybackHandler(
    createBuildEnsureMpvConnectedForJellyfinPlaybackMainDepsHandler({
      ...options.ensureMpvConnectedForJellyfinPlaybackMainDeps,
      waitForMpvConnected: (timeoutMs) => waitForMpvConnected(timeoutMs),
      launchMpvIdleForJellyfinPlayback: () => launchMpvIdleForJellyfinPlayback(),
    })(),
  );

  const preloadJellyfinExternalSubtitles = createPreloadJellyfinExternalSubtitlesHandler(
    createBuildPreloadJellyfinExternalSubtitlesMainDepsHandler(
      options.preloadJellyfinExternalSubtitlesMainDeps,
    )(),
  );
  const playJellyfinItemInMpv = createPlayJellyfinItemInMpvHandler(
    createBuildPlayJellyfinItemInMpvMainDepsHandler({
      ...options.playJellyfinItemInMpvMainDeps,
      ensureMpvConnectedForPlayback: () => ensureMpvConnectedForJellyfinPlayback(),
      preloadExternalSubtitles: (params) => {
        void preloadJellyfinExternalSubtitles(params);
      },
    })(),
  );

  const {
    reportJellyfinRemoteProgress,
    reportJellyfinRemoteStopped,
    handleJellyfinRemotePlay,
    handleJellyfinRemotePlaystate,
    handleJellyfinRemoteGeneralCommand,
  } = composeJellyfinRemoteHandlers({
    ...options.remoteComposerOptions,
    getClientInfo: () => getJellyfinClientInfo(),
    getJellyfinConfig: () => getResolvedJellyfinConfig(),
    playJellyfinItem: (params) =>
      playJellyfinItemInMpv(params as Parameters<typeof playJellyfinItemInMpv>[0]),
  });

  const handleJellyfinAuthCommands = createHandleJellyfinAuthCommands(
    createBuildHandleJellyfinAuthCommandsMainDepsHandler(
      options.handleJellyfinAuthCommandsMainDeps,
    )(),
  );
  const handleJellyfinListCommands = createHandleJellyfinListCommands(
    createBuildHandleJellyfinListCommandsMainDepsHandler(
      options.handleJellyfinListCommandsMainDeps,
    )(),
  );
  const handleJellyfinPlayCommand = createHandleJellyfinPlayCommand(
    createBuildHandleJellyfinPlayCommandMainDepsHandler({
      ...options.handleJellyfinPlayCommandMainDeps,
      playJellyfinItemInMpv: (params) =>
        playJellyfinItemInMpv(params as Parameters<typeof playJellyfinItemInMpv>[0]),
    })(),
  );

  let startJellyfinRemoteSession!: ReturnType<typeof createStartJellyfinRemoteSessionHandler>;
  const handleJellyfinRemoteAnnounceCommand = createHandleJellyfinRemoteAnnounceCommand(
    createBuildHandleJellyfinRemoteAnnounceCommandMainDepsHandler({
      ...options.handleJellyfinRemoteAnnounceCommandMainDeps,
      startJellyfinRemoteSession: (startOptions) => startJellyfinRemoteSession(startOptions),
    })(),
  );

  startJellyfinRemoteSession = createStartJellyfinRemoteSessionHandler(
    createBuildStartJellyfinRemoteSessionMainDepsHandler({
      ...options.startJellyfinRemoteSessionMainDeps,
      getJellyfinConfig: () => getResolvedJellyfinConfig(),
      getClientInfo: () => getJellyfinClientInfo(),
      handlePlay: (payload) => handleJellyfinRemotePlay(payload),
      handlePlaystate: (payload) => handleJellyfinRemotePlaystate(payload),
      handleGeneralCommand: (payload) => handleJellyfinRemoteGeneralCommand(payload),
    })(),
  );

  const stopJellyfinRemoteSession = createStopJellyfinRemoteSessionHandler(
    createBuildStopJellyfinRemoteSessionMainDepsHandler(
      options.stopJellyfinRemoteSessionMainDeps,
    )(),
  );

  const runJellyfinCommand = createRunJellyfinCommandHandler(
    createBuildRunJellyfinCommandMainDepsHandler({
      ...options.runJellyfinCommandMainDeps,
      getJellyfinConfig: () => getResolvedJellyfinConfig(),
      getJellyfinClientInfo: (jellyfinConfig) => getJellyfinClientInfo(jellyfinConfig),
      handleAuthCommands: (params) => handleJellyfinAuthCommands(params),
      handleRemoteAnnounceCommand: (args) => handleJellyfinRemoteAnnounceCommand(args),
      handleListCommands: (params) => handleJellyfinListCommands(params),
      handlePlayCommand: (params) => handleJellyfinPlayCommand(params),
    })(),
  );

  const maybeFocusExistingJellyfinSetupWindow = createMaybeFocusExistingJellyfinSetupWindowHandler(
    options.maybeFocusExistingJellyfinSetupWindowMainDeps,
  );
  const restartJellyfinRemoteSessionAfterSetupLogin =
    createRestartJellyfinRemoteSessionAfterSetupLoginHandler({
      getCurrentSession: () => options.startJellyfinRemoteSessionMainDeps.getCurrentSession(),
      startJellyfinRemoteSession: (startOptions) => startJellyfinRemoteSession(startOptions),
    });
  const openJellyfinSetupWindow = createOpenJellyfinSetupWindowHandler(
    createBuildOpenJellyfinSetupWindowMainDepsHandler({
      ...options.openJellyfinSetupWindowMainDeps,
      maybeFocusExistingSetupWindow: maybeFocusExistingJellyfinSetupWindow,
      getResolvedJellyfinConfig: () => getResolvedJellyfinConfig(),
      getJellyfinClientInfo: () => getJellyfinClientInfo(),
      restartRemoteSession: () => restartJellyfinRemoteSessionAfterSetupLogin(),
      stopRemoteSession: () => stopJellyfinRemoteSession(),
    })(),
  );

  return {
    getResolvedJellyfinConfig,
    getJellyfinClientInfo,
    // Shared so other playback sources (the anime browser) reuse the same
    // auto-launch in-flight guard instead of racing a second mpv launch.
    ensureMpvConnectedForPlayback: ensureMpvConnectedForJellyfinPlayback,
    reportJellyfinRemoteProgress,
    reportJellyfinRemoteStopped,
    handleJellyfinRemotePlay,
    handleJellyfinRemotePlaystate,
    handleJellyfinRemoteGeneralCommand,
    playJellyfinItemInMpv,
    cleanupJellyfinSubtitleCache: () => preloadJellyfinExternalSubtitles.cleanupCachedSubtitles(),
    startJellyfinRemoteSession,
    stopJellyfinRemoteSession,
    runJellyfinCommand,
    openJellyfinSetupWindow,
  };
}

export { buildJellyfinSetupFormHtml, parseJellyfinSetupSubmissionUrl };
