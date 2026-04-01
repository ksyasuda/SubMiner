import type {
  AnkiJimakuIpcRuntimeServiceDepsParams,
  MainIpcRuntimeServiceDepsParams,
  RuntimeOptionsIpcDepsParams,
} from './dependencies';
import { createAnkiJimakuIpcRuntimeServiceDeps } from './dependencies';
import {
  handleMpvCommandFromIpcRuntime,
  type MpvCommandFromIpcRuntimeDeps,
} from './ipc-mpv-command';
import { registerIpcRuntimeServices } from './ipc-runtime-services';
import { composeIpcRuntimeHandlers } from './runtime/composers/ipc-runtime-composer';

export interface RegisterIpcRuntimeServicesParams {
  runtimeOptions: RuntimeOptionsIpcDepsParams;
  mainDeps: Omit<MainIpcRuntimeServiceDepsParams, 'setRuntimeOption' | 'cycleRuntimeOption'>;
  ankiJimakuDeps: AnkiJimakuIpcRuntimeServiceDepsParams;
}

export interface IpcRuntimeMainInput {
  window: Pick<
    RegisterIpcRuntimeServicesParams['mainDeps'],
    | 'getMainWindow'
    | 'getVisibleOverlayVisibility'
    | 'focusMainWindow'
    | 'onOverlayModalClosed'
    | 'onOverlayModalOpened'
    | 'onYoutubePickerResolve'
    | 'openYomitanSettings'
    | 'quitApp'
    | 'toggleVisibleOverlay'
  >;
  subtitle: Pick<
    RegisterIpcRuntimeServicesParams['mainDeps'],
    | 'tokenizeCurrentSubtitle'
    | 'getCurrentSubtitleRaw'
    | 'getCurrentSubtitleAss'
    | 'getSubtitleSidebarSnapshot'
    | 'getPlaybackPaused'
    | 'getSubtitlePosition'
    | 'getSubtitleStyle'
    | 'saveSubtitlePosition'
    | 'getMecabTokenizer'
    | 'getKeybindings'
    | 'getConfiguredShortcuts'
    | 'getStatsToggleKey'
    | 'getMarkWatchedKey'
    | 'getSecondarySubMode'
  >;
  controller: Pick<
    RegisterIpcRuntimeServicesParams['mainDeps'],
    'getControllerConfig' | 'saveControllerConfig' | 'saveControllerPreference'
  >;
  runtime: Pick<
    RegisterIpcRuntimeServicesParams['mainDeps'],
    'getMpvClient' | 'getAnkiConnectStatus' | 'getRuntimeOptions' | 'reportOverlayContentBounds'
  > &
    Partial<Pick<RegisterIpcRuntimeServicesParams['mainDeps'], 'getImmersionTracker'>>;
  anilist: {
    getStatus: RegisterIpcRuntimeServicesParams['mainDeps']['getAnilistStatus'];
    clearToken: RegisterIpcRuntimeServicesParams['mainDeps']['clearAnilistToken'];
    openSetup: RegisterIpcRuntimeServicesParams['mainDeps']['openAnilistSetup'];
    getQueueStatus: RegisterIpcRuntimeServicesParams['mainDeps']['getAnilistQueueStatus'];
    retryQueueNow: RegisterIpcRuntimeServicesParams['mainDeps']['retryAnilistQueueNow'];
  };
  mining: {
    appendClipboardVideoToQueue: RegisterIpcRuntimeServicesParams['mainDeps']['appendClipboardVideoToQueue'];
  };
}

export interface IpcRuntimeRegistrationInput {
  runtimeOptions: RuntimeOptionsIpcDepsParams;
  main: IpcRuntimeMainInput;
  ankiJimaku: AnkiJimakuIpcRuntimeServiceDepsParams;
  registerIpcRuntimeServices: (params: RegisterIpcRuntimeServicesParams) => void;
}

export interface IpcRuntimeInput {
  mpv: {
    mainDeps: MpvCommandFromIpcRuntimeDeps;
    handleMpvCommandFromIpcRuntime: (
      command: (string | number)[],
      deps: MpvCommandFromIpcRuntimeDeps,
    ) => void;
    runSubsyncManualFromIpc: MainIpcRuntimeServiceDepsParams['runSubsyncManual'];
  };
  registration: IpcRuntimeRegistrationInput;
}

export interface IpcRuntime {
  registerIpcRuntimeHandlers: () => void;
}

export interface IpcRuntimeFromMainStateInput {
  mpv: {
    mainDeps: MpvCommandFromIpcRuntimeDeps;
    runSubsyncManualFromIpc: MainIpcRuntimeServiceDepsParams['runSubsyncManual'];
  };
  runtimeOptions: RuntimeOptionsIpcDepsParams;
  main: IpcRuntimeMainInput;
  ankiJimaku: AnkiJimakuIpcRuntimeServiceDepsParams;
}

export function createIpcRuntime(input: IpcRuntimeInput): IpcRuntime {
  const { registerIpcRuntimeHandlers } = composeIpcRuntimeHandlers({
    mpvCommandMainDeps: input.mpv.mainDeps,
    handleMpvCommandFromIpcRuntime: input.mpv.handleMpvCommandFromIpcRuntime,
    runSubsyncManualFromIpc: input.mpv.runSubsyncManualFromIpc,
    registration: {
      runtimeOptions: input.registration.runtimeOptions,
      mainDeps: {
        ...input.registration.main.window,
        ...input.registration.main.subtitle,
        ...input.registration.main.controller,
        ...input.registration.main.runtime,
        getAnilistStatus: input.registration.main.anilist.getStatus,
        clearAnilistToken: input.registration.main.anilist.clearToken,
        openAnilistSetup: input.registration.main.anilist.openSetup,
        getAnilistQueueStatus: input.registration.main.anilist.getQueueStatus,
        retryAnilistQueueNow: input.registration.main.anilist.retryQueueNow,
        appendClipboardVideoToQueue: input.registration.main.mining.appendClipboardVideoToQueue,
      },
      ankiJimakuDeps: createAnkiJimakuIpcRuntimeServiceDeps(input.registration.ankiJimaku),
      registerIpcRuntimeServices: (params) => input.registration.registerIpcRuntimeServices(params),
    },
  });

  return {
    registerIpcRuntimeHandlers,
  };
}

export function createIpcRuntimeFromMainState(input: IpcRuntimeFromMainStateInput): IpcRuntime {
  return createIpcRuntime({
    mpv: {
      mainDeps: input.mpv.mainDeps,
      handleMpvCommandFromIpcRuntime,
      runSubsyncManualFromIpc: input.mpv.runSubsyncManualFromIpc,
    },
    registration: {
      runtimeOptions: input.runtimeOptions,
      main: input.main,
      ankiJimaku: input.ankiJimaku,
      registerIpcRuntimeServices,
    },
  });
}
