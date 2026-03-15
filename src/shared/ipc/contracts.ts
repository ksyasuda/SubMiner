import type { OverlayContentMeasurement, RuntimeOptionId, RuntimeOptionValue } from '../../types';

export const OVERLAY_HOSTED_MODALS = [
  'runtime-options',
  'subsync',
  'jimaku',
  'kiku',
  'controller-select',
  'controller-debug',
] as const;
export type OverlayHostedModal = (typeof OVERLAY_HOSTED_MODALS)[number];

export const IPC_CHANNELS = {
  command: {
    setIgnoreMouseEvents: 'set-ignore-mouse-events',
    overlayModalClosed: 'overlay:modal-closed',
    openYomitanSettings: 'open-yomitan-settings',
    quitApp: 'quit-app',
    toggleDevTools: 'toggle-dev-tools',
    toggleOverlay: 'toggle-overlay',
    saveSubtitlePosition: 'save-subtitle-position',
    saveControllerConfig: 'save-controller-config',
    saveControllerPreference: 'save-controller-preference',
    setMecabEnabled: 'set-mecab-enabled',
    mpvCommand: 'mpv-command',
    setAnkiConnectEnabled: 'set-anki-connect-enabled',
    clearAnkiConnectHistory: 'clear-anki-connect-history',
    refreshKnownWords: 'anki:refresh-known-words',
    kikuFieldGroupingRespond: 'kiku:field-grouping-respond',
    reportOverlayContentBounds: 'overlay-content-bounds:report',
    overlayModalOpened: 'overlay:modal-opened',
    toggleStatsOverlay: 'stats:toggle-overlay',
  },
  request: {
    getVisibleOverlayVisibility: 'get-visible-overlay-visibility',
    getCurrentSubtitle: 'get-current-subtitle',
    getCurrentSubtitleRaw: 'get-current-subtitle-raw',
    getCurrentSubtitleAss: 'get-current-subtitle-ass',
    getPlaybackPaused: 'get-playback-paused',
    getSubtitlePosition: 'get-subtitle-position',
    getSubtitleStyle: 'get-subtitle-style',
    getMecabStatus: 'get-mecab-status',
    getKeybindings: 'get-keybindings',
    getConfigShortcuts: 'get-config-shortcuts',
    getStatsToggleKey: 'get-stats-toggle-key',
    getControllerConfig: 'get-controller-config',
    getSecondarySubMode: 'get-secondary-sub-mode',
    getCurrentSecondarySub: 'get-current-secondary-sub',
    focusMainWindow: 'focus-main-window',
    runSubsyncManual: 'subsync:run-manual',
    getAnkiConnectStatus: 'get-anki-connect-status',
    getRuntimeOptions: 'runtime-options:get',
    setRuntimeOption: 'runtime-options:set',
    cycleRuntimeOption: 'runtime-options:cycle',
    getAnilistStatus: 'anilist:get-status',
    clearAnilistToken: 'anilist:clear-token',
    openAnilistSetup: 'anilist:open-setup',
    getAnilistQueueStatus: 'anilist:get-queue-status',
    retryAnilistNow: 'anilist:retry-now',
    appendClipboardVideoToQueue: 'clipboard:append-video-to-queue',
    jimakuGetMediaInfo: 'jimaku:get-media-info',
    jimakuSearchEntries: 'jimaku:search-entries',
    jimakuListFiles: 'jimaku:list-files',
    jimakuDownloadFile: 'jimaku:download-file',
    kikuBuildMergePreview: 'kiku:build-merge-preview',
    statsGetOverview: 'stats:get-overview',
    statsGetDailyRollups: 'stats:get-daily-rollups',
    statsGetMonthlyRollups: 'stats:get-monthly-rollups',
    statsGetSessions: 'stats:get-sessions',
    statsGetSessionTimeline: 'stats:get-session-timeline',
    statsGetSessionEvents: 'stats:get-session-events',
    statsGetVocabulary: 'stats:get-vocabulary',
    statsGetKanji: 'stats:get-kanji',
    statsGetMediaLibrary: 'stats:get-media-library',
    statsGetMediaDetail: 'stats:get-media-detail',
    statsGetMediaSessions: 'stats:get-media-sessions',
    statsGetMediaDailyRollups: 'stats:get-media-daily-rollups',
    statsGetMediaCover: 'stats:get-media-cover',
  },
  event: {
    subtitleSet: 'subtitle:set',
    subtitleVisibility: 'mpv:subVisibility',
    subtitlePositionSet: 'subtitle-position:set',
    subtitleAssSet: 'subtitle-ass:set',
    secondarySubtitleSet: 'secondary-subtitle:set',
    secondarySubtitleMode: 'secondary-subtitle:mode',
    subsyncOpenManual: 'subsync:open-manual',
    kikuFieldGroupingRequest: 'kiku:field-grouping-request',
    runtimeOptionsChanged: 'runtime-options:changed',
    runtimeOptionsOpen: 'runtime-options:open',
    jimakuOpen: 'jimaku:open',
    keyboardModeToggleRequested: 'keyboard-mode-toggle:requested',
    lookupWindowToggleRequested: 'lookup-window-toggle:requested',
    configHotReload: 'config:hot-reload',
  },
} as const;

export type RuntimeOptionsSetRequest = {
  id: RuntimeOptionId;
  value: RuntimeOptionValue;
};

export type RuntimeOptionsCycleRequest = {
  id: RuntimeOptionId;
  direction: 1 | -1;
};

export type OverlayContentBoundsReportRequest = {
  measurement: OverlayContentMeasurement;
};
