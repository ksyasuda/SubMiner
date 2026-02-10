export const IPC_CHANNELS = {
  rendererToMainInvoke: {
    getOverlayVisibility: "get-overlay-visibility",
    getVisibleOverlayVisibility: "get-visible-overlay-visibility",
    getInvisibleOverlayVisibility: "get-invisible-overlay-visibility",
    getCurrentSubtitle: "get-current-subtitle",
    getCurrentSubtitleAss: "get-current-subtitle-ass",
    getMpvSubtitleRenderMetrics: "get-mpv-subtitle-render-metrics",
    getSubtitlePosition: "get-subtitle-position",
    getSubtitleStyle: "get-subtitle-style",
    getMecabStatus: "get-mecab-status",
    getKeybindings: "get-keybindings",
    getSecondarySubMode: "get-secondary-sub-mode",
    getCurrentSecondarySub: "get-current-secondary-sub",
    runSubsyncManual: "subsync:run-manual",
    getAnkiConnectStatus: "get-anki-connect-status",
    runtimeOptionsGet: "runtime-options:get",
    runtimeOptionsSet: "runtime-options:set",
    runtimeOptionsCycle: "runtime-options:cycle",
    kikuBuildMergePreview: "kiku:build-merge-preview",
    jimakuGetMediaInfo: "jimaku:get-media-info",
    jimakuSearchEntries: "jimaku:search-entries",
    jimakuListFiles: "jimaku:list-files",
    jimakuDownloadFile: "jimaku:download-file",
  },
  rendererToMainSend: {
    setIgnoreMouseEvents: "set-ignore-mouse-events",
    overlayModalClosed: "overlay:modal-closed",
    openYomitanSettings: "open-yomitan-settings",
    quitApp: "quit-app",
    toggleDevTools: "toggle-dev-tools",
    toggleOverlay: "toggle-overlay",
    saveSubtitlePosition: "save-subtitle-position",
    setMecabEnabled: "set-mecab-enabled",
    mpvCommand: "mpv-command",
    setAnkiConnectEnabled: "set-anki-connect-enabled",
    clearAnkiConnectHistory: "clear-anki-connect-history",
    kikuFieldGroupingRespond: "kiku:field-grouping-respond",
  },
  mainToRendererEvent: {
    subtitleSet: "subtitle:set",
    mpvSubVisibility: "mpv:subVisibility",
    subtitlePositionSet: "subtitle-position:set",
    mpvSubtitleRenderMetricsSet: "mpv-subtitle-render-metrics:set",
    subtitleAssSet: "subtitle-ass:set",
    overlayDebugVisualizationSet: "overlay-debug-visualization:set",
    secondarySubtitleSet: "secondary-subtitle:set",
    secondarySubtitleMode: "secondary-subtitle:mode",
    subsyncOpenManual: "subsync:open-manual",
    kikuFieldGroupingRequest: "kiku:field-grouping-request",
    runtimeOptionsChanged: "runtime-options:changed",
    runtimeOptionsOpen: "runtime-options:open",
  },
} as const;

export type RendererToMainInvokeChannel =
  (typeof IPC_CHANNELS.rendererToMainInvoke)[keyof typeof IPC_CHANNELS.rendererToMainInvoke];
export type RendererToMainSendChannel =
  (typeof IPC_CHANNELS.rendererToMainSend)[keyof typeof IPC_CHANNELS.rendererToMainSend];
export type MainToRendererEventChannel =
  (typeof IPC_CHANNELS.mainToRendererEvent)[keyof typeof IPC_CHANNELS.mainToRendererEvent];
