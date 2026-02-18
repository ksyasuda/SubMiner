export { Texthooker } from './texthooker';
export { hasMpvWebsocketPlugin, SubtitleWebSocket } from './subtitle-ws';
export { registerGlobalShortcuts } from './shortcut';
export { createIpcDepsRuntime, registerIpcHandlers } from './ipc';
export { shortcutMatchesInputForLocalFallback } from './shortcut-fallback';
export {
  refreshOverlayShortcutsRuntime,
  registerOverlayShortcuts,
  syncOverlayShortcutsRuntime,
  unregisterOverlayShortcutsRuntime,
} from './overlay-shortcut';
export { createOverlayShortcutRuntimeHandlers } from './overlay-shortcut-handler';
export { createCliCommandDepsRuntime, handleCliCommand } from './cli-command';
export {
  copyCurrentSubtitle,
  handleMineSentenceDigit,
  handleMultiCopyDigit,
  markLastCardAsAudioCard,
  mineSentenceCard,
  triggerFieldGrouping,
  updateLastCardFromClipboard,
} from './mining';
export { createAppLifecycleDepsRuntime, startAppLifecycle } from './app-lifecycle';
export { cycleSecondarySubMode } from './subtitle-position';
export {
  getInitialInvisibleOverlayVisibility,
  isAutoUpdateEnabledRuntime,
  shouldAutoInitializeOverlayRuntimeFromConfig,
  shouldBindVisibleOverlayToMpvSubVisibility,
} from './startup';
export { openYomitanSettingsWindow } from './yomitan-settings';
export { createTokenizerDepsRuntime, tokenizeSubtitle } from './tokenizer';
export { createFrequencyDictionaryLookup } from './frequency-dictionary';
export { createJlptVocabularyLookup } from './jlpt-vocab';
export {
  getIgnoredPos1Entries,
  JlptIgnoredPos1Entry,
  JLPT_EXCLUDED_TERMS,
  JLPT_IGNORED_MECAB_POS1,
  JLPT_IGNORED_MECAB_POS1_ENTRIES,
  JLPT_IGNORED_MECAB_POS1_LIST,
  shouldIgnoreJlptByTerm,
  shouldIgnoreJlptForMecabPos1,
} from './jlpt-token-filter';
export { loadYomitanExtension } from './yomitan-extension-loader';
export {
  getJimakuLanguagePreference,
  getJimakuMaxEntryResults,
  jimakuFetchJson,
  resolveJimakuApiKey,
} from './jimaku';
export {
  loadSubtitlePosition,
  saveSubtitlePosition,
  updateCurrentMediaPath,
} from './subtitle-position';
export {
  createOverlayWindow,
  enforceOverlayLayerOrder,
  ensureOverlayWindowLevel,
  updateOverlayWindowBounds,
} from './overlay-window';
export { initializeOverlayRuntime } from './overlay-runtime-init';
export {
  setInvisibleOverlayVisible,
  setVisibleOverlayVisible,
  syncInvisibleOverlayMousePassthrough,
  updateInvisibleOverlayVisibility,
  updateVisibleOverlayVisibility,
} from './overlay-visibility';
export {
  MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY,
  MpvIpcClient,
  MpvRuntimeClientLike,
  MpvTrackProperty,
  playNextSubtitleRuntime,
  replayCurrentSubtitleRuntime,
  resolveCurrentAudioStreamIndex,
  sendMpvCommandRuntime,
  setMpvSubVisibilityRuntime,
  showMpvOsdRuntime,
} from './mpv';
export {
  applyMpvSubtitleRenderMetricsPatch,
  DEFAULT_MPV_SUBTITLE_RENDER_METRICS,
  sanitizeMpvSubtitleRenderMetrics,
} from './mpv-render-metrics';
export { createOverlayContentMeasurementStore } from './overlay-content-measurement';
export { handleMpvCommandFromIpc } from './ipc-command';
export { createFieldGroupingOverlayRuntime } from './field-grouping-overlay';
export { createNumericShortcutRuntime } from './numeric-shortcut';
export { runStartupBootstrapRuntime } from './startup';
export { runSubsyncManualFromIpcRuntime, triggerSubsyncFromConfigRuntime } from './subsync-runner';
export { registerAnkiJimakuIpcRuntime } from './anki-jimaku';
export { ImmersionTrackerService } from './immersion-tracker-service';
export {
  authenticateWithPassword as authenticateWithPasswordRuntime,
  listItems as listJellyfinItemsRuntime,
  listLibraries as listJellyfinLibrariesRuntime,
  listSubtitleTracks as listJellyfinSubtitleTracksRuntime,
  resolvePlaybackPlan as resolveJellyfinPlaybackPlanRuntime,
  ticksToSeconds as jellyfinTicksToSecondsRuntime,
} from './jellyfin';
export { buildJellyfinTimelinePayload, JellyfinRemoteSessionService } from './jellyfin-remote';
export {
  broadcastRuntimeOptionsChangedRuntime,
  createOverlayManager,
  setOverlayDebugVisualizationEnabledRuntime,
} from './overlay-manager';
