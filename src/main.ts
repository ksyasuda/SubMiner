/*
  SubMiner - All-in-one sentence mining overlay
  Copyright (C) 2024 sudacode

  This program is free software: you can redistribute it and/or modify
  it under the terms of the GNU General Public License as published by
  the Free Software Foundation, either version 3 of the License, or
  (at your option) any later version.

  This program is distributed in the hope that it will be useful,
  but WITHOUT ANY WARRANTY; without even the implied warranty of
  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
  GNU General Public License for more details.

  You should have received a copy of the GNU General Public License
  along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/
import {
  app,
  BrowserWindow,
  globalShortcut,
  clipboard,
  shell,
  protocol,
  Extension,
  Menu,
  Tray,
  nativeImage,
  dialog,
} from 'electron';

protocol.registerSchemesAsPrivileged([
  {
    scheme: 'chrome-extension',
    privileges: {
      standard: true,
      secure: true,
      supportFetchAPI: true,
      corsEnabled: true,
      bypassCSP: true,
    },
  },
]);

import * as path from 'path';
import * as os from 'os';
import * as fs from 'fs';
import { spawn } from 'node:child_process';
import { MecabTokenizer } from './mecab-tokenizer';
import type {
  JimakuApiResponse,
  SubtitleData,
  SubtitlePosition,
  WindowGeometry,
  SecondarySubMode,
  SubsyncManualRunRequest,
  SubsyncResult,
  KikuFieldGroupingChoice,
  RuntimeOptionState,
  MpvSubtitleRenderMetrics,
  ResolvedConfig,
} from './types';
import { SubtitleTimingTracker } from './subtitle-timing-tracker';
import { AnkiIntegration } from './anki-integration';
import { RuntimeOptionsManager } from './runtime-options';
import { downloadToFile, isRemoteMediaPath, parseMediaInfo } from './jimaku/utils';
import { createLogger, setLogLevel, type LogLevelSource } from './logger';
import { commandNeedsOverlayRuntime, parseArgs, shouldStartApp } from './cli/args';
import type { CliArgs, CliCommandSource } from './cli/args';
import { printHelp } from './cli/help';
import {
  createCriticalConfigErrorHandler,
  createReloadConfigHandler,
} from './main/runtime/startup-config';
import { buildConfigWarningNotificationBody } from './main/config-validation';
import { createImmersionTrackerStartupHandler } from './main/runtime/immersion-startup';
import { createBuildImmersionTrackerStartupMainDepsHandler } from './main/runtime/immersion-startup-main-deps';
import { createImmersionMediaRuntime } from './main/runtime/immersion-media';
import { createAnilistStateRuntime } from './main/runtime/anilist-state';
import { createConfigDerivedRuntime } from './main/runtime/config-derived';
import { appendClipboardVideoToQueueRuntime } from './main/runtime/clipboard-queue';
import { createMainSubsyncRuntime } from './main/runtime/subsync-runtime';
import {
  buildAnilistSetupUrl,
  consumeAnilistSetupCallbackUrl,
  findAnilistSetupDeepLinkArgvUrl,
  isAnilistTrackingEnabled,
  loadAnilistManualTokenEntry,
  loadAnilistSetupFallback,
  openAnilistSetupInBrowser,
} from './main/runtime/anilist-setup';
import {
  createConsumeAnilistSetupTokenFromUrlHandler,
  createHandleAnilistSetupProtocolUrlHandler,
  createNotifyAnilistSetupHandler,
  createRegisterSubminerProtocolClientHandler,
} from './main/runtime/anilist-setup-protocol';
import {
  createBuildConsumeAnilistSetupTokenFromUrlMainDepsHandler,
  createBuildHandleAnilistSetupProtocolUrlMainDepsHandler,
  createBuildNotifyAnilistSetupMainDepsHandler,
  createBuildRegisterSubminerProtocolClientMainDepsHandler,
} from './main/runtime/anilist-setup-protocol-main-deps';
import { createRefreshAnilistClientSecretStateHandler } from './main/runtime/anilist-token-refresh';
import { createBuildRefreshAnilistClientSecretStateMainDepsHandler } from './main/runtime/anilist-token-refresh-main-deps';
import {
  createHandleJellyfinRemoteGeneralCommand,
  createHandleJellyfinRemotePlay,
  createHandleJellyfinRemotePlaystate,
  getConfiguredJellyfinSession,
  type ActiveJellyfinRemotePlaybackState,
} from './main/runtime/jellyfin-remote-commands';
import {
  createReportJellyfinRemoteProgressHandler,
  createReportJellyfinRemoteStoppedHandler,
} from './main/runtime/jellyfin-remote-playback';
import {
  createBuildHandleJellyfinRemoteGeneralCommandMainDepsHandler,
  createBuildHandleJellyfinRemotePlayMainDepsHandler,
  createBuildHandleJellyfinRemotePlaystateMainDepsHandler,
  createBuildReportJellyfinRemoteProgressMainDepsHandler,
  createBuildReportJellyfinRemoteStoppedMainDepsHandler,
} from './main/runtime/jellyfin-remote-main-deps';
import { createBuildSubtitleProcessingControllerMainDepsHandler } from './main/runtime/subtitle-processing-main-deps';
import {
  createBuildAnilistStateRuntimeMainDepsHandler,
  createBuildConfigDerivedRuntimeMainDepsHandler,
  createBuildImmersionMediaRuntimeMainDepsHandler,
  createBuildMainSubsyncRuntimeMainDepsHandler,
} from './main/runtime/runtime-bootstrap-main-deps';
import {
  createBuildOverlayContentMeasurementStoreMainDepsHandler,
  createBuildOverlayModalRuntimeMainDepsHandler,
} from './main/runtime/overlay-bootstrap-main-deps';
import {
  createEnsureMpvConnectedForJellyfinPlaybackHandler,
  createLaunchMpvIdleForJellyfinPlaybackHandler,
  createWaitForMpvConnectedHandler,
} from './main/runtime/jellyfin-remote-connection';
import {
  createBuildEnsureMpvConnectedForJellyfinPlaybackMainDepsHandler,
  createBuildLaunchMpvIdleForJellyfinPlaybackMainDepsHandler,
  createBuildWaitForMpvConnectedMainDepsHandler,
} from './main/runtime/jellyfin-remote-connection-main-deps';
import {
  buildJellyfinSetupFormHtml,
  createOpenJellyfinSetupWindowHandler,
  createMaybeFocusExistingJellyfinSetupWindowHandler,
  parseJellyfinSetupSubmissionUrl,
} from './main/runtime/jellyfin-setup-window';
import { createBuildOpenJellyfinSetupWindowMainDepsHandler } from './main/runtime/jellyfin-setup-window-main-deps';
import {
  createMaybeFocusExistingAnilistSetupWindowHandler,
  createOpenAnilistSetupWindowHandler,
} from './main/runtime/anilist-setup-window';
import { createBuildOpenAnilistSetupWindowMainDepsHandler } from './main/runtime/anilist-setup-window-main-deps';
import {
  createEnsureAnilistMediaGuessHandler,
  createMaybeProbeAnilistDurationHandler,
} from './main/runtime/anilist-media-guess';
import {
  createBuildEnsureAnilistMediaGuessMainDepsHandler,
  createBuildMaybeProbeAnilistDurationMainDepsHandler,
} from './main/runtime/anilist-media-guess-main-deps';
import {
  createGetAnilistMediaGuessRuntimeStateHandler,
  createGetCurrentAnilistMediaKeyHandler,
  createResetAnilistMediaGuessStateHandler,
  createResetAnilistMediaTrackingHandler,
  createSetAnilistMediaGuessRuntimeStateHandler,
} from './main/runtime/anilist-media-state';
import {
  createBuildGetAnilistMediaGuessRuntimeStateMainDepsHandler,
  createBuildGetCurrentAnilistMediaKeyMainDepsHandler,
  createBuildResetAnilistMediaGuessStateMainDepsHandler,
  createBuildResetAnilistMediaTrackingMainDepsHandler,
  createBuildSetAnilistMediaGuessRuntimeStateMainDepsHandler,
} from './main/runtime/anilist-media-state-main-deps';
import {
  buildAnilistAttemptKey,
  createMaybeRunAnilistPostWatchUpdateHandler,
  createProcessNextAnilistRetryUpdateHandler,
  rememberAnilistAttemptedUpdateKey,
} from './main/runtime/anilist-post-watch';
import {
  createBuildMaybeRunAnilistPostWatchUpdateMainDepsHandler,
  createBuildProcessNextAnilistRetryUpdateMainDepsHandler,
} from './main/runtime/anilist-post-watch-main-deps';
import {
  createLoadSubtitlePositionHandler,
  createSaveSubtitlePositionHandler,
} from './main/runtime/subtitle-position';
import {
  createBuildLoadSubtitlePositionMainDepsHandler,
  createBuildSaveSubtitlePositionMainDepsHandler,
} from './main/runtime/subtitle-position-main-deps';
import { registerProtocolUrlHandlers } from './main/runtime/protocol-url-handlers';
import { createBuildRegisterProtocolUrlHandlersMainDepsHandler } from './main/runtime/protocol-url-handlers-main-deps';
import { createHandleJellyfinAuthCommands } from './main/runtime/jellyfin-cli-auth';
import { createRunJellyfinCommandHandler } from './main/runtime/jellyfin-command-dispatch';
import { createBuildRunJellyfinCommandMainDepsHandler } from './main/runtime/jellyfin-command-dispatch-main-deps';
import { createHandleJellyfinListCommands } from './main/runtime/jellyfin-cli-list';
import { createHandleJellyfinPlayCommand } from './main/runtime/jellyfin-cli-play';
import { createHandleJellyfinRemoteAnnounceCommand } from './main/runtime/jellyfin-cli-remote-announce';
import {
  createBuildHandleJellyfinAuthCommandsMainDepsHandler,
  createBuildHandleJellyfinListCommandsMainDepsHandler,
  createBuildHandleJellyfinPlayCommandMainDepsHandler,
  createBuildHandleJellyfinRemoteAnnounceCommandMainDepsHandler,
} from './main/runtime/jellyfin-cli-main-deps';
import {
  createGetJellyfinClientInfoHandler,
  createGetResolvedJellyfinConfigHandler,
} from './main/runtime/jellyfin-client-info';
import {
  createBuildGetJellyfinClientInfoMainDepsHandler,
  createBuildGetResolvedJellyfinConfigMainDepsHandler,
} from './main/runtime/jellyfin-client-info-main-deps';
import {
  createApplyJellyfinMpvDefaultsHandler,
  createGetDefaultSocketPathHandler,
} from './main/runtime/mpv-jellyfin-defaults';
import {
  createBuildApplyJellyfinMpvDefaultsMainDepsHandler,
  createBuildGetDefaultSocketPathMainDepsHandler,
} from './main/runtime/mpv-jellyfin-defaults-main-deps';
import { createBuildMediaRuntimeMainDepsHandler } from './main/runtime/media-runtime-main-deps';
import {
  createBuildDictionaryRootsMainHandler,
  createBuildFrequencyDictionaryRootsMainHandler,
  createBuildFrequencyDictionaryRuntimeMainDepsHandler,
  createBuildJlptDictionaryRuntimeMainDepsHandler,
} from './main/runtime/dictionary-runtime-main-deps';
import { createPlayJellyfinItemInMpvHandler } from './main/runtime/jellyfin-playback-launch';
import { createBuildPlayJellyfinItemInMpvMainDepsHandler } from './main/runtime/jellyfin-playback-launch-main-deps';
import { createPreloadJellyfinExternalSubtitlesHandler } from './main/runtime/jellyfin-subtitle-preload';
import { createBuildPreloadJellyfinExternalSubtitlesMainDepsHandler } from './main/runtime/jellyfin-subtitle-preload-main-deps';
import {
  createStartJellyfinRemoteSessionHandler,
  createStopJellyfinRemoteSessionHandler,
} from './main/runtime/jellyfin-remote-session-lifecycle';
import {
  createBuildStartJellyfinRemoteSessionMainDepsHandler,
  createBuildStopJellyfinRemoteSessionMainDepsHandler,
} from './main/runtime/jellyfin-remote-session-main-deps';
import { createInitialArgsRuntimeHandler } from './main/runtime/initial-args-runtime-handler';
import { createHandleTexthookerOnlyModeTransitionHandler } from './main/runtime/cli-command-prechecks';
import { createBuildHandleTexthookerOnlyModeTransitionMainDepsHandler } from './main/runtime/cli-command-prechecks-main-deps';
import {
  createGetFieldGroupingResolverHandler,
  createSetFieldGroupingResolverHandler,
} from './main/runtime/field-grouping-resolver';
import {
  createBuildGetFieldGroupingResolverMainDepsHandler,
  createBuildSetFieldGroupingResolverMainDepsHandler,
} from './main/runtime/field-grouping-resolver-main-deps';
import { createBuildFieldGroupingOverlayMainDepsHandler } from './main/runtime/field-grouping-overlay-main-deps';
import { createCliCommandContextFactory } from './main/runtime/cli-command-context-factory';
import { createBindMpvMainEventHandlersHandler } from './main/runtime/mpv-main-event-bindings';
import { createBuildBindMpvMainEventHandlersMainDepsHandler } from './main/runtime/mpv-main-event-main-deps';
import { createBuildMpvClientRuntimeServiceFactoryDepsHandler } from './main/runtime/mpv-client-runtime-service-main-deps';
import { createMpvClientRuntimeServiceFactory } from './main/runtime/mpv-client-runtime-service';
import { createUpdateMpvSubtitleRenderMetricsHandler } from './main/runtime/mpv-subtitle-render-metrics';
import { createBuildUpdateMpvSubtitleRenderMetricsMainDepsHandler } from './main/runtime/mpv-subtitle-render-metrics-main-deps';
import {
  createBuildTokenizerDepsMainHandler,
  createCreateMecabTokenizerAndCheckMainHandler,
  createPrewarmSubtitleDictionariesMainHandler,
} from './main/runtime/subtitle-tokenization-main-deps';
import {
  createLaunchBackgroundWarmupTaskHandler,
  createStartBackgroundWarmupsHandler,
} from './main/runtime/startup-warmups';
import {
  createBuildLaunchBackgroundWarmupTaskMainDepsHandler,
  createBuildStartBackgroundWarmupsMainDepsHandler,
} from './main/runtime/startup-warmups-main-deps';
import {
  createEnforceOverlayLayerOrderHandler,
  createEnsureOverlayWindowLevelHandler,
  createUpdateInvisibleOverlayBoundsHandler,
  createUpdateVisibleOverlayBoundsHandler,
} from './main/runtime/overlay-window-layout';
import {
  createBuildEnforceOverlayLayerOrderMainDepsHandler,
  createBuildEnsureOverlayWindowLevelMainDepsHandler,
  createBuildUpdateInvisibleOverlayBoundsMainDepsHandler,
  createBuildUpdateVisibleOverlayBoundsMainDepsHandler,
} from './main/runtime/overlay-window-layout-main-deps';
import { buildTrayMenuTemplateRuntime, resolveTrayIconPathRuntime } from './main/runtime/tray-runtime';
import {
  createGetConfiguredShortcutsHandler,
  createRefreshGlobalAndOverlayShortcutsHandler,
  createRegisterGlobalShortcutsHandler,
} from './main/runtime/global-shortcuts';
import {
  createBuildGetConfiguredShortcutsMainDepsHandler,
  createBuildRefreshGlobalAndOverlayShortcutsMainDepsHandler,
  createBuildRegisterGlobalShortcutsMainDepsHandler,
} from './main/runtime/global-shortcuts-main-deps';
import { createAppendToMpvLogHandler, createShowMpvOsdHandler } from './main/runtime/mpv-osd-log';
import {
  createBuildAppendToMpvLogMainDepsHandler,
  createBuildShowMpvOsdMainDepsHandler,
} from './main/runtime/mpv-osd-log-main-deps';
import { createBuildCycleSecondarySubModeMainDepsHandler } from './main/runtime/secondary-sub-mode-main-deps';
import {
  createCancelNumericShortcutSessionHandler,
  createStartNumericShortcutSessionHandler,
} from './main/runtime/numeric-shortcut-session-handlers';
import {
  createBuildCancelNumericShortcutSessionMainDepsHandler,
  createBuildStartNumericShortcutSessionMainDepsHandler,
} from './main/runtime/numeric-shortcut-session-main-deps';
import { createBuildNumericShortcutRuntimeMainDepsHandler } from './main/runtime/numeric-shortcut-runtime-main-deps';
import {
  createRefreshOverlayShortcutsHandler,
  createRegisterOverlayShortcutsHandler,
  createSyncOverlayShortcutsHandler,
  createUnregisterOverlayShortcutsHandler,
} from './main/runtime/overlay-shortcuts-lifecycle';
import {
  createBuildRefreshOverlayShortcutsMainDepsHandler,
  createBuildRegisterOverlayShortcutsMainDepsHandler,
  createBuildSyncOverlayShortcutsMainDepsHandler,
  createBuildUnregisterOverlayShortcutsMainDepsHandler,
} from './main/runtime/overlay-shortcuts-lifecycle-main-deps';
import { createBuildOverlayShortcutsRuntimeMainDepsHandler } from './main/runtime/overlay-shortcuts-runtime-main-deps';
import {
  createMarkLastCardAsAudioCardHandler,
  createMineSentenceCardHandler,
  createRefreshKnownWordCacheHandler,
  createTriggerFieldGroupingHandler,
  createUpdateLastCardFromClipboardHandler,
} from './main/runtime/anki-actions';
import {
  createBuildMarkLastCardAsAudioCardMainDepsHandler,
  createBuildMineSentenceCardMainDepsHandler,
  createBuildRefreshKnownWordCacheMainDepsHandler,
  createBuildTriggerFieldGroupingMainDepsHandler,
  createBuildUpdateLastCardFromClipboardMainDepsHandler,
} from './main/runtime/anki-actions-main-deps';
import {
  createCopyCurrentSubtitleHandler,
  createHandleMineSentenceDigitHandler,
  createHandleMultiCopyDigitHandler,
} from './main/runtime/mining-actions';
import {
  createBuildCopyCurrentSubtitleMainDepsHandler,
  createBuildHandleMineSentenceDigitMainDepsHandler,
  createBuildHandleMultiCopyDigitMainDepsHandler,
} from './main/runtime/mining-actions-main-deps';
import { createBuildOverlayVisibilityRuntimeMainDepsHandler } from './main/runtime/overlay-visibility-runtime-main-deps';
import { createOverlayVisibilityRuntime } from './main/runtime/overlay-visibility-runtime';
import {
  createAppendClipboardVideoToQueueHandler,
  createHandleOverlayModalClosedHandler,
} from './main/runtime/overlay-main-actions';
import {
  createBuildAppendClipboardVideoToQueueMainDepsHandler,
  createBuildHandleOverlayModalClosedMainDepsHandler,
} from './main/runtime/overlay-main-actions-main-deps';
import {
  createBroadcastRuntimeOptionsChangedHandler,
  createGetRuntimeOptionsStateHandler,
  createOpenRuntimeOptionsPaletteHandler,
  createRestorePreviousSecondarySubVisibilityHandler,
  createSendToActiveOverlayWindowHandler,
  createSetOverlayDebugVisualizationEnabledHandler,
} from './main/runtime/overlay-runtime-main-actions';
import {
  createBuildBroadcastRuntimeOptionsChangedMainDepsHandler,
  createBuildGetRuntimeOptionsStateMainDepsHandler,
  createBuildOpenRuntimeOptionsPaletteMainDepsHandler,
  createBuildRestorePreviousSecondarySubVisibilityMainDepsHandler,
  createBuildSendToActiveOverlayWindowMainDepsHandler,
  createBuildSetOverlayDebugVisualizationEnabledMainDepsHandler,
} from './main/runtime/overlay-runtime-main-actions-main-deps';
import { createIpcRuntimeHandlers } from './main/runtime/ipc-runtime-handlers';
import { createBuildMpvCommandFromIpcRuntimeMainDepsHandler } from './main/runtime/ipc-mpv-command-main-deps';
import { createOverlayWindowRuntimeHandlers } from './main/runtime/overlay-window-runtime-handlers';
import { createOverlayRuntimeBootstrapHandlers } from './main/runtime/overlay-runtime-bootstrap-handlers';
import { createTrayRuntimeHandlers } from './main/runtime/tray-runtime-handlers';
import { createYomitanExtensionRuntime } from './main/runtime/yomitan-extension-runtime';
import { createYomitanSettingsRuntime } from './main/runtime/yomitan-settings-runtime';
import {
  createOnWillQuitCleanupHandler,
  createRestoreWindowsOnActivateHandler,
  createShouldRestoreWindowsOnActivateHandler,
} from './main/runtime/app-lifecycle-actions';
import { createBuildOnWillQuitCleanupDepsHandler } from './main/runtime/app-lifecycle-main-cleanup';
import {
  createBuildRestoreWindowsOnActivateMainDepsHandler,
  createBuildShouldRestoreWindowsOnActivateMainDepsHandler,
} from './main/runtime/app-lifecycle-main-activate';
import {
  buildRestartRequiredConfigMessage,
  createConfigHotReloadAppliedHandler,
  createConfigHotReloadMessageHandler,
  resolveSubtitleStyleForRenderer,
} from './main/runtime/config-hot-reload-handlers';
import {
  createBuildConfigHotReloadMessageMainDepsHandler,
  createBuildConfigHotReloadAppliedMainDepsHandler,
  createBuildConfigHotReloadRuntimeMainDepsHandler,
  createBuildWatchConfigPathMainDepsHandler,
  createWatchConfigPathHandler,
} from './main/runtime/config-hot-reload-main-deps';
import {
  createBuildCriticalConfigErrorMainDepsHandler,
  createBuildReloadConfigMainDepsHandler,
} from './main/runtime/startup-config-main-deps';
import { createBuildAppReadyRuntimeMainDepsHandler } from './main/runtime/app-ready-main-deps';
import { createStartupRuntimeHandlers } from './main/runtime/startup-runtime-handlers';
import {
  enforceUnsupportedWaylandMode,
  forceX11Backend,
  generateDefaultConfigFile,
  resolveConfiguredShortcuts,
  resolveKeybindings,
  showDesktopNotification,
} from './core/utils';
import {
  MpvIpcClient,
  SubtitleWebSocket,
  Texthooker,
  applyMpvSubtitleRenderMetricsPatch,
  broadcastRuntimeOptionsChangedRuntime,
  copyCurrentSubtitle as copyCurrentSubtitleCore,
  createOverlayManager,
  createFieldGroupingOverlayRuntime,
  createNumericShortcutRuntime,
  createOverlayContentMeasurementStore,
  createSubtitleProcessingController,
  createOverlayWindow as createOverlayWindowCore,
  createTokenizerDepsRuntime,
  cycleSecondarySubMode as cycleSecondarySubModeCore,
  enforceOverlayLayerOrder as enforceOverlayLayerOrderCore,
  ensureOverlayWindowLevel as ensureOverlayWindowLevelCore,
  handleMineSentenceDigit as handleMineSentenceDigitCore,
  handleMultiCopyDigit as handleMultiCopyDigitCore,
  hasMpvWebsocketPlugin,
  initializeOverlayRuntime as initializeOverlayRuntimeCore,
  loadSubtitlePosition as loadSubtitlePositionCore,
  loadYomitanExtension as loadYomitanExtensionCore,
  listJellyfinItemsRuntime,
  listJellyfinLibrariesRuntime,
  listJellyfinSubtitleTracksRuntime,
  markLastCardAsAudioCard as markLastCardAsAudioCardCore,
  DEFAULT_MPV_SUBTITLE_RENDER_METRICS,
  ImmersionTrackerService,
  JellyfinRemoteSessionService,
  mineSentenceCard as mineSentenceCardCore,
  openYomitanSettingsWindow,
  playNextSubtitleRuntime,
  registerGlobalShortcuts as registerGlobalShortcutsCore,
  replayCurrentSubtitleRuntime,
  runStartupBootstrapRuntime,
  saveSubtitlePosition as saveSubtitlePositionCore,
  authenticateWithPasswordRuntime,
  createConfigHotReloadRuntime,
  resolveJellyfinPlaybackPlanRuntime,
  jellyfinTicksToSecondsRuntime,
  sendMpvCommandRuntime,
  setInvisibleOverlayVisible as setInvisibleOverlayVisibleCore,
  setMpvSubVisibilityRuntime,
  setOverlayDebugVisualizationEnabledRuntime,
  setVisibleOverlayVisible as setVisibleOverlayVisibleCore,
  showMpvOsdRuntime,
  tokenizeSubtitle as tokenizeSubtitleCore,
  triggerFieldGrouping as triggerFieldGroupingCore,
  updateLastCardFromClipboard as updateLastCardFromClipboardCore,
} from './core/services';
import {
  guessAnilistMediaInfo,
  type AnilistMediaGuess,
  updateAnilistPostWatchProgress,
} from './core/services/anilist/anilist-updater';
import { createAnilistTokenStore } from './core/services/anilist/anilist-token-store';
import { createAnilistUpdateQueue } from './core/services/anilist/anilist-update-queue';
import { createJellyfinTokenStore } from './core/services/jellyfin-token-store';
import { applyRuntimeOptionResultRuntime } from './core/services/runtime-options-ipc';
import { createAppReadyRuntimeRunner } from './main/app-lifecycle';
import { createStartupBootstrapRuntimeDeps } from './main/startup';
import { createAppLifecycleRuntimeRunner } from './main/startup-lifecycle';
import { handleMpvCommandFromIpcRuntime } from './main/ipc-mpv-command';
import { registerIpcRuntimeServices } from './main/ipc-runtime';
import { createAnkiJimakuIpcRuntimeServiceDeps } from './main/dependencies';
import { handleCliCommandRuntimeServiceWithContext } from './main/cli-runtime';
import { createOverlayModalRuntimeService, type OverlayHostedModal } from './main/overlay-runtime';
import { createOverlayShortcutsRuntimeService } from './main/overlay-shortcuts-runtime';
import {
  createJlptDictionaryRuntimeService,
  getJlptDictionarySearchPaths,
} from './main/jlpt-runtime';
import {
  createFrequencyDictionaryRuntimeService,
  getFrequencyDictionarySearchPaths,
} from './main/frequency-dictionary-runtime';
import { createMediaRuntimeService } from './main/media-runtime';
import { createOverlayVisibilityRuntimeService } from './main/overlay-visibility-runtime';
import { type AppState, type StartupState, applyStartupState, createAppState } from './main/state';
import {
  isAllowedAnilistExternalUrl,
  isAllowedAnilistSetupNavigationUrl,
} from './main/anilist-url-guard';
import {
  ConfigService,
  DEFAULT_CONFIG,
  DEFAULT_KEYBINDINGS,
  generateConfigTemplate,
} from './config';
import { resolveConfigDir } from './config/path-resolution';

if (process.platform === 'linux') {
  app.commandLine.appendSwitch('enable-features', 'GlobalShortcutsPortal');
}

app.setName('SubMiner');

const DEFAULT_TEXTHOOKER_PORT = 5174;
const DEFAULT_MPV_LOG_FILE = path.join(os.homedir(), '.cache', 'SubMiner', 'mp.log');
const ANILIST_SETUP_CLIENT_ID_URL = 'https://anilist.co/api/v2/oauth/authorize';
const ANILIST_SETUP_RESPONSE_TYPE = 'token';
const ANILIST_DEFAULT_CLIENT_ID = '36084';
const ANILIST_REDIRECT_URI = 'https://anilist.subminer.moe/';
const ANILIST_DEVELOPER_SETTINGS_URL = 'https://anilist.co/settings/developer';
const ANILIST_UPDATE_MIN_WATCH_RATIO = 0.85;
const ANILIST_UPDATE_MIN_WATCH_SECONDS = 10 * 60;
const ANILIST_DURATION_RETRY_INTERVAL_MS = 15_000;
const ANILIST_MAX_ATTEMPTED_UPDATE_KEYS = 1000;
const ANILIST_TOKEN_STORE_FILE = 'anilist-token-store.json';
const JELLYFIN_TOKEN_STORE_FILE = 'jellyfin-token-store.json';
const ANILIST_RETRY_QUEUE_FILE = 'anilist-retry-queue.json';
const TRAY_TOOLTIP = 'SubMiner';

let anilistCurrentMediaKey: string | null = null;
let anilistCurrentMediaDurationSec: number | null = null;
let anilistCurrentMediaGuess: AnilistMediaGuess | null = null;
let anilistCurrentMediaGuessPromise: Promise<AnilistMediaGuess | null> | null = null;
let anilistLastDurationProbeAtMs = 0;
let anilistUpdateInFlight = false;
const anilistAttemptedUpdateKeys = new Set<string>();
let anilistCachedAccessToken: string | null = null;
let jellyfinPlayQuitOnDisconnectArmed = false;
const JELLYFIN_LANG_PREF = 'ja,jp,jpn,japanese,en,eng,english,enUS,en-US';
const JELLYFIN_TICKS_PER_SECOND = 10_000_000;
const JELLYFIN_REMOTE_PROGRESS_INTERVAL_MS = 3000;
const JELLYFIN_MPV_CONNECT_TIMEOUT_MS = 3000;
const JELLYFIN_MPV_AUTO_LAUNCH_TIMEOUT_MS = 20000;
const MPV_JELLYFIN_DEFAULT_ARGS = [
  '--sub-auto=fuzzy',
  '--sub-file-paths=.;subs;subtitles',
  '--sid=auto',
  '--secondary-sid=auto',
  '--secondary-sub-visibility=no',
  '--alang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
  '--slang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
] as const;

let activeJellyfinRemotePlayback: ActiveJellyfinRemotePlaybackState | null = null;
let jellyfinRemoteLastProgressAtMs = 0;
let jellyfinMpvAutoLaunchInFlight: Promise<boolean> | null = null;
let backgroundWarmupsStarted = false;
let yomitanLoadInFlight: Promise<Extension | null> | null = null;

const buildApplyJellyfinMpvDefaultsMainDepsHandler =
  createBuildApplyJellyfinMpvDefaultsMainDepsHandler({
  sendMpvCommandRuntime: (client, command) => sendMpvCommandRuntime(client as never, command),
  jellyfinLangPref: JELLYFIN_LANG_PREF,
});
const applyJellyfinMpvDefaultsMainDeps = buildApplyJellyfinMpvDefaultsMainDepsHandler();
const applyJellyfinMpvDefaultsHandler = createApplyJellyfinMpvDefaultsHandler(
  applyJellyfinMpvDefaultsMainDeps,
);

function applyJellyfinMpvDefaults(client: MpvIpcClient): void {
  applyJellyfinMpvDefaultsHandler(client);
}

const CONFIG_DIR = resolveConfigDir({
  xdgConfigHome: process.env.XDG_CONFIG_HOME,
  homeDir: os.homedir(),
  existsSync: fs.existsSync,
});
const USER_DATA_PATH = CONFIG_DIR;
const DEFAULT_MPV_LOG_PATH = process.env.SUBMINER_MPV_LOG?.trim() || DEFAULT_MPV_LOG_FILE;
const DEFAULT_IMMERSION_DB_PATH = path.join(USER_DATA_PATH, 'immersion.sqlite');
const configService = new ConfigService(CONFIG_DIR);
const anilistTokenStore = createAnilistTokenStore(
  path.join(USER_DATA_PATH, ANILIST_TOKEN_STORE_FILE),
  {
    info: (message: string) => console.info(message),
    warn: (message: string, details?: unknown) => console.warn(message, details),
    error: (message: string, details?: unknown) => console.error(message, details),
  },
);
const jellyfinTokenStore = createJellyfinTokenStore(
  path.join(USER_DATA_PATH, JELLYFIN_TOKEN_STORE_FILE),
  {
    info: (message: string) => console.info(message),
    warn: (message: string, details?: unknown) => console.warn(message, details),
    error: (message: string, details?: unknown) => console.error(message, details),
  },
);
const anilistUpdateQueue = createAnilistUpdateQueue(
  path.join(USER_DATA_PATH, ANILIST_RETRY_QUEUE_FILE),
  {
    info: (message: string) => console.info(message),
    warn: (message: string, details?: unknown) => console.warn(message, details),
    error: (message: string, details?: unknown) => console.error(message, details),
  },
);
const isDev = process.argv.includes('--dev') || process.argv.includes('--debug');
const texthookerService = new Texthooker();
const subtitleWsService = new SubtitleWebSocket();
const logger = createLogger('main');
const appLogger = {
  logInfo: (message: string) => {
    logger.info(message);
  },
  logWarning: (message: string) => {
    logger.warn(message);
  },
  logError: (message: string, details: unknown) => {
    logger.error(message, details);
  },
  logNoRunningInstance: () => {
    logger.error('No running instance. Use --start to launch the app.');
  },
  logConfigWarning: (warning: {
    path: string;
    message: string;
    value: unknown;
    fallback: unknown;
  }) => {
    logger.warn(
      `[config] ${warning.path}: ${warning.message} value=${JSON.stringify(warning.value)} fallback=${JSON.stringify(warning.fallback)}`,
    );
  },
};

const buildGetDefaultSocketPathMainDepsHandler = createBuildGetDefaultSocketPathMainDepsHandler({
  platform: process.platform,
});
const getDefaultSocketPathMainDeps = buildGetDefaultSocketPathMainDepsHandler();
const getDefaultSocketPathHandler = createGetDefaultSocketPathHandler(
  getDefaultSocketPathMainDeps,
);

function getDefaultSocketPath(): string {
  return getDefaultSocketPathHandler();
}

if (!fs.existsSync(USER_DATA_PATH)) {
  fs.mkdirSync(USER_DATA_PATH, { recursive: true });
}
app.setPath('userData', USER_DATA_PATH);

process.on('SIGINT', () => {
  app.quit();
});
process.on('SIGTERM', () => {
  app.quit();
});

const overlayManager = createOverlayManager();
const buildOverlayContentMeasurementStoreMainDepsHandler =
  createBuildOverlayContentMeasurementStoreMainDepsHandler({
  now: () => Date.now(),
  warn: (message: string) => logger.warn(message),
});
const buildOverlayModalRuntimeMainDepsHandler = createBuildOverlayModalRuntimeMainDepsHandler({
  getMainWindow: () => overlayManager.getMainWindow(),
  getInvisibleWindow: () => overlayManager.getInvisibleWindow(),
});
const overlayContentMeasurementStoreMainDeps =
  buildOverlayContentMeasurementStoreMainDepsHandler();
const overlayContentMeasurementStore = createOverlayContentMeasurementStore(
  overlayContentMeasurementStoreMainDeps,
);
const overlayModalRuntime = createOverlayModalRuntimeService(buildOverlayModalRuntimeMainDepsHandler());
const appState = createAppState({
  mpvSocketPath: getDefaultSocketPath(),
  texthookerPort: DEFAULT_TEXTHOOKER_PORT,
});
const buildImmersionMediaRuntimeMainDepsHandler = createBuildImmersionMediaRuntimeMainDepsHandler({
  getResolvedConfig: () => getResolvedConfig(),
  defaultImmersionDbPath: DEFAULT_IMMERSION_DB_PATH,
  getTracker: () => appState.immersionTracker,
  getMpvClient: () => appState.mpvClient,
  getCurrentMediaPath: () => appState.currentMediaPath,
  getCurrentMediaTitle: () => appState.currentMediaTitle,
  logDebug: (message) => logger.debug(message),
  logInfo: (message) => logger.info(message),
});
const buildAnilistStateRuntimeMainDepsHandler = createBuildAnilistStateRuntimeMainDepsHandler({
  getClientSecretState: () => appState.anilistClientSecretState,
  setClientSecretState: (next) => {
    appState.anilistClientSecretState = next;
  },
  getRetryQueueState: () => appState.anilistRetryQueueState,
  setRetryQueueState: (next) => {
    appState.anilistRetryQueueState = next;
  },
  getUpdateQueueSnapshot: () => anilistUpdateQueue.getSnapshot(),
  clearStoredToken: () => anilistTokenStore.clearToken(),
  clearCachedAccessToken: () => {
    anilistCachedAccessToken = null;
  },
});
const buildConfigDerivedRuntimeMainDepsHandler = createBuildConfigDerivedRuntimeMainDepsHandler({
  getResolvedConfig: () => getResolvedConfig(),
  getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
  platform: process.platform,
  defaultJimakuLanguagePreference: DEFAULT_CONFIG.jimaku.languagePreference,
  defaultJimakuMaxEntryResults: DEFAULT_CONFIG.jimaku.maxEntryResults,
  defaultJimakuApiBaseUrl: DEFAULT_CONFIG.jimaku.apiBaseUrl,
});
const buildMainSubsyncRuntimeMainDepsHandler = createBuildMainSubsyncRuntimeMainDepsHandler({
  getMpvClient: () => appState.mpvClient,
  getResolvedConfig: () => getResolvedConfig(),
  getSubsyncInProgress: () => appState.subsyncInProgress,
  setSubsyncInProgress: (inProgress) => {
    appState.subsyncInProgress = inProgress;
  },
  showMpvOsd: (text) => showMpvOsd(text),
  openManualPicker: (payload) => {
    sendToActiveOverlayWindow('subsync:open-manual', payload, {
      restoreOnModalClose: 'subsync',
    });
  },
});
const immersionMediaRuntime = createImmersionMediaRuntime(buildImmersionMediaRuntimeMainDepsHandler());
const anilistStateRuntime = createAnilistStateRuntime(buildAnilistStateRuntimeMainDepsHandler());
const configDerivedRuntime = createConfigDerivedRuntime(buildConfigDerivedRuntimeMainDepsHandler());
const subsyncRuntime = createMainSubsyncRuntime(buildMainSubsyncRuntimeMainDepsHandler());
let appTray: Tray | null = null;
const buildSubtitleProcessingControllerMainDepsHandler =
  createBuildSubtitleProcessingControllerMainDepsHandler({
  tokenizeSubtitle: async (text: string) => {
    if (getOverlayWindows().length === 0 && !subtitleWsService.hasClients()) {
      return null;
    }
    return await tokenizeSubtitle(text);
  },
  emitSubtitle: (payload) => {
    broadcastToOverlayWindows('subtitle:set', payload);
    subtitleWsService.broadcast(payload, {
      enabled: getResolvedConfig().subtitleStyle.frequencyDictionary.enabled,
      topX: getResolvedConfig().subtitleStyle.frequencyDictionary.topX,
      mode: getResolvedConfig().subtitleStyle.frequencyDictionary.mode,
    });
  },
  logDebug: (message) => {
    logger.debug(`[subtitle-processing] ${message}`);
  },
  now: () => Date.now(),
});
const subtitleProcessingControllerMainDeps =
  buildSubtitleProcessingControllerMainDepsHandler();
const subtitleProcessingController = createSubtitleProcessingController(
  subtitleProcessingControllerMainDeps,
);
const overlayShortcutsRuntime = createOverlayShortcutsRuntimeService(
  createBuildOverlayShortcutsRuntimeMainDepsHandler({
    getConfiguredShortcuts: () => getConfiguredShortcuts(),
    getShortcutsRegistered: () => appState.shortcutsRegistered,
    setShortcutsRegistered: (registered) => {
      appState.shortcutsRegistered = registered;
    },
    isOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
    showMpvOsd: (text: string) => showMpvOsd(text),
    openRuntimeOptionsPalette: () => {
      openRuntimeOptionsPalette();
    },
    openJimaku: () => {
      sendToActiveOverlayWindow('jimaku:open', undefined, {
        restoreOnModalClose: 'jimaku',
      });
    },
    markAudioCard: () => markLastCardAsAudioCard(),
    copySubtitleMultiple: (timeoutMs) => {
      startPendingMultiCopy(timeoutMs);
    },
    copySubtitle: () => {
      copyCurrentSubtitle();
    },
    toggleSecondarySubMode: () => cycleSecondarySubMode(),
    updateLastCardFromClipboard: () => updateLastCardFromClipboard(),
    triggerFieldGrouping: () => triggerFieldGrouping(),
    triggerSubsyncFromConfig: () => triggerSubsyncFromConfig(),
    mineSentenceCard: () => mineSentenceCard(),
    mineSentenceMultiple: (timeoutMs) => {
      startPendingMineSentenceMultiple(timeoutMs);
    },
    cancelPendingMultiCopy: () => {
      cancelPendingMultiCopy();
    },
    cancelPendingMineSentenceMultiple: () => {
      cancelPendingMineSentenceMultiple();
    },
  })(),
);

const buildConfigHotReloadMessageMainDepsHandler = createBuildConfigHotReloadMessageMainDepsHandler({
  showMpvOsd: (message) => showMpvOsd(message),
  showDesktopNotification: (title, options) => showDesktopNotification(title, options),
});
const configHotReloadMessageMainDeps = buildConfigHotReloadMessageMainDepsHandler();
const notifyConfigHotReloadMessage = createConfigHotReloadMessageHandler(
  configHotReloadMessageMainDeps,
);
const buildWatchConfigPathMainDepsHandler = createBuildWatchConfigPathMainDepsHandler({
  fileExists: (targetPath) => fs.existsSync(targetPath),
  dirname: (targetPath) => path.dirname(targetPath),
  watchPath: (targetPath, listener) => fs.watch(targetPath, listener),
});
const watchConfigPathHandler = createWatchConfigPathHandler(buildWatchConfigPathMainDepsHandler());
const buildConfigHotReloadAppliedMainDepsHandler = createBuildConfigHotReloadAppliedMainDepsHandler({
  setKeybindings: (keybindings) => {
    appState.keybindings = keybindings as never;
  },
  refreshGlobalAndOverlayShortcuts: () => {
    refreshGlobalAndOverlayShortcuts();
  },
  setSecondarySubMode: (mode) => {
    appState.secondarySubMode = mode as never;
  },
  broadcastToOverlayWindows: (channel, payload) => {
    broadcastToOverlayWindows(channel, payload);
  },
  applyAnkiRuntimeConfigPatch: (patch) => {
    if (appState.ankiIntegration) {
      appState.ankiIntegration.applyRuntimeConfigPatch(patch as never);
    }
  },
});
const buildConfigHotReloadRuntimeMainDepsHandler = createBuildConfigHotReloadRuntimeMainDepsHandler({
  getCurrentConfig: () => getResolvedConfig(),
  reloadConfigStrict: () => configService.reloadConfigStrict(),
  watchConfigPath: (configPath, onChange) => watchConfigPathHandler(configPath, onChange),
  setTimeout: (callback, delayMs) => setTimeout(callback, delayMs),
  clearTimeout: (timeout) => clearTimeout(timeout),
  debounceMs: 250,
  onHotReloadApplied: createConfigHotReloadAppliedHandler(buildConfigHotReloadAppliedMainDepsHandler()),
  onRestartRequired: (fields) => notifyConfigHotReloadMessage(buildRestartRequiredConfigMessage(fields)),
  onInvalidConfig: notifyConfigHotReloadMessage,
  onValidationWarnings: (configPath, warnings) => {
    showDesktopNotification('SubMiner', {
      body: buildConfigWarningNotificationBody(configPath, warnings),
    });
  },
});
const buildHandleJellyfinRemotePlayMainDepsHandler =
  createBuildHandleJellyfinRemotePlayMainDepsHandler({
  getConfiguredSession: () => getConfiguredJellyfinSession(getResolvedJellyfinConfig()),
  getClientInfo: () => getJellyfinClientInfo(),
  getJellyfinConfig: () => getResolvedJellyfinConfig(),
  playJellyfinItem: (params) =>
    playJellyfinItemInMpv(params as Parameters<typeof playJellyfinItemInMpv>[0]),
  logWarn: (message) => logger.warn(message),
});
const buildHandleJellyfinRemotePlaystateMainDepsHandler =
  createBuildHandleJellyfinRemotePlaystateMainDepsHandler({
  getMpvClient: () => appState.mpvClient,
  sendMpvCommand: (client, command) => sendMpvCommandRuntime(client as MpvIpcClient, command),
  reportJellyfinRemoteProgress: (force) => reportJellyfinRemoteProgress(force),
  reportJellyfinRemoteStopped: () => reportJellyfinRemoteStopped(),
  jellyfinTicksToSeconds: (ticks) => jellyfinTicksToSecondsRuntime(ticks),
});
const buildHandleJellyfinRemoteGeneralCommandMainDepsHandler =
  createBuildHandleJellyfinRemoteGeneralCommandMainDepsHandler({
  getMpvClient: () => appState.mpvClient,
  sendMpvCommand: (client, command) => sendMpvCommandRuntime(client as MpvIpcClient, command),
  getActivePlayback: () => activeJellyfinRemotePlayback,
  reportJellyfinRemoteProgress: (force) => reportJellyfinRemoteProgress(force),
  logDebug: (message) => logger.debug(message),
});
const buildReportJellyfinRemoteProgressMainDepsHandler =
  createBuildReportJellyfinRemoteProgressMainDepsHandler({
  getActivePlayback: () => activeJellyfinRemotePlayback,
  clearActivePlayback: () => {
    activeJellyfinRemotePlayback = null;
  },
  getSession: () => appState.jellyfinRemoteSession,
  getMpvClient: () => appState.mpvClient,
  getNow: () => Date.now(),
  getLastProgressAtMs: () => jellyfinRemoteLastProgressAtMs,
  setLastProgressAtMs: (value) => {
    jellyfinRemoteLastProgressAtMs = value;
  },
  progressIntervalMs: JELLYFIN_REMOTE_PROGRESS_INTERVAL_MS,
  ticksPerSecond: JELLYFIN_TICKS_PER_SECOND,
  logDebug: (message, error) => logger.debug(message, error),
});
const buildReportJellyfinRemoteStoppedMainDepsHandler =
  createBuildReportJellyfinRemoteStoppedMainDepsHandler({
  getActivePlayback: () => activeJellyfinRemotePlayback,
  clearActivePlayback: () => {
    activeJellyfinRemotePlayback = null;
  },
  getSession: () => appState.jellyfinRemoteSession,
  logDebug: (message, error) => logger.debug(message, error),
});
const reportJellyfinRemoteProgressMainDeps =
  buildReportJellyfinRemoteProgressMainDepsHandler();
const reportJellyfinRemoteStoppedMainDeps =
  buildReportJellyfinRemoteStoppedMainDepsHandler();
const reportJellyfinRemoteProgress = createReportJellyfinRemoteProgressHandler(
  reportJellyfinRemoteProgressMainDeps,
);
const reportJellyfinRemoteStopped = createReportJellyfinRemoteStoppedHandler(
  reportJellyfinRemoteStoppedMainDeps,
);
const handleJellyfinRemotePlayMainDeps = buildHandleJellyfinRemotePlayMainDepsHandler();
const handleJellyfinRemotePlaystateMainDeps =
  buildHandleJellyfinRemotePlaystateMainDepsHandler();
const handleJellyfinRemoteGeneralCommandMainDeps =
  buildHandleJellyfinRemoteGeneralCommandMainDepsHandler();
const handleJellyfinRemotePlay = createHandleJellyfinRemotePlay(
  handleJellyfinRemotePlayMainDeps,
);
const handleJellyfinRemotePlaystate = createHandleJellyfinRemotePlaystate(
  handleJellyfinRemotePlaystateMainDeps,
);
const handleJellyfinRemoteGeneralCommand = createHandleJellyfinRemoteGeneralCommand(
  handleJellyfinRemoteGeneralCommandMainDeps,
);

const configHotReloadRuntime = createConfigHotReloadRuntime(buildConfigHotReloadRuntimeMainDepsHandler());

const buildDictionaryRootsHandler = createBuildDictionaryRootsMainHandler({
  dirname: __dirname,
  appPath: app.getAppPath(),
  resourcesPath: process.resourcesPath,
  userDataPath: USER_DATA_PATH,
  appUserDataPath: app.getPath('userData'),
  homeDir: os.homedir(),
  cwd: process.cwd(),
  joinPath: (...parts) => path.join(...parts),
});
const buildFrequencyDictionaryRootsHandler = createBuildFrequencyDictionaryRootsMainHandler({
  dirname: __dirname,
  appPath: app.getAppPath(),
  resourcesPath: process.resourcesPath,
  userDataPath: USER_DATA_PATH,
  appUserDataPath: app.getPath('userData'),
  homeDir: os.homedir(),
  cwd: process.cwd(),
  joinPath: (...parts) => path.join(...parts),
});

const jlptDictionaryRuntime = createJlptDictionaryRuntimeService(
  createBuildJlptDictionaryRuntimeMainDepsHandler({
    isJlptEnabled: () => getResolvedConfig().subtitleStyle.enableJlpt,
    getDictionaryRoots: () => buildDictionaryRootsHandler(),
    getJlptDictionarySearchPaths,
    setJlptLevelLookup: (lookup) => {
      appState.jlptLevelLookup = lookup as never;
    },
    logInfo: (message) => logger.info(message),
  })(),
);

const frequencyDictionaryRuntime = createFrequencyDictionaryRuntimeService(
  createBuildFrequencyDictionaryRuntimeMainDepsHandler({
    isFrequencyDictionaryEnabled: () => getResolvedConfig().subtitleStyle.frequencyDictionary.enabled,
    getDictionaryRoots: () => buildFrequencyDictionaryRootsHandler(),
    getFrequencyDictionarySearchPaths,
    getSourcePath: () => getResolvedConfig().subtitleStyle.frequencyDictionary.sourcePath,
    setFrequencyRankLookup: (lookup) => {
      appState.frequencyRankLookup = lookup as never;
    },
    logInfo: (message) => logger.info(message),
  })(),
);

const buildGetFieldGroupingResolverMainDepsHandler = createBuildGetFieldGroupingResolverMainDepsHandler(
  {
  getResolver: () => appState.fieldGroupingResolver,
  },
);
const getFieldGroupingResolverMainDeps =
  buildGetFieldGroupingResolverMainDepsHandler();
const getFieldGroupingResolverHandler = createGetFieldGroupingResolverHandler(
  getFieldGroupingResolverMainDeps,
);

function getFieldGroupingResolver(): ((choice: KikuFieldGroupingChoice) => void) | null {
  return getFieldGroupingResolverHandler();
}

const buildSetFieldGroupingResolverMainDepsHandler = createBuildSetFieldGroupingResolverMainDepsHandler(
  {
  setResolver: (resolver) => {
    appState.fieldGroupingResolver = resolver;
  },
  nextSequence: () => {
    appState.fieldGroupingResolverSequence += 1;
    return appState.fieldGroupingResolverSequence;
  },
  getSequence: () => appState.fieldGroupingResolverSequence,
  },
);
const setFieldGroupingResolverMainDeps =
  buildSetFieldGroupingResolverMainDepsHandler();
const setFieldGroupingResolverHandler = createSetFieldGroupingResolverHandler(
  setFieldGroupingResolverMainDeps,
);

function setFieldGroupingResolver(
  resolver: ((choice: KikuFieldGroupingChoice) => void) | null,
): void {
  setFieldGroupingResolverHandler(resolver);
}

const fieldGroupingOverlayRuntime = createFieldGroupingOverlayRuntime<OverlayHostedModal>(
  createBuildFieldGroupingOverlayMainDepsHandler<
    OverlayHostedModal,
    KikuFieldGroupingChoice
  >({
    getMainWindow: () => overlayManager.getMainWindow(),
    getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    getInvisibleOverlayVisible: () => overlayManager.getInvisibleOverlayVisible(),
    setVisibleOverlayVisible: (visible) => setVisibleOverlayVisible(visible),
    setInvisibleOverlayVisible: (visible) => setInvisibleOverlayVisible(visible),
    getResolver: () => getFieldGroupingResolver(),
    setResolver: (resolver) => setFieldGroupingResolver(resolver),
    getRestoreVisibleOverlayOnModalClose: () =>
      overlayModalRuntime.getRestoreVisibleOverlayOnModalClose(),
    sendToActiveOverlayWindow: (channel, payload, runtimeOptions) =>
      overlayModalRuntime.sendToActiveOverlayWindow(channel, payload, runtimeOptions),
  })(),
);
const createFieldGroupingCallback = fieldGroupingOverlayRuntime.createFieldGroupingCallback;

const SUBTITLE_POSITIONS_DIR = path.join(CONFIG_DIR, 'subtitle-positions');

const mediaRuntime = createMediaRuntimeService(
  createBuildMediaRuntimeMainDepsHandler({
    isRemoteMediaPath: (mediaPath) => isRemoteMediaPath(mediaPath),
    loadSubtitlePosition: () => loadSubtitlePosition(),
    getCurrentMediaPath: () => appState.currentMediaPath,
    getPendingSubtitlePosition: () => appState.pendingSubtitlePosition,
    getSubtitlePositionsDir: () => SUBTITLE_POSITIONS_DIR,
    setCurrentMediaPath: (nextPath: string | null) => {
      appState.currentMediaPath = nextPath;
    },
    clearPendingSubtitlePosition: () => {
      appState.pendingSubtitlePosition = null;
    },
    setSubtitlePosition: (position: SubtitlePosition | null) => {
      appState.subtitlePosition = position;
    },
    broadcastToOverlayWindows: (channel, payload) => {
      broadcastToOverlayWindows(channel, payload);
    },
    getCurrentMediaTitle: () => appState.currentMediaTitle,
    setCurrentMediaTitle: (title) => {
      appState.currentMediaTitle = title;
    },
  })(),
);

const overlayVisibilityRuntime = createOverlayVisibilityRuntimeService(
  createBuildOverlayVisibilityRuntimeMainDepsHandler({
    getMainWindow: () => overlayManager.getMainWindow(),
    getInvisibleWindow: () => overlayManager.getInvisibleWindow(),
    getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    getInvisibleOverlayVisible: () => overlayManager.getInvisibleOverlayVisible(),
    getWindowTracker: () => appState.windowTracker,
    getTrackerNotReadyWarningShown: () => appState.trackerNotReadyWarningShown,
    setTrackerNotReadyWarningShown: (shown: boolean) => {
      appState.trackerNotReadyWarningShown = shown;
    },
    updateVisibleOverlayBounds: (geometry: WindowGeometry) => updateVisibleOverlayBounds(geometry),
    updateInvisibleOverlayBounds: (geometry: WindowGeometry) =>
      updateInvisibleOverlayBounds(geometry),
    ensureOverlayWindowLevel: (window) => {
      ensureOverlayWindowLevel(window);
    },
    enforceOverlayLayerOrder: () => {
      enforceOverlayLayerOrder();
    },
    syncOverlayShortcuts: () => {
      overlayShortcutsRuntime.syncOverlayShortcuts();
    },
  })(),
);

const buildGetRuntimeOptionsStateMainDepsHandler = createBuildGetRuntimeOptionsStateMainDepsHandler({
  getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
});
const getRuntimeOptionsStateMainDeps = buildGetRuntimeOptionsStateMainDepsHandler();
const getRuntimeOptionsStateHandler = createGetRuntimeOptionsStateHandler(
  getRuntimeOptionsStateMainDeps,
);

function getRuntimeOptionsState(): RuntimeOptionState[] {
  return getRuntimeOptionsStateHandler();
}

function getOverlayWindows(): BrowserWindow[] {
  return overlayManager.getOverlayWindows();
}

const buildRestorePreviousSecondarySubVisibilityMainDepsHandler =
  createBuildRestorePreviousSecondarySubVisibilityMainDepsHandler({
    getMpvClient: () => appState.mpvClient,
  });
const restorePreviousSecondarySubVisibilityMainDeps =
  buildRestorePreviousSecondarySubVisibilityMainDepsHandler();
const restorePreviousSecondarySubVisibilityHandler = createRestorePreviousSecondarySubVisibilityHandler(
  restorePreviousSecondarySubVisibilityMainDeps,
);

function restorePreviousSecondarySubVisibility(): void {
  restorePreviousSecondarySubVisibilityHandler();
}

function broadcastToOverlayWindows(channel: string, ...args: unknown[]): void {
  overlayManager.broadcastToOverlayWindows(channel, ...args);
}

const buildBroadcastRuntimeOptionsChangedMainDepsHandler =
  createBuildBroadcastRuntimeOptionsChangedMainDepsHandler({
    broadcastRuntimeOptionsChangedRuntime,
    getRuntimeOptionsState: () => getRuntimeOptionsState(),
    broadcastToOverlayWindows: (channel, ...args) => broadcastToOverlayWindows(channel, ...args),
  });
const broadcastRuntimeOptionsChangedMainDeps =
  buildBroadcastRuntimeOptionsChangedMainDepsHandler();
const broadcastRuntimeOptionsChangedHandler = createBroadcastRuntimeOptionsChangedHandler(
  broadcastRuntimeOptionsChangedMainDeps,
);

function broadcastRuntimeOptionsChanged(): void {
  broadcastRuntimeOptionsChangedHandler();
}

const buildSendToActiveOverlayWindowMainDepsHandler =
  createBuildSendToActiveOverlayWindowMainDepsHandler({
    sendToActiveOverlayWindowRuntime: (channel, payload, runtimeOptions) =>
      overlayModalRuntime.sendToActiveOverlayWindow(channel, payload, runtimeOptions),
  });
const sendToActiveOverlayWindowMainDeps =
  buildSendToActiveOverlayWindowMainDepsHandler();
const sendToActiveOverlayWindowHandler = createSendToActiveOverlayWindowHandler(
  sendToActiveOverlayWindowMainDeps,
);

function sendToActiveOverlayWindow(
  channel: string,
  payload?: unknown,
  runtimeOptions?: { restoreOnModalClose?: OverlayHostedModal },
): boolean {
  return sendToActiveOverlayWindowHandler(channel, payload, runtimeOptions);
}

const buildSetOverlayDebugVisualizationEnabledMainDepsHandler =
  createBuildSetOverlayDebugVisualizationEnabledMainDepsHandler({
    setOverlayDebugVisualizationEnabledRuntime,
    getCurrentEnabled: () => appState.overlayDebugVisualizationEnabled,
    setCurrentEnabled: (next) => {
      appState.overlayDebugVisualizationEnabled = next;
    },
    broadcastToOverlayWindows: (channel, ...args) => broadcastToOverlayWindows(channel, ...args),
  });
const setOverlayDebugVisualizationEnabledMainDeps =
  buildSetOverlayDebugVisualizationEnabledMainDepsHandler();
const setOverlayDebugVisualizationEnabledHandler = createSetOverlayDebugVisualizationEnabledHandler(
  setOverlayDebugVisualizationEnabledMainDeps,
);

function setOverlayDebugVisualizationEnabled(enabled: boolean): void {
  setOverlayDebugVisualizationEnabledHandler(enabled);
}

const buildOpenRuntimeOptionsPaletteMainDepsHandler =
  createBuildOpenRuntimeOptionsPaletteMainDepsHandler({
    openRuntimeOptionsPaletteRuntime: () => overlayModalRuntime.openRuntimeOptionsPalette(),
  });
const openRuntimeOptionsPaletteMainDeps =
  buildOpenRuntimeOptionsPaletteMainDepsHandler();
const openRuntimeOptionsPaletteHandler = createOpenRuntimeOptionsPaletteHandler(
  openRuntimeOptionsPaletteMainDeps,
);

function openRuntimeOptionsPalette(): void {
  openRuntimeOptionsPaletteHandler();
}

function getResolvedConfig() {
  return configService.getConfig();
}

const buildGetResolvedJellyfinConfigMainDepsHandler =
  createBuildGetResolvedJellyfinConfigMainDepsHandler({
  getResolvedConfig: () => getResolvedConfig(),
  loadStoredToken: () => jellyfinTokenStore.loadToken(),
});
const getResolvedJellyfinConfigMainDeps =
  buildGetResolvedJellyfinConfigMainDepsHandler();
const getResolvedJellyfinConfigHandler = createGetResolvedJellyfinConfigHandler(
  getResolvedJellyfinConfigMainDeps,
);

function getResolvedJellyfinConfig() {
  return getResolvedJellyfinConfigHandler();
}

const buildGetJellyfinClientInfoMainDepsHandler = createBuildGetJellyfinClientInfoMainDepsHandler({
  getResolvedJellyfinConfig: () => getResolvedJellyfinConfig(),
  getDefaultJellyfinConfig: () => DEFAULT_CONFIG.jellyfin,
});
const getJellyfinClientInfoMainDeps =
  buildGetJellyfinClientInfoMainDepsHandler();
const getJellyfinClientInfoHandler = createGetJellyfinClientInfoHandler(
  getJellyfinClientInfoMainDeps,
);

function getJellyfinClientInfo(config = getResolvedJellyfinConfig()) {
  return getJellyfinClientInfoHandler(config);
}

const buildWaitForMpvConnectedMainDepsHandler = createBuildWaitForMpvConnectedMainDepsHandler({
  getMpvClient: () => appState.mpvClient,
  now: () => Date.now(),
  sleep: (delayMs) => new Promise((resolve) => setTimeout(resolve, delayMs)),
});
const waitForMpvConnectedMainDeps = buildWaitForMpvConnectedMainDepsHandler();
const waitForMpvConnected = createWaitForMpvConnectedHandler(
  waitForMpvConnectedMainDeps,
);

const buildLaunchMpvIdleForJellyfinPlaybackMainDepsHandler =
  createBuildLaunchMpvIdleForJellyfinPlaybackMainDepsHandler({
    getSocketPath: () => appState.mpvSocketPath,
    platform: process.platform,
    execPath: process.execPath,
    defaultMpvLogPath: DEFAULT_MPV_LOG_PATH,
    defaultMpvArgs: MPV_JELLYFIN_DEFAULT_ARGS,
    removeSocketPath: (socketPath) => {
      fs.rmSync(socketPath, { force: true });
    },
    spawnMpv: (args) =>
      spawn('mpv', args, {
        detached: true,
        stdio: 'ignore',
      }),
    logWarn: (message, error) => logger.warn(message, error),
    logInfo: (message) => logger.info(message),
  });
const launchMpvIdleForJellyfinPlaybackMainDeps =
  buildLaunchMpvIdleForJellyfinPlaybackMainDepsHandler();
const launchMpvIdleForJellyfinPlayback = createLaunchMpvIdleForJellyfinPlaybackHandler(
  launchMpvIdleForJellyfinPlaybackMainDeps,
);

const buildEnsureMpvConnectedForJellyfinPlaybackMainDepsHandler =
  createBuildEnsureMpvConnectedForJellyfinPlaybackMainDepsHandler({
    getMpvClient: () => appState.mpvClient,
    setMpvClient: (client) => {
      appState.mpvClient = client as MpvIpcClient | null;
    },
    createMpvClient: () => createMpvClientRuntimeService(),
    waitForMpvConnected: (timeoutMs) => waitForMpvConnected(timeoutMs),
    launchMpvIdleForJellyfinPlayback: () => launchMpvIdleForJellyfinPlayback(),
    getAutoLaunchInFlight: () => jellyfinMpvAutoLaunchInFlight,
    setAutoLaunchInFlight: (promise) => {
      jellyfinMpvAutoLaunchInFlight = promise;
    },
    connectTimeoutMs: JELLYFIN_MPV_CONNECT_TIMEOUT_MS,
    autoLaunchTimeoutMs: JELLYFIN_MPV_AUTO_LAUNCH_TIMEOUT_MS,
  });
const ensureMpvConnectedForJellyfinPlaybackMainDeps =
  buildEnsureMpvConnectedForJellyfinPlaybackMainDepsHandler();
const ensureMpvConnectedForJellyfinPlayback = createEnsureMpvConnectedForJellyfinPlaybackHandler(
  ensureMpvConnectedForJellyfinPlaybackMainDeps,
);

const buildPreloadJellyfinExternalSubtitlesMainDepsHandler =
  createBuildPreloadJellyfinExternalSubtitlesMainDepsHandler({
    listJellyfinSubtitleTracks: (session, clientInfo, itemId) =>
      listJellyfinSubtitleTracksRuntime(session, clientInfo, itemId),
    getMpvClient: () => appState.mpvClient,
    sendMpvCommand: (command) => {
      sendMpvCommandRuntime(appState.mpvClient, command);
    },
    wait: (ms) => new Promise<void>((resolve) => setTimeout(resolve, ms)),
    logDebug: (message, error) => {
      logger.debug(message, error);
    },
  });
const preloadJellyfinExternalSubtitlesMainDeps =
  buildPreloadJellyfinExternalSubtitlesMainDepsHandler();
const preloadJellyfinExternalSubtitles = createPreloadJellyfinExternalSubtitlesHandler(
  preloadJellyfinExternalSubtitlesMainDeps,
);

const buildPlayJellyfinItemInMpvMainDepsHandler = createBuildPlayJellyfinItemInMpvMainDepsHandler({
  ensureMpvConnectedForPlayback: () => ensureMpvConnectedForJellyfinPlayback(),
  getMpvClient: () => appState.mpvClient,
  resolvePlaybackPlan: (params) =>
    resolveJellyfinPlaybackPlanRuntime(
      params.session,
      params.clientInfo,
      params.jellyfinConfig as ReturnType<typeof getResolvedJellyfinConfig>,
      {
        itemId: params.itemId,
        audioStreamIndex: params.audioStreamIndex ?? undefined,
        subtitleStreamIndex: params.subtitleStreamIndex ?? undefined,
      },
    ),
  applyJellyfinMpvDefaults: (mpvClient) =>
    applyJellyfinMpvDefaults((mpvClient as unknown) as MpvIpcClient),
  sendMpvCommand: (command) => sendMpvCommandRuntime(appState.mpvClient, command),
  armQuitOnDisconnect: () => {
    jellyfinPlayQuitOnDisconnectArmed = false;
    setTimeout(() => {
      jellyfinPlayQuitOnDisconnectArmed = true;
    }, 3000);
  },
  schedule: (callback, delayMs) => {
    setTimeout(callback, delayMs);
  },
  convertTicksToSeconds: (ticks) => jellyfinTicksToSecondsRuntime(ticks),
  preloadExternalSubtitles: (params) => {
    void preloadJellyfinExternalSubtitles(params);
  },
  setActivePlayback: (state) => {
    activeJellyfinRemotePlayback = state as ActiveJellyfinRemotePlaybackState;
  },
  setLastProgressAtMs: (value) => {
    jellyfinRemoteLastProgressAtMs = value;
  },
  reportPlaying: (payload) => {
    void appState.jellyfinRemoteSession?.reportPlaying(payload);
  },
  showMpvOsd: (text) => {
    showMpvOsd(text);
  },
});
const playJellyfinItemInMpvMainDeps = buildPlayJellyfinItemInMpvMainDepsHandler();
const playJellyfinItemInMpv = createPlayJellyfinItemInMpvHandler(
  playJellyfinItemInMpvMainDeps,
);

const buildHandleJellyfinAuthCommandsMainDepsHandler =
  createBuildHandleJellyfinAuthCommandsMainDepsHandler({
  patchRawConfig: (patch) => {
    configService.patchRawConfig(patch);
  },
  authenticateWithPassword: (serverUrl, username, password, clientInfo) =>
    authenticateWithPasswordRuntime(serverUrl, username, password, clientInfo),
  saveStoredToken: (token) => jellyfinTokenStore.saveToken(token),
  clearStoredToken: () => jellyfinTokenStore.clearToken(),
  logInfo: (message) => logger.info(message),
});
const handleJellyfinAuthCommandsMainDeps =
  buildHandleJellyfinAuthCommandsMainDepsHandler();
const handleJellyfinAuthCommands = createHandleJellyfinAuthCommands(
  handleJellyfinAuthCommandsMainDeps,
);

const buildHandleJellyfinListCommandsMainDepsHandler =
  createBuildHandleJellyfinListCommandsMainDepsHandler({
  listJellyfinLibraries: (session, clientInfo) => listJellyfinLibrariesRuntime(session, clientInfo),
  listJellyfinItems: (session, clientInfo, params) =>
    listJellyfinItemsRuntime(session, clientInfo, params),
  listJellyfinSubtitleTracks: (session, clientInfo, itemId) =>
    listJellyfinSubtitleTracksRuntime(session, clientInfo, itemId),
  logInfo: (message) => logger.info(message),
});
const handleJellyfinListCommandsMainDeps =
  buildHandleJellyfinListCommandsMainDepsHandler();
const handleJellyfinListCommands = createHandleJellyfinListCommands(
  handleJellyfinListCommandsMainDeps,
);

const buildHandleJellyfinPlayCommandMainDepsHandler = createBuildHandleJellyfinPlayCommandMainDepsHandler(
  {
  playJellyfinItemInMpv: (params) =>
    playJellyfinItemInMpv(params as Parameters<typeof playJellyfinItemInMpv>[0]),
  logWarn: (message) => logger.warn(message),
  },
);
const handleJellyfinPlayCommandMainDeps =
  buildHandleJellyfinPlayCommandMainDepsHandler();
const handleJellyfinPlayCommand = createHandleJellyfinPlayCommand(
  handleJellyfinPlayCommandMainDeps,
);

const buildHandleJellyfinRemoteAnnounceCommandMainDepsHandler =
  createBuildHandleJellyfinRemoteAnnounceCommandMainDepsHandler({
  startJellyfinRemoteSession: () => startJellyfinRemoteSession(),
  getRemoteSession: () => appState.jellyfinRemoteSession,
  logInfo: (message) => logger.info(message),
  logWarn: (message) => logger.warn(message),
});
const handleJellyfinRemoteAnnounceCommandMainDeps =
  buildHandleJellyfinRemoteAnnounceCommandMainDepsHandler();
const handleJellyfinRemoteAnnounceCommand = createHandleJellyfinRemoteAnnounceCommand(
  handleJellyfinRemoteAnnounceCommandMainDeps,
);

const buildStartJellyfinRemoteSessionMainDepsHandler =
  createBuildStartJellyfinRemoteSessionMainDepsHandler({
  getJellyfinConfig: () => getResolvedJellyfinConfig(),
  getCurrentSession: () => appState.jellyfinRemoteSession,
  setCurrentSession: (session) => {
    appState.jellyfinRemoteSession = session as typeof appState.jellyfinRemoteSession;
  },
  createRemoteSessionService: (options) => new JellyfinRemoteSessionService(options),
  defaultDeviceId: DEFAULT_CONFIG.jellyfin.deviceId,
  defaultClientName: DEFAULT_CONFIG.jellyfin.clientName,
  defaultClientVersion: DEFAULT_CONFIG.jellyfin.clientVersion,
  handlePlay: (payload) => handleJellyfinRemotePlay(payload),
  handlePlaystate: (payload) => handleJellyfinRemotePlaystate(payload),
  handleGeneralCommand: (payload) => handleJellyfinRemoteGeneralCommand(payload),
  logInfo: (message) => logger.info(message),
  logWarn: (message, details) => logger.warn(message, details),
});
const startJellyfinRemoteSessionMainDeps =
  buildStartJellyfinRemoteSessionMainDepsHandler();
const startJellyfinRemoteSession = createStartJellyfinRemoteSessionHandler(
  startJellyfinRemoteSessionMainDeps,
);

const buildStopJellyfinRemoteSessionMainDepsHandler =
  createBuildStopJellyfinRemoteSessionMainDepsHandler({
  getCurrentSession: () => appState.jellyfinRemoteSession,
  setCurrentSession: (session) => {
    appState.jellyfinRemoteSession = session as typeof appState.jellyfinRemoteSession;
  },
  clearActivePlayback: () => {
    activeJellyfinRemotePlayback = null;
  },
});
const stopJellyfinRemoteSessionMainDeps =
  buildStopJellyfinRemoteSessionMainDepsHandler();
const stopJellyfinRemoteSession = createStopJellyfinRemoteSessionHandler(
  stopJellyfinRemoteSessionMainDeps,
);

const buildRunJellyfinCommandMainDepsHandler = createBuildRunJellyfinCommandMainDepsHandler({
  getJellyfinConfig: () => getResolvedJellyfinConfig(),
  defaultServerUrl: DEFAULT_CONFIG.jellyfin.serverUrl,
  getJellyfinClientInfo: (jellyfinConfig) => getJellyfinClientInfo(jellyfinConfig),
  handleAuthCommands: (params) => handleJellyfinAuthCommands(params),
  handleRemoteAnnounceCommand: (args) => handleJellyfinRemoteAnnounceCommand(args),
  handleListCommands: (params) => handleJellyfinListCommands(params),
  handlePlayCommand: (params) => handleJellyfinPlayCommand(params),
});
const runJellyfinCommandMainDeps = buildRunJellyfinCommandMainDepsHandler();
const runJellyfinCommand = createRunJellyfinCommandHandler(
  runJellyfinCommandMainDeps,
);

const buildNotifyAnilistSetupMainDepsHandler = createBuildNotifyAnilistSetupMainDepsHandler({
  hasMpvClient: () => Boolean(appState.mpvClient),
  showMpvOsd: (message) => showMpvOsd(message),
  showDesktopNotification: (title, options) => showDesktopNotification(title, options),
  logInfo: (message) => logger.info(message),
});
const notifyAnilistSetupMainDeps = buildNotifyAnilistSetupMainDepsHandler();
const notifyAnilistSetup = createNotifyAnilistSetupHandler(
  notifyAnilistSetupMainDeps,
);

const buildConsumeAnilistSetupTokenFromUrlMainDepsHandler =
  createBuildConsumeAnilistSetupTokenFromUrlMainDepsHandler({
    consumeAnilistSetupCallbackUrl,
    saveToken: (token) => anilistTokenStore.saveToken(token),
    setCachedToken: (token) => {
      anilistCachedAccessToken = token;
    },
    setResolvedState: (resolvedAt) => {
      anilistStateRuntime.setClientSecretState({
        status: 'resolved',
        source: 'stored',
        message: 'saved token from AniList login',
        resolvedAt,
        errorAt: null,
      });
    },
    setSetupPageOpened: (opened) => {
      appState.anilistSetupPageOpened = opened;
    },
    onSuccess: () => {
      notifyAnilistSetup('AniList login success');
    },
    closeWindow: () => {
      if (appState.anilistSetupWindow && !appState.anilistSetupWindow.isDestroyed()) {
        appState.anilistSetupWindow.close();
      }
    },
  });
const consumeAnilistSetupTokenFromUrlMainDeps =
  buildConsumeAnilistSetupTokenFromUrlMainDepsHandler();
const consumeAnilistSetupTokenFromUrl = createConsumeAnilistSetupTokenFromUrlHandler(
  consumeAnilistSetupTokenFromUrlMainDeps,
);

const buildHandleAnilistSetupProtocolUrlMainDepsHandler =
  createBuildHandleAnilistSetupProtocolUrlMainDepsHandler({
    consumeAnilistSetupTokenFromUrl: (rawUrl) => consumeAnilistSetupTokenFromUrl(rawUrl),
    logWarn: (message, details) => logger.warn(message, details),
  });
const handleAnilistSetupProtocolUrlMainDeps =
  buildHandleAnilistSetupProtocolUrlMainDepsHandler();
const handleAnilistSetupProtocolUrl = createHandleAnilistSetupProtocolUrlHandler(
  handleAnilistSetupProtocolUrlMainDeps,
);

const buildRegisterSubminerProtocolClientMainDepsHandler =
  createBuildRegisterSubminerProtocolClientMainDepsHandler({
    isDefaultApp: () => Boolean(process.defaultApp),
    getArgv: () => process.argv,
    execPath: process.execPath,
    resolvePath: (value) => path.resolve(value),
    setAsDefaultProtocolClient: (scheme, appPath, args) =>
      appPath
        ? app.setAsDefaultProtocolClient(scheme, appPath, args)
        : app.setAsDefaultProtocolClient(scheme),
    logWarn: (message, details) => logger.warn(message, details),
  });
const registerSubminerProtocolClientMainDeps =
  buildRegisterSubminerProtocolClientMainDepsHandler();
const registerSubminerProtocolClient = createRegisterSubminerProtocolClientHandler(
  registerSubminerProtocolClientMainDeps,
);

const maybeFocusExistingAnilistSetupWindow = createMaybeFocusExistingAnilistSetupWindowHandler({
  getSetupWindow: () => appState.anilistSetupWindow,
});
const buildOpenAnilistSetupWindowMainDepsHandler = createBuildOpenAnilistSetupWindowMainDepsHandler({
  maybeFocusExistingSetupWindow: maybeFocusExistingAnilistSetupWindow,
  createSetupWindow: () =>
    new BrowserWindow({
      width: 1000,
      height: 760,
      title: 'Anilist Setup',
      show: true,
      autoHideMenuBar: true,
      webPreferences: {
        nodeIntegration: false,
        contextIsolation: true,
      },
    }),
  buildAuthorizeUrl: () =>
    buildAnilistSetupUrl({
      authorizeUrl: ANILIST_SETUP_CLIENT_ID_URL,
      clientId: ANILIST_DEFAULT_CLIENT_ID,
      responseType: ANILIST_SETUP_RESPONSE_TYPE,
    }),
  consumeCallbackUrl: (rawUrl) => consumeAnilistSetupTokenFromUrl(rawUrl),
  openSetupInBrowser: (authorizeUrl) =>
    openAnilistSetupInBrowser({
      authorizeUrl,
      openExternal: (url) => shell.openExternal(url),
      logError: (message, error) => logger.error(message, error),
    }),
  loadManualTokenEntry: (setupWindow, authorizeUrl) =>
    loadAnilistManualTokenEntry({
      setupWindow: setupWindow as BrowserWindow,
      authorizeUrl,
      developerSettingsUrl: ANILIST_DEVELOPER_SETTINGS_URL,
      logWarn: (message, data) => logger.warn(message, data),
    }),
  redirectUri: ANILIST_REDIRECT_URI,
  developerSettingsUrl: ANILIST_DEVELOPER_SETTINGS_URL,
  isAllowedExternalUrl: (url) => isAllowedAnilistExternalUrl(url),
  isAllowedNavigationUrl: (url) => isAllowedAnilistSetupNavigationUrl(url),
  logWarn: (message, details) => logger.warn(message, details),
  logError: (message, details) => logger.error(message, details),
  clearSetupWindow: () => {
    appState.anilistSetupWindow = null;
  },
  setSetupPageOpened: (opened) => {
    appState.anilistSetupPageOpened = opened;
  },
  setSetupWindow: (setupWindow) => {
    appState.anilistSetupWindow = setupWindow as BrowserWindow;
  },
  openExternal: (url) => {
    void shell.openExternal(url);
  },
});

function openAnilistSetupWindow(): void {
  createOpenAnilistSetupWindowHandler(buildOpenAnilistSetupWindowMainDepsHandler())();
}

const maybeFocusExistingJellyfinSetupWindow = createMaybeFocusExistingJellyfinSetupWindowHandler({
  getSetupWindow: () => appState.jellyfinSetupWindow,
});
const buildOpenJellyfinSetupWindowMainDepsHandler =
  createBuildOpenJellyfinSetupWindowMainDepsHandler({
    maybeFocusExistingSetupWindow: maybeFocusExistingJellyfinSetupWindow,
    createSetupWindow: () =>
      new BrowserWindow({
        width: 520,
        height: 560,
        title: 'Jellyfin Setup',
        show: true,
        autoHideMenuBar: true,
        webPreferences: {
          nodeIntegration: false,
          contextIsolation: true,
        },
      }),
    getResolvedJellyfinConfig: () => getResolvedJellyfinConfig(),
    buildSetupFormHtml: (defaultServer, defaultUser) =>
      buildJellyfinSetupFormHtml(defaultServer, defaultUser),
    parseSubmissionUrl: (rawUrl) => parseJellyfinSetupSubmissionUrl(rawUrl),
    authenticateWithPassword: (server, username, password, clientInfo) =>
      authenticateWithPasswordRuntime(server, username, password, clientInfo),
    getJellyfinClientInfo: () => getJellyfinClientInfo(),
    saveStoredToken: (token) => jellyfinTokenStore.saveToken(token),
    patchJellyfinConfig: (session) => {
      configService.patchRawConfig({
        jellyfin: {
          enabled: true,
          serverUrl: session.serverUrl,
          username: session.username,
          accessToken: '',
          userId: session.userId,
        },
      });
    },
    logInfo: (message) => logger.info(message),
    logError: (message, error) => logger.error(message, error),
    showMpvOsd: (message) => showMpvOsd(message),
    clearSetupWindow: () => {
      appState.jellyfinSetupWindow = null;
    },
    setSetupWindow: (window) => {
      appState.jellyfinSetupWindow = window as BrowserWindow;
    },
    encodeURIComponent: (value) => encodeURIComponent(value),
  });

function openJellyfinSetupWindow(): void {
  createOpenJellyfinSetupWindowHandler(buildOpenJellyfinSetupWindowMainDepsHandler())();
}

const buildRefreshAnilistClientSecretStateMainDepsHandler =
  createBuildRefreshAnilistClientSecretStateMainDepsHandler({
    getResolvedConfig: () => getResolvedConfig(),
    isAnilistTrackingEnabled: (config) => isAnilistTrackingEnabled(config as ResolvedConfig),
    getCachedAccessToken: () => anilistCachedAccessToken,
    setCachedAccessToken: (token) => {
      anilistCachedAccessToken = token;
    },
    saveStoredToken: (token) => {
      anilistTokenStore.saveToken(token);
    },
    loadStoredToken: () => anilistTokenStore.loadToken(),
    setClientSecretState: (state) => {
      anilistStateRuntime.setClientSecretState(state);
    },
    getAnilistSetupPageOpened: () => appState.anilistSetupPageOpened,
    setAnilistSetupPageOpened: (opened) => {
      appState.anilistSetupPageOpened = opened;
    },
    openAnilistSetupWindow: () => {
      openAnilistSetupWindow();
    },
    now: () => Date.now(),
  });
const refreshAnilistClientSecretStateMainDeps =
  buildRefreshAnilistClientSecretStateMainDepsHandler();
const refreshAnilistClientSecretState = createRefreshAnilistClientSecretStateHandler(
  refreshAnilistClientSecretStateMainDeps,
);

const buildGetCurrentAnilistMediaKeyMainDepsHandler =
  createBuildGetCurrentAnilistMediaKeyMainDepsHandler({
  getCurrentMediaPath: () => appState.currentMediaPath,
});
const getCurrentAnilistMediaKeyMainDeps =
  buildGetCurrentAnilistMediaKeyMainDepsHandler();
const getCurrentAnilistMediaKey = createGetCurrentAnilistMediaKeyHandler(
  getCurrentAnilistMediaKeyMainDeps,
);

const buildResetAnilistMediaTrackingMainDepsHandler =
  createBuildResetAnilistMediaTrackingMainDepsHandler({
  setMediaKey: (value) => {
    anilistCurrentMediaKey = value;
  },
  setMediaDurationSec: (value) => {
    anilistCurrentMediaDurationSec = value;
  },
  setMediaGuess: (value) => {
    anilistCurrentMediaGuess = value;
  },
  setMediaGuessPromise: (value) => {
    anilistCurrentMediaGuessPromise = value;
  },
  setLastDurationProbeAtMs: (value) => {
    anilistLastDurationProbeAtMs = value;
  },
});
const resetAnilistMediaTrackingMainDeps =
  buildResetAnilistMediaTrackingMainDepsHandler();
const resetAnilistMediaTracking = createResetAnilistMediaTrackingHandler(
  resetAnilistMediaTrackingMainDeps,
);

const buildGetAnilistMediaGuessRuntimeStateMainDepsHandler =
  createBuildGetAnilistMediaGuessRuntimeStateMainDepsHandler({
  getMediaKey: () => anilistCurrentMediaKey,
  getMediaDurationSec: () => anilistCurrentMediaDurationSec,
  getMediaGuess: () => anilistCurrentMediaGuess,
  getMediaGuessPromise: () => anilistCurrentMediaGuessPromise,
  getLastDurationProbeAtMs: () => anilistLastDurationProbeAtMs,
});
const getAnilistMediaGuessRuntimeStateMainDeps =
  buildGetAnilistMediaGuessRuntimeStateMainDepsHandler();
const getAnilistMediaGuessRuntimeState = createGetAnilistMediaGuessRuntimeStateHandler(
  getAnilistMediaGuessRuntimeStateMainDeps,
);

const buildSetAnilistMediaGuessRuntimeStateMainDepsHandler =
  createBuildSetAnilistMediaGuessRuntimeStateMainDepsHandler({
  setMediaKey: (value) => {
    anilistCurrentMediaKey = value;
  },
  setMediaDurationSec: (value) => {
    anilistCurrentMediaDurationSec = value;
  },
  setMediaGuess: (value) => {
    anilistCurrentMediaGuess = value;
  },
  setMediaGuessPromise: (value) => {
    anilistCurrentMediaGuessPromise = value;
  },
  setLastDurationProbeAtMs: (value) => {
    anilistLastDurationProbeAtMs = value;
  },
});
const setAnilistMediaGuessRuntimeStateMainDeps =
  buildSetAnilistMediaGuessRuntimeStateMainDepsHandler();
const setAnilistMediaGuessRuntimeState = createSetAnilistMediaGuessRuntimeStateHandler(
  setAnilistMediaGuessRuntimeStateMainDeps,
);

const buildResetAnilistMediaGuessStateMainDepsHandler =
  createBuildResetAnilistMediaGuessStateMainDepsHandler({
  setMediaGuess: (value) => {
    anilistCurrentMediaGuess = value;
  },
  setMediaGuessPromise: (value) => {
    anilistCurrentMediaGuessPromise = value;
  },
});
const resetAnilistMediaGuessStateMainDeps =
  buildResetAnilistMediaGuessStateMainDepsHandler();
const resetAnilistMediaGuessState = createResetAnilistMediaGuessStateHandler(
  resetAnilistMediaGuessStateMainDeps,
);

const buildMaybeProbeAnilistDurationMainDepsHandler =
  createBuildMaybeProbeAnilistDurationMainDepsHandler({
  getState: () => getAnilistMediaGuessRuntimeState(),
  setState: (state) => {
    setAnilistMediaGuessRuntimeState(state);
  },
  durationRetryIntervalMs: ANILIST_DURATION_RETRY_INTERVAL_MS,
  now: () => Date.now(),
  requestMpvDuration: async () => appState.mpvClient?.requestProperty('duration'),
  logWarn: (message, error) => logger.warn(message, error),
});
const maybeProbeAnilistDurationMainDeps =
  buildMaybeProbeAnilistDurationMainDepsHandler();
const maybeProbeAnilistDuration = createMaybeProbeAnilistDurationHandler(
  maybeProbeAnilistDurationMainDeps,
);

const buildEnsureAnilistMediaGuessMainDepsHandler = createBuildEnsureAnilistMediaGuessMainDepsHandler(
  {
  getState: () => getAnilistMediaGuessRuntimeState(),
  setState: (state) => {
    setAnilistMediaGuessRuntimeState(state);
  },
  resolveMediaPathForJimaku: (currentMediaPath) => mediaRuntime.resolveMediaPathForJimaku(currentMediaPath),
  getCurrentMediaPath: () => appState.currentMediaPath,
  getCurrentMediaTitle: () => appState.currentMediaTitle,
  guessAnilistMediaInfo: (mediaPath, mediaTitle) => guessAnilistMediaInfo(mediaPath, mediaTitle),
},
);
const ensureAnilistMediaGuessMainDeps =
  buildEnsureAnilistMediaGuessMainDepsHandler();
const ensureAnilistMediaGuess = createEnsureAnilistMediaGuessHandler(
  ensureAnilistMediaGuessMainDeps,
);

const rememberAnilistAttemptedUpdate = (key: string): void => {
  rememberAnilistAttemptedUpdateKey(anilistAttemptedUpdateKeys, key, ANILIST_MAX_ATTEMPTED_UPDATE_KEYS);
};

const buildProcessNextAnilistRetryUpdateMainDepsHandler =
  createBuildProcessNextAnilistRetryUpdateMainDepsHandler({
  nextReady: () => anilistUpdateQueue.nextReady(),
  refreshRetryQueueState: () => anilistStateRuntime.refreshRetryQueueState(),
  setLastAttemptAt: (value) => {
    appState.anilistRetryQueueState.lastAttemptAt = value;
  },
  setLastError: (value) => {
    appState.anilistRetryQueueState.lastError = value;
  },
  refreshAnilistClientSecretState: () => refreshAnilistClientSecretState(),
  updateAnilistPostWatchProgress: (accessToken, title, episode) =>
    updateAnilistPostWatchProgress(accessToken, title, episode),
  markSuccess: (key) => {
    anilistUpdateQueue.markSuccess(key);
  },
  rememberAttemptedUpdateKey: (key) => {
    rememberAnilistAttemptedUpdate(key);
  },
  markFailure: (key, message) => {
    anilistUpdateQueue.markFailure(key, message);
  },
  logInfo: (message) => logger.info(message),
  now: () => Date.now(),
});
const processNextAnilistRetryUpdateMainDeps =
  buildProcessNextAnilistRetryUpdateMainDepsHandler();
const processNextAnilistRetryUpdate = createProcessNextAnilistRetryUpdateHandler(
  processNextAnilistRetryUpdateMainDeps,
);

const buildMaybeRunAnilistPostWatchUpdateMainDepsHandler =
  createBuildMaybeRunAnilistPostWatchUpdateMainDepsHandler({
  getInFlight: () => anilistUpdateInFlight,
  setInFlight: (value) => {
    anilistUpdateInFlight = value;
  },
  getResolvedConfig: () => getResolvedConfig(),
  isAnilistTrackingEnabled: (config) => isAnilistTrackingEnabled(config as ResolvedConfig),
  getCurrentMediaKey: () => getCurrentAnilistMediaKey(),
  hasMpvClient: () => Boolean(appState.mpvClient),
  getTrackedMediaKey: () => anilistCurrentMediaKey,
  resetTrackedMedia: (mediaKey) => {
    resetAnilistMediaTracking(mediaKey);
  },
  getWatchedSeconds: () => appState.mpvClient?.currentTimePos ?? Number.NaN,
  maybeProbeAnilistDuration: (mediaKey) => maybeProbeAnilistDuration(mediaKey),
  ensureAnilistMediaGuess: (mediaKey) => ensureAnilistMediaGuess(mediaKey),
  hasAttemptedUpdateKey: (key) => anilistAttemptedUpdateKeys.has(key),
  processNextAnilistRetryUpdate: () => processNextAnilistRetryUpdate(),
  refreshAnilistClientSecretState: () => refreshAnilistClientSecretState(),
  enqueueRetry: (key, title, episode) => {
    anilistUpdateQueue.enqueue(key, title, episode);
  },
  markRetryFailure: (key, message) => {
    anilistUpdateQueue.markFailure(key, message);
  },
  markRetrySuccess: (key) => {
    anilistUpdateQueue.markSuccess(key);
  },
  refreshRetryQueueState: () => anilistStateRuntime.refreshRetryQueueState(),
  updateAnilistPostWatchProgress: (accessToken, title, episode) =>
    updateAnilistPostWatchProgress(accessToken, title, episode),
  rememberAttemptedUpdateKey: (key) => {
    rememberAnilistAttemptedUpdate(key);
  },
  showMpvOsd: (message) => showMpvOsd(message),
  logInfo: (message) => logger.info(message),
  logWarn: (message) => logger.warn(message),
  minWatchSeconds: ANILIST_UPDATE_MIN_WATCH_SECONDS,
  minWatchRatio: ANILIST_UPDATE_MIN_WATCH_RATIO,
});
const maybeRunAnilistPostWatchUpdateMainDeps =
  buildMaybeRunAnilistPostWatchUpdateMainDepsHandler();
const maybeRunAnilistPostWatchUpdate = createMaybeRunAnilistPostWatchUpdateHandler(
  maybeRunAnilistPostWatchUpdateMainDeps,
);

const buildLoadSubtitlePositionMainDepsHandler = createBuildLoadSubtitlePositionMainDepsHandler({
  loadSubtitlePositionCore: () =>
    loadSubtitlePositionCore({
      currentMediaPath: appState.currentMediaPath,
      fallbackPosition: getResolvedConfig().subtitlePosition,
      subtitlePositionsDir: SUBTITLE_POSITIONS_DIR,
    }),
  setSubtitlePosition: (position) => {
    appState.subtitlePosition = position;
  },
});
const loadSubtitlePositionMainDeps = buildLoadSubtitlePositionMainDepsHandler();
const loadSubtitlePosition = createLoadSubtitlePositionHandler(
  loadSubtitlePositionMainDeps,
);

const buildSaveSubtitlePositionMainDepsHandler = createBuildSaveSubtitlePositionMainDepsHandler({
  saveSubtitlePositionCore: (position) => {
    saveSubtitlePositionCore({
      position,
      currentMediaPath: appState.currentMediaPath,
      subtitlePositionsDir: SUBTITLE_POSITIONS_DIR,
      onQueuePending: (queued) => {
        appState.pendingSubtitlePosition = queued;
      },
      onPersisted: () => {
        appState.pendingSubtitlePosition = null;
      },
    });
  },
  setSubtitlePosition: (position) => {
    appState.subtitlePosition = position;
  },
});
const saveSubtitlePositionMainDeps = buildSaveSubtitlePositionMainDepsHandler();
const saveSubtitlePosition = createSaveSubtitlePositionHandler(
  saveSubtitlePositionMainDeps,
);

registerSubminerProtocolClient();

const buildRegisterProtocolUrlHandlersMainDepsHandler =
  createBuildRegisterProtocolUrlHandlersMainDepsHandler({
  registerOpenUrl: (listener) => {
    app.on('open-url', listener);
  },
  registerSecondInstance: (listener) => {
    app.on('second-instance', listener);
  },
  handleAnilistSetupProtocolUrl: (rawUrl) => handleAnilistSetupProtocolUrl(rawUrl),
  findAnilistSetupDeepLinkArgvUrl: (argv) => findAnilistSetupDeepLinkArgvUrl(argv),
  logUnhandledOpenUrl: (rawUrl) => {
    logger.warn('Unhandled app protocol URL', { rawUrl });
  },
  logUnhandledSecondInstanceUrl: (rawUrl) => {
    logger.warn('Unhandled second-instance protocol URL', { rawUrl });
  },
});
const registerProtocolUrlHandlersMainDeps = buildRegisterProtocolUrlHandlersMainDepsHandler();
registerProtocolUrlHandlers(registerProtocolUrlHandlersMainDeps);

const buildOnWillQuitCleanupDepsHandler = createBuildOnWillQuitCleanupDepsHandler({
    destroyTray: () => destroyTray(),
    stopConfigHotReload: () => configHotReloadRuntime.stop(),
    restorePreviousSecondarySubVisibility: () => restorePreviousSecondarySubVisibility(),
    unregisterAllGlobalShortcuts: () => globalShortcut.unregisterAll(),
    stopSubtitleWebsocket: () => subtitleWsService.stop(),
    stopTexthookerService: () => texthookerService.stop(),
    getYomitanParserWindow: () => appState.yomitanParserWindow,
    clearYomitanParserState: () => {
      appState.yomitanParserWindow = null;
      appState.yomitanParserReadyPromise = null;
      appState.yomitanParserInitPromise = null;
    },
    getWindowTracker: () => appState.windowTracker,
    getMpvSocket: () => appState.mpvClient?.socket ?? null,
    getReconnectTimer: () => appState.reconnectTimer,
    clearReconnectTimerRef: () => {
      appState.reconnectTimer = null;
    },
    getSubtitleTimingTracker: () => appState.subtitleTimingTracker,
    getImmersionTracker: () => appState.immersionTracker,
    clearImmersionTracker: () => {
      appState.immersionTracker = null;
    },
    getAnkiIntegration: () => appState.ankiIntegration,
    getAnilistSetupWindow: () => appState.anilistSetupWindow,
    clearAnilistSetupWindow: () => {
      appState.anilistSetupWindow = null;
    },
    getJellyfinSetupWindow: () => appState.jellyfinSetupWindow,
    clearJellyfinSetupWindow: () => {
      appState.jellyfinSetupWindow = null;
    },
    stopJellyfinRemoteSession: () => stopJellyfinRemoteSession(),
  });
const onWillQuitCleanupHandler = createOnWillQuitCleanupHandler(buildOnWillQuitCleanupDepsHandler());

const buildShouldRestoreWindowsOnActivateMainDepsHandler =
  createBuildShouldRestoreWindowsOnActivateMainDepsHandler({
    isOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
    getAllWindowCount: () => BrowserWindow.getAllWindows().length,
  });
const shouldRestoreWindowsOnActivateMainDeps =
  buildShouldRestoreWindowsOnActivateMainDepsHandler();
const shouldRestoreWindowsOnActivateHandler = createShouldRestoreWindowsOnActivateHandler(
  shouldRestoreWindowsOnActivateMainDeps,
);

const buildRestoreWindowsOnActivateMainDepsHandler =
  createBuildRestoreWindowsOnActivateMainDepsHandler({
    createMainWindow: () => {
      createMainWindow();
    },
    createInvisibleWindow: () => {
      createInvisibleWindow();
    },
    updateVisibleOverlayVisibility: () => {
      overlayVisibilityRuntime.updateVisibleOverlayVisibility();
    },
    updateInvisibleOverlayVisibility: () => {
      overlayVisibilityRuntime.updateInvisibleOverlayVisibility();
    },
  });
const restoreWindowsOnActivateMainDeps = buildRestoreWindowsOnActivateMainDepsHandler();
const restoreWindowsOnActivateHandler = createRestoreWindowsOnActivateHandler(
  restoreWindowsOnActivateMainDeps,
);

const buildReloadConfigMainDepsHandler = createBuildReloadConfigMainDepsHandler({
    reloadConfigStrict: () => configService.reloadConfigStrict(),
    logInfo: (message) => appLogger.logInfo(message),
    logWarning: (message) => appLogger.logWarning(message),
    showDesktopNotification: (title, options) => showDesktopNotification(title, options),
    startConfigHotReload: () => configHotReloadRuntime.start(),
    refreshAnilistClientSecretState: (options) => refreshAnilistClientSecretState(options),
    failHandlers: {
      logError: (details) => logger.error(details),
      showErrorBox: (title, details) => dialog.showErrorBox(title, details),
      quit: () => app.quit(),
    },
  });
const reloadConfigHandler = createReloadConfigHandler(buildReloadConfigMainDepsHandler());

const buildCriticalConfigErrorMainDepsHandler = createBuildCriticalConfigErrorMainDepsHandler({
    getConfigPath: () => configService.getConfigPath(),
    failHandlers: {
      logError: (message) => logger.error(message),
      showErrorBox: (title, message) => dialog.showErrorBox(title, message),
      quit: () => app.quit(),
    },
  });
const criticalConfigErrorMainDeps = buildCriticalConfigErrorMainDepsHandler();
const criticalConfigErrorHandler = createCriticalConfigErrorHandler(
  criticalConfigErrorMainDeps,
);

const buildAppReadyRuntimeMainDepsHandler = createBuildAppReadyRuntimeMainDepsHandler({
    loadSubtitlePosition: () => loadSubtitlePosition(),
    resolveKeybindings: () => {
      appState.keybindings = resolveKeybindings(getResolvedConfig(), DEFAULT_KEYBINDINGS);
    },
    createMpvClient: () => {
      appState.mpvClient = createMpvClientRuntimeService();
    },
    reloadConfig: reloadConfigHandler,
    getResolvedConfig: () => getResolvedConfig(),
    getConfigWarnings: () => configService.getWarnings(),
    logConfigWarning: (warning) => appLogger.logConfigWarning(warning),
    setLogLevel: (level: string, source: LogLevelSource) => setLogLevel(level, source),
    initRuntimeOptionsManager: () => {
      appState.runtimeOptionsManager = new RuntimeOptionsManager(
        () => configService.getConfig().ankiConnect,
        {
          applyAnkiPatch: (patch) => {
            if (appState.ankiIntegration) {
              appState.ankiIntegration.applyRuntimeConfigPatch(patch);
            }
          },
          onOptionsChanged: () => {
            broadcastRuntimeOptionsChanged();
            refreshOverlayShortcuts();
          },
        },
      );
    },
    setSecondarySubMode: (mode: SecondarySubMode) => {
      appState.secondarySubMode = mode;
    },
    defaultSecondarySubMode: 'hover',
    defaultWebsocketPort: DEFAULT_CONFIG.websocket.port,
    hasMpvWebsocketPlugin: () => hasMpvWebsocketPlugin(),
    startSubtitleWebsocket: (port: number) => {
      subtitleWsService.start(port, () => appState.currentSubText);
    },
    log: (message) => appLogger.logInfo(message),
    createMecabTokenizerAndCheck: async () => {
      await createMecabTokenizerAndCheck();
    },
    createSubtitleTimingTracker: () => {
      const tracker = new SubtitleTimingTracker();
      appState.subtitleTimingTracker = tracker;
    },
    createImmersionTracker: createImmersionTrackerStartupHandler(
      createBuildImmersionTrackerStartupMainDepsHandler({
        getResolvedConfig: () => getResolvedConfig(),
        getConfiguredDbPath: () => immersionMediaRuntime.getConfiguredDbPath(),
        createTrackerService: (params) => new ImmersionTrackerService(params),
        setTracker: (tracker) => {
          appState.immersionTracker = tracker as ImmersionTrackerService | null;
        },
        getMpvClient: () => appState.mpvClient,
        seedTrackerFromCurrentMedia: () => {
          void immersionMediaRuntime.seedFromCurrentMedia();
        },
        logInfo: (message) => logger.info(message),
        logDebug: (message) => logger.debug(message),
        logWarn: (message, details) => logger.warn(message, details),
      })(),
    ),
    loadYomitanExtension: async () => {
      await loadYomitanExtension();
    },
    startJellyfinRemoteSession: async () => {
      await startJellyfinRemoteSession();
    },
    prewarmSubtitleDictionaries: async () => {
      await prewarmSubtitleDictionaries();
    },
    startBackgroundWarmups: () => {
      startBackgroundWarmups();
    },
    texthookerOnlyMode: appState.texthookerOnlyMode,
    shouldAutoInitializeOverlayRuntimeFromConfig: () =>
      appState.backgroundMode
        ? false
        : configDerivedRuntime.shouldAutoInitializeOverlayRuntimeFromConfig(),
    initializeOverlayRuntime: () => initializeOverlayRuntime(),
    handleInitialArgs: () => handleInitialArgs(),
    onCriticalConfigErrors: criticalConfigErrorHandler,
    logDebug: (message: string) => {
      logger.debug(message);
    },
    now: () => Date.now(),
  });
const appReadyRuntimeRunner = createAppReadyRuntimeRunner(buildAppReadyRuntimeMainDepsHandler());

const { appLifecycleRuntimeRunner, runAndApplyStartupState } = createStartupRuntimeHandlers<
  CliArgs,
  StartupState,
  ReturnType<typeof createStartupBootstrapRuntimeDeps>
>({
  appLifecycleRuntimeRunnerMainDeps: {
    app,
    platform: process.platform,
    shouldStartApp: (nextArgs: CliArgs) => shouldStartApp(nextArgs),
    parseArgs: (argv: string[]) => parseArgs(argv),
    handleCliCommand: (nextArgs: CliArgs, source: CliCommandSource) =>
      handleCliCommand(nextArgs, source),
    printHelp: () => printHelp(DEFAULT_TEXTHOOKER_PORT),
    logNoRunningInstance: () => appLogger.logNoRunningInstance(),
    onReady: appReadyRuntimeRunner,
    onWillQuitCleanup: () => onWillQuitCleanupHandler(),
    shouldRestoreWindowsOnActivate: () => shouldRestoreWindowsOnActivateHandler(),
    restoreWindowsOnActivate: () => restoreWindowsOnActivateHandler(),
    shouldQuitOnWindowAllClosed: () => !appState.backgroundMode,
  },
  createAppLifecycleRuntimeRunner: (params) => createAppLifecycleRuntimeRunner(params),
  buildStartupBootstrapMainDeps: (startAppLifecycle) => ({
    argv: process.argv,
    parseArgs: (argv: string[]) => parseArgs(argv),
    setLogLevel: (level: string, source: LogLevelSource) => {
      setLogLevel(level, source);
    },
    forceX11Backend: (args: CliArgs) => {
      forceX11Backend(args);
    },
    enforceUnsupportedWaylandMode: (args: CliArgs) => {
      enforceUnsupportedWaylandMode(args);
    },
    shouldStartApp: (args: CliArgs) => shouldStartApp(args),
    getDefaultSocketPath: () => getDefaultSocketPath(),
    defaultTexthookerPort: DEFAULT_TEXTHOOKER_PORT,
    configDir: CONFIG_DIR,
    defaultConfig: DEFAULT_CONFIG,
    generateConfigTemplate: (config: ResolvedConfig) => generateConfigTemplate(config),
    generateDefaultConfigFile: (
      args: CliArgs,
      options: {
        configDir: string;
        defaultConfig: unknown;
        generateTemplate: (config: unknown) => string;
      },
    ) => generateDefaultConfigFile(args, options),
    setExitCode: (code) => {
      process.exitCode = code;
    },
    quitApp: () => app.quit(),
    logGenerateConfigError: (message) => logger.error(message),
    startAppLifecycle,
  }),
  createStartupBootstrapRuntimeDeps: (deps) => createStartupBootstrapRuntimeDeps(deps),
  runStartupBootstrapRuntime,
  applyStartupState: (startupState) => applyStartupState(appState, startupState),
});

runAndApplyStartupState();
void refreshAnilistClientSecretState({ force: true });
anilistStateRuntime.refreshRetryQueueState();

const handleTexthookerOnlyModeTransitionHandler = createHandleTexthookerOnlyModeTransitionHandler(
  createBuildHandleTexthookerOnlyModeTransitionMainDepsHandler({
    isTexthookerOnlyMode: () => appState.texthookerOnlyMode,
    setTexthookerOnlyMode: (enabled) => {
      appState.texthookerOnlyMode = enabled;
    },
    commandNeedsOverlayRuntime: (inputArgs) => commandNeedsOverlayRuntime(inputArgs),
    startBackgroundWarmups: () => startBackgroundWarmups(),
    logInfo: (message: string) => logger.info(message),
  })(),
);

function handleCliCommand(args: CliArgs, source: CliCommandSource = 'initial'): void {
  handleTexthookerOnlyModeTransitionHandler(args);

  const cliContext = createCliCommandContextHandler();
  handleCliCommandRuntimeServiceWithContext(args, source, cliContext);
}

const handleInitialArgsRuntimeHandler = createInitialArgsRuntimeHandler({
  getInitialArgs: () => appState.initialArgs,
  isBackgroundMode: () => appState.backgroundMode,
  ensureTray: () => ensureTray(),
  isTexthookerOnlyMode: () => appState.texthookerOnlyMode,
  hasImmersionTracker: () => Boolean(appState.immersionTracker),
  getMpvClient: () => appState.mpvClient,
  logInfo: (message) => logger.info(message),
  handleCliCommand: (args, source) => handleCliCommand(args, source),
});

function handleInitialArgs(): void {
  handleInitialArgsRuntimeHandler();
}

const buildBindMpvMainEventHandlersMainDepsHandler =
  createBuildBindMpvMainEventHandlersMainDepsHandler({
    appState,
    getQuitOnDisconnectArmed: () => jellyfinPlayQuitOnDisconnectArmed,
    scheduleQuitCheck: (callback) => {
      setTimeout(callback, 500);
    },
    quitApp: () => app.quit(),
    reportJellyfinRemoteStopped: () => {
      void reportJellyfinRemoteStopped();
    },
    maybeRunAnilistPostWatchUpdate: () => maybeRunAnilistPostWatchUpdate(),
    logSubtitleTimingError: (message, error) => logger.error(message, error),
    broadcastToOverlayWindows: (channel, payload) => {
      broadcastToOverlayWindows(channel, payload);
    },
    onSubtitleChange: (text) => {
      subtitleProcessingController.onSubtitleChange(text);
    },
    updateCurrentMediaPath: (path) => {
      mediaRuntime.updateCurrentMediaPath(path);
    },
    getCurrentAnilistMediaKey: () => getCurrentAnilistMediaKey(),
    resetAnilistMediaTracking: (mediaKey) => {
      resetAnilistMediaTracking(mediaKey);
    },
    maybeProbeAnilistDuration: (mediaKey) => {
      void maybeProbeAnilistDuration(mediaKey);
    },
    ensureAnilistMediaGuess: (mediaKey) => {
      void ensureAnilistMediaGuess(mediaKey);
    },
    syncImmersionMediaState: () => {
      immersionMediaRuntime.syncFromCurrentMediaState();
    },
    updateCurrentMediaTitle: (title) => {
      mediaRuntime.updateCurrentMediaTitle(title);
    },
    resetAnilistMediaGuessState: () => {
      resetAnilistMediaGuessState();
    },
    reportJellyfinRemoteProgress: (forceImmediate) => {
      void reportJellyfinRemoteProgress(forceImmediate);
    },
    updateSubtitleRenderMetrics: (patch) => {
      updateMpvSubtitleRenderMetrics(patch as Partial<MpvSubtitleRenderMetrics>);
    },
  });
const bindMpvMainEventHandlersMainDeps = buildBindMpvMainEventHandlersMainDepsHandler();
const bindMpvClientEventHandlers = createBindMpvMainEventHandlersHandler(
  bindMpvMainEventHandlersMainDeps,
);

const buildMpvClientRuntimeServiceFactoryMainDepsHandler =
  createBuildMpvClientRuntimeServiceFactoryDepsHandler({
    createClient: MpvIpcClient,
    getSocketPath: () => appState.mpvSocketPath,
    getResolvedConfig: () => getResolvedConfig(),
    isAutoStartOverlayEnabled: () => appState.autoStartOverlay,
    setOverlayVisible: (visible: boolean) => setOverlayVisible(visible),
    shouldBindVisibleOverlayToMpvSubVisibility: () =>
      configDerivedRuntime.shouldBindVisibleOverlayToMpvSubVisibility(),
    isVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    getReconnectTimer: () => appState.reconnectTimer,
    setReconnectTimer: (timer: ReturnType<typeof setTimeout> | null) => {
      appState.reconnectTimer = timer;
    },
    bindEventHandlers: (client) => bindMpvClientEventHandlers(client),
  });

function createMpvClientRuntimeService(): MpvIpcClient {
  return createMpvClientRuntimeServiceFactory(buildMpvClientRuntimeServiceFactoryMainDepsHandler())();
}

const buildUpdateMpvSubtitleRenderMetricsMainDepsHandler =
  createBuildUpdateMpvSubtitleRenderMetricsMainDepsHandler({
  getCurrentMetrics: () => appState.mpvSubtitleRenderMetrics,
  setCurrentMetrics: (metrics) => {
    appState.mpvSubtitleRenderMetrics = metrics;
  },
  applyPatch: (current, patch) => applyMpvSubtitleRenderMetricsPatch(current, patch),
  broadcastMetrics: (metrics) => {
    broadcastToOverlayWindows('mpv-subtitle-render-metrics:set', metrics);
  },
});
const updateMpvSubtitleRenderMetricsMainDeps =
  buildUpdateMpvSubtitleRenderMetricsMainDepsHandler();
const updateMpvSubtitleRenderMetricsRuntime = createUpdateMpvSubtitleRenderMetricsHandler(
  updateMpvSubtitleRenderMetricsMainDeps,
);

function updateMpvSubtitleRenderMetrics(patch: Partial<MpvSubtitleRenderMetrics>): void {
  updateMpvSubtitleRenderMetricsRuntime(patch);
}

const buildTokenizerDepsHandler = createBuildTokenizerDepsMainHandler({
  getYomitanExt: () => appState.yomitanExt,
  getYomitanParserWindow: () => appState.yomitanParserWindow,
  setYomitanParserWindow: (window) => {
    appState.yomitanParserWindow = window as BrowserWindow | null;
  },
  getYomitanParserReadyPromise: () => appState.yomitanParserReadyPromise,
  setYomitanParserReadyPromise: (promise) => {
    appState.yomitanParserReadyPromise = promise;
  },
  getYomitanParserInitPromise: () => appState.yomitanParserInitPromise,
  setYomitanParserInitPromise: (promise) => {
    appState.yomitanParserInitPromise = promise;
  },
  isKnownWord: (text) => Boolean(appState.ankiIntegration?.isKnownWord(text)),
  recordLookup: (hit) => {
    appState.immersionTracker?.recordLookup(hit);
  },
  getKnownWordMatchMode: () =>
    appState.ankiIntegration?.getKnownWordMatchMode() ??
    getResolvedConfig().ankiConnect.nPlusOne.matchMode,
  getMinSentenceWordsForNPlusOne: () => getResolvedConfig().ankiConnect.nPlusOne.minSentenceWords,
  getJlptLevel: (text) => appState.jlptLevelLookup(text),
  getJlptEnabled: () => getResolvedConfig().subtitleStyle.enableJlpt,
  getFrequencyDictionaryEnabled: () => getResolvedConfig().subtitleStyle.frequencyDictionary.enabled,
  getFrequencyRank: (text) => appState.frequencyRankLookup(text),
  getYomitanGroupDebugEnabled: () => appState.overlayDebugVisualizationEnabled,
  getMecabTokenizer: () => appState.mecabTokenizer,
});

const buildCreateMecabTokenizerAndCheckMainDepsHandler = createCreateMecabTokenizerAndCheckMainHandler(
  {
  getMecabTokenizer: () => appState.mecabTokenizer,
  setMecabTokenizer: (tokenizer) => {
    appState.mecabTokenizer = tokenizer;
  },
  createMecabTokenizer: () => new MecabTokenizer(),
  checkAvailability: async (tokenizer) => tokenizer.checkAvailability(),
},
);
const createMecabTokenizerAndCheckHandler = buildCreateMecabTokenizerAndCheckMainDepsHandler;

const buildPrewarmSubtitleDictionariesMainDepsHandler = createPrewarmSubtitleDictionariesMainHandler(
  {
  ensureJlptDictionaryLookup: () => jlptDictionaryRuntime.ensureJlptDictionaryLookup(),
  ensureFrequencyDictionaryLookup: () => frequencyDictionaryRuntime.ensureFrequencyDictionaryLookup(),
},
);
const prewarmSubtitleDictionariesHandler = buildPrewarmSubtitleDictionariesMainDepsHandler;

async function tokenizeSubtitle(text: string): Promise<SubtitleData> {
  await jlptDictionaryRuntime.ensureJlptDictionaryLookup();
  await frequencyDictionaryRuntime.ensureFrequencyDictionaryLookup();
  return tokenizeSubtitleCore(text, createTokenizerDepsRuntime(buildTokenizerDepsHandler()));
}

async function createMecabTokenizerAndCheck(): Promise<void> {
  await createMecabTokenizerAndCheckHandler();
}

async function prewarmSubtitleDictionaries(): Promise<void> {
  await prewarmSubtitleDictionariesHandler();
}

const buildLaunchBackgroundWarmupTaskMainDepsHandler =
  createBuildLaunchBackgroundWarmupTaskMainDepsHandler({
  now: () => Date.now(),
  logDebug: (message) => logger.debug(message),
  logWarn: (message) => logger.warn(message),
});
const launchBackgroundWarmupTaskMainDeps =
  buildLaunchBackgroundWarmupTaskMainDepsHandler();
const launchBackgroundWarmupTask = createLaunchBackgroundWarmupTaskHandler(
  launchBackgroundWarmupTaskMainDeps,
);

const buildStartBackgroundWarmupsMainDepsHandler = createBuildStartBackgroundWarmupsMainDepsHandler(
  {
  getStarted: () => backgroundWarmupsStarted,
  setStarted: (started) => {
    backgroundWarmupsStarted = started;
  },
  isTexthookerOnlyMode: () => appState.texthookerOnlyMode,
  launchTask: (label, task) => launchBackgroundWarmupTask(label, task),
  createMecabTokenizerAndCheck: () => createMecabTokenizerAndCheck(),
  ensureYomitanExtensionLoaded: () => ensureYomitanExtensionLoaded().then(() => {}),
  prewarmSubtitleDictionaries: () => prewarmSubtitleDictionaries(),
  shouldAutoConnectJellyfinRemote: () => getResolvedConfig().jellyfin.remoteControlAutoConnect,
  startJellyfinRemoteSession: () => startJellyfinRemoteSession(),
},
);
const startBackgroundWarmupsMainDeps =
  buildStartBackgroundWarmupsMainDepsHandler();
const startBackgroundWarmups = createStartBackgroundWarmupsHandler(
  startBackgroundWarmupsMainDeps,
);

const buildUpdateVisibleOverlayBoundsMainDepsHandler =
  createBuildUpdateVisibleOverlayBoundsMainDepsHandler({
  setOverlayWindowBounds: (layer, geometry) => overlayManager.setOverlayWindowBounds(layer, geometry),
});
const updateVisibleOverlayBoundsMainDeps =
  buildUpdateVisibleOverlayBoundsMainDepsHandler();
const updateVisibleOverlayBounds = createUpdateVisibleOverlayBoundsHandler(
  updateVisibleOverlayBoundsMainDeps,
);

const buildUpdateInvisibleOverlayBoundsMainDepsHandler =
  createBuildUpdateInvisibleOverlayBoundsMainDepsHandler({
  setOverlayWindowBounds: (layer, geometry) => overlayManager.setOverlayWindowBounds(layer, geometry),
});
const updateInvisibleOverlayBoundsMainDeps =
  buildUpdateInvisibleOverlayBoundsMainDepsHandler();
const updateInvisibleOverlayBounds = createUpdateInvisibleOverlayBoundsHandler(
  updateInvisibleOverlayBoundsMainDeps,
);

const buildEnsureOverlayWindowLevelMainDepsHandler =
  createBuildEnsureOverlayWindowLevelMainDepsHandler({
  ensureOverlayWindowLevelCore: (window) => ensureOverlayWindowLevelCore(window as BrowserWindow),
});
const ensureOverlayWindowLevelMainDeps =
  buildEnsureOverlayWindowLevelMainDepsHandler();
const ensureOverlayWindowLevel = createEnsureOverlayWindowLevelHandler(
  ensureOverlayWindowLevelMainDeps,
);

const buildEnforceOverlayLayerOrderMainDepsHandler =
  createBuildEnforceOverlayLayerOrderMainDepsHandler({
  enforceOverlayLayerOrderCore: (params) =>
    enforceOverlayLayerOrderCore({
      visibleOverlayVisible: params.visibleOverlayVisible,
      invisibleOverlayVisible: params.invisibleOverlayVisible,
      mainWindow: params.mainWindow as BrowserWindow | null,
      invisibleWindow: params.invisibleWindow as BrowserWindow | null,
      ensureOverlayWindowLevel: (window) => params.ensureOverlayWindowLevel(window as BrowserWindow),
    }),
  getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
  getInvisibleOverlayVisible: () => overlayManager.getInvisibleOverlayVisible(),
  getMainWindow: () => overlayManager.getMainWindow(),
  getInvisibleWindow: () => overlayManager.getInvisibleWindow(),
  ensureOverlayWindowLevel: (window) => ensureOverlayWindowLevel(window as BrowserWindow),
});
const enforceOverlayLayerOrderMainDeps = buildEnforceOverlayLayerOrderMainDepsHandler();
const enforceOverlayLayerOrder = createEnforceOverlayLayerOrderHandler(
  enforceOverlayLayerOrderMainDeps,
);

async function loadYomitanExtension(): Promise<Extension | null> {
  return yomitanExtensionRuntime.loadYomitanExtension();
}

async function ensureYomitanExtensionLoaded(): Promise<Extension | null> {
  return yomitanExtensionRuntime.ensureYomitanExtensionLoaded();
}

function createOverlayWindow(kind: 'visible' | 'invisible'): BrowserWindow {
  return createOverlayWindowHandler(kind);
}

function createMainWindow(): BrowserWindow {
  return createMainWindowHandler();
}
function createInvisibleWindow(): BrowserWindow {
  return createInvisibleWindowHandler();
}

function resolveTrayIconPath(): string | null {
  return resolveTrayIconPathHandler();
}

function buildTrayMenu(): Menu {
  return buildTrayMenuHandler();
}

function ensureTray(): void {
  ensureTrayHandler();
}

function destroyTray(): void {
  destroyTrayHandler();
}

function initializeOverlayRuntime(): void {
  initializeOverlayRuntimeHandler();
}

function openYomitanSettings(): void {
  openYomitanSettingsHandler();
}

const buildGetConfiguredShortcutsMainDepsHandler = createBuildGetConfiguredShortcutsMainDepsHandler(
  {
    getResolvedConfig: () => getResolvedConfig(),
    defaultConfig: DEFAULT_CONFIG,
    resolveConfiguredShortcuts,
  },
);
const getConfiguredShortcutsMainDeps = buildGetConfiguredShortcutsMainDepsHandler();
const getConfiguredShortcutsHandler = createGetConfiguredShortcutsHandler(
  getConfiguredShortcutsMainDeps,
);

const buildRegisterGlobalShortcutsMainDepsHandler =
  createBuildRegisterGlobalShortcutsMainDepsHandler({
    getConfiguredShortcuts: () => getConfiguredShortcuts(),
    registerGlobalShortcutsCore,
    toggleVisibleOverlay: () => toggleVisibleOverlay(),
    toggleInvisibleOverlay: () => toggleInvisibleOverlay(),
    openYomitanSettings: () => openYomitanSettings(),
    isDev,
    getMainWindow: () => overlayManager.getMainWindow(),
  });
const registerGlobalShortcutsMainDeps = buildRegisterGlobalShortcutsMainDepsHandler();
const registerGlobalShortcutsHandler = createRegisterGlobalShortcutsHandler(
  registerGlobalShortcutsMainDeps,
);

const buildRefreshGlobalAndOverlayShortcutsMainDepsHandler =
  createBuildRefreshGlobalAndOverlayShortcutsMainDepsHandler({
    unregisterAllGlobalShortcuts: () => globalShortcut.unregisterAll(),
    registerGlobalShortcuts: () => registerGlobalShortcuts(),
    syncOverlayShortcuts: () => syncOverlayShortcuts(),
  });
const refreshGlobalAndOverlayShortcutsMainDeps =
  buildRefreshGlobalAndOverlayShortcutsMainDepsHandler();
const refreshGlobalAndOverlayShortcutsHandler = createRefreshGlobalAndOverlayShortcutsHandler(
  refreshGlobalAndOverlayShortcutsMainDeps,
);

function registerGlobalShortcuts(): void {
  registerGlobalShortcutsHandler();
}

function refreshGlobalAndOverlayShortcuts(): void {
  refreshGlobalAndOverlayShortcutsHandler();
}

function getConfiguredShortcuts() {
  return getConfiguredShortcutsHandler();
}

const buildCycleSecondarySubModeMainDepsHandler = createBuildCycleSecondarySubModeMainDepsHandler(
  {
    getSecondarySubMode: () => appState.secondarySubMode,
    setSecondarySubMode: (mode: SecondarySubMode) => {
      appState.secondarySubMode = mode;
    },
    getLastSecondarySubToggleAtMs: () => appState.lastSecondarySubToggleAtMs,
    setLastSecondarySubToggleAtMs: (timestampMs: number) => {
      appState.lastSecondarySubToggleAtMs = timestampMs;
    },
    broadcastToOverlayWindows: (channel, mode) => {
      broadcastToOverlayWindows(channel, mode);
    },
    showMpvOsd: (text: string) => showMpvOsd(text),
  },
);

function cycleSecondarySubMode(): void {
  cycleSecondarySubModeCore(buildCycleSecondarySubModeMainDepsHandler());
}

const buildAppendToMpvLogMainDepsHandler = createBuildAppendToMpvLogMainDepsHandler({
  logPath: DEFAULT_MPV_LOG_PATH,
  dirname: (targetPath) => path.dirname(targetPath),
  mkdirSync: (targetPath, options) => fs.mkdirSync(targetPath, options),
  appendFileSync: (targetPath, data, options) => fs.appendFileSync(targetPath, data, options),
  now: () => new Date(),
});
const appendToMpvLogMainDeps = buildAppendToMpvLogMainDepsHandler();
const appendToMpvLogHandler = createAppendToMpvLogHandler(appendToMpvLogMainDeps);

const buildShowMpvOsdMainDepsHandler = createBuildShowMpvOsdMainDepsHandler({
  appendToMpvLog: (message) => appendToMpvLog(message),
  showMpvOsdRuntime: (mpvClient, text, fallbackLog) =>
    showMpvOsdRuntime(mpvClient as never, text, fallbackLog),
  getMpvClient: () => appState.mpvClient,
  logInfo: (line) => logger.info(line),
});
const showMpvOsdMainDeps = buildShowMpvOsdMainDepsHandler();
const showMpvOsdHandler = createShowMpvOsdHandler(showMpvOsdMainDeps);

function showMpvOsd(text: string): void {
  showMpvOsdHandler(text);
}

function appendToMpvLog(message: string): void {
  appendToMpvLogHandler(message);
}

const buildNumericShortcutRuntimeMainDepsHandler = createBuildNumericShortcutRuntimeMainDepsHandler({
  globalShortcut,
  showMpvOsd: (text) => showMpvOsd(text),
  setTimer: (handler, timeoutMs) => setTimeout(handler, timeoutMs),
  clearTimer: (timer) => clearTimeout(timer),
});
const numericShortcutRuntimeMainDeps = buildNumericShortcutRuntimeMainDepsHandler();
const numericShortcutRuntime = createNumericShortcutRuntime(
  numericShortcutRuntimeMainDeps,
);
const multiCopySession = numericShortcutRuntime.createSession();
const mineSentenceSession = numericShortcutRuntime.createSession();
const buildCancelPendingMultiCopyMainDepsHandler =
  createBuildCancelNumericShortcutSessionMainDepsHandler({
  session: multiCopySession,
});
const cancelPendingMultiCopyMainDeps =
  buildCancelPendingMultiCopyMainDepsHandler();
const cancelPendingMultiCopyHandler = createCancelNumericShortcutSessionHandler(
  cancelPendingMultiCopyMainDeps,
);

const buildStartPendingMultiCopyMainDepsHandler =
  createBuildStartNumericShortcutSessionMainDepsHandler({
  session: multiCopySession,
  onDigit: (count) => handleMultiCopyDigit(count),
  messages: {
    prompt: 'Copy how many lines? Press 1-9 (Esc to cancel)',
    timeout: 'Copy timeout',
    cancelled: 'Cancelled',
  },
});
const startPendingMultiCopyMainDeps =
  buildStartPendingMultiCopyMainDepsHandler();
const startPendingMultiCopyHandler = createStartNumericShortcutSessionHandler(
  startPendingMultiCopyMainDeps,
);

const buildCancelPendingMineSentenceMultipleMainDepsHandler =
  createBuildCancelNumericShortcutSessionMainDepsHandler({
  session: mineSentenceSession,
});
const cancelPendingMineSentenceMultipleMainDeps =
  buildCancelPendingMineSentenceMultipleMainDepsHandler();
const cancelPendingMineSentenceMultipleHandler = createCancelNumericShortcutSessionHandler(
  cancelPendingMineSentenceMultipleMainDeps,
);

const buildStartPendingMineSentenceMultipleMainDepsHandler =
  createBuildStartNumericShortcutSessionMainDepsHandler({
  session: mineSentenceSession,
  onDigit: (count) => handleMineSentenceDigit(count),
  messages: {
    prompt: 'Mine how many lines? Press 1-9 (Esc to cancel)',
    timeout: 'Mine sentence timeout',
    cancelled: 'Cancelled',
  },
});
const startPendingMineSentenceMultipleMainDeps =
  buildStartPendingMineSentenceMultipleMainDepsHandler();
const startPendingMineSentenceMultipleHandler = createStartNumericShortcutSessionHandler(
  startPendingMineSentenceMultipleMainDeps,
);

const buildRegisterOverlayShortcutsMainDepsHandler =
  createBuildRegisterOverlayShortcutsMainDepsHandler({
  overlayShortcutsRuntime,
});
const registerOverlayShortcutsMainDeps =
  buildRegisterOverlayShortcutsMainDepsHandler();
const registerOverlayShortcutsHandler = createRegisterOverlayShortcutsHandler(
  registerOverlayShortcutsMainDeps,
);

const buildUnregisterOverlayShortcutsMainDepsHandler =
  createBuildUnregisterOverlayShortcutsMainDepsHandler({
  overlayShortcutsRuntime,
});
const unregisterOverlayShortcutsMainDeps =
  buildUnregisterOverlayShortcutsMainDepsHandler();
const unregisterOverlayShortcutsHandler = createUnregisterOverlayShortcutsHandler(
  unregisterOverlayShortcutsMainDeps,
);

const buildSyncOverlayShortcutsMainDepsHandler = createBuildSyncOverlayShortcutsMainDepsHandler({
  overlayShortcutsRuntime,
});
const syncOverlayShortcutsMainDeps = buildSyncOverlayShortcutsMainDepsHandler();
const syncOverlayShortcutsHandler = createSyncOverlayShortcutsHandler(
  syncOverlayShortcutsMainDeps,
);

const buildRefreshOverlayShortcutsMainDepsHandler =
  createBuildRefreshOverlayShortcutsMainDepsHandler({
  overlayShortcutsRuntime,
});
const refreshOverlayShortcutsMainDeps =
  buildRefreshOverlayShortcutsMainDepsHandler();
const refreshOverlayShortcutsHandler = createRefreshOverlayShortcutsHandler(
  refreshOverlayShortcutsMainDeps,
);

async function triggerSubsyncFromConfig(): Promise<void> {
  await subsyncRuntime.triggerFromConfig();
}

function cancelPendingMultiCopy(): void {
  cancelPendingMultiCopyHandler();
}

function startPendingMultiCopy(timeoutMs: number): void {
  startPendingMultiCopyHandler(timeoutMs);
}

function handleMultiCopyDigit(count: number): void {
  handleMultiCopyDigitHandler(count);
}

function copyCurrentSubtitle(): void {
  copyCurrentSubtitleHandler();
}

const buildUpdateLastCardFromClipboardMainDepsHandler =
  createBuildUpdateLastCardFromClipboardMainDepsHandler({
  getAnkiIntegration: () => appState.ankiIntegration,
  readClipboardText: () => clipboard.readText(),
  showMpvOsd: (text) => showMpvOsd(text),
  updateLastCardFromClipboardCore,
});
const updateLastCardFromClipboardMainDeps =
  buildUpdateLastCardFromClipboardMainDepsHandler();
const updateLastCardFromClipboardHandler = createUpdateLastCardFromClipboardHandler(
  updateLastCardFromClipboardMainDeps,
);

const buildRefreshKnownWordCacheMainDepsHandler = createBuildRefreshKnownWordCacheMainDepsHandler({
  getAnkiIntegration: () => appState.ankiIntegration,
  missingIntegrationMessage: 'AnkiConnect integration not enabled',
});
const refreshKnownWordCacheMainDeps = buildRefreshKnownWordCacheMainDepsHandler();
const refreshKnownWordCacheHandler = createRefreshKnownWordCacheHandler(
  refreshKnownWordCacheMainDeps,
);

const buildTriggerFieldGroupingMainDepsHandler = createBuildTriggerFieldGroupingMainDepsHandler({
  getAnkiIntegration: () => appState.ankiIntegration,
  showMpvOsd: (text) => showMpvOsd(text),
  triggerFieldGroupingCore,
});
const triggerFieldGroupingMainDeps = buildTriggerFieldGroupingMainDepsHandler();
const triggerFieldGroupingHandler = createTriggerFieldGroupingHandler(
  triggerFieldGroupingMainDeps,
);

const buildMarkLastCardAsAudioCardMainDepsHandler =
  createBuildMarkLastCardAsAudioCardMainDepsHandler({
  getAnkiIntegration: () => appState.ankiIntegration,
  showMpvOsd: (text) => showMpvOsd(text),
  markLastCardAsAudioCardCore,
});
const markLastCardAsAudioCardMainDeps =
  buildMarkLastCardAsAudioCardMainDepsHandler();
const markLastCardAsAudioCardHandler = createMarkLastCardAsAudioCardHandler(
  markLastCardAsAudioCardMainDeps,
);

const buildMineSentenceCardMainDepsHandler = createBuildMineSentenceCardMainDepsHandler({
  getAnkiIntegration: () => appState.ankiIntegration,
  getMpvClient: () => appState.mpvClient,
  showMpvOsd: (text) => showMpvOsd(text),
  mineSentenceCardCore,
  recordCardsMined: (count) => {
    appState.immersionTracker?.recordCardsMined(count);
  },
});
const mineSentenceCardHandler = createMineSentenceCardHandler(buildMineSentenceCardMainDepsHandler());

const buildHandleMultiCopyDigitMainDepsHandler = createBuildHandleMultiCopyDigitMainDepsHandler({
  getSubtitleTimingTracker: () => appState.subtitleTimingTracker,
  writeClipboardText: (text) => clipboard.writeText(text),
  showMpvOsd: (text) => showMpvOsd(text),
  handleMultiCopyDigitCore,
});
const handleMultiCopyDigitMainDeps = buildHandleMultiCopyDigitMainDepsHandler();
const handleMultiCopyDigitHandler = createHandleMultiCopyDigitHandler(
  handleMultiCopyDigitMainDeps,
);

const buildCopyCurrentSubtitleMainDepsHandler = createBuildCopyCurrentSubtitleMainDepsHandler({
  getSubtitleTimingTracker: () => appState.subtitleTimingTracker,
  writeClipboardText: (text) => clipboard.writeText(text),
  showMpvOsd: (text) => showMpvOsd(text),
  copyCurrentSubtitleCore,
});
const copyCurrentSubtitleMainDeps = buildCopyCurrentSubtitleMainDepsHandler();
const copyCurrentSubtitleHandler = createCopyCurrentSubtitleHandler(
  copyCurrentSubtitleMainDeps,
);

const buildHandleMineSentenceDigitMainDepsHandler =
  createBuildHandleMineSentenceDigitMainDepsHandler({
  getSubtitleTimingTracker: () => appState.subtitleTimingTracker,
  getAnkiIntegration: () => appState.ankiIntegration,
  getCurrentSecondarySubText: () => appState.mpvClient?.currentSecondarySubText || undefined,
  showMpvOsd: (text) => showMpvOsd(text),
  logError: (message, err) => {
    logger.error(message, err);
  },
  onCardsMined: (cards) => {
    appState.immersionTracker?.recordCardsMined(cards);
  },
  handleMineSentenceDigitCore,
});
const handleMineSentenceDigitMainDeps = buildHandleMineSentenceDigitMainDepsHandler();
const handleMineSentenceDigitHandler = createHandleMineSentenceDigitHandler(
  handleMineSentenceDigitMainDeps,
);
const {
  setVisibleOverlayVisible: setVisibleOverlayVisibleHandler,
  setInvisibleOverlayVisible: setInvisibleOverlayVisibleHandler,
  toggleVisibleOverlay: toggleVisibleOverlayHandler,
  toggleInvisibleOverlay: toggleInvisibleOverlayHandler,
  setOverlayVisible: setOverlayVisibleHandler,
  toggleOverlay: toggleOverlayHandler,
} = createOverlayVisibilityRuntime({
  setVisibleOverlayVisibleDeps: {
    setVisibleOverlayVisibleCore,
    setVisibleOverlayVisibleState: (nextVisible) => {
      overlayManager.setVisibleOverlayVisible(nextVisible);
    },
    updateVisibleOverlayVisibility: () => overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
    updateInvisibleOverlayVisibility: () =>
      overlayVisibilityRuntime.updateInvisibleOverlayVisibility(),
    syncInvisibleOverlayMousePassthrough: () =>
      overlayVisibilityRuntime.syncInvisibleOverlayMousePassthrough(),
    shouldBindVisibleOverlayToMpvSubVisibility: () =>
      configDerivedRuntime.shouldBindVisibleOverlayToMpvSubVisibility(),
    isMpvConnected: () => Boolean(appState.mpvClient && appState.mpvClient.connected),
    setMpvSubVisibility: (mpvSubVisible) => {
      setMpvSubVisibilityRuntime(appState.mpvClient, mpvSubVisible);
    },
  },
  setInvisibleOverlayVisibleDeps: {
    setInvisibleOverlayVisibleCore,
    setInvisibleOverlayVisibleState: (nextVisible) => {
      overlayManager.setInvisibleOverlayVisible(nextVisible);
    },
    updateInvisibleOverlayVisibility: () =>
      overlayVisibilityRuntime.updateInvisibleOverlayVisibility(),
    syncInvisibleOverlayMousePassthrough: () =>
      overlayVisibilityRuntime.syncInvisibleOverlayMousePassthrough(),
  },
  getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
  getInvisibleOverlayVisible: () => overlayManager.getInvisibleOverlayVisible(),
});

const buildHandleOverlayModalClosedMainDepsHandler =
  createBuildHandleOverlayModalClosedMainDepsHandler({
  handleOverlayModalClosedRuntime: (modal) => overlayModalRuntime.handleOverlayModalClosed(modal),
});
const handleOverlayModalClosedMainDeps =
  buildHandleOverlayModalClosedMainDepsHandler();
const handleOverlayModalClosedHandler = createHandleOverlayModalClosedHandler(
  handleOverlayModalClosedMainDeps,
);

const buildAppendClipboardVideoToQueueMainDepsHandler =
  createBuildAppendClipboardVideoToQueueMainDepsHandler({
  appendClipboardVideoToQueueRuntime,
  getMpvClient: () => appState.mpvClient,
  readClipboardText: () => clipboard.readText(),
  showMpvOsd: (text) => showMpvOsd(text),
  sendMpvCommand: (command) => {
    sendMpvCommandRuntime(appState.mpvClient, command);
  },
});
const appendClipboardVideoToQueueMainDeps =
  buildAppendClipboardVideoToQueueMainDepsHandler();
const appendClipboardVideoToQueueHandler = createAppendClipboardVideoToQueueHandler(
  appendClipboardVideoToQueueMainDeps,
);

const buildMpvCommandFromIpcRuntimeMainDepsHandler =
  createBuildMpvCommandFromIpcRuntimeMainDepsHandler({
    triggerSubsyncFromConfig: () => triggerSubsyncFromConfig(),
    openRuntimeOptionsPalette: () => openRuntimeOptionsPalette(),
    cycleRuntimeOption: (id, direction) => {
      if (!appState.runtimeOptionsManager) {
        return { ok: false, error: 'Runtime options manager unavailable' };
      }
      return applyRuntimeOptionResultRuntime(
        appState.runtimeOptionsManager.cycleOption(id, direction),
        (text) => showMpvOsd(text),
      );
    },
    showMpvOsd: (text: string) => showMpvOsd(text),
    replayCurrentSubtitle: () => replayCurrentSubtitleRuntime(appState.mpvClient),
    playNextSubtitle: () => playNextSubtitleRuntime(appState.mpvClient),
    sendMpvCommand: (rawCommand: (string | number)[]) =>
      sendMpvCommandRuntime(appState.mpvClient, rawCommand),
    isMpvConnected: () => Boolean(appState.mpvClient && appState.mpvClient.connected),
    hasRuntimeOptionsManager: () => appState.runtimeOptionsManager !== null,
  });

const mpvCommandFromIpcRuntimeMainDeps = buildMpvCommandFromIpcRuntimeMainDepsHandler();
const { handleMpvCommandFromIpc: handleMpvCommandFromIpcHandler, runSubsyncManualFromIpc: runSubsyncManualFromIpcHandler } =
  createIpcRuntimeHandlers<SubsyncManualRunRequest, Awaited<ReturnType<typeof subsyncRuntime.runManualFromIpc>>>({
    handleMpvCommandFromIpcDeps: {
      handleMpvCommandFromIpcRuntime,
      buildMpvCommandDeps: () => mpvCommandFromIpcRuntimeMainDeps,
    },
    runSubsyncManualFromIpcDeps: {
      runManualFromIpc: (request) => subsyncRuntime.runManualFromIpc(request),
    },
  });
const createCliCommandContextHandler = createCliCommandContextFactory({
    appState,
    texthookerService,
    getResolvedConfig: () => getResolvedConfig(),
    openExternal: (url: string) => shell.openExternal(url),
    logBrowserOpenError: (url: string, error: unknown) =>
      logger.error(`Failed to open browser for texthooker URL: ${url}`, error),
    showMpvOsd: (text: string) => showMpvOsd(text),
    initializeOverlayRuntime: () => initializeOverlayRuntime(),
    toggleVisibleOverlay: () => toggleVisibleOverlay(),
    toggleInvisibleOverlay: () => toggleInvisibleOverlay(),
    setVisibleOverlayVisible: (visible: boolean) => setVisibleOverlayVisible(visible),
    setInvisibleOverlayVisible: (visible: boolean) => setInvisibleOverlayVisible(visible),
    copyCurrentSubtitle: () => copyCurrentSubtitle(),
    startPendingMultiCopy: (timeoutMs: number) => startPendingMultiCopy(timeoutMs),
    mineSentenceCard: () => mineSentenceCard(),
    startPendingMineSentenceMultiple: (timeoutMs: number) =>
      startPendingMineSentenceMultiple(timeoutMs),
    updateLastCardFromClipboard: () => updateLastCardFromClipboard(),
    refreshKnownWordCache: () => refreshKnownWordCache(),
    triggerFieldGrouping: () => triggerFieldGrouping(),
    triggerSubsyncFromConfig: () => triggerSubsyncFromConfig(),
    markLastCardAsAudioCard: () => markLastCardAsAudioCard(),
    getAnilistStatus: () => anilistStateRuntime.getStatusSnapshot(),
    clearAnilistToken: () => anilistStateRuntime.clearTokenState(),
    openAnilistSetupWindow: () => openAnilistSetupWindow(),
    openJellyfinSetupWindow: () => openJellyfinSetupWindow(),
    getAnilistQueueStatus: () => anilistStateRuntime.getQueueStatusSnapshot(),
    processNextAnilistRetryUpdate: () => processNextAnilistRetryUpdate(),
    runJellyfinCommand: (argsFromCommand: CliArgs) => runJellyfinCommand(argsFromCommand),
    openYomitanSettings: () => openYomitanSettings(),
    cycleSecondarySubMode: () => cycleSecondarySubMode(),
    openRuntimeOptionsPalette: () => openRuntimeOptionsPalette(),
    printHelp: () => printHelp(DEFAULT_TEXTHOOKER_PORT),
    stopApp: () => app.quit(),
    hasMainWindow: () => Boolean(overlayManager.getMainWindow()),
    getMultiCopyTimeoutMs: () => getConfiguredShortcuts().multiCopyTimeoutMs,
    schedule: (fn: () => void, delayMs: number) => setTimeout(fn, delayMs),
    logInfo: (message: string) => logger.info(message),
    logWarn: (message: string) => logger.warn(message),
    logError: (message: string, err: unknown) => logger.error(message, err),
  });
const {
  createOverlayWindow: createOverlayWindowHandler,
  createMainWindow: createMainWindowHandler,
  createInvisibleWindow: createInvisibleWindowHandler,
} = createOverlayWindowRuntimeHandlers<BrowserWindow>({
  createOverlayWindowDeps: {
    createOverlayWindowCore: (kind, options) => createOverlayWindowCore(kind, options),
    isDev,
    getOverlayDebugVisualizationEnabled: () => appState.overlayDebugVisualizationEnabled,
    ensureOverlayWindowLevel: (window) => ensureOverlayWindowLevel(window),
    onRuntimeOptionsChanged: () => broadcastRuntimeOptionsChanged(),
    setOverlayDebugVisualizationEnabled: (enabled) =>
      setOverlayDebugVisualizationEnabled(enabled),
    isOverlayVisible: (windowKind) =>
      windowKind === 'visible'
        ? overlayManager.getVisibleOverlayVisible()
        : overlayManager.getInvisibleOverlayVisible(),
    tryHandleOverlayShortcutLocalFallback: (input) =>
      overlayShortcutsRuntime.tryHandleOverlayShortcutLocalFallback(input),
    onWindowClosed: (windowKind) => {
      if (windowKind === 'visible') {
        overlayManager.setMainWindow(null);
      } else {
        overlayManager.setInvisibleWindow(null);
      }
    },
  },
  setMainWindow: (window) => overlayManager.setMainWindow(window),
  setInvisibleWindow: (window) => overlayManager.setInvisibleWindow(window),
});
const {
  resolveTrayIconPath: resolveTrayIconPathHandler,
  buildTrayMenu: buildTrayMenuHandler,
  ensureTray: ensureTrayHandler,
  destroyTray: destroyTrayHandler,
} = createTrayRuntimeHandlers({
  resolveTrayIconPathDeps: {
    resolveTrayIconPathRuntime,
    platform: process.platform,
    resourcesPath: process.resourcesPath,
    appPath: app.getAppPath(),
    dirname: __dirname,
    joinPath: (...parts) => path.join(...parts),
    fileExists: (candidate) => fs.existsSync(candidate),
  },
  buildTrayMenuTemplateDeps: {
    buildTrayMenuTemplateRuntime,
    initializeOverlayRuntime: () => initializeOverlayRuntime(),
    isOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
    setVisibleOverlayVisible: (visible) => setVisibleOverlayVisible(visible),
    openYomitanSettings: () => openYomitanSettings(),
    openRuntimeOptionsPalette: () => openRuntimeOptionsPalette(),
    openJellyfinSetupWindow: () => openJellyfinSetupWindow(),
    openAnilistSetupWindow: () => openAnilistSetupWindow(),
    quitApp: () => app.quit(),
  },
  ensureTrayDeps: {
    getTray: () => appTray,
    setTray: (tray) => {
      appTray = tray as Tray | null;
    },
    createImageFromPath: (iconPath) => nativeImage.createFromPath(iconPath),
    createEmptyImage: () => nativeImage.createEmpty(),
    createTray: (icon) => new Tray(icon as never),
    trayTooltip: TRAY_TOOLTIP,
    platform: process.platform,
    logWarn: (message) => logger.warn(message),
    initializeOverlayRuntime: () => initializeOverlayRuntime(),
    isOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
    setVisibleOverlayVisible: (visible) => setVisibleOverlayVisible(visible),
  },
  destroyTrayDeps: {
    getTray: () => appTray,
    setTray: (tray) => {
      appTray = tray as Tray | null;
    },
  },
  buildMenuFromTemplate: (template) => Menu.buildFromTemplate(template),
});
const yomitanExtensionRuntime = createYomitanExtensionRuntime({
  loadYomitanExtensionCore,
  userDataPath: USER_DATA_PATH,
  getYomitanParserWindow: () => appState.yomitanParserWindow,
  setYomitanParserWindow: (window) => {
    appState.yomitanParserWindow = window as BrowserWindow | null;
  },
  setYomitanParserReadyPromise: (promise) => {
    appState.yomitanParserReadyPromise = promise;
  },
  setYomitanParserInitPromise: (promise) => {
    appState.yomitanParserInitPromise = promise;
  },
  setYomitanExtension: (extension) => {
    appState.yomitanExt = extension;
  },
  getYomitanExtension: () => appState.yomitanExt,
  getLoadInFlight: () => yomitanLoadInFlight,
  setLoadInFlight: (promise) => {
    yomitanLoadInFlight = promise;
  },
});
const { initializeOverlayRuntime: initializeOverlayRuntimeHandler } =
  createOverlayRuntimeBootstrapHandlers({
    initializeOverlayRuntimeMainDeps: {
      appState,
      overlayManager: {
        getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
        getInvisibleOverlayVisible: () => overlayManager.getInvisibleOverlayVisible(),
      },
      overlayVisibilityRuntime: {
        updateVisibleOverlayVisibility: () => overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
        updateInvisibleOverlayVisibility: () =>
          overlayVisibilityRuntime.updateInvisibleOverlayVisibility(),
      },
      overlayShortcutsRuntime: {
        syncOverlayShortcuts: () => overlayShortcutsRuntime.syncOverlayShortcuts(),
      },
      getInitialInvisibleOverlayVisibility: () =>
        configDerivedRuntime.getInitialInvisibleOverlayVisibility(),
      createMainWindow: () => createMainWindow(),
      createInvisibleWindow: () => createInvisibleWindow(),
      registerGlobalShortcuts: () => registerGlobalShortcuts(),
      updateVisibleOverlayBounds: (geometry) => updateVisibleOverlayBounds(geometry),
      updateInvisibleOverlayBounds: (geometry) => updateInvisibleOverlayBounds(geometry),
      getOverlayWindows: () => getOverlayWindows(),
      getResolvedConfig: () => getResolvedConfig(),
      showDesktopNotification,
      createFieldGroupingCallback: () => createFieldGroupingCallback() as never,
      getKnownWordCacheStatePath: () => path.join(USER_DATA_PATH, 'known-words-cache.json'),
    },
    initializeOverlayRuntimeBootstrapDeps: {
      isOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
      initializeOverlayRuntimeCore: (options) => initializeOverlayRuntimeCore(options as never),
      setInvisibleOverlayVisible: (visible) => {
        overlayManager.setInvisibleOverlayVisible(visible);
      },
      setOverlayRuntimeInitialized: (initialized) => {
        appState.overlayRuntimeInitialized = initialized;
      },
      startBackgroundWarmups: () => startBackgroundWarmups(),
    },
  });
const { openYomitanSettings: openYomitanSettingsHandler } = createYomitanSettingsRuntime({
  ensureYomitanExtensionLoaded: () => ensureYomitanExtensionLoaded(),
  openYomitanSettingsWindow: ({ yomitanExt, getExistingWindow, setWindow }) => {
    openYomitanSettingsWindow({
      yomitanExt: yomitanExt as Extension,
      getExistingWindow: () => getExistingWindow() as BrowserWindow | null,
      setWindow: (window) => setWindow(window as BrowserWindow | null),
    });
  },
  getExistingWindow: () => appState.yomitanSettingsWindow,
  setWindow: (window) => {
    appState.yomitanSettingsWindow = window as BrowserWindow | null;
  },
  logWarn: (message) => logger.warn(message),
  logError: (message, error) => logger.error(message, error),
});

async function updateLastCardFromClipboard(): Promise<void> {
  await updateLastCardFromClipboardHandler();
}

async function refreshKnownWordCache(): Promise<void> {
  await refreshKnownWordCacheHandler();
}

async function triggerFieldGrouping(): Promise<void> {
  await triggerFieldGroupingHandler();
}

async function markLastCardAsAudioCard(): Promise<void> {
  await markLastCardAsAudioCardHandler();
}

async function mineSentenceCard(): Promise<void> {
  await mineSentenceCardHandler();
}

function cancelPendingMineSentenceMultiple(): void {
  cancelPendingMineSentenceMultipleHandler();
}

function startPendingMineSentenceMultiple(timeoutMs: number): void {
  startPendingMineSentenceMultipleHandler(timeoutMs);
}

function handleMineSentenceDigit(count: number): void {
  handleMineSentenceDigitHandler(count);
}

function registerOverlayShortcuts(): void {
  registerOverlayShortcutsHandler();
}

function unregisterOverlayShortcuts(): void {
  unregisterOverlayShortcutsHandler();
}

function syncOverlayShortcuts(): void {
  syncOverlayShortcutsHandler();
}
function refreshOverlayShortcuts(): void {
  refreshOverlayShortcutsHandler();
}

function setVisibleOverlayVisible(visible: boolean): void {
  setVisibleOverlayVisibleHandler(visible);
}

function setInvisibleOverlayVisible(visible: boolean): void {
  setInvisibleOverlayVisibleHandler(visible);
}

function toggleVisibleOverlay(): void {
  toggleVisibleOverlayHandler();
}
function toggleInvisibleOverlay(): void {
  toggleInvisibleOverlayHandler();
}
function setOverlayVisible(visible: boolean): void {
  setOverlayVisibleHandler(visible);
}
function toggleOverlay(): void {
  toggleOverlayHandler();
}
function handleOverlayModalClosed(modal: OverlayHostedModal): void {
  handleOverlayModalClosedHandler(modal);
}

function handleMpvCommandFromIpc(command: (string | number)[]): void {
  handleMpvCommandFromIpcHandler(command);
}

async function runSubsyncManualFromIpc(request: SubsyncManualRunRequest): Promise<SubsyncResult> {
  return runSubsyncManualFromIpcHandler(request);
}

function appendClipboardVideoToQueue(): { ok: boolean; message: string } {
  return appendClipboardVideoToQueueHandler();
}

registerIpcRuntimeServices({
  runtimeOptions: {
    getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
    showMpvOsd: (text: string) => showMpvOsd(text),
  },
  mainDeps: {
    getInvisibleWindow: () => overlayManager.getInvisibleWindow(),
    getMainWindow: () => overlayManager.getMainWindow(),
    getVisibleOverlayVisibility: () => overlayManager.getVisibleOverlayVisible(),
    getInvisibleOverlayVisibility: () => overlayManager.getInvisibleOverlayVisible(),
    focusMainWindow: () => {
      const mainWindow = overlayManager.getMainWindow();
      if (!mainWindow || mainWindow.isDestroyed()) return;
      if (!mainWindow.isFocused()) {
        mainWindow.focus();
      }
    },
    onOverlayModalClosed: (modal: string) => {
      handleOverlayModalClosed(modal as OverlayHostedModal);
    },
    openYomitanSettings: () => openYomitanSettings(),
    quitApp: () => app.quit(),
    toggleVisibleOverlay: () => toggleVisibleOverlay(),
    tokenizeCurrentSubtitle: () => tokenizeSubtitle(appState.currentSubText),
    getCurrentSubtitleRaw: () => appState.currentSubText,
    getCurrentSubtitleAss: () => appState.currentSubAssText,
    getMpvSubtitleRenderMetrics: () => appState.mpvSubtitleRenderMetrics,
    getSubtitlePosition: () => loadSubtitlePosition(),
    getSubtitleStyle: () => {
      const resolvedConfig = getResolvedConfig();
      return resolveSubtitleStyleForRenderer(resolvedConfig);
    },
    saveSubtitlePosition: (position: unknown) => saveSubtitlePosition(position as SubtitlePosition),
    getMecabTokenizer: () => appState.mecabTokenizer,
    handleMpvCommand: (command: (string | number)[]) => handleMpvCommandFromIpc(command),
    getKeybindings: () => appState.keybindings,
    getConfiguredShortcuts: () => getConfiguredShortcuts(),
    getSecondarySubMode: () => appState.secondarySubMode,
    getMpvClient: () => appState.mpvClient,
    runSubsyncManual: (request: unknown) =>
      runSubsyncManualFromIpc(request as SubsyncManualRunRequest),
    getAnkiConnectStatus: () => appState.ankiIntegration !== null,
    getRuntimeOptions: () => getRuntimeOptionsState(),
    reportOverlayContentBounds: (payload: unknown) => {
      overlayContentMeasurementStore.report(payload);
    },
    getAnilistStatus: () => anilistStateRuntime.getStatusSnapshot(),
    clearAnilistToken: () => anilistStateRuntime.clearTokenState(),
    openAnilistSetup: () => openAnilistSetupWindow(),
    getAnilistQueueStatus: () => anilistStateRuntime.getQueueStatusSnapshot(),
    retryAnilistQueueNow: () => processNextAnilistRetryUpdate(),
    appendClipboardVideoToQueue: () => appendClipboardVideoToQueue(),
  },
  ankiJimakuDeps: createAnkiJimakuIpcRuntimeServiceDeps({
    patchAnkiConnectEnabled: (enabled: boolean) => {
      configService.patchRawConfig({ ankiConnect: { enabled } });
    },
    getResolvedConfig: () => getResolvedConfig(),
    getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
    getSubtitleTimingTracker: () => appState.subtitleTimingTracker,
    getMpvClient: () => appState.mpvClient,
    getAnkiIntegration: () => appState.ankiIntegration,
    setAnkiIntegration: (integration: AnkiIntegration | null) => {
      appState.ankiIntegration = integration;
    },
    getKnownWordCacheStatePath: () => path.join(USER_DATA_PATH, 'known-words-cache.json'),
    showDesktopNotification,
    createFieldGroupingCallback: () => createFieldGroupingCallback(),
    broadcastRuntimeOptionsChanged: () => broadcastRuntimeOptionsChanged(),
    getFieldGroupingResolver: () => getFieldGroupingResolver(),
    setFieldGroupingResolver: (resolver: ((choice: KikuFieldGroupingChoice) => void) | null) =>
      setFieldGroupingResolver(resolver),
    parseMediaInfo: (mediaPath: string | null) =>
      parseMediaInfo(mediaRuntime.resolveMediaPathForJimaku(mediaPath)),
    getCurrentMediaPath: () => appState.currentMediaPath,
    jimakuFetchJson: <T>(
      endpoint: string,
      query?: Record<string, string | number | boolean | null | undefined>,
    ): Promise<JimakuApiResponse<T>> => configDerivedRuntime.jimakuFetchJson<T>(endpoint, query),
    getJimakuMaxEntryResults: () => configDerivedRuntime.getJimakuMaxEntryResults(),
    getJimakuLanguagePreference: () => configDerivedRuntime.getJimakuLanguagePreference(),
    resolveJimakuApiKey: () => configDerivedRuntime.resolveJimakuApiKey(),
    isRemoteMediaPath: (mediaPath: string) => isRemoteMediaPath(mediaPath),
    downloadToFile: (url: string, destPath: string, headers: Record<string, string>) =>
      downloadToFile(url, destPath, headers),
  }),
});
