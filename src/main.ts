/*
  SubMiner - All-in-one sentence mining overlay
  Copyright (C) 2026 sudacode

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
import { generateSentenceFurigana } from './core/services/tokenizer/sentence-furigana';
import {
  app,
  BrowserWindow,
  clipboard,
  globalShortcut,
  ipcMain,
  shell,
  protocol,
  Extension,
  Session,
  Menu,
  nativeImage,
  Tray,
  dialog,
  screen,
} from 'electron';
import { applyControllerConfigUpdate } from './main/controller-config-update.js';
import {
  createHachidoriExtensionRuntime,
  getHachidoriSession,
} from './core/services/hachidori-extension';
import {
  DICTIONARY_EXTERNAL_LINK_CHANNEL,
  parseDictionaryExternalUrl,
} from './shared/dictionary-external-link';
import { readMpvInputBindings } from './main/runtime/mpv-input-bindings';
import { createAniSkipRuntime } from './main/runtime/aniskip-runtime';
import { resolveAniSkipMetadataForFile } from './main/runtime/aniskip-metadata';
import { createDiscordRpcClient } from './main/runtime/discord-rpc-client.js';
import { startAppControlServer } from './main/runtime/app-control-server';
import { createEnsureBackgroundStatsServerHandler } from './main/runtime/background-stats-startup';
import {
  markJellyfinRemotePlaybackLoaded as markJellyfinRemotePlaybackLoadedState,
  shouldAutoLoadSecondarySubTrackForJellyfinPlayback,
} from './main/runtime/jellyfin-remote-playback';
import { getAppControlSocketPath } from './shared/app-control';
import {
  type CancelLinuxMpvFullscreenOverlayRefreshBurst,
  clearLinuxMpvFullscreenOverlayRefreshTimeouts,
  updateLinuxMpvFullscreenOverlayRefreshBurst,
} from './main/runtime/linux-mpv-fullscreen-overlay-refresh';
import { createLinuxOverlayModeRuntime } from './main/runtime/linux-overlay-mode-runtime';
import { shouldRunLinuxOverlayZOrderKeepAlive } from './main/runtime/linux-overlay-zorder-keepalive';
import { focusMacOSOverlayWindow } from './main/runtime/macos-overlay-window-focus';
import { restoreMacOSMpvFocusAfterModalClose } from './main/runtime/macos-modal-focus-handoff';
import { resolveFreshPlaybackPaused } from './main/runtime/playback-paused-state';
import { mergeAiConfig } from './ai/config';
import {
  getDefaultPasswordStore,
  getPasswordStoreArg,
  normalizePasswordStoreArg,
} from './main/password-store-args';

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

import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { MecabTokenizer } from './mecab-tokenizer';
import type {
  JimakuApiResponse,
  KikuFieldGroupingChoice,
  MpvSubtitleRenderMetrics,
  ResolvedConfig,
  SessionActionDispatchRequest,
  SecondarySubMode,
  SubtitleData,
  SubtitleCue,
  SubtitleMiningContext,
  SubtitlePosition,
  SubtitleSidebarPlaybackState,
  WindowGeometry,
} from './types';
import { OPEN_ANKI_CARD_ACTION_ID } from './types';
import { AnkiIntegration } from './anki-integration';
import { AnkiConnectClient } from './anki-connect';
import { SubtitleTimingTracker } from './subtitle-timing-tracker';
import { RuntimeOptionsManager } from './runtime-options';
import { downloadToFile, isRemoteMediaPath, parseMediaInfo } from './jimaku/utils';
import {
  createLogger,
  setLogFileToggles,
  setLogLevel,
  setLogRotation,
  type LogLevelSource,
} from './logger';
import {
  isLogFileEnabled,
  pruneLogDirectoryForPath,
  resolveDefaultLogFilePath,
} from './shared/log-files';
import { createFatalErrorReporter } from './main/fatal-error';
import { ensureWindowsOverlayTransparency } from './window-trackers/windows-helper';
import {
  commandNeedsOverlayStartupPrereqs,
  commandNeedsOverlayRuntime,
  isHeadlessInitialCommand,
  parseArgs,
  shouldStartApp,
  type CliArgs,
  type CliCommandSource,
} from './cli/args';
import { printHelp } from './cli/help';
import { IPC_CHANNELS, type OverlayHostedModal } from './shared/ipc/contracts';
import { buildMpvLoggingArgs } from './shared/mpv-logging-args';
import {
  getStartupModeFlags,
  shouldHandleInitialArgsBeforeDeferredOverlayWarmup,
  shouldRefreshAnilistOnConfigReload,
  shouldStartAutomaticUpdateChecks,
} from './main/runtime/startup-mode-flags';
import { buildConfigParseErrorDetails, failStartupFromConfig } from './main/config-validation';
import {
  findAnilistSetupDeepLinkArgvUrl,
  isAnilistTrackingEnabled,
} from './main/runtime/domains/anilist';
import { shouldShowTexthookerTrayEntry } from './main/runtime/tray-main-actions';
import { createGetDefaultSocketPathHandler } from './main/runtime/domains/jellyfin';
import {
  createTrayRuntimeHandlers,
  createOverlayVisibilityRuntime,
  createGetRuntimeOptionsStateHandler,
  createGetFieldGroupingResolverHandler,
  createSetFieldGroupingResolverHandler,
  createLoadSubtitlePositionHandler,
  createSaveSubtitlePositionHandler,
  createAppendClipboardVideoToQueueHandler,
  createHandleOverlayModalClosedHandler,
  buildTrayMenuTemplateRuntime,
  resolveTrayIconPathRuntime,
  createYomitanExtensionRuntime,
  createYomitanSettingsRuntime,
  resolveSubtitleStyleForRenderer,
} from './main/runtime/domains/overlay';
import {
  createImmersionMediaRuntime,
  createConfigDerivedRuntime,
  appendClipboardVideoToQueueRuntime,
  createMainSubsyncRuntime,
} from './main/runtime/domains/startup';
import {
  createMpvOsdRuntimeHandlers,
  createCycleSecondarySubModeRuntimeHandler,
} from './main/runtime/domains/mpv';
import { buildSubtitleTrackDiagnostics } from './main/runtime/mpv-track-diagnostics';
import {
  resolveCanonicalPrimarySubtitle,
  resolvePrimarySubtitle,
} from './main/runtime/primary-subtitle-text';
import {
  createMarkLastCardAsAudioCardHandler,
  createMineSentenceCardHandler,
  createRefreshKnownWordCacheHandler,
  createTriggerFieldGroupingHandler,
  createUpdateLastCardFromClipboardHandler,
  createCopyCurrentSubtitleHandler,
  createHandleMineSentenceDigitHandler,
  createHandleMultiCopyDigitHandler,
} from './main/runtime/domains/mining';
import {
  enforceUnsupportedWaylandMode,
  forceX11Backend,
  generateDefaultConfigFile,
  resolveConfiguredShortcuts,
  resolveKeybindings,
  showDesktopNotification,
} from './core/utils';
import { ensureDefaultConfigBootstrap, getDefaultConfigFilePaths } from './shared/setup-state';
import {
  ImmersionTrackerService,
  MpvIpcClient,
  SubtitleWebSocket,
  Texthooker,
  applyMpvSubtitleRenderMetricsPatch,
  broadcastRuntimeOptionsChangedRuntime,
  captureLiveSubtitleMiningContext,
  copyCurrentSubtitle as copyCurrentSubtitleCore,
  createDiscordPresenceService,
  createFieldGroupingOverlayRuntime,
  createOverlayContentMeasurementStore,
  createOverlayManager,
  createOverlayWindow as createOverlayWindowCore,
  createTokenizerDepsRuntime,
  cycleSecondarySubMode as cycleSecondarySubModeCore,
  destroyYomitanSettingsWindow,
  handleMineSentenceDigit as handleMineSentenceDigitCore,
  handleMultiCopyDigit as handleMultiCopyDigitCore,
  hasMpvWebsocketPlugin,
  initializeOverlayRuntime as initializeOverlayRuntimeCore,
  isSubtitleAnnotationUpgrade,
  isOverlayWindowContentReady,
  loadSubtitlePosition as loadSubtitlePositionCore,
  loadYomitanExtension as loadYomitanExtensionCore,
  markLastCardAsAudioCard as markLastCardAsAudioCardCore,
  mineSentenceCard as mineSentenceCardCore,
  openYomitanSettingsWindow,
  playNextSubtitleRuntime,
  promoteSettingsWindowAboveOverlay,
  registerGlobalShortcuts as registerGlobalShortcutsCore,
  replayCurrentSubtitleRuntime,
  resolveSanitizedSubtitleSeekCommand,
  toJellyfinOriginalStreamUrl,
  runStartupBootstrapRuntime,
  saveSubtitlePosition as saveSubtitlePositionCore,
  clearYomitanParserCachesForWindow,
  getYomitanCurrentAnkiDeckName as getYomitanCurrentAnkiDeckNameCore,
  sendMpvCommandRuntime,
  setMpvSubVisibilityRuntime,
  setOverlayDebugVisualizationEnabledRuntime,
  setVisibleOverlayVisible as setVisibleOverlayVisibleCore,
  shiftSubtitleRangeToVideoClock,
  showMpvOsdRuntime,
  tokenizeSubtitle as tokenizeSubtitleCore,
  toSubtitleClock,
  triggerFieldGrouping as triggerFieldGroupingCore,
  updateLastCardFromClipboard as updateLastCardFromClipboardCore,
} from './core/services';
import { applyOverlayClickThrough } from './core/services/overlay-click-through';
import {
  destroyStatsWindow,
  registerStatsOverlayToggle,
  toggleStatsOverlay as toggleStatsOverlayWindow,
  withStatsWindowLayerSuspendedForNativeDialog,
} from './core/services/stats-window.js';
import { shouldAutoOpenFirstRunSetup } from './main/runtime/first-run-setup-service';
import { createCharacterDictionaryAssembly } from './main/assemblies/character-dictionary';
import { createDictionariesAssembly } from './main/assemblies/dictionaries';
import { createJellyfinAssembly } from './main/assemblies/jellyfin';
import { createAnilistAssembly } from './main/assemblies/anilist';
import { createSettingsAssembly } from './main/assemblies/settings';
import { createSyncAssembly } from './main/assemblies/sync';
import { createYoutubeAssembly } from './main/assemblies/youtube';
import { createOverlayModalsAssembly } from './main/assemblies/overlay-modals';
import { createFirstRunSetupAssembly } from './main/assemblies/first-run-setup';
import {
  createAutoplayReadyGate,
  type AutoplayReadySignal,
} from './main/runtime/autoplay-ready-gate';
import { createAutoplaySubtitlePrimingRuntime } from './main/runtime/autoplay-subtitle-priming-runtime';
import { createAutoplayTokenizationWarmRelease } from './main/runtime/autoplay-tokenization-warm-release';
import { isVisibleOverlayAutoplayTargetReady } from './main/runtime/visible-overlay-autoplay-readiness';
import { createManagedLocalSubtitleSelectionRuntime } from './main/runtime/local-subtitle-selection';
import { createWindowsMpvPluginDetectionRuntime } from './main/runtime/windows-mpv-plugin-detection';
import { refreshManagedCommandLineLauncher } from './main/runtime/command-line-launcher';
import { applyLoginShellPath } from './main/runtime/login-shell-path';
import { createWindowsMpvLaunchDeps, launchWindowsMpv } from './main/runtime/windows-mpv-launch';
import { findMpvExecutablePath } from './main/runtime/mpv-process';
import {
  isJellyfinConfiguredForTray as isJellyfinConfiguredForTrayRuntime,
  toggleJellyfinDiscoveryFromTray as toggleJellyfinDiscoveryFromTrayRuntime,
} from './main/runtime/jellyfin-tray-discovery';
import {
  shouldEnsureTrayOnStartupForInitialArgs,
  shouldQuitOnMpvShutdownForTrayState,
  shouldQuitOnWindowAllClosedForTrayState,
} from './main/runtime/startup-tray-policy';
import { createLogExportTrayRuntime } from './main/runtime/log-export-tray';
import {
  createImmersionTrackerStartupHandler,
  type ImmersionTrackerStartupDeps,
} from './main/runtime/immersion-startup';
import {
  createRunStatsCliCommandHandler,
  writeStatsCliCommandResponse,
} from './main/runtime/stats-cli-command';
import { createStatsServerRuntime } from './main/runtime/stats-server-runtime';
import { createForceQuitHandler } from './main/runtime/app-lifecycle-actions';
import { resolveLegacyVocabularyPosFromTokens } from './core/services/immersion-tracker/legacy-vocabulary-pos';
import { createAnilistUpdateQueue } from './core/services/anilist/anilist-update-queue';
import { createCoverArtFetcher } from './core/services/anilist/cover-art-fetcher';
import { createAnilistRateLimiter } from './core/services/anilist/rate-limiter';
import { createLiveActionMetadataResolver } from './core/services/tmdb/live-action-resolver';
import { createTmdbClient, createTmdbApiKeyResolver } from './core/services/tmdb/tmdb-client';
import { readBundledTmdbApiKey } from './core/services/tmdb/bundled-api-key';
import { createJellyfinTokenStore } from './core/services/jellyfin-token-store';
import { applyRuntimeOptionResultRuntime } from './core/services/runtime-options-ipc';
import { createAnilistTokenStore } from './core/services/anilist/anilist-token-store';
import { dispatchSessionAction as dispatchSessionActionCore } from './core/services/session-actions';
import { createMainRuntimeRegistry } from './main/runtime/registry';
import {
  createEnsureOverlayMpvSubtitlesHiddenHandler,
  createRestoreOverlayMpvSubtitlesHandler,
} from './main/runtime/overlay-mpv-sub-visibility';
import {
  composeAppReadyRuntime,
  composeCliStartupHandlers,
  composeHeadlessStartupHandlers,
  composeIpcRuntimeHandlers,
  composeMpvRuntimeHandlers,
  composeOverlayVisibilityRuntime,
  composeShortcutRuntimes,
  composeStartupLifecycleHandlers,
} from './main/runtime/composers';
import { createOverlayWindowRuntimeHandlers } from './main/runtime/overlay-window-runtime-handlers';
import { tryBeginVisibleOverlayNumericSelection } from './main/runtime/overlay-numeric-selection';
import { createStartupBootstrapRuntimeDeps } from './main/startup';
import { createAppLifecycleRuntimeRunner } from './main/startup-lifecycle';
import {
  registerSecondInstanceHandlerEarly,
  requestSingleInstanceLockEarly,
  shouldBypassSingleInstanceLockForArgv,
} from './main/early-single-instance';
import { handleMpvCommandFromIpcRuntime } from './main/ipc-mpv-command';
import { registerIpcRuntimeServices } from './main/ipc-runtime';
import { createSubtitleGenerationRuntime } from './main/runtime/subtitle-generation-runtime';
import { registerSubtitleGenerationIpc } from './main/runtime/subtitle-generation-ipc';
import {
  createSubtitleSelectionRuntime,
  openSubtitleSelectionModal,
  registerSubtitleSelectionIpc,
} from './main/runtime/subtitle-selection';
import { openSubtitleGenerationModal } from './main/runtime/subtitle-generation-open';
import { createAnkiJimakuIpcRuntimeServiceDeps } from './main/dependencies';
import { createMainBootServices, type MainBootServicesResult } from './main/boot/services';
import { handleCliCommandRuntimeServiceWithContext } from './main/cli-runtime';
import { createOverlayModalRuntimeService } from './main/overlay-runtime';
import { createOverlayModalInputState } from './main/runtime/overlay-modal-input-state';
import { MediaTimingPreviewSession } from './core/services/media-timing-preview';
import { getSharedRemoteMediaWindowCache } from './core/services/remote-media-window-cache';
import { resolveMediaGenerationInput } from './anki-integration/media-source';
import { generateSpeechWaveform } from './core/services/media-timing-waveform';
import { createMediaTimingFrameExtractor } from './core/services/media-timing-frame';
import {
  collectMediaTimingContextLines,
  createMediaTimingReviewRuntime,
} from './main/runtime/media-timing-review';
import { openMediaTimingReviewModal } from './main/runtime/media-timing-review-open';
import { openSubsyncManualModal as openSubsyncManualModalRuntime } from './main/runtime/subsync-open';
import { createChangelogRuntime } from './main/runtime/changelog/changelog-runtime';
import {
  createPluginHost,
  PLUGIN_DEACTIVATE_TIMEOUT_MS,
} from './core/services/plugins/plugin-host';
import { buildPluginSubtitleContext } from './core/services/plugins/subtitle-context';
import type { PluginSubtitleLine } from './core/services/plugins/types';
import {
  createPluginPanelRuntime,
  openPluginPanelModal,
} from './main/runtime/plugin-panel-runtime';
import { createPluginWindowRuntime } from './main/runtime/plugin-window-runtime';
import { createPlaylistBrowserIpcRuntime } from './main/runtime/playlist-browser-ipc';
import { createSessionBindingsRuntime } from './main/runtime/session-bindings-runtime';
import { createOverlayShortcutsRuntimeService } from './main/overlay-shortcuts-runtime';
import { createMediaRuntimeService } from './main/media-runtime';
import { createPrefetchAwareSubtitleProcessingController } from './main/runtime/subtitle-processing-runtime';
import { createOverlayVisibilityRuntimeService } from './main/overlay-visibility-runtime';
import { createDiscordPresenceRuntime } from './main/runtime/discord-presence-runtime';
import {
  getCharacterDictionaryManagerSnapshot,
  moveCharacterDictionaryManagedEntry,
  removeCharacterDictionaryManagedEntry,
  replaceCharacterDictionaryManagedEntry,
} from './main/runtime/character-dictionary-auto-sync';
import { createCurrentMediaTokenizationGate } from './main/runtime/current-media-tokenization-gate';
import { resolveCurrentSubtitleForRenderer } from './main/runtime/current-subtitle-snapshot';
import { createOverlayGeometryRuntime } from './main/runtime/overlay-geometry-runtime';
import { createVisibleOverlayInteractionRuntime } from './main/runtime/visible-overlay-interaction-runtime';
import { createStartupOsdSequencer } from './main/runtime/startup-osd-sequencer';
import {
  INSTALL_UPDATE_ACTION_ID,
  UPDATE_AVAILABLE_NOTIFICATION_ID,
  VIEW_CHANGELOG_ACTION_ID,
} from './main/runtime/update/update-notifications';
import { createOverlayNotificationsRuntime } from './main/runtime/overlay-notifications-runtime';
import {
  runUpdateCliCommand,
  writeUpdateCliCommandResponse,
} from './main/runtime/update/update-cli-command';
import {
  runEnsureLinuxRuntimePluginAssetsCliCommand,
  writeEnsureLinuxRuntimePluginAssetsCliCommandResponse,
} from './main/runtime/linux-runtime-plugin-assets-cli-command';
import { ensureLinuxRuntimePluginAssets } from './main/runtime/linux-runtime-plugin-assets';
import { createUpdateServiceRuntime } from './main/runtime/update/update-service-runtime';
import {
  createRefreshSubtitlePrefetchFromActiveTrackHandler,
  createResolveActiveSubtitleSidebarSourceHandler,
} from './main/runtime/subtitle-prefetch-runtime';
import { createSecondarySubtitleTrackController } from './main/runtime/secondary-subtitle-track';
import { shouldSuppressVisibleOverlayRaiseForSeparateWindow } from './main/runtime/settings-window-z-order';
import {
  isSameYoutubeMediaPath,
  isYoutubeMediaPath,
  shouldUseCachedYoutubeParsedCues,
} from './main/runtime/youtube-playback';
import { createYomitanProfilePolicy } from './main/runtime/yomitan-profile-policy';
import { reloadOverlayWindowsForYomitanContentScripts } from './main/runtime/yomitan-extension-overlay-reload';
import { createYomitanAnkiServerSyncRuntime } from './main/runtime/yomitan-anki-server-sync';
import { type StartupState, applyStartupState, createAppState } from './main/state';
import {
  ConfigService,
  DEFAULT_CONFIG,
  DEFAULT_KEYBINDINGS,
  generateConfigTemplate,
} from './config';
import { resolveConfigDir } from './config/path-resolution';
import { parseSubtitleCues } from './core/services/subtitle-cue-parser';
import {
  createSubtitlePrefetchService,
  type SubtitlePrefetchService,
} from './core/services/subtitle-prefetch';
import { createSubtitlePrefetchInitController } from './main/runtime/subtitle-prefetch-init';
import {
  createCachedInternalSubtitleTrackExtractor,
  loadSubtitleSourceText,
} from './main/runtime/internal-subtitle-extraction';
import { createRemoteMediaPathDetector } from './main/runtime/network-media-path';
import { applyCharacterDictionarySelection } from './main/character-dictionary-selection';

if (process.platform === 'linux') {
  app.commandLine.appendSwitch('enable-features', 'GlobalShortcutsPortal');
  const passwordStore = normalizePasswordStoreArg(
    getPasswordStoreArg(process.argv) ?? getDefaultPasswordStore(),
  );
  app.commandLine.appendSwitch('password-store', passwordStore);
  createLogger('main').debug(`Applied --password-store ${passwordStore}`);
}

app.setName('SubMiner');

const DEFAULT_TEXTHOOKER_PORT = 5174;
const DEFAULT_MPV_LOG_FILE = resolveDefaultLogFilePath('mpv', {
  platform: process.platform,
  homeDir: os.homedir(),
  appDataDir: process.env.APPDATA,
});
const TRAY_TOOLTIP = 'SubMiner';

const DISCORD_PRESENCE_APP_ID = '1475264834730856619';

let backgroundWarmupsStarted = false;
let yomitanLoadInFlight: Promise<Extension | null> | null = null;
let notifyAnilistTokenStoreWarning: (message: string) => void = () => {};

const isDev = process.argv.includes('--dev') || process.argv.includes('--debug');
const texthookerService = new Texthooker(() => {
  const config = configService.getConfig();
  const characterDictionaryEnabled =
    config.subtitleStyle.nameMatchEnabled && yomitanProfilePolicy.isCharacterDictionaryEnabled();
  const knownWordColoringEnabled = getRuntimeBooleanOption(
    'subtitle.annotation.knownWords.highlightEnabled',
    config.ankiConnect.knownWords.highlightEnabled,
  );
  const nPlusOneColoringEnabled = getRuntimeBooleanOption(
    'subtitle.annotation.nPlusOne',
    config.ankiConnect.nPlusOne.enabled,
  );

  return {
    enableKnownWordColoring: knownWordColoringEnabled,
    enableNPlusOneColoring: nPlusOneColoringEnabled,
    enableNameMatchColoring: config.subtitleStyle.nameMatchEnabled && characterDictionaryEnabled,
    enableFrequencyColoring: getRuntimeBooleanOption(
      'subtitle.annotation.frequency',
      config.subtitleStyle.frequencyDictionary.enabled,
    ),
    enableJlptColoring: getRuntimeBooleanOption(
      'subtitle.annotation.jlpt',
      config.subtitleStyle.enableJlpt,
    ),
    characterDictionaryEnabled,
    knownWordColor: config.subtitleStyle.knownWordColor,
    nPlusOneColor: config.subtitleStyle.nPlusOneColor,
    nameMatchColor: config.subtitleStyle.nameMatchColor,
    hoverTokenColor: config.subtitleStyle.hoverTokenColor,
    hoverTokenBackgroundColor: config.subtitleStyle.hoverTokenBackgroundColor,
    jlptColors: config.subtitleStyle.jlptColors,
    frequencyDictionary: {
      singleColor: config.subtitleStyle.frequencyDictionary.singleColor,
      bandedColors: config.subtitleStyle.frequencyDictionary.bandedColors,
    },
  };
});
let syncOverlayShortcutsForModal: (isActive: boolean) => void = () => {};
let syncOverlayVisibilityForModal: () => void = () => {};
const getDefaultSocketPathHandler = createGetDefaultSocketPathHandler({
  platform: process.platform,
});

type BootServices = MainBootServicesResult<
  ConfigService,
  ReturnType<typeof createAnilistTokenStore>,
  ReturnType<typeof createJellyfinTokenStore>,
  ReturnType<typeof createAnilistUpdateQueue>,
  SubtitleWebSocket,
  ReturnType<typeof createLogger>,
  ReturnType<typeof createMainRuntimeRegistry>,
  ReturnType<typeof createOverlayManager>,
  ReturnType<typeof createOverlayModalInputState>,
  ReturnType<typeof createOverlayContentMeasurementStore>,
  ReturnType<typeof createOverlayModalRuntimeService>,
  ReturnType<typeof createAppState>,
  {
    requestSingleInstanceLock: () => boolean;
    quit: () => void;
    exit: (code?: number) => void;
    on: (event: string, listener: (...args: unknown[]) => void) => unknown;
    whenReady: () => Promise<void>;
  }
>;

const bootServices = createMainBootServices({
  platform: process.platform,
  argv: process.argv,
  configDir: app.getPath('userData'),
  appDataDir: process.env.APPDATA,
  xdgConfigHome: process.env.XDG_CONFIG_HOME,
  homeDir: os.homedir(),
  defaultMpvLogFile: DEFAULT_MPV_LOG_FILE,
  envMpvLog: process.env.SUBMINER_MPV_LOG,
  defaultTexthookerPort: DEFAULT_TEXTHOOKER_PORT,
  getDefaultSocketPath: () => getDefaultSocketPathHandler(),
  resolveConfigDir,
  existsSync: fs.existsSync,
  mkdirSync: fs.mkdirSync,
  joinPath: (...parts) => path.join(...parts),
  app,
  shouldBypassSingleInstanceLock: () => shouldBypassSingleInstanceLockForArgv(process.argv),
  requestSingleInstanceLockEarly: () => requestSingleInstanceLockEarly(app),
  registerSecondInstanceHandlerEarly: (listener) => {
    registerSecondInstanceHandlerEarly(app, listener);
  },
  onConfigStartupParseError: (error) => {
    failStartupFromConfig(
      'SubMiner config parse error',
      buildConfigParseErrorDetails(error.path, error.parseError),
      {
        logError: (details) => console.error(details),
        showErrorBox: (title, details) => dialog.showErrorBox(title, details),
        quit: () => requestAppQuit(),
      },
    );
  },
  createConfigService: (configDir) => new ConfigService(configDir),
  createAnilistTokenStore: (targetPath) =>
    createAnilistTokenStore(targetPath, {
      info: (message: string) => console.info(message),
      warn: (message: string, details?: unknown) => console.warn(message, details),
      error: (message: string, details?: unknown) => console.error(message, details),
      warnUser: (message: string) => notifyAnilistTokenStoreWarning(message),
    }),
  createJellyfinTokenStore: (targetPath) =>
    createJellyfinTokenStore(targetPath, {
      info: (message: string) => console.info(message),
      warn: (message: string, details?: unknown) => console.warn(message, details),
      error: (message: string, details?: unknown) => console.error(message, details),
    }),
  createAnilistUpdateQueue: (targetPath) =>
    createAnilistUpdateQueue(targetPath, {
      info: (message: string) => console.info(message),
      warn: (message: string, details?: unknown) => console.warn(message, details),
      error: (message: string, details?: unknown) => console.error(message, details),
    }),
  createSubtitleWebSocket: (payloadMode) => new SubtitleWebSocket(payloadMode),
  createLogger,
  createMainRuntimeRegistry,
  createOverlayManager: () =>
    createOverlayManager({
      shouldPromoteWindowOnBoundsUpdate: (window) =>
        !shouldSuppressVisibleOverlayRaiseForSeparateWindow({
          window,
          mainWindow: overlayManager.getMainWindow(),
          separateWindows: getOverlayForegroundSeparateWindows(),
        }),
    }),
  createOverlayModalInputState,
  createOverlayContentMeasurementStore: ({ logger }) => {
    return createOverlayContentMeasurementStore({
      now: () => Date.now(),
      warn: (message: string) => logger.warn(message),
    });
  },
  getSyncOverlayShortcutsForModal: () => syncOverlayShortcutsForModal,
  getSyncOverlayVisibilityForModal: () => syncOverlayVisibilityForModal,
  createOverlayModalRuntime: ({ overlayManager, overlayModalInputState }) => {
    return createOverlayModalRuntimeService(
      {
        getMainWindow: () => overlayManager.getMainWindow(),
        getModalWindow: () => overlayManager.getModalWindow(),
        createModalWindow: () => createModalWindow(),
        getModalGeometry: () => overlayGeometryRuntime.getCurrentOverlayGeometry(),
        setModalWindowBounds: (geometry) => overlayManager.setModalWindowBounds(geometry),
      },
      {
        onModalStateChange: (isActive: boolean) =>
          overlayModalInputState.handleModalInputStateChange(isActive),
        // On macOS, after the last modal closes the post-close visibility sync hides the overlay
        // when neither it nor mpv is focused, and keyboard focus is left in limbo (mpv keys like
        // shift-to-unpause stop working). Programmatically focusing mpv from our background helper
        // is refused by macOS (most visibly in native fullscreen), so instead resign SubMiner's
        // active status — exactly what a manual click does — handing focus back to the previously
        // active app (mpv). The overlay is already hidden at this point, so app.hide() hides nothing
        // visible; once mpv is focused the tracker re-shows the overlay above it.
        onFinalModalClosed: () => {
          void restoreMacOSMpvFocusAfterModalClose({
            platform: process.platform,
            focusMpv: async () => {
              app.hide();
            },
            getWindowTracker: () => appState.windowTracker,
            updateVisibleOverlayVisibility: () =>
              overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
            warn: (message, details) => logger.warn(message, details),
          });
        },
      },
    );
  },
  createAppState,
}) as BootServices;
const {
  configDir: CONFIG_DIR,
  userDataPath: USER_DATA_PATH,
  defaultMpvLogPath: DEFAULT_MPV_LOG_PATH,
  defaultImmersionDbPath: DEFAULT_IMMERSION_DB_PATH,
  configService,
  anilistTokenStore,
  jellyfinTokenStore,
  anilistUpdateQueue,
  subtitleWsService,
  annotationSubtitleWsService,
  logger,
  runtimeRegistry,
  overlayManager,
  overlayModalInputState,
  overlayContentMeasurementStore,
  overlayModalRuntime,
  appState,
  appLifecycleApp,
} = bootServices;
// Backend changes take effect on restart; windows and parser must share one session.
const activeDictionaryBackend = configService.getConfig().dictionaryBackend;
const hachidoriExtensionRuntime = createHachidoriExtensionRuntime(USER_DATA_PATH);
let hachidoriSettingsWindow: BrowserWindow | null = null;
let inactiveYomitanExtension: Extension | null = null;
let inactiveYomitanSettingsWindow: BrowserWindow | null = null;
let inactiveYomitanLoad: Promise<Extension | null> | null = null;
let pendingSubtitleMiningContext: SubtitleMiningContext | null = null;

ipcMain.handle(DICTIONARY_EXTERNAL_LINK_CHANNEL, async (event, value: unknown) => {
  if (
    activeDictionaryBackend !== 'hachidori' ||
    event.senderFrame !== event.sender.mainFrame ||
    !overlayManager.getOverlayWindows().some((window) => window.webContents === event.sender)
  ) {
    throw new Error('Dictionary links are only available from the active overlay');
  }
  await shell.openExternal(parseDictionaryExternalUrl(value));
});

function getOverlayForegroundSeparateWindows(): BrowserWindow[] {
  return [
    appState.configSettingsWindow,
    appState.syncUiWindow,
    appState.yomitanSettingsWindow,
    hachidoriSettingsWindow,
    inactiveYomitanSettingsWindow,
    appState.anilistSetupWindow,
    appState.jellyfinSetupWindow,
    appState.firstRunSetupWindow,
  ].filter((window): window is BrowserWindow => Boolean(window));
}

notifyAnilistTokenStoreWarning = (message: string) => {
  logger.warn(`[AniList] ${message}`);
  try {
    showDesktopNotification('SubMiner AniList', {
      body: message,
    });
  } catch {
    // Notification may fail if desktop notifications are unavailable early in startup.
  }
};
const appLogger = {
  logInfo: (message: string) => {
    logger.info(message);
  },
  logDebug: (message: string) => {
    logger.debug(message);
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
const reportFatalError = createFatalErrorReporter({
  showErrorBox: (title, details) => dialog.showErrorBox(title, details),
  consoleError: (message, error) => logger.error(message, error),
});

let forceQuitTimer: ReturnType<typeof setTimeout> | null = null;
const statsDistPath = path.join(__dirname, '..', 'stats', 'dist');
// Release builds stage a project TMDB key next to the compiled main process.
const bundledTmdbApiKey = readBundledTmdbApiKey(__dirname);
const statsPreloadPath = path.join(__dirname, 'preload-stats.js');
const statsServerRuntime = createStatsServerRuntime({
  userDataPath: USER_DATA_PATH,
  statsDistPath,
  getResolvedConfig: () => configService.getConfig(),
  getImmersionTracker: () => appState.immersionTracker,
  setAppStateStatsServer: (server) => {
    appState.statsServer = server;
  },
  getMpvSocketPath: () => appState.mpvSocketPath,
  getYomitanExt: () => appState.yomitanExt,
  getYomitanSession: () => appState.yomitanSession,
  getYomitanParserWindow: () => appState.yomitanParserWindow,
  setYomitanParserWindow: (w) => {
    appState.yomitanParserWindow = w;
  },
  getYomitanParserReadyPromise: () => appState.yomitanParserReadyPromise,
  setYomitanParserReadyPromise: (p) => {
    appState.yomitanParserReadyPromise = p;
  },
  getYomitanParserInitPromise: () => appState.yomitanParserInitPromise,
  setYomitanParserInitPromise: (p) => {
    appState.yomitanParserInitPromise = p;
  },
  getYomitanAnkiDeckName: () => getCurrentYomitanAnkiDeckNameForRuntime(),
  getAnilistRateLimiter: () => anilistRateLimiter,
  getBundledTmdbApiKey: () => bundledTmdbApiKey,
  resolveAnkiNoteId: (noteId) => appState.ankiIntegration?.resolveCurrentNoteId(noteId) ?? noteId,
  trackDuplicateNoteIdsForNote: (noteId, duplicateNoteIds) => {
    appState.ankiIntegration?.trackDuplicateNoteIdsForNote(noteId, duplicateNoteIds);
  },
  resolveSentenceSearchHeadwords: (term) => resolveSentenceSearchHeadwords(term),
  ensureImmersionTrackerStarted: () => ensureImmersionTrackerStarted(),
  setStatsStartupInProgress: (inProgress) => {
    appState.statsStartupInProgress = inProgress;
  },
});
const {
  stopStatsServer,
  ensureStatsServerStarted,
  ensureBackgroundStatsServerStarted,
  stopBackgroundStatsServer,
} = statsServerRuntime;

function requestAppQuit(): void {
  destroyYomitanSettingsWindow(hachidoriSettingsWindow);
  destroyYomitanSettingsWindow(inactiveYomitanSettingsWindow);
  destroyYomitanSettingsWindow(appState.yomitanSettingsWindow);
  appState.yomitanSettingsWindow = null;
  destroyStatsWindow();
  void stopStatsServer().catch((error: unknown) => {
    logger.warn('Failed to stop stats server while quitting.', error);
  });
  if (!forceQuitTimer) {
    // Quit cleanup unloads plugins first, so allow their budget plus 2 s for the rest.
    forceQuitTimer = setTimeout(() => {
      logger.warn('App quit timed out; forcing process exit.');
      void createForceQuitHandler({
        destroyImmersionTracker: () => appState.immersionTracker?.destroy(),
        logError: (error) => logger.error('Failed to finalize stats before forced exit.', error),
        exit: () => app.exit(0),
      })();
    }, PLUGIN_DEACTIVATE_TIMEOUT_MS + 2000);
  }
  app.quit();
}

process.on('SIGINT', () => {
  requestAppQuit();
});
process.on('SIGTERM', () => {
  requestAppQuit();
});

// Managed Whisper models live next to the config file.
const getSubtitleGenerationModelDirectory = (): string =>
  path.join(path.dirname(configService.getConfigPath()), 'models', 'whisper');
const youtube = createYoutubeAssembly({
  appState,
  getConfig: () => configService.getConfig(),
  logger,
  getWhisperModelDirectory: getSubtitleGenerationModelDirectory,
  getMainOverlayWindow: () => overlayManager.getMainWindow(),
  overlayModalRuntime,
  getOverlayNotifications: () => overlayNotificationsRuntime,
  getOverlayGeometry: () => overlayGeometryRuntime,
  showMpvOsd: (text) => showMpvOsd(text),
  showDesktopNotification: (title, options) => showDesktopNotification(title, options),
  refreshCurrentSubtitle: (text: string) => {
    subtitleProcessingController.refreshCurrentSubtitle(text);
  },
  refreshSubtitleSidebarSource: async (sourcePath: string, mediaPath?: string) => {
    await subtitlePrefetchRuntime.refreshSubtitleSidebarFromSource(sourcePath, mediaPath);
  },
  startTokenizationWarmups: async () => {
    await startTokenizationWarmups();
  },
  waitForMediaTokenizationReady: (mediaPath) =>
    currentMediaTokenizationGate.waitUntilReady(mediaPath),
  generateYoutubeSubtitles: (request) =>
    subtitleGenerationRuntime.generateYoutubeSubtitles(request),
  openSubtitleGenerationModal: () =>
    openSubtitleGenerationModal(overlayModals.createOverlayHostedModalOpenDeps()),
  invalidatePendingAutoplayReadyFallbacks: () =>
    autoplayReadyGate.invalidatePendingAutoplayReadyFallbacks(),
  ensurePlaybackRuntimeReady: () => ensureYoutubePlaybackRuntimeReady(),
  launchWindowsMpv: (playbackUrl, args) => {
    const config = configService.getConfig();
    setLogFileToggles(config.logging.files);
    const mpvLogPath = isLogFileEnabled('mpv') ? resolveDefaultLogFilePath('mpv') : '';
    if (mpvLogPath) {
      pruneLogDirectoryForPath(mpvLogPath, config.logging.rotation);
    }
    const mpvArgs = [...args, ...buildMpvLoggingArgs(config.logging.level, mpvLogPath, args)];
    return launchWindowsMpv(
      [playbackUrl],
      createWindowsMpvLaunchDeps({
        showError: (title, content) => dialog.showErrorBox(title, content),
      }),
      mpvArgs,
      process.execPath,
      resolveBundledMpvRuntimePluginEntrypoint(),
      config.mpv.executablePath,
      config.mpv.launchMode,
      {
        detectInstalledMpvPlugin: detectWindowsInstalledMpvPlugin,
        notifyInstalledPluginDetected: logInstalledMpvPluginDetected,
        resolveInstalledPluginBeforeLaunch: (detection, mpvPath) =>
          promptForLegacyMpvPluginRemovalBeforeWindowsLaunch(mpvPath, detection),
      },
      getMpvPluginRuntimeConfig(),
    );
  },
  ensureMpvConnectedForPlayback: () => jellyfin.ensureMpvConnectedForPlayback(),
  requestAppQuit: () => requestAppQuit(),
  browserWindow: {
    createBrowserWindow: (options) => new BrowserWindow(options),
    preloadPath: path.join(__dirname, 'preload-youtube-browser.js'),
    ipcMain,
    showContextMenu: (window, template) => {
      Menu.buildFromTemplate(template).popup({ window });
    },
    openExternal: (url) => {
      void shell.openExternal(url);
    },
  },
});
const POST_WARM_AUTOPLAY_SUBTITLE_PRIME_DELAYS_MS = [0, 75, 200, 500, 1_000] as const;

function schedulePostWarmAutoplaySubtitlePrime(signal: AutoplayReadySignal): void {
  if (signal.payload.text.trim() !== '__warm__') {
    return;
  }

  const mediaPath = signal.mediaPath;
  for (const delayMs of POST_WARM_AUTOPLAY_SUBTITLE_PRIME_DELAYS_MS) {
    const timer = setTimeout(() => {
      const currentMediaPath =
        appState.currentMediaPath?.trim() || appState.mpvClient?.currentVideoPath?.trim() || null;
      if (currentMediaPath !== mediaPath || !overlayManager.getVisibleOverlayVisible()) {
        return;
      }
      void autoplaySubtitlePrimingRuntime
        .primeCurrentSubtitleForAutoplay(mediaPath)
        .catch((error) => {
          logger.debug(
            `[autoplay-subtitle-prime] failed to prime current subtitle after warm readiness: ${
              error instanceof Error ? error.message : String(error)
            }`,
          );
        });
    }, delayMs);
    timer.unref?.();
  }
}

const autoplayReadyGate = createAutoplayReadyGate({
  isAppOwnedFlowInFlight: () => youtube.primarySubtitleNotification.isAppOwnedFlowInFlight(),
  getCurrentMediaPath: () => appState.currentMediaPath,
  getCurrentVideoPath: () => appState.mpvClient?.currentVideoPath ?? null,
  getPlaybackPaused: () => appState.playbackPaused,
  getMpvClient: () => appState.mpvClient,
  signalPluginAutoplayReady: () => {
    sendMpvCommandRuntime(appState.mpvClient, ['script-message', 'subminer-autoplay-ready']);
  },
  // Deferred: isTokenizationWarmupReady is assigned during composeMpvRuntimeHandlers below.
  isTokenizationReady: () => isTokenizationWarmupReady(),
  requestOverlayPointerRecovery: () => {
    if (process.platform !== 'darwin' || !overlayManager.getVisibleOverlayVisible()) {
      return;
    }
    // Renderer-side recovery alone cannot wake subtitle hover on macOS: the overlay only
    // receives mouse-moved events while it is the key window of the frontmost app, and mpv
    // is frontmost once playback starts. Activate the overlay window first so the broadcast's
    // setIgnoreMouseEvents toggling actually takes effect, mirroring a manual subtitle click.
    focusMacOSOverlayWindow({
      platform: process.platform,
      getOverlayWindow: () => overlayManager.getMainWindow(),
      stealAppFocus: () => app.focus({ steal: true }),
      warn: (message, details) => logger.warn(message, details),
    });
    overlayManager.broadcastToOverlayWindows(IPC_CHANNELS.event.overlayPointerRecoveryRequest);
  },
  onAutoplayReadyReleased: (signal) => {
    schedulePostWarmAutoplaySubtitlePrime(signal);
  },
  isSignalTargetReady: (signal) =>
    isVisibleOverlayAutoplayTargetReady(
      {
        getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
        isOverlayWindowReady: () => {
          const overlayWindow = overlayManager.getMainWindow();
          return Boolean(overlayWindow && isOverlayWindowContentReady(overlayWindow));
        },
        getLatestVisibleMeasurement: () =>
          overlayContentMeasurementStore.getLatestByLayer('visible'),
      },
      signal,
    ),
  schedule: (callback, delayMs) => setTimeout(callback, delayMs),
  logDebug: (message) => logger.debug(message),
});
const managedLocalSubtitleSelectionRuntime = createManagedLocalSubtitleSelectionRuntime({
  getCurrentMediaPath: () => appState.currentMediaPath,
  getMpvClient: () => appState.mpvClient,
  getPrimarySubtitleLanguages: () => configService.getConfig().youtube.primarySubLanguages,
  getSecondarySubtitleLanguages: () => configService.getConfig().secondarySub.secondarySubLanguages,
  sendMpvCommand: (command) => {
    sendMpvCommandRuntime(appState.mpvClient, command);
  },
  schedule: (callback, delayMs) => setTimeout(callback, delayMs),
  clearScheduled: (timer) => clearTimeout(timer),
});

const {
  resolveBundledMpvRuntimePluginEntrypoint,
  detectWindowsInstalledMpvPlugin,
  logInstalledMpvPluginDetected,
  promptForLegacyMpvPluginRemovalBeforeWindowsLaunch,
} = createWindowsMpvPluginDetectionRuntime({
  mainDirname: __dirname,
  logWarn: (message) => logger.warn(message),
});

function getMpvPluginRuntimeConfig() {
  const config = configService.getConfig();
  return {
    socketPath: appState.mpvSocketPath,
    binaryPath: config.mpv.subminerBinaryPath,
    backend: config.mpv.backend,
    autoStart: config.mpv.autoStartSubMiner,
    autoStartVisibleOverlay: config.auto_start_overlay,
    autoStartPauseUntilReady: config.mpv.pauseUntilOverlayReady,
    osdMessages:
      config.ankiConnect.behavior.notificationType === 'osd' ||
      config.ankiConnect.behavior.notificationType === 'osd-system',
    texthookerEnabled: config.texthooker.launchAtStartup,
  };
}

// Finder/Dock launches inherit launchd's minimal PATH; pick up the user's shell PATH so
// launcher/Bun detection and spawned tools match what their terminal sees.
const loginShellPathReady: Promise<void> =
  process.platform === 'darwin'
    ? applyLoginShellPath({ env: process.env })
        .then((applied) => {
          if (!applied) logger.warn('Login shell PATH was empty; using inherited PATH');
        })
        .catch((error) => {
          logger.warn('Failed to read login shell PATH; using inherited PATH', error);
        })
    : Promise.resolve();
const createCommandLineLauncherRuntimeOptions = async () => {
  await loginShellPathReady;
  return {
    platform: process.platform,
    env: process.env,
    homeDir: os.homedir(),
    localAppData: process.env.LOCALAPPDATA,
    userProfile: process.env.USERPROFILE,
    cwd: process.cwd(),
    resourcesPath: process.resourcesPath,
    appExePath: process.execPath,
    appVersion: app.getVersion(),
    bundledBunPath: app.isPackaged
      ? path.join(process.resourcesPath, 'bun', process.platform === 'win32' ? 'bun.exe' : 'bun')
      : undefined,
  };
};
const firstRunSetup = createFirstRunSetupAssembly({
  appState,
  configService,
  logger,
  configDir: CONFIG_DIR,
  userDataPath: USER_DATA_PATH,
  dictionaryBackend: activeDictionaryBackend,
  ensureYomitanExtensionLoaded: () => ensureYomitanExtensionLoaded(),
  getYomitanParserRuntimeDeps: () => getYomitanParserRuntimeDeps(),
  createCommandLineLauncherRuntimeOptions: () => createCommandLineLauncherRuntimeOptions(),
  getAppPath: (name) => app.getPath(name),
  trashItem: (targetPath) => shell.trashItem(targetPath),
  writeShortcutLink: (shortcutPath, operation, details) =>
    shell.writeShortcutLink(shortcutPath, operation, details),
  createBrowserWindow: (options) => new BrowserWindow(options),
  openDictionarySettings: () => openDictionarySettings(),
  openConfigSettingsWindow: () => settings.window.openWindow(),
  onSetupStateChanged: () => {
    if (appTray) {
      ensureTrayHandler();
    }
  },
  requestAppQuit: () => requestAppQuit(),
});
const discordPresenceSessionStartedAtMs = Date.now();
let discordPresenceMediaDurationSec: number | null = null;
const discordPresenceRuntime = createDiscordPresenceRuntime({
  getDiscordPresenceService: () => appState.discordPresenceService,
  isDiscordPresenceEnabled: () => configService.getConfig().discordPresence.enabled === true,
  getMpvClient: () => appState.mpvClient,
  getCurrentMediaTitle: () => appState.currentMediaTitle,
  getCurrentMediaPath: () => appState.currentMediaPath,
  getCurrentSubtitleText: () => appState.currentSubText,
  getPlaybackPaused: () => appState.playbackPaused,
  getFallbackMediaDurationSec: () => anilist.getMediaDurationSec(),
  getSessionStartedAtMs: () => discordPresenceSessionStartedAtMs,
  getMediaDurationSec: () => discordPresenceMediaDurationSec,
  setMediaDurationSec: (next) => {
    discordPresenceMediaDurationSec = next;
  },
});

async function initializeDiscordPresenceService(): Promise<void> {
  if (configService.getConfig().discordPresence.enabled !== true) {
    appState.discordPresenceService = null;
    return;
  }

  appState.discordPresenceService = createDiscordPresenceService({
    config: configService.getConfig().discordPresence,
    createClient: () => createDiscordRpcClient(DISCORD_PRESENCE_APP_ID),
    logDebug: (message, meta) => logger.debug(message, meta),
  });
  await appState.discordPresenceService.start();
  discordPresenceRuntime.publishDiscordPresence();
}
const ensureOverlayMpvSubtitlesHidden = createEnsureOverlayMpvSubtitlesHiddenHandler({
  getMpvClient: () => appState.mpvClient,
  getSavedSubVisibility: () => appState.overlaySavedMpvSubVisibility,
  setSavedSubVisibility: (visible) => {
    appState.overlaySavedMpvSubVisibility = visible;
  },
  getRevision: () => appState.overlayMpvSubVisibilityRevision,
  setRevision: (revision) => {
    appState.overlayMpvSubVisibilityRevision = revision;
  },
  setMpvSubVisibility: (visible) => {
    setMpvSubVisibilityRuntime(appState.mpvClient, visible);
  },
  logWarn: (message, error) => {
    logger.warn(message, error);
  },
});
const restoreOverlayMpvSubtitles = createRestoreOverlayMpvSubtitlesHandler({
  getSavedSubVisibility: () => appState.overlaySavedMpvSubVisibility,
  setSavedSubVisibility: (visible) => {
    appState.overlaySavedMpvSubVisibility = visible;
  },
  getRevision: () => appState.overlayMpvSubVisibilityRevision,
  setRevision: (revision) => {
    appState.overlayMpvSubVisibilityRevision = revision;
  },
  isMpvConnected: () => Boolean(appState.mpvClient?.connected),
  shouldKeepSuppressedFromVisibleOverlayBinding: () => overlayManager.getVisibleOverlayVisible(),
  setMpvSubVisibility: (visible) => {
    setMpvSubVisibilityRuntime(appState.mpvClient, visible);
  },
});

function syncOverlayMpvSubtitleSuppression(): void {
  if (overlayManager.getVisibleOverlayVisible()) {
    void ensureOverlayMpvSubtitlesHidden();
    return;
  }

  restoreOverlayMpvSubtitles();
}

const immersionMediaRuntime = createImmersionMediaRuntime({
  getResolvedConfig: () => configService.getConfig(),
  defaultImmersionDbPath: DEFAULT_IMMERSION_DB_PATH,
  getTracker: () => appState.immersionTracker,
  getMpvClient: () => appState.mpvClient,
  getCurrentMediaPath: () => appState.currentMediaPath,
  getCurrentMediaTitle: () => appState.currentMediaTitle,
  logDebug: (message) => logger.debug(message),
  logInfo: (message) => logger.info(message),
});
const anilistRateLimiter = createAnilistRateLimiter();
const statsCoverArtFetcher = createCoverArtFetcher(
  anilistRateLimiter,
  createLogger('main:stats-cover-art'),
  {
    liveAction: createLiveActionMetadataResolver(
      createTmdbClient({
        resolveApiKey: createTmdbApiKeyResolver(
          () => configService.getConfig().tmdb,
          () => bundledTmdbApiKey,
        ),
      }),
      createLogger('main:tmdb'),
    ),
  },
);
const configDerivedRuntime = createConfigDerivedRuntime({
  getResolvedConfig: () => configService.getConfig(),
  getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
  defaultJimakuLanguagePreference: DEFAULT_CONFIG.jimaku.languagePreference,
  defaultJimakuMaxEntryResults: DEFAULT_CONFIG.jimaku.maxEntryResults,
  defaultJimakuApiBaseUrl: DEFAULT_CONFIG.jimaku.apiBaseUrl,
});
const subsyncRuntime = createMainSubsyncRuntime({
  getMpvClient: () => appState.mpvClient,
  getResolvedConfig: () => configService.getConfig(),
  getSubsyncInProgress: () => appState.subsyncInProgress,
  setSubsyncInProgress: (inProgress) => {
    appState.subsyncInProgress = inProgress;
  },
  showMpvOsd: (text) => overlayNotificationsRuntime.showSubsyncStatusNotification(text),
  openManualPicker: (payload) => {
    overlayModals.openOverlayHostedModalWithOsd(
      (deps) => openSubsyncManualModalRuntime(deps, payload),
      'Subsync overlay unavailable.',
      'Failed to open subsync overlay.',
    );
  },
});
const currentMediaTokenizationGate = createCurrentMediaTokenizationGate();
const startupOsdSequencer = createStartupOsdSequencer({
  getNotificationType: () => overlayNotificationsRuntime.getConfiguredStatusNotificationType(),
  showOsd: (message) => showMpvOsd(message),
  showOverlayNotification: (payload) =>
    overlayNotificationsRuntime.showOverlayNotification(payload),
  showDesktopNotification: (title, options) => showDesktopNotification(title, options),
});
let appTray: Tray | null = null;
let tokenizeSubtitleDeferred: ((text: string) => Promise<SubtitleData>) | null = null;
function withCurrentSubtitleTiming(payload: SubtitleData): SubtitleData {
  const canonical = resolveCanonicalPrimarySubtitle({
    liveText: payload.text,
    currentTimeSec: toSubtitleClock(
      Number(appState.mpvClient?.currentTimePos),
      appState.mpvClient?.currentSubDelay ?? 0,
    ),
    cues: appState.activeParsedSubtitleCues,
  });
  return {
    ...payload,
    startTime: canonical?.startTime ?? appState.mpvClient?.currentSubStart ?? null,
    endTime: canonical?.endTime ?? appState.mpvClient?.currentSubEnd ?? null,
  };
}

function captureCurrentPrimarySubtitleMiningContext(): SubtitleMiningContext | null {
  // Mine what the overlay shows, not raw mpv `sub-text`: the raw text lists every active
  // event, so a finished caption row lingering beside a fresh line would end up on the
  // card. The parsed view also carries the cue's own timings for the clip range.
  // Cues sit on the subtitle clock, so the lookup undoes `sub-delay` first.
  const subDelay = appState.mpvClient?.currentSubDelay ?? 0;
  const resolved = resolvePrimarySubtitle({
    liveText: appState.mpvClient?.currentSubText ?? '',
    currentTimeSec: toSubtitleClock(Number(appState.mpvClient?.currentTimePos), subDelay),
    cues: appState.activeParsedSubtitleCues,
  });
  // Same validity bar as the live capture path: an unusable resolved span must fall
  // back rather than hand mining an empty line or an inverted range.
  const resolvedText = resolved?.text.replace(/\n{2,}/g, '\n').trim();
  if (
    !resolved ||
    !resolvedText ||
    !Number.isFinite(resolved.startTime) ||
    !Number.isFinite(resolved.endTime) ||
    resolved.endTime <= resolved.startTime
  ) {
    return captureLiveSubtitleMiningContext(appState.mpvClient);
  }
  return shiftSubtitleRangeToVideoClock<SubtitleMiningContext>(
    {
      source: 'overlay',
      text: resolvedText,
      startTime: resolved.startTime,
      endTime: resolved.endTime,
      capturedAtMs: Date.now(),
    },
    subDelay,
  );
}
function emitSubtitlePayload(payload: SubtitleData, options?: { resumePrefetch?: boolean }): void {
  const timedPayload = withCurrentSubtitleTiming(payload);
  const currentSubtitleData = appState.currentSubtitleData;
  const isAnnotationUpgrade = isSubtitleAnnotationUpgrade(currentSubtitleData, timedPayload);
  const frequencyDictionary = configService.getConfig().subtitleStyle.frequencyDictionary;
  const frequencyOptions = {
    enabled: frequencyDictionary.enabled,
    topX: frequencyDictionary.topX,
    mode: frequencyDictionary.mode,
  };
  appState.currentSubtitleData = timedPayload;
  overlayManager.broadcastToOverlayWindows('subtitle:set', timedPayload);
  if (!isAnnotationUpgrade) {
    subtitleWsService.broadcast(timedPayload, frequencyOptions);
  }
  annotationSubtitleWsService.broadcast(timedPayload, frequencyOptions);
  autoplayReadyGate.maybeSignalPluginAutoplayReady(timedPayload, { forceWhilePaused: true });
  // resumePrefetch: false marks an emit that is not the end of the work for
  // this line; prefetch stays paused until the subtitle processing controller
  // settles so it does not compete with the on-screen line for the single
  // Yomitan parser window.
  if (options?.resumePrefetch !== false) {
    subtitlePrefetchService?.resume();
  }
}
function getCurrentAutoplaySubtitlePayload(): SubtitleData | null {
  const payload = appState.currentSubtitleData;
  if (!payload?.text.trim()) {
    return null;
  }
  if (payload.text !== appState.currentSubText) {
    return null;
  }
  return payload;
}

async function resolveSentenceSearchHeadwords(term: string): Promise<string[]> {
  const fallback = term.trim() ? [term.trim()] : [];
  try {
    const tokenized = tokenizeSubtitleDeferred ? await tokenizeSubtitleDeferred(term) : null;
    const tokens = tokenized?.tokens ?? [];
    if (tokens.length === 0) return fallback;

    const seen = new Set<string>();
    const headwords: string[] = [];
    for (const token of tokens) {
      const headword = (token.headword || token.surface).trim();
      if (!headword || seen.has(headword)) continue;
      seen.add(headword);
      headwords.push(headword);
    }
    return headwords.length > 0 ? headwords : fallback;
  } catch (error) {
    logger.debug(
      'Failed to resolve sentence-search headwords:',
      error instanceof Error ? error.message : String(error),
    );
    return fallback;
  }
}

function signalCurrentSubtitleAutoplayReady(): void {
  autoplayReadyGate.flushPendingAutoplayReadySignal();
  const payload = getCurrentAutoplaySubtitlePayload();
  if (payload) {
    autoplayReadyGate.maybeSignalPluginAutoplayReady(payload, { forceWhilePaused: true });
    return;
  }
  if (!appState.currentSubText.trim()) {
    autoplayReadyGate.maybeSignalPluginAutoplayReady(
      { text: '__warm__', tokens: null },
      { forceWhilePaused: true },
    );
  }
}
const subtitleProcessingController = createPrefetchAwareSubtitleProcessingController({
  tokenizeSubtitle: async (text: string) =>
    tokenizeSubtitleDeferred ? await tokenizeSubtitleDeferred(text) : { text, tokens: null },
  emitSubtitlePayload: (payload, options) => emitSubtitlePayload(payload, options),
  resumePrefetch: () => {
    subtitlePrefetchService?.resume();
  },
  logDebug: (message) => {
    logger.debug(`[subtitle-processing] ${message}`);
  },
  now: () => Date.now(),
});

let subtitlePrefetchService: SubtitlePrefetchService | null = null;
let lastObservedTimePos = 0;
let lastObservedPrimarySubtitleTrackId: number | null = null;
let cancelLinuxMpvFullscreenOverlayRefreshBurst: CancelLinuxMpvFullscreenOverlayRefreshBurst | null =
  null;
const linuxOverlayModeRuntime = createLinuxOverlayModeRuntime({
  isEnabled: shouldRunLinuxOverlayZOrderKeepAlive,
  isVisible: () => overlayManager.getVisibleOverlayVisible(),
  getWindow: () => overlayManager.getMainWindow(),
  clearWindow: () => overlayManager.setMainWindow(null),
  createWindow: () => {
    visibleOverlayInteractionRuntime.resetVisibleOverlayInputState();
    createMainWindow();
  },
  refreshWindow: () => {
    const trackedGeometry = overlayGeometryRuntime.getCurrentTrackedOverlayGeometry();
    if (trackedGeometry) overlayManager.setOverlayWindowBounds(trackedGeometry);
    overlayVisibilityRuntime.updateVisibleOverlayVisibility();
    void ensureOverlayMpvSubtitlesHidden();
    if (appState.currentSubText.trim()) {
      subtitleProcessingController.refreshCurrentSubtitle(appState.currentSubText);
    }
  },
  now: Date.now,
  logDebug: (message) => logger.debug(message),
});
let subtitleSidebarRequestedOpen = false;
let subtitleSidebarCuesRef: SubtitleCue[] | null = null;
let subtitleSidebarCuesRevision = 0;

/**
 * Cheap sidebar state polled ~12x/s while the sidebar is open. Reads only cached state, no
 * mpv round-trips. `cuesRevision` bumps whenever a new parsed cue array is swapped in, so a
 * re-parse of the same source still tells the renderer to fetch the full snapshot.
 */
function getSubtitleSidebarPlaybackState(): SubtitleSidebarPlaybackState {
  if (appState.activeParsedSubtitleCues !== subtitleSidebarCuesRef) {
    subtitleSidebarCuesRef = appState.activeParsedSubtitleCues;
    subtitleSidebarCuesRevision += 1;
  }
  const subDelay = appState.mpvClient?.currentSubDelay ?? 0;
  return {
    sourceKey: JSON.stringify([
      appState.activeParsedSubtitleMediaPath,
      appState.activeParsedSubtitleSource,
    ]),
    cuesRevision: subtitleSidebarCuesRevision,
    currentTimeSec:
      appState.mpvClient?.currentTimePos === undefined
        ? null
        : toSubtitleClock(appState.mpvClient.currentTimePos, subDelay),
    subDelay,
    currentSubtitle: {
      text: appState.currentSubText,
      startTime: appState.mpvClient?.currentSubStart ?? null,
      endTime: appState.mpvClient?.currentSubEnd ?? null,
    },
  };
}
const SEEK_THRESHOLD_SECONDS = 3;
const EXPLICIT_SEEK_INTENT_TTL_MS = 2000;
let explicitSeekIntentExpiresAtMs = 0;

function isExplicitMpvSeekCommand(command: readonly (string | number)[]): boolean {
  return command[0] === 'seek' || command[0] === 'sub-seek';
}

function getMpvPauseCommandValue(command: readonly (string | number)[]): 'yes' | 'no' | null {
  if ((command[0] !== 'set_property' && command[0] !== 'set') || command[1] !== 'pause') {
    return null;
  }
  return command[2] === 'yes' || command[2] === 'no' ? command[2] : null;
}

function sendRendererMpvCommand(rawCommand: (string | number)[]): void {
  // Overlay auto-pause releases (popup closed, hover left) must not resume playback
  // behind an open timing review; the review applies them when it closes. A pause
  // request during the review keeps playback paused after it closes.
  const pauseValue = getMpvPauseCommandValue(rawCommand);
  if (pauseValue === 'no' && mediaTimingReviewRuntime.deferPlaybackResume()) {
    return;
  }
  if (pauseValue === 'yes') {
    mediaTimingReviewRuntime.cancelPlaybackResume();
  }
  const command =
    resolveSanitizedSubtitleSeekCommand(
      rawCommand,
      appState.activeParsedSubtitleCues,
      appState.mpvClient?.currentTimePos ?? Number.NaN,
      appState.mpvClient?.currentSubDelay ?? 0,
    ) ?? rawCommand;
  if (isExplicitMpvSeekCommand(command)) {
    explicitSeekIntentExpiresAtMs = Date.now() + EXPLICIT_SEEK_INTENT_TTL_MS;
  }
  sendMpvCommandRuntime(appState.mpvClient, command);
}

function consumeExplicitSeekIntent(): boolean {
  const pending = explicitSeekIntentExpiresAtMs >= Date.now();
  explicitSeekIntentExpiresAtMs = 0;
  return pending;
}

const autoplaySubtitlePrimingRuntime = createAutoplaySubtitlePrimingRuntime({
  getCurrentMediaPath: () => appState.currentMediaPath,
  getMpvClient: () => appState.mpvClient,
  setCurrentSubText: (text) => {
    appState.currentSubText = text;
  },
  getCurrentSubText: () => appState.currentSubText,
  getCurrentSubtitleData: () => appState.currentSubtitleData,
  getActiveParsedSubtitleCues: () => appState.activeParsedSubtitleCues,
  setActiveParsedSubtitleMediaPath: (mediaPath) => {
    appState.activeParsedSubtitleMediaPath = mediaPath;
  },
  subtitleProcessingController,
  emitSubtitlePayload: (payload, options) => emitSubtitlePayload(payload, options),
  getSubtitlePrefetchService: () => subtitlePrefetchService,
  getLastObservedTimePos: () => lastObservedTimePos,
  getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
  emitSecondarySubtitle: (text) => {
    secondarySubtitleTrackController.handleLiveText(text);
  },
  initSubtitlePrefetch: (sourcePath, currentTimePos, sourceKey) =>
    subtitlePrefetchInitController.initSubtitlePrefetch(sourcePath, currentTimePos, sourceKey),
  refreshSubtitlePrefetchFromActiveTrack: () => refreshSubtitlePrefetchFromActiveTrackHandler(),
  logDebug: (message) => {
    logger.debug(message);
  },
});

function cancelPendingLinuxMpvFullscreenOverlayRefreshBurst(): void {
  cancelLinuxMpvFullscreenOverlayRefreshBurst?.();
  cancelLinuxMpvFullscreenOverlayRefreshBurst = null;
}

const subtitlePrefetchInitController = createSubtitlePrefetchInitController({
  getCurrentService: () => subtitlePrefetchService,
  setCurrentService: (service) => {
    subtitlePrefetchService = service;
  },
  loadSubtitleSourceText,
  parseSubtitleCues: (content, filename) => parseSubtitleCues(content, filename),
  createSubtitlePrefetchService: (deps) => createSubtitlePrefetchService(deps),
  tokenizeSubtitle: async (text) =>
    tokenizeSubtitleDeferred ? await tokenizeSubtitleDeferred(text) : null,
  preCacheTokenization: (text, data) => {
    subtitleProcessingController.preCacheTokenization(text, data);
  },
  hasCachedTokenization: (text) => subtitleProcessingController.hasCachedSubtitle(text),
  getCacheGeneration: () => subtitleProcessingController.getCacheGeneration(),
  logInfo: (message) => logger.info(message),
  logWarn: (message) => logger.warn(message),
  onParsedSubtitleCuesChanged: (cues, sourceKey) => {
    appState.activeParsedSubtitleCues = cues ?? [];
    appState.activeParsedSubtitleSource = sourceKey;
    if (!cues?.length) {
      appState.activeParsedSubtitleMediaPath = null;
    }
    const mediaPath = autoplaySubtitlePrimingRuntime.getCurrentAutoplayMediaPath();
    if (mediaPath && cues?.length) {
      void autoplaySubtitlePrimingRuntime
        .primeAutoplaySubtitleFromParsedCues(mediaPath, cues)
        .catch((error) => {
          logger.debug(
            `[autoplay-subtitle-prime] failed to prime from parsed cues: ${
              error instanceof Error ? error.message : String(error)
            }`,
          );
        });
    }
  },
});
const cachedInternalSubtitleTrackExtractor = createCachedInternalSubtitleTrackExtractor();
const detectRemoteMediaPath = createRemoteMediaPathDetector();
const resolveActiveSubtitleSidebarSourceHandler = createResolveActiveSubtitleSidebarSourceHandler({
  getFfmpegPath: () => configService.getConfig().subsync.ffmpeg_path.trim() || 'ffmpeg',
  extractInternalSubtitleTrack: (ffmpegPath, videoPath, track) =>
    cachedInternalSubtitleTrackExtractor.extract(ffmpegPath, videoPath, track),
  logDebug: (message) => logger.debug(message),
});

const secondarySubtitleTrackController = createSecondarySubtitleTrackController({
  getMpvClient: () => appState.mpvClient,
  getCurrentTimePos: () => appState.mpvClient?.currentTimePos ?? lastObservedTimePos,
  resolveSubtitleSource: (input) => resolveActiveSubtitleSidebarSourceHandler(input),
  loadSubtitleSourceText,
  parseSubtitleCues: (content, filename) => parseSubtitleCues(content, filename),
  setCurrentSecondaryText: (text) => {
    if (appState.mpvClient) {
      appState.mpvClient.currentSecondarySubText = text;
    }
  },
  broadcastSecondaryText: (text) => {
    overlayManager.broadcastToOverlayWindows('secondary-subtitle:set', text);
  },
  logDebug: (message) => logger.debug(message),
  logWarn: (message, error) => logger.warn(message, error),
});

const refreshSubtitlePrefetchFromActiveTrackHandler =
  createRefreshSubtitlePrefetchFromActiveTrackHandler({
    getMpvClient: () => appState.mpvClient,
    getLastObservedTimePos: () => lastObservedTimePos,
    // Remote media has no extractable on-disk track to fall back to, so a transient
    // resolve miss (sid briefly 'no', a cycle onto an embedded stream track) would
    // otherwise drop a working cue list for the rest of the episode.
    shouldKeepExistingCuesOnMissingSource: async (videoPath) =>
      isYoutubeMediaPath(videoPath) || (await detectRemoteMediaPath(videoPath)),
    subtitlePrefetchInitController,
    resolveActiveSubtitleSidebarSource: (input) => resolveActiveSubtitleSidebarSourceHandler(input),
    logDebug: (message) => logger.debug(message),
    logWarn: (message) => logger.warn(message),
  });

const subtitlePrefetchRuntime = {
  cancelPendingInit: () => subtitlePrefetchInitController.cancelPendingInit(),
  initSubtitlePrefetch: subtitlePrefetchInitController.initSubtitlePrefetch,
  refreshSubtitleSidebarFromSource: (sourcePath: string, mediaPath?: string) =>
    autoplaySubtitlePrimingRuntime.refreshSubtitleSidebarFromSource(sourcePath, mediaPath),
  refreshSubtitlePrefetchFromActiveTrack: () => refreshSubtitlePrefetchFromActiveTrackHandler(),
  scheduleSubtitlePrefetchRefresh: (delayMs?: number) =>
    autoplaySubtitlePrimingRuntime.scheduleSubtitlePrefetchRefresh(delayMs),
  clearScheduledSubtitlePrefetchRefresh: () =>
    autoplaySubtitlePrimingRuntime.clearScheduledSubtitlePrefetchRefresh(),
} as const;

const overlayShortcutsRuntime = createOverlayShortcutsRuntimeService({
  getConfiguredShortcuts: () => getConfiguredShortcuts(),
  getShortcutsRegistered: () => appState.shortcutsRegistered,
  setShortcutsRegistered: (registered: boolean) => {
    appState.shortcutsRegistered = registered;
  },
  isOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
  isOverlayShortcutContextActive: () => {
    if (process.platform !== 'win32') {
      return true;
    }

    if (!overlayManager.getVisibleOverlayVisible()) {
      return false;
    }

    const windowTracker = appState.windowTracker;
    if (!windowTracker || !windowTracker.isTracking()) {
      return false;
    }

    return windowTracker.isTargetWindowFocused();
  },
  showMpvOsd: (text: string) => overlayNotificationsRuntime.showConfiguredStatusNotification(text),
  openRuntimeOptionsPalette: () => {
    overlayModals.openRuntimeOptionsPalette();
  },
  openCharacterDictionaryManager: () => {
    overlayModals.openCharacterDictionaryManagerOverlay();
  },
  openJimaku: () => {
    overlayModals.openJimakuOverlay();
  },
  openTsukihime: () => {
    overlayModals.openTsukihimeOverlay();
  },
  markAudioCard: () => markLastCardAsAudioCardHandler(),
  copySubtitleMultiple: (timeoutMs: number) => {
    startPendingMultiCopy(timeoutMs);
  },
  copySubtitle: () => {
    copyCurrentSubtitleHandler();
  },
  toggleSecondarySubMode: () => cycleSecondarySubMode(),
  updateLastCardFromClipboard: () => updateLastCardFromClipboardHandler(),
  triggerFieldGrouping: () => triggerFieldGroupingHandler(),
  triggerSubsyncFromConfig: () => subsyncRuntime.triggerFromConfig(),
  mineSentenceCard: () => mineSentenceCardHandler(),
  mineSentenceMultiple: (timeoutMs: number) => {
    startPendingMineSentenceMultiple(timeoutMs);
  },
  cancelPendingMultiCopy: () => {
    cancelPendingMultiCopy();
  },
  cancelPendingMineSentenceMultiple: () => {
    cancelPendingMineSentenceMultiple();
  },
});
syncOverlayShortcutsForModal = (isActive: boolean): void => {
  if (isActive) {
    overlayShortcutsRuntime.unregisterOverlayShortcuts();
  } else {
    overlayShortcutsRuntime.syncOverlayShortcuts();
  }
};

async function getCurrentYomitanAnkiDeckNameForRuntime(): Promise<string> {
  await yomitanExtensionRuntime.ensureYomitanExtensionLoaded();
  return getYomitanCurrentAnkiDeckNameCore(getYomitanParserRuntimeDeps(), {
    error: (message, ...args) => {
      logger.error(message, ...args);
    },
    info: (message, ...args) => {
      logger.info(message, ...args);
    },
  });
}

const settings = createSettingsAssembly({
  appState,
  configService,
  logger,
  getOverlayNotifications: () => overlayNotificationsRuntime,
  showMpvOsd: (message) => showMpvOsd(message),
  showDesktopNotification: (title, options) => showDesktopNotification(title, options),
  hotReloadApplied: {
    setKeybindings: (keybindings) => {
      appState.keybindings = keybindings;
    },
    setSessionBindings: (sessionBindings, sessionBindingWarnings) => {
      persistSessionBindings(sessionBindings, sessionBindingWarnings);
    },
    refreshGlobalAndOverlayShortcuts: () => {
      refreshGlobalAndOverlayShortcuts();
    },
    setSecondarySubMode: (mode) => {
      setSecondarySubMode(mode);
    },
    broadcastToOverlayWindows: (channel, payload) => {
      overlayManager.broadcastToOverlayWindows(channel, payload);
    },
    applyAnkiRuntimeConfigPatch: (patch) => {
      if (appState.ankiIntegration) {
        appState.ankiIntegration.applyRuntimeConfigPatch(patch);
      }
      if (activeDictionaryBackend === 'hachidori' && appState.yomitanExt) {
        void syncYomitanDefaultProfileAnkiServer().catch((error: unknown) =>
          logger.error('Failed to auto-populate Hachidori Anki settings', error),
        );
      }
    },
    invalidateTokenizationCache: () => {
      subtitleProcessingController.invalidateTokenizationCache();
    },
    refreshSubtitlePrefetch: () => {
      subtitlePrefetchService?.onSeek(lastObservedTimePos);
    },
    refreshCurrentSubtitle: () => {
      subtitleProcessingController.refreshCurrentSubtitle(appState.currentSubText);
    },
    setLogLevel: (level) => {
      setLogLevel(level, 'config');
    },
    setLogRotation: (rotation) => {
      setLogRotation(rotation);
    },
    setLogFileToggles: (files) => {
      setLogFileToggles(files);
    },
    applyAniSkipConfig: () => {
      aniSkipRuntime.applyConfigChange();
    },
  },
  getYomitanAnkiDeckName: getCurrentYomitanAnkiDeckNameForRuntime,
  showErrorBox: (title, content) => dialog.showErrorBox(title, content),
  createBrowserWindow: (options) => new BrowserWindow(options),
  mainDirname: __dirname,
  openPath: (targetPath) => shell.openPath(targetPath),
  ipcMain,
});

settings.window.registerHandlers();

const sync = createSyncAssembly({
  appState,
  configService,
  logger,
  userDataPath: USER_DATA_PATH,
  defaultImmersionDbPath: DEFAULT_IMMERSION_DB_PATH,
  getOverlayNotifications: () => overlayNotificationsRuntime,
  requestAppQuit: () => requestAppQuit(),
  createBrowserWindow: (options) => new BrowserWindow(options),
  mainDirname: __dirname,
  ipcMain,
  isPackaged: () => app.isPackaged,
  getAppPath: () => app.getAppPath(),
  pickSnapshotFile: async () => {
    const result = await dialog.showOpenDialog({
      title: 'Choose a SubMiner stats snapshot',
      filters: [{ name: 'SQLite snapshots', extensions: ['sqlite'] }],
      properties: ['openFile'],
    });
    return result.canceled ? null : (result.filePaths[0] ?? null);
  },
  revealPath: (targetPath) => shell.showItemInFolder(targetPath),
});
sync.runtime.registerHandlers();
sync.autoScheduler.start();

const dictionaries = createDictionariesAssembly({
  roots: {
    platform: process.platform,
    dirname: __dirname,
    appPath: app.getAppPath(),
    resourcesPath: process.resourcesPath,
    userDataPath: USER_DATA_PATH,
    appUserDataPath: app.getPath('userData'),
    homeDir: os.homedir(),
    appDataDir: process.env.APPDATA,
    cwd: process.cwd(),
    joinPath: (...parts: string[]) => path.join(...parts),
  },
  appState,
  getConfig: () => configService.getConfig(),
  logger,
});

const getFieldGroupingResolverHandler = createGetFieldGroupingResolverHandler({
  getResolver: () => appState.fieldGroupingResolver,
});

const setFieldGroupingResolverHandler = createSetFieldGroupingResolverHandler({
  setResolver: (resolver) => {
    appState.fieldGroupingResolver = resolver;
  },
  nextSequence: () => {
    appState.fieldGroupingResolverSequence += 1;
    return appState.fieldGroupingResolverSequence;
  },
  getSequence: () => appState.fieldGroupingResolverSequence,
});

const fieldGroupingOverlayRuntime = createFieldGroupingOverlayRuntime<OverlayHostedModal>({
  getMainWindow: () => overlayManager.getMainWindow(),
  getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
  setVisibleOverlayVisible: (visible) => setVisibleOverlayVisible(visible),
  getResolver: () => getFieldGroupingResolverHandler(),
  setResolver: (resolver) => setFieldGroupingResolverHandler(resolver),
  getRestoreVisibleOverlayOnModalClose: () =>
    overlayModalRuntime.getRestoreVisibleOverlayOnModalClose(),
  // These are optional on the runtime options, so a missing forward compiles silently, but
  // dropping them leaves the field grouping modal with no modal-open ack/retry, no teardown on
  // failure, and no warn logging.
  waitForModalOpen: (modal, timeoutMs) => overlayModalRuntime.waitForModalOpen(modal, timeoutMs),
  handleOverlayModalClosed: (modal) => overlayModalRuntime.handleOverlayModalClosed(modal),
  logWarn: (message) => logger.warn(message),
  ensureOverlayStartupPrereqs: () => ensureOverlayStartupPrereqs(),
  ensureOverlayWindowsReadyForVisibilityActions: () =>
    ensureOverlayWindowsReadyForVisibilityActions(),
  sendToVisibleOverlay: (channel, payload, runtimeOptions) =>
    overlayModalRuntime.sendToActiveOverlayWindow(channel, payload, runtimeOptions),
});
const createFieldGroupingCallback = fieldGroupingOverlayRuntime.createFieldGroupingCallback;

const SUBTITLE_POSITIONS_DIR = path.join(CONFIG_DIR, 'subtitle-positions');

const mediaRuntime = createMediaRuntimeService({
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
  broadcastSubtitlePosition: (position) => {
    overlayManager.broadcastToOverlayWindows('subtitle-position:set', position);
  },
  getCurrentMediaTitle: () => appState.currentMediaTitle,
  setCurrentMediaTitle: (title) => {
    appState.currentMediaTitle = title;
  },
});

const characterDictionary = createCharacterDictionaryAssembly({
  userDataPath: USER_DATA_PATH,
  appState,
  getConfig: () => configService.getConfig(),
  logger,
  startupOsdSequencer,
  resolveMediaPathForJimaku: (mediaPath) => mediaRuntime.resolveMediaPathForJimaku(mediaPath),
  getYomitanProfilePolicy: () => yomitanProfilePolicy,
  isYoutubePlaybackActive: () => youtube.isPlaybackActiveNow(),
  waitForMediaTokenizationReady: (mediaPath) =>
    currentMediaTokenizationGate.waitUntilReady(mediaPath),
  ensureYomitanExtensionLoaded: () => ensureYomitanExtensionLoaded(),
  getYomitanParserRuntimeDeps: () => getYomitanParserRuntimeDeps(),
  getOverlayNotifications: () => overlayNotificationsRuntime,
  showMpvOsd: (message) => showMpvOsd(message),
  showDesktopNotification: (title, options) => showDesktopNotification(title, options),
  invalidateTokenizationCache: () => subtitleProcessingController.invalidateTokenizationCache(),
  refreshCurrentSubtitle: (text) => subtitleProcessingController.refreshCurrentSubtitle(text),
  refreshSubtitlePrefetch: () => {
    subtitlePrefetchService?.onSeek(lastObservedTimePos);
  },
  refreshCurrentSubtitleAnnotations: () => refreshCurrentSubtitleAnnotations(),
});

const overlayVisibilityRuntime = createOverlayVisibilityRuntimeService({
  getMainWindow: () => overlayManager.getMainWindow(),
  getModalActive: () => overlayModalInputState.getModalInputExclusive(),
  getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
  getForceMousePassthrough: () => appState.statsOverlayVisible,
  getNonNativeInputRegionActive: () =>
    process.platform === 'linux' &&
    visibleOverlayInteractionRuntime.getLinuxOverlayInputShapeActive(),
  getSuspendVisibleOverlay: () => appState.statsOverlayVisible,
  getOverlayInteractionActive: () =>
    visibleOverlayInteractionRuntime.getVisibleOverlayInteractionActive(),
  getWindowTracker: () => appState.windowTracker,
  getMacOSForegroundProbeActive: () =>
    visibleOverlayInteractionRuntime.getMacOSVisibleOverlayForegroundProbeActive(),
  getTrackerNotReadyWarningShown: () => appState.trackerNotReadyWarningShown,
  setTrackerNotReadyWarningShown: (shown: boolean) => {
    appState.trackerNotReadyWarningShown = shown;
  },
  updateVisibleOverlayBounds: (geometry: WindowGeometry) =>
    overlayGeometryRuntime.updateVisibleOverlayBounds(geometry),
  ensureOverlayWindowLevel: (window) => {
    overlayGeometryRuntime.ensureOverlayWindowLevel(window);
  },
  syncWindowsOverlayToMpvZOrder: (_window) => {
    visibleOverlayInteractionRuntime.requestWindowsVisibleOverlayZOrderSync();
  },
  syncPrimaryOverlayWindowLayer: (layer) => {
    overlayGeometryRuntime.syncPrimaryOverlayWindowLayer(layer);
  },
  enforceOverlayLayerOrder: () => {
    overlayGeometryRuntime.enforceOverlayLayerOrder();
  },
  syncOverlayShortcuts: () => {
    overlayShortcutsRuntime.syncOverlayShortcuts();
  },
  isMacOSPlatform: () => process.platform === 'darwin',
  isWindowsPlatform: () => process.platform === 'win32',
  showOverlayLoadingOsd: (_message: string) => {
    overlayNotificationsRuntime.showOverlayLoadingStatusNotification();
  },
  dismissOverlayLoadingOsd: () => {
    overlayNotificationsRuntime.dismissOverlayLoadingStatusNotification();
  },
  hideNonNativeOverlayWhenTargetUnfocused: () =>
    shouldRunLinuxOverlayZOrderKeepAlive() &&
    linuxOverlayModeRuntime.mode === 'fullscreen-override',
  resolveFallbackBounds: () => {
    const cursorPoint = screen.getCursorScreenPoint();
    const display = screen.getDisplayNearestPoint(cursorPoint);
    const fallbackBounds = display.workArea;
    return {
      x: fallbackBounds.x,
      y: fallbackBounds.y,
      width: fallbackBounds.width,
      height: fallbackBounds.height,
    };
  },
});

const visibleOverlayInteractionRuntime = createVisibleOverlayInteractionRuntime({
  overlayManager: {
    getMainWindow: () => overlayManager.getMainWindow(),
    getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
  },
  overlayContentMeasurementStore: {
    clear: (layer) => overlayContentMeasurementStore.clear(layer),
    getLatestByLayer: (layer) => overlayContentMeasurementStore.getLatestByLayer(layer),
  },
  logger: {
    info: (message, ...args) => logger.info(message, ...args),
    warn: (message, ...args) => logger.warn(message, ...args),
    debug: (message, ...args) => logger.debug(message, ...args),
  },
  updateVisibleOverlayVisibility: () => overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
  getModalInputExclusive: () => overlayModalInputState.getModalInputExclusive(),
  getStatsOverlayVisible: () => appState.statsOverlayVisible,
  setStatsOverlayVisible: (visible) => {
    appState.statsOverlayVisible = visible;
  },
  getWindowTracker: () => appState.windowTracker,
  setWindowTracker: (tracker) => {
    appState.windowTracker = tracker;
  },
  setTrackerNotReadyWarningShown: (shown) => {
    appState.trackerNotReadyWarningShown = shown;
  },
  getMpvSocketPath: () => appState.mpvSocketPath,
  getBackendOverride: () => appState.backendOverride,
  getInitialArgs: () => appState.initialArgs,
  getOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
  getLinuxVisibleOverlayWindowMode: () => linuxOverlayModeRuntime.mode,
  setLinuxVisibleOverlayOwnerBindingKey: (key) => {
    linuxOverlayModeRuntime.ownerBindingKey = key;
  },
  bindVisibleOverlayToTrackedX11Window: (window) =>
    overlayGeometryRuntime.bindVisibleOverlayToTrackedX11Window(window),
  updateVisibleOverlayBounds: (geometry) =>
    overlayGeometryRuntime.updateVisibleOverlayBounds(geometry),
  refreshCurrentSubtitle: () => {
    subtitleProcessingController.refreshCurrentSubtitle(appState.currentSubText);
  },
  getOverlayWindows: () => overlayManager.getOverlayWindows(),
  syncOverlayShortcuts: () => overlayShortcutsRuntime.syncOverlayShortcuts(),
  resetLastOverlayWindowGeometry: () => overlayGeometryRuntime.resetLastOverlayWindowGeometry(),
  enforceOverlayLayerOrder: () => {
    overlayGeometryRuntime.enforceOverlayLayerOrder();
  },
  getOverlayForegroundSeparateWindows: () => getOverlayForegroundSeparateWindows(),
});

const getRuntimeOptionsStateHandler = createGetRuntimeOptionsStateHandler({
  getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
});

syncOverlayVisibilityForModal = () => {
  overlayVisibilityRuntime.updateVisibleOverlayVisibility();
};

const overlayNotificationsRuntime = createOverlayNotificationsRuntime({
  getResolvedConfig: () => configService.getConfig(),
  getMainOverlayWindow: () => overlayManager.getMainWindow(),
  getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
  broadcastToOverlayWindows: (channel, ...args) =>
    overlayManager.broadcastToOverlayWindows(channel, ...args),
  showMpvOsd: (message) => showMpvOsd(message),
  getMpvClient: () => appState.mpvClient,
  getAnkiIntegration: () => appState.ankiIntegration,
  getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
});
const {
  flushQueuedOverlayNotifications,
  flushQueuedMpvOsdNotifications,
  openAnkiCardFromNotification,
  toggleNotificationHistoryPanel,
  showConfiguredPlaybackFeedback,
  maybeStartOverlayLoadingOsd,
} = overlayNotificationsRuntime;

const overlayVisibilityComposer = composeOverlayVisibilityRuntime({
  overlayVisibilityRuntime,
  restorePreviousSecondarySubVisibilityMainDeps: {
    getMpvClient: () => appState.mpvClient,
  },
  broadcastRuntimeOptionsChangedMainDeps: {
    broadcastRuntimeOptionsChangedRuntime,
    getRuntimeOptionsState: () => getRuntimeOptionsStateHandler(),
    broadcastToOverlayWindows: (channel, ...args) =>
      overlayManager.broadcastToOverlayWindows(channel, ...args),
  },
  sendToActiveOverlayWindowMainDeps: {
    sendToActiveOverlayWindowRuntime: (channel, payload, runtimeOptions) =>
      overlayModalRuntime.sendToActiveOverlayWindow(channel, payload, runtimeOptions),
  },
  setOverlayDebugVisualizationEnabledMainDeps: {
    setOverlayDebugVisualizationEnabledRuntime,
    getCurrentEnabled: () => appState.overlayDebugVisualizationEnabled,
    setCurrentEnabled: (next) => {
      appState.overlayDebugVisualizationEnabled = next;
    },
  },
  openRuntimeOptionsPaletteMainDeps: {
    openRuntimeOptionsPaletteRuntime: () => overlayModalRuntime.openRuntimeOptionsPalette(),
  },
});

const overlayModals = createOverlayModalsAssembly({
  appState,
  getConfig: () => configService.getConfig(),
  logger,
  overlayModalRuntime,
  getOverlayNotifications: () => overlayNotificationsRuntime,
  ensureOverlayStartupPrereqs: () => ensureOverlayStartupPrereqs(),
  ensureOverlayWindowsReadyForVisibilityActions: () =>
    ensureOverlayWindowsReadyForVisibilityActions(),
  sendToActiveOverlayWindow: (channel, payload, runtimeOptions) =>
    overlayVisibilityComposer.sendToActiveOverlayWindow(channel, payload, runtimeOptions),
  showMpvOsd: (text) => showMpvOsd(text),
  showDesktopNotification: (title, options) => showDesktopNotification(title, options),
});

const pluginPanelRuntime = createPluginPanelRuntime({
  openModal: (getSnapshot, signal) =>
    openPluginPanelModal(overlayModals.createOverlayHostedModalOpenDeps(), getSnapshot, signal),
  closeModal: () => overlayModalRuntime.handleOverlayModalClosed('plugin-panel'),
  // The panel may live in either window; broadcastToOverlayWindows only reaches the main one.
  broadcast: (channel, payload) => {
    for (const window of [overlayManager.getMainWindow(), overlayManager.getModalWindow()]) {
      if (window && !window.isDestroyed()) window.webContents.send(channel, payload);
    }
  },
  onOpenFailed: () =>
    overlayNotificationsRuntime.showConfiguredStatusNotification('Plugin panel unavailable.', {
      variant: 'warning',
    }),
});

function getPluginCurrentSubtitle(): PluginSubtitleLine | null {
  const text = appState.currentSubText.trim();
  if (!text) return null;
  // Same view as mining: the parsed line with its cue timings on the subtitle file's timeline.
  const resolved = resolvePrimarySubtitle({
    liveText: appState.mpvClient?.currentSubText ?? '',
    currentTimeSec: toSubtitleClock(
      Number(appState.mpvClient?.currentTimePos),
      appState.mpvClient?.currentSubDelay ?? 0,
    ),
    cues: appState.activeParsedSubtitleCues,
  });
  const resolvedText = resolved?.text.replace(/\n{2,}/g, '\n').trim();
  if (
    resolved &&
    resolvedText &&
    Number.isFinite(resolved.startTime) &&
    Number.isFinite(resolved.endTime) &&
    resolved.endTime > resolved.startTime
  ) {
    return { text: resolvedText, startTime: resolved.startTime, endTime: resolved.endTime };
  }
  const client = appState.mpvClient;
  // The client reports a cleared sub-start as zero, so only trust timing for a live line.
  const hasTiming = client !== null && client.currentSubEnd > client.currentSubStart;
  return {
    text,
    startTime: hasTiming ? client.currentSubStart : null,
    endTime: hasTiming ? client.currentSubEnd : null,
  };
}

const pluginWindowRuntime = createPluginWindowRuntime({
  createBrowserWindow: (options) => new BrowserWindow(options),
  promoteAboveOverlay: (window) => promoteSettingsWindowAboveOverlay(window),
  logError: (message) => logger.error(message),
});

const pluginHost = createPluginHost({
  pluginsDir: path.join(CONFIG_DIR, 'plugins'),
  pluginDataDir: path.join(USER_DATA_PATH, 'plugin-data'),
  getPluginConfig: () => configService.getConfig().plugins,
  createLogger,
  isMpvConnected: () => appState.mpvClient?.connected === true,
  sendMpvCommand: (command) => sendMpvCommandRuntime(appState.mpvClient, command),
  requestMpvProperty: async (name) => {
    if (!appState.mpvClient) throw new Error('mpv is not connected.');
    return await appState.mpvClient.requestProperty(name);
  },
  requestMpvCommand: async (command) => {
    if (!appState.mpvClient) throw new Error('mpv is not connected.');
    return await appState.mpvClient.requestCommand(command);
  },
  ensureMpvRunning: () => jellyfin.ensureMpvConnectedForPlayback(),
  invokeAnkiConnect: (action, params, options) => {
    const ankiConfig = configService.getConfig().ankiConnect;
    const url = (
      appState.runtimeOptionsManager?.getEffectiveAnkiConnectConfig(ankiConfig) ?? ankiConfig
    ).url;
    if (!url) return Promise.reject(new Error('ankiConnect.url is not configured.'));
    // No retries: a retried `sync` or `addNote` could run twice.
    return new AnkiConnectClient(url).invoke(action, params, {
      timeout: options.timeoutMs,
      maxRetries: 0,
    });
  },
  dispatchSessionAction: (request) => dispatchSessionAction(request),
  getCurrentSubtitle: () => getPluginCurrentSubtitle(),
  getSubtitleContext: (options) =>
    buildPluginSubtitleContext({
      cues: appState.activeParsedSubtitleCues,
      current: getPluginCurrentSubtitle(),
      recentLines: (appState.subtitleTimingTracker?.getRecentEntries(40) ?? []).map((entry) => ({
        text: entry.displayText,
        startTime: entry.startTime,
        endTime: entry.endTime,
      })),
      ...options,
    }),
  notify: (message, variant) =>
    overlayNotificationsRuntime.showConfiguredStatusNotification(message, { variant }),
  openPanel: (content) => pluginPanelRuntime.open(content),
  openWindow: (request) => pluginWindowRuntime.open(request),
  unstable: { appState, configService, overlayManager },
});

const mediaTimingFrameExtractor = createMediaTimingFrameExtractor();
const mediaTimingReviewRuntime = createMediaTimingReviewRuntime({
  getMpvClient: () => appState.mpvClient,
  getCurrentMediaPath: () =>
    appState.currentMediaPath?.trim() || appState.mpvClient?.currentVideoPath?.trim() || null,
  // An invalid configured path is passed through so the preview fails instead of using PATH mpv.
  getMpvExecutablePath: () => {
    const configuredPath = configService.getConfig().mpv.executablePath;
    return findMpvExecutablePath(configuredPath) || configuredPath;
  },
  createPreviewSession: () => new MediaTimingPreviewSession(),
  generateWaveform: (options) => generateSpeechWaveform(options),
  generateFrame: (options) => mediaTimingFrameExtractor.generate(options),
  clearFrameCache: () => mediaTimingFrameExtractor.clear(),
  resolveVideoSource: () =>
    resolveMediaGenerationInput(appState.mpvClient, 'video', {
      getCachedMediaPath: (currentVideoPath, kind) =>
        youtube.getCachedMediaPathForCurrentPlayback(currentVideoPath, kind),
      remoteCacheMode: youtube.shouldRequireMediaCacheForCurrentPlayback()
        ? 'required'
        : 'optional',
    }),
  resolveMediaSource: async () => {
    const resolved = await resolveMediaGenerationInput(appState.mpvClient, 'audio', {
      getCachedMediaPath: (currentVideoPath, kind) =>
        youtube.getCachedMediaPathForCurrentPlayback(currentVideoPath, kind),
      remoteCacheMode: youtube.shouldRequireMediaCacheForCurrentPlayback()
        ? 'required'
        : 'optional',
    });
    return resolved
      ? {
          path: resolved.path,
          ...(resolved.inputOptions ? { inputOptions: resolved.inputOptions } : {}),
          singleResolvedStream: resolved.singleResolvedStream,
        }
      : null;
  },
  acquireMediaWindow: (source, range) => getSharedRemoteMediaWindowCache().acquire(source, range),
  getSubtitleContextLines: (range) =>
    collectMediaTimingContextLines({
      // Review ranges and tracker entries are on the video clock; parsed cues are not.
      cues: appState.activeParsedSubtitleCues.map((cue) =>
        shiftSubtitleRangeToVideoClock(cue, appState.mpvClient?.currentSubDelay ?? 0),
      ),
      fallbackPrevious: appState.subtitleTimingTracker?.getRecentEntries(40) ?? [],
      startTime: range.startTime,
      endTime: range.endTime,
    }),
  openModal: (payload, signal) =>
    openMediaTimingReviewModal(overlayModals.createOverlayHostedModalOpenDeps(), payload, signal),
  onPreviewEnded: (reviewId) => {
    // The review may live in either overlay window; the renderer ignores foreign review ids.
    for (const window of [overlayManager.getMainWindow(), overlayManager.getModalWindow()]) {
      if (window && !window.isDestroyed()) {
        window.webContents.send(IPC_CHANNELS.event.mediaTimingReviewPreviewEnded, reviewId);
      }
    }
  },
  onPreviewPosition: (position) => {
    for (const window of [overlayManager.getMainWindow(), overlayManager.getModalWindow()]) {
      if (window && !window.isDestroyed()) {
        window.webContents.send(IPC_CHANNELS.event.mediaTimingReviewPreviewPosition, position);
      }
    }
  },
  showStatus: (message) =>
    overlayNotificationsRuntime.showConfiguredStatusNotification(message, { variant: 'warning' }),
});

function getRuntimeBooleanOption(
  id:
    | 'subtitle.annotation.knownWords.highlightEnabled'
    | 'subtitle.annotation.nPlusOne'
    | 'subtitle.annotation.jlpt'
    | 'subtitle.annotation.frequency',
  fallback: boolean,
): boolean {
  const value = appState.runtimeOptionsManager?.getOptionValue(id);
  return typeof value === 'boolean' ? value : fallback;
}

function shouldInitializeMecabForAnnotations(): boolean {
  const config = configService.getConfig();
  const knownWordsEnabled = getRuntimeBooleanOption(
    'subtitle.annotation.knownWords.highlightEnabled',
    config.ankiConnect.knownWords.highlightEnabled,
  );
  const nPlusOneEnabled = getRuntimeBooleanOption(
    'subtitle.annotation.nPlusOne',
    config.ankiConnect.nPlusOne.enabled,
  );
  const jlptEnabled = getRuntimeBooleanOption(
    'subtitle.annotation.jlpt',
    config.subtitleStyle.enableJlpt,
  );
  const frequencyEnabled = getRuntimeBooleanOption(
    'subtitle.annotation.frequency',
    config.subtitleStyle.frequencyDictionary.enabled,
  );
  return knownWordsEnabled || nPlusOneEnabled || jlptEnabled || frequencyEnabled;
}

const jellyfin = createJellyfinAssembly({
  appState,
  configService,
  jellyfinTokenStore,
  logger,
  defaultMpvLogPath: DEFAULT_MPV_LOG_PATH,
  getAppDataPath: () => app.getPath('appData'),
  getRuntimePluginEntrypoint: () => resolveBundledMpvRuntimePluginEntrypoint(),
  getMpvPluginRuntimeConfig: () => getMpvPluginRuntimeConfig(),
  createMpvClient: () => createMpvClientRuntimeService(),
  initSubtitlePrefetch: (sourcePath) =>
    subtitlePrefetchRuntime.refreshSubtitleSidebarFromSource(sourcePath),
  showVisibleOverlay: () => setVisibleOverlayVisible(true),
  getOverlayNotifications: () => overlayNotificationsRuntime,
  updateCurrentMediaTitle: (title) => mediaRuntime.updateCurrentMediaTitle(title),
  ensureImmersionTrackerStarted: () => ensureImmersionTrackerStarted(),
  refreshTrayMenu: () => refreshTrayMenuIfPresent(),
  setupPreloadPath: path.join(__dirname, 'preload-jellyfin-setup.js'),
  createBrowserWindow: (options) => new BrowserWindow(options),
  ipcMain,
});

const anilist = createAnilistAssembly({
  appState,
  configService,
  logger,
  anilistTokenStore,
  anilistUpdateQueue,
  anilistRateLimiter,
  getOverlayNotifications: () => overlayNotificationsRuntime,
  showDesktopNotification: (title, options) => showDesktopNotification(title, options),
  resolveMediaPathForJimaku: (currentMediaPath) =>
    mediaRuntime.resolveMediaPathForJimaku(currentMediaPath),
  resolvePinnedMediaId: (request) => characterDictionary.runtime.resolvePinnedMediaId(request),
  createBrowserWindow: (options) => new BrowserWindow(options),
  shell,
  app,
});

const loadSubtitlePosition = createLoadSubtitlePositionHandler({
  loadSubtitlePositionCore: () =>
    loadSubtitlePositionCore({
      currentMediaPath: appState.currentMediaPath,
      fallbackPosition: configService.getConfig().subtitlePosition,
      subtitlePositionsDir: SUBTITLE_POSITIONS_DIR,
    }),
  setSubtitlePosition: (position) => {
    appState.subtitlePosition = position;
  },
});

const saveSubtitlePosition = createSaveSubtitlePositionHandler({
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

anilist.registerSubminerProtocolClient();
let flushPendingMpvLogWrites = (): void => {};
const {
  registerProtocolUrlHandlers: registerProtocolUrlHandlersHandler,
  onWillQuitCleanup: onWillQuitCleanupHandler,
  shouldRestoreWindowsOnActivate: shouldRestoreWindowsOnActivateHandler,
  restoreWindowsOnActivate: restoreWindowsOnActivateHandler,
} = composeStartupLifecycleHandlers({
  registerProtocolUrlHandlersMainDeps: {
    registerOpenUrl: (listener) => {
      app.on('open-url', listener);
    },
    registerSecondInstance: (listener) => {
      registerSecondInstanceHandlerEarly(app, listener);
    },
    handleAnilistSetupProtocolUrl: (rawUrl) => anilist.handleAnilistSetupProtocolUrl(rawUrl),
    findAnilistSetupDeepLinkArgvUrl: (argv) => findAnilistSetupDeepLinkArgvUrl(argv),
    logUnhandledOpenUrl: (rawUrl) => {
      logger.warn('Unhandled app protocol URL', { rawUrl });
    },
    logUnhandledSecondInstanceUrl: (rawUrl) => {
      logger.warn('Unhandled second-instance protocol URL', { rawUrl });
    },
  },
  onWillQuitCleanupMainDeps: {
    destroyTray: () => destroyTrayHandler(),
    stopConfigHotReload: () => settings.hotReload.stop(),
    restorePreviousSecondarySubVisibility: () =>
      overlayVisibilityComposer.restorePreviousSecondarySubVisibility(),
    restoreMpvSubVisibility: () => {
      restoreOverlayMpvSubtitles({ force: true });
    },
    isAppReady: () => app.isReady(),
    unregisterAllGlobalShortcuts: () => globalShortcut.unregisterAll(),
    stopSubtitleWebsocket: () => {
      subtitleWsService.stop();
      annotationSubtitleWsService.stop();
    },
    stopTexthookerService: () => texthookerService.stop(),
    stopSyncAutoScheduler: async () => {
      sync.autoScheduler.stop();
      await sync.runtime.shutdown();
    },
    clearWindowsVisibleOverlayForegroundPollLoop: () =>
      visibleOverlayInteractionRuntime.clearWindowsVisibleOverlayForegroundPollLoop(),
    clearLinuxMpvFullscreenOverlayRefreshTimeouts: () => {
      linuxOverlayModeRuntime.cancelPendingTransition();
      cancelLinuxMpvFullscreenOverlayRefreshBurst = null;
      clearLinuxMpvFullscreenOverlayRefreshTimeouts();
    },
    getMainOverlayWindow: () => overlayManager.getMainWindow(),
    clearMainOverlayWindow: () => overlayManager.setMainWindow(null),
    getModalOverlayWindow: () => overlayManager.getModalWindow(),
    clearModalOverlayWindow: () => overlayManager.setModalWindow(null),
    getYomitanParserWindow: () => appState.yomitanParserWindow,
    clearYomitanParserState: () => {
      appState.yomitanParserWindow = null;
      appState.yomitanParserReadyPromise = null;
      appState.yomitanParserInitPromise = null;
      appState.yomitanSession = null;
    },
    getWindowTracker: () => appState.windowTracker,
    flushMpvLog: () => flushPendingMpvLogWrites(),
    getMpvSocket: () => appState.mpvClient?.socket ?? null,
    getReconnectTimer: () => appState.reconnectTimer,
    clearReconnectTimerRef: () => {
      appState.reconnectTimer = null;
    },
    getSubtitleTimingTracker: () => appState.subtitleTimingTracker,
    getImmersionTracker: () => appState.immersionTracker,
    stopStatsServer: () => stopStatsServer(),
    clearImmersionTracker: () => {
      appState.statsServer = null;
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
    getFirstRunSetupWindow: () => appState.firstRunSetupWindow,
    clearFirstRunSetupWindow: () => {
      appState.firstRunSetupWindow = null;
    },
    getYomitanSettingsWindow: () => appState.yomitanSettingsWindow,
    clearYomitanSettingsWindow: () => {
      appState.yomitanSettingsWindow = null;
    },
    stopJellyfinRemoteSession: () => jellyfin.stopJellyfinRemoteSession(),
    cleanupInternalSubtitleTrackCache: () => cachedInternalSubtitleTrackExtractor.clear(),
    cleanupYoutubeSubtitleTempDirs: () => {
      // Stops yt-dlp/whisper-cli and deletes the audio they were working on.
      youtube.flow.cancelWhisperGeneration();
      youtube.whisperSubtitles.removeActiveAudio();
      youtube.flow.cleanupSubtitleTempDirs();
    },
    cleanupYoutubeMediaCache: () => youtube.mediaCache.cleanup(),
    cleanupRemoteMediaWindows: () => getSharedRemoteMediaWindowCache().cleanup(),
    cleanupJellyfinSubtitleCache: () => jellyfin.cleanupJellyfinSubtitleCache(),
    stopDiscordPresenceService: () => {
      void appState.discordPresenceService?.stop();
      appState.discordPresenceService = null;
    },
    unloadPlugins: () => pluginHost.unloadAll(),
  },
  shouldRestoreWindowsOnActivateMainDeps: {
    isOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
    getAllWindowCount: () => BrowserWindow.getAllWindows().length,
  },
  restoreWindowsOnActivateMainDeps: {
    createMainWindow: () => {
      createMainWindow();
    },
    updateVisibleOverlayVisibility: () => {
      overlayVisibilityRuntime.updateVisibleOverlayVisibility();
    },
    syncOverlayMpvSubtitleSuppression: () => {
      syncOverlayMpvSubtitleSuppression();
    },
  },
});
registerProtocolUrlHandlersHandler();

const resolveLegacyVocabularyPos = async (row: {
  headword: string;
  word: string;
  reading: string | null;
}) => {
  const tokenizer = appState.mecabTokenizer;
  if (!tokenizer) {
    return null;
  }

  const lookupTexts = [...new Set([row.headword, row.word, row.reading ?? ''])]
    .map((value) => value.trim())
    .filter((value) => value.length > 0);

  for (const lookupText of lookupTexts) {
    const tokens = await tokenizer.tokenize(lookupText);
    const resolved = resolveLegacyVocabularyPosFromTokens(lookupText, tokens);
    if (resolved) {
      return resolved;
    }
  }

  return null;
};

const immersionTrackerStartupMainDeps: ImmersionTrackerStartupDeps = {
  getResolvedConfig: () => configService.getConfig(),
  getConfiguredDbPath: () => immersionMediaRuntime.getConfiguredDbPath(),
  createTrackerService: (params) =>
    new ImmersionTrackerService({
      ...params,
      resolveLegacyVocabularyPos,
    }),
  setTracker: (tracker) => {
    const trackerHasChanged =
      appState.immersionTracker !== null && appState.immersionTracker !== tracker;
    if (trackerHasChanged && appState.statsServer) {
      void stopStatsServer().catch((error: unknown) => {
        logger.warn('Failed to stop stats server while replacing immersion tracker.', error);
      });
      appState.statsServer = null;
    }

    appState.immersionTracker = tracker as ImmersionTrackerService | null;
    appState.immersionTracker?.setCoverArtFetcher(statsCoverArtFetcher);
    if (tracker) {
      // Start HTTP stats server
      if (!appState.statsServer) {
        const config = configService.getConfig();
        if (config.stats.autoStartServer) {
          void ensureStatsServerStarted().catch((error: unknown) => {
            logger.warn('Failed to auto-start stats server.', error);
          });
        }
      }

      // Register stats overlay toggle IPC handler (idempotent)
      registerStatsOverlayToggle({
        preloadPath: statsPreloadPath,
        getApiBaseUrl: async () => (await ensureStatsServerStarted()).url,
        onStartupError: (error) =>
          overlayNotificationsRuntime.showConfiguredStatusNotification(
            `Stats server startup failed: ${error instanceof Error ? error.message : String(error)}`,
            { title: 'Stats' },
          ),
        getToggleKey: () => configService.getConfig().stats.toggleKey,
        resolveBounds: () => overlayGeometryRuntime.getCurrentOverlayGeometry(),
        onVisibilityChanged: (visible) => {
          visibleOverlayInteractionRuntime.handleStatsOverlayVisibilityChanged(visible);
        },
      });
    }
  },
  getMpvClient: () => appState.mpvClient,
  shouldAutoConnectMpv: () => !appState.statsStartupInProgress,
  seedTrackerFromCurrentMedia: () => {
    void immersionMediaRuntime.seedFromCurrentMedia();
  },
  logInfo: (message) => logger.info(message),
  logDebug: (message) => logger.debug(message),
  logWarn: (message, details) => logger.warn(message, details),
};
const createImmersionTrackerStartup = createImmersionTrackerStartupHandler(
  immersionTrackerStartupMainDeps,
);
const recordSubtitleMiningContext = (context: SubtitleMiningContext | null): void => {
  pendingSubtitleMiningContext = context;
};
const consumePendingSubtitleMiningContext = (): SubtitleMiningContext | null => {
  const context = pendingSubtitleMiningContext;
  pendingSubtitleMiningContext = null;
  return context;
};
const recordTrackedCardsMined = (count: number, noteIds?: number[]): void => {
  ensureImmersionTrackerStarted();
  appState.immersionTracker?.recordCardsMined(count, noteIds);
};
function refreshCurrentSubtitleAnnotations(): void {
  const hasCurrentSubtitle = appState.currentSubText.trim().length > 0;
  if (hasCurrentSubtitle) {
    subtitlePrefetchService?.pause();
  }
  subtitleProcessingController.invalidateTokenizationCache();
  subtitlePrefetchService?.onSeek(lastObservedTimePos);
  if (!subtitleProcessingController.refreshCurrentSubtitle(appState.currentSubText)) {
    // Idle controller: no settle is coming to release the pause above.
    subtitlePrefetchService?.resume();
  }
}
let hasAttemptedImmersionTrackerStartup = false;
const ensureImmersionTrackerStarted = (): void => {
  if (hasAttemptedImmersionTrackerStartup || appState.immersionTracker) {
    return;
  }
  hasAttemptedImmersionTrackerStartup = true;
  createImmersionTrackerStartup();
};
const statsStartupRuntime = {
  ensureStatsServerStarted: () => ensureStatsServerStarted(),
  ensureBackgroundStatsServerStarted: () => ensureBackgroundStatsServerStarted(),
  stopBackgroundStatsServer: () => stopBackgroundStatsServer(),
  ensureImmersionTrackerStarted: () => {
    appState.statsStartupInProgress = true;
    try {
      ensureImmersionTrackerStarted();
    } finally {
      appState.statsStartupInProgress = false;
    }
  },
} as const;
const ensureBackgroundStatsServer = createEnsureBackgroundStatsServerHandler({
  isStatsAutoStartEnabled: () => configService.getConfig().stats.autoStartServer,
  isImmersionTrackingEnabled: () => configService.getConfig().immersionTracking?.enabled !== false,
  ensureBackgroundStatsServerStarted: () =>
    statsStartupRuntime.ensureBackgroundStatsServerStarted(),
  logInfo: (message) => logger.info(message),
  logWarn: (message, error) => logger.warn(message, error),
});

const runStatsCliCommand = createRunStatsCliCommandHandler({
  getResolvedConfig: () => configService.getConfig(),
  ensureImmersionTrackerStarted: () => statsStartupRuntime.ensureImmersionTrackerStarted(),
  ensureVocabularyCleanupTokenizerReady: async () => {
    await createMecabTokenizerAndCheck();
  },
  getImmersionTracker: () => appState.immersionTracker,
  ensureStatsServerStarted: async () => (await statsStartupRuntime.ensureStatsServerStarted()).url,
  ensureBackgroundStatsServerStarted: () =>
    statsStartupRuntime.ensureBackgroundStatsServerStarted(),
  stopBackgroundStatsServer: () => statsStartupRuntime.stopBackgroundStatsServer(),
  openExternal: (url: string) => shell.openExternal(url),
  writeResponse: (responsePath, payload) => {
    writeStatsCliCommandResponse(responsePath, payload);
  },
  exitAppWithCode: (code) => {
    process.exitCode = code;
    requestAppQuit();
  },
  logInfo: (message) => logger.info(message),
  logWarn: (message, error) => logger.warn(message, error),
  logError: (message, error) => logger.error(message, error),
});

async function runHeadlessInitialCommand(): Promise<void> {
  if (!appState.initialArgs?.refreshKnownWords) {
    handleInitialArgs();
    return;
  }

  const resolvedConfig = configService.getConfig();
  if (resolvedConfig.ankiConnect.enabled !== true) {
    logger.error('Headless known-word refresh failed: AnkiConnect integration not enabled');
    process.exitCode = 1;
    requestAppQuit();
    return;
  }

  const effectiveAnkiConfig =
    appState.runtimeOptionsManager?.getEffectiveAnkiConnectConfig(resolvedConfig.ankiConnect) ??
    resolvedConfig.ankiConnect;
  const integration = new AnkiIntegration(
    effectiveAnkiConfig,
    new SubtitleTimingTracker(),
    { send: () => undefined } as never,
    undefined,
    undefined,
    async () => ({
      keepNoteId: 0,
      deleteNoteId: 0,
      deleteDuplicate: false,
      cancelled: true,
    }),
    path.join(USER_DATA_PATH, 'known-words-cache.json'),
    mergeAiConfig(resolvedConfig.ai, resolvedConfig.ankiConnect?.ai),
  );

  try {
    await integration.refreshKnownWordCache();
  } catch (error) {
    logger.error('Headless known-word refresh failed:', error);
    process.exitCode = 1;
  } finally {
    integration.stop();
    requestAppQuit();
  }
}

const { appReadyRuntimeRunner } = composeAppReadyRuntime({
  reloadConfigMainDeps: {
    reloadConfigStrict: () => configService.reloadConfigStrict(),
    logInfo: (message) => appLogger.logInfo(message),
    logDebug: (message) => appLogger.logDebug(message),
    logWarning: (message) => appLogger.logWarning(message),
    showDesktopNotification: (title, options) => showDesktopNotification(title, options),
    startConfigHotReload: () => settings.hotReload.start(),
    shouldRefreshAnilistClientSecretState: () =>
      shouldRefreshAnilistOnConfigReload(appState.initialArgs),
    refreshAnilistClientSecretState: (options) =>
      anilist.refreshClientSecretStateIfEnabled(options),
    failHandlers: {
      logError: (details) => logger.error(details),
      showErrorBox: (title, details) => dialog.showErrorBox(title, details),
      quit: () => requestAppQuit(),
    },
  },
  criticalConfigErrorMainDeps: {
    getConfigPath: () => configService.getConfigPath(),
    failHandlers: {
      logError: (message) => logger.error(message),
      showErrorBox: (title, message) => dialog.showErrorBox(title, message),
      quit: () => requestAppQuit(),
    },
  },
  appReadyRuntimeMainDeps: {
    ensureDefaultConfigBootstrap: () => {
      ensureDefaultConfigBootstrap({
        configDir: CONFIG_DIR,
        configFilePaths: getDefaultConfigFilePaths(CONFIG_DIR),
        generateTemplate: () => generateConfigTemplate(DEFAULT_CONFIG),
      });
    },
    loadSubtitlePosition: () => loadSubtitlePosition(),
    resolveKeybindings: () => {
      appState.keybindings = resolveKeybindings(configService.getConfig(), DEFAULT_KEYBINDINGS);
      refreshCurrentSessionBindings();
    },
    createMpvClient: () => {
      appState.mpvClient = createMpvClientRuntimeService();
    },
    startPlugins: () => {
      void pluginHost.loadAll();
    },
    getResolvedConfig: () => configService.getConfig(),
    getConfigWarnings: () => configService.getWarnings(),
    logConfigWarning: (warning) => appLogger.logConfigWarning(warning),
    setLogLevel: (level: string, source: LogLevelSource) => setLogLevel(level, source),
    setLogRotation: (rotation: number) => setLogRotation(rotation),
    setLogFileToggles: (files) => setLogFileToggles(files),
    initRuntimeOptionsManager: () => {
      appState.runtimeOptionsManager = new RuntimeOptionsManager(
        () => configService.getConfig().ankiConnect,
        {
          applyAnkiPatch: (patch) => {
            if (appState.ankiIntegration) {
              appState.ankiIntegration.applyRuntimeConfigPatch(patch);
            }
          },
          getSubtitleStyleConfig: () => configService.getConfig().subtitleStyle,
          onOptionsChanged: () => {
            subtitleProcessingController.invalidateTokenizationCache();
            subtitlePrefetchService?.onSeek(lastObservedTimePos);
            overlayVisibilityComposer.broadcastRuntimeOptionsChanged();
            refreshOverlayShortcuts();
          },
        },
      );
    },
    setSecondarySubMode: (mode: SecondarySubMode) => {
      setSecondarySubMode(mode);
    },
    defaultSecondarySubMode: 'hover',
    defaultWebsocketPort: DEFAULT_CONFIG.websocket.port,
    defaultAnnotationWebsocketPort: DEFAULT_CONFIG.annotationWebsocket.port,
    defaultTexthookerPort: DEFAULT_TEXTHOOKER_PORT,
    hasMpvWebsocketPlugin: () => hasMpvWebsocketPlugin(),
    startSubtitleWebsocket: (port: number) => {
      subtitleWsService.start(
        port,
        () =>
          appState.currentSubtitleData ??
          (appState.currentSubText
            ? {
                text: appState.currentSubText,
                tokens: null,
                startTime: appState.mpvClient?.currentSubStart ?? null,
                endTime: appState.mpvClient?.currentSubEnd ?? null,
              }
            : null),
        () => {
          const frequencyDictionary = configService.getConfig().subtitleStyle.frequencyDictionary;
          return {
            enabled: frequencyDictionary.enabled,
            topX: frequencyDictionary.topX,
            mode: frequencyDictionary.mode,
          };
        },
      );
    },
    startAnnotationWebsocket: (port: number) => {
      annotationSubtitleWsService.start(
        port,
        () =>
          appState.currentSubtitleData ??
          (appState.currentSubText
            ? {
                text: appState.currentSubText,
                tokens: null,
                startTime: appState.mpvClient?.currentSubStart ?? null,
                endTime: appState.mpvClient?.currentSubEnd ?? null,
              }
            : null),
        () => {
          const frequencyDictionary = configService.getConfig().subtitleStyle.frequencyDictionary;
          return {
            enabled: frequencyDictionary.enabled,
            topX: frequencyDictionary.topX,
            mode: frequencyDictionary.mode,
          };
        },
      );
    },
    startTexthooker: (port: number, websocketUrl?: string) => {
      if (!texthookerService.isRunning()) {
        texthookerService.start(port, websocketUrl);
      }
    },
    log: (message) => appLogger.logInfo(message),
    createMecabTokenizerAndCheck: async () => {
      await createMecabTokenizerAndCheck();
    },
    createSubtitleTimingTracker: () => {
      const tracker = new SubtitleTimingTracker();
      appState.subtitleTimingTracker = tracker;
    },
    loadYomitanExtension: async () => {
      await loadYomitanExtension();
    },
    ensureYomitanExtensionLoaded: async () => {
      await ensureYomitanExtensionLoaded();
    },
    handleFirstRunSetup: async () => {
      const snapshot = await firstRunSetup.service.ensureSetupStateInitialized();
      appState.firstRunSetupCompleted = snapshot.state.status === 'completed';
      const args = appState.initialArgs;
      if (args && shouldAutoOpenFirstRunSetup(args)) {
        const force = Boolean(args.setup);
        if (force || snapshot.state.status !== 'completed') {
          firstRunSetup.openWindow(force);
        }
      }
    },
    startJellyfinRemoteSession: async () => {
      await jellyfin.startJellyfinRemoteSession();
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
    setVisibleOverlayVisible: (visible: boolean) => setVisibleOverlayVisible(visible),
    initializeOverlayRuntime: () => initializeOverlayRuntime(),
    runHeadlessInitialCommand: () => runHeadlessInitialCommand(),
    handleInitialArgs: () => handleInitialArgs(),
    shouldRunHeadlessInitialCommand: () =>
      Boolean(appState.initialArgs && isHeadlessInitialCommand(appState.initialArgs)),
    shouldUseMinimalStartup: () =>
      getStartupModeFlags(appState.initialArgs).shouldUseMinimalStartup,
    shouldSkipHeavyStartup: () => getStartupModeFlags(appState.initialArgs).shouldSkipHeavyStartup,
    shouldHandleInitialArgsBeforeDeferredOverlayWarmup: () =>
      shouldHandleInitialArgsBeforeDeferredOverlayWarmup(appState.initialArgs),
    createImmersionTracker: () => {
      ensureImmersionTrackerStarted();
    },
    logDebug: (message: string) => {
      logger.debug(message);
    },
    now: () => Date.now(),
  },
  immersionTrackerStartupMainDeps,
});

async function runAppReadyRuntimeWithFatalReporting(): Promise<void> {
  try {
    await appReadyRuntimeRunner();
  } catch (error) {
    reportFatalError(error, {
      title: 'SubMiner startup failed',
      context: 'SubMiner failed during app-ready startup.',
    });
    process.exitCode = 1;
    requestAppQuit();
    return;
  }
}

function ensureOverlayStartupPrereqs(): void {
  if (appState.subtitlePosition === null) {
    loadSubtitlePosition();
  }
  if (appState.keybindings.length === 0) {
    appState.keybindings = resolveKeybindings(configService.getConfig(), DEFAULT_KEYBINDINGS);
    refreshCurrentSessionBindings();
  } else if (!appState.sessionBindingsInitialized) {
    refreshCurrentSessionBindings();
  }
  if (!appState.mpvClient) {
    appState.mpvClient = createMpvClientRuntimeService();
  }
  if (!appState.runtimeOptionsManager) {
    appState.runtimeOptionsManager = new RuntimeOptionsManager(
      () => configService.getConfig().ankiConnect,
      {
        applyAnkiPatch: (patch) => {
          if (appState.ankiIntegration) {
            appState.ankiIntegration.applyRuntimeConfigPatch(patch);
          }
        },
        getSubtitleStyleConfig: () => configService.getConfig().subtitleStyle,
        onOptionsChanged: () => {
          subtitleProcessingController.invalidateTokenizationCache();
          subtitlePrefetchService?.onSeek(lastObservedTimePos);
          overlayVisibilityComposer.broadcastRuntimeOptionsChanged();
          refreshOverlayShortcuts();
        },
      },
    );
  }
  if (!appState.subtitleTimingTracker) {
    appState.subtitleTimingTracker = new SubtitleTimingTracker();
  }
}

async function ensureYoutubePlaybackRuntimeReady(): Promise<void> {
  ensureOverlayStartupPrereqs();
  await ensureYomitanExtensionLoaded();
  if (!appState.overlayRuntimeInitialized) {
    initializeOverlayRuntime();
    return;
  }
  ensureOverlayWindowsReadyForVisibilityActions();
}

let signalAutoplayReadyFromWarmTokenization: ((path: string | null | undefined) => void) | null =
  null;

const {
  createMpvClientRuntimeService: createMpvClientRuntimeServiceHandler,
  updateMpvSubtitleRenderMetrics: updateMpvSubtitleRenderMetricsHandler,
  tokenizeSubtitle: tokenizeSubtitleRuntime,
  createMecabTokenizerAndCheck,
  prewarmSubtitleDictionaries,
  startBackgroundWarmups,
  startTokenizationWarmups,
  isTokenizationWarmupReady,
} = composeMpvRuntimeHandlers<
  MpvIpcClient,
  ReturnType<typeof createTokenizerDepsRuntime>,
  SubtitleData
>({
  bindMpvMainEventHandlersMainDeps: {
    appState,
    getQuitOnDisconnectArmed: () =>
      jellyfin.getQuitOnDisconnectArmed() || youtube.playback.getQuitOnDisconnectArmed(),
    scheduleQuitCheck: (callback) => {
      setTimeout(callback, 500);
    },
    quitApp: () => requestAppQuit(),
    reportJellyfinRemoteStopped: () => {
      void jellyfin.reportJellyfinRemoteStopped();
    },
    onMpvConnected: () => {
      maybeStartOverlayLoadingOsd();
      flushQueuedMpvOsdNotifications();
      secondarySubtitleTrackController.scheduleRefresh(0);
      void refreshMpvSessionBindings();
      if (appState.sessionBindingsInitialized) {
        sendMpvCommandRuntime(appState.mpvClient, [
          'script-message',
          'subminer-reload-session-bindings',
        ]);
      }
      if (appState.currentSubText.trim()) {
        subtitleProcessingController.refreshCurrentSubtitle(appState.currentSubText);
      }
    },
    maybeRunAnilistPostWatchUpdate: (options) => anilist.maybeRunAnilistPostWatchUpdate(options),
    recordAnilistMediaDuration: (durationSec) => {
      anilist.recordAnilistMediaDuration(durationSec);
    },
    logSubtitleTimingError: (message, error) => logger.error(message, error),
    broadcastToOverlayWindows: (channel, payload) => {
      overlayManager.broadcastToOverlayWindows(channel, payload);
    },
    onSecondarySubtitleChange: (text) => {
      secondarySubtitleTrackController.handleLiveText(text);
    },
    getImmediateSubtitlePayload: (text) => subtitleProcessingController.consumeCachedSubtitle(text),
    emitImmediateSubtitle: (payload) => {
      emitSubtitlePayload(payload);
    },
    onSubtitleChange: (text) => {
      // Pause only; restarting the prefetch run here would discard in-flight
      // tokenization work on every line. Real seeks restart via onTimePosUpdate.
      subtitlePrefetchService?.pause();
      if (!subtitleProcessingController.onSubtitleChange(text)) {
        // Repeat of the current text: the controller is idle, so no settle is
        // coming to release the pause. Resume now instead of idling prefetch
        // for the rest of the cue.
        subtitlePrefetchService?.resume();
      }
    },
    refreshDiscordPresence: () => {
      discordPresenceRuntime.publishDiscordPresence();
    },
    logSubtitleProcessingDebug: (message: string) => logger.debug(message),
    ensureImmersionTrackerInitialized: () => {
      ensureImmersionTrackerStarted();
    },
    tokenizeSubtitleForImmersion: async (text): Promise<SubtitleData | null> =>
      tokenizeSubtitleDeferred ? await tokenizeSubtitleDeferred(text) : null,
    updateCurrentMediaPath: (path) => {
      const normalizedPath = path.trim();
      maybeStartOverlayLoadingOsd(normalizedPath);
      const previousPath = appState.currentMediaPath?.trim() || null;
      const preserveParsedSubtitleCues = isSameYoutubeMediaPath(
        normalizedPath,
        appState.activeParsedSubtitleMediaPath,
      );
      if ((normalizedPath || null) !== previousPath) {
        cachedInternalSubtitleTrackExtractor.clear();
        secondarySubtitleTrackController.reset();
        const resetSubtitlePayload = { text: '', tokens: null };
        const frequencyDictionary = configService.getConfig().subtitleStyle.frequencyDictionary;
        const frequencyOptions = {
          enabled: frequencyDictionary.enabled,
          topX: frequencyDictionary.topX,
          mode: frequencyDictionary.mode,
        };
        autoplaySubtitlePrimingRuntime.resetAutoplaySubtitlePrime();
        lastObservedTimePos = 0;
        appState.currentSubText = '';
        appState.currentSubtitleData = null;
        if (!preserveParsedSubtitleCues) {
          appState.activeParsedSubtitleCues = [];
          appState.activeParsedSubtitleSource = null;
          appState.activeParsedSubtitleMediaPath = null;
        }
        overlayManager.broadcastToOverlayWindows('subtitle:set', resetSubtitlePayload);
        subtitleWsService.broadcast(resetSubtitlePayload, frequencyOptions);
        annotationSubtitleWsService.broadcast(resetSubtitlePayload, frequencyOptions);
        autoplayReadyGate.invalidatePendingAutoplayReadyFallbacks();
      }
      currentMediaTokenizationGate.updateCurrentMediaPath(path);
      managedLocalSubtitleSelectionRuntime.handleMediaPathChange(path);
      startupOsdSequencer.reset();
      subtitlePrefetchRuntime.clearScheduledSubtitlePrefetchRefresh();
      if (!preserveParsedSubtitleCues) {
        subtitlePrefetchRuntime.cancelPendingInit();
      }
      youtube.primarySubtitleNotification.handleMediaPathChange(path);
      void youtube.mediaCachePlayback.handleMediaPathChange(path);
      youtube.browserPlayback.handleMediaPathChange(path);
      youtube.flow.handleMediaPathChange(path);
      subtitleGenerationRuntime.handleMediaPathChange(path);
      if (path) {
        ensureImmersionTrackerStarted();
        secondarySubtitleTrackController.scheduleRefresh();
        void subtitlePrefetchRuntime.refreshSubtitlePrefetchFromActiveTrack();
        // Retry after a short delay because MPV can populate track-list after path.
        subtitlePrefetchRuntime.scheduleSubtitlePrefetchRefresh(500);
      }
      mediaRuntime.updateCurrentMediaPath(path);
    },
    restoreMpvSubVisibility: () => {
      restoreOverlayMpvSubtitles();
    },
    resetSubtitleSidebarEmbeddedLayout: () => {
      resetSubtitleSidebarEmbeddedLayoutRuntime();
    },
    getCurrentAnilistMediaKey: () => anilist.getCurrentAnilistMediaKey(),
    resetAnilistMediaTracking: (mediaKey) => {
      anilist.resetAnilistMediaTracking(mediaKey);
    },
    maybeProbeAnilistDuration: (mediaKey) => {
      void anilist.maybeProbeAnilistDuration(mediaKey);
    },
    ensureAnilistMediaGuess: (mediaKey) => {
      void anilist.ensureAnilistMediaGuess(mediaKey);
    },
    syncImmersionMediaState: () => {
      immersionMediaRuntime.syncFromCurrentMediaState();
    },
    signalAutoplayReadyIfWarm: (path) => signalAutoplayReadyFromWarmTokenization?.(path),
    markJellyfinRemotePlaybackLoaded: (path) => {
      markJellyfinRemotePlaybackLoadedState(jellyfin.getActivePlayback(), path);
    },
    scheduleCharacterDictionarySync: () => {
      if (
        !configService.getConfig().subtitleStyle.nameMatchEnabled ||
        !yomitanProfilePolicy.isCharacterDictionaryEnabled() ||
        youtube.isPlaybackActiveNow()
      ) {
        return;
      }
      characterDictionary.autoSync.scheduleSync();
    },
    updateCurrentMediaTitle: (title) => {
      mediaRuntime.updateCurrentMediaTitle(title);
    },
    resetAnilistMediaGuessState: () => {
      anilist.resetAnilistMediaGuessState();
    },
    reportJellyfinRemoteProgress: (forceImmediate) => {
      void jellyfin.reportJellyfinRemoteProgress(forceImmediate);
    },
    consumeExplicitSeek: () => consumeExplicitSeekIntent(),
    onTimePosUpdate: (time) => {
      const delta = time - lastObservedTimePos;
      if (subtitlePrefetchService && (delta > SEEK_THRESHOLD_SECONDS || delta < 0)) {
        subtitlePrefetchService.onSeek(time);
      }
      lastObservedTimePos = time;
      secondarySubtitleTrackController.handleTimePos(time);
    },
    onFullscreenChange: (fullscreen) => {
      cancelLinuxMpvFullscreenOverlayRefreshBurst = updateLinuxMpvFullscreenOverlayRefreshBurst(
        fullscreen,
        {
          overlayManager: {
            getMainWindow: () => overlayManager.getMainWindow(),
            getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
          },
          overlayVisibilityRuntime,
          syncVisibleOverlayMpvFullscreenMode: (nextFullscreen) =>
            linuxOverlayModeRuntime.sync(nextFullscreen),
          getOverlayInteractionActive: () =>
            visibleOverlayInteractionRuntime.getVisibleOverlayInteractionActive() ||
            visibleOverlayInteractionRuntime.getLinuxOverlayInputShapeActive(),
          ensureOverlayWindowLevel: (window) =>
            overlayGeometryRuntime.ensureOverlayWindowLevel(window),
        },
        cancelLinuxMpvFullscreenOverlayRefreshBurst,
      );
    },
    onSubtitleTrackChange: (sid) => {
      lastObservedPrimarySubtitleTrackId = sid;
      logger.info('[mpv-subtitles] primary subtitle track changed', { sid });
      autoplaySubtitlePrimingRuntime.scheduleSubtitlePrefetchRefresh();
      youtube.primarySubtitleNotification.handleSubtitleTrackChange(sid);
    },
    onSecondarySubtitleTrackChange: () => {
      secondarySubtitleTrackController.handleTrackChange();
      secondarySubtitleTrackController.scheduleRefresh(0);
    },
    onSecondarySubtitleDelayChange: (delay) => {
      secondarySubtitleTrackController.handleDelayChange(delay);
    },
    onSubtitleTrackListChange: (trackList) => {
      const diagnostics = buildSubtitleTrackDiagnostics(
        lastObservedPrimarySubtitleTrackId,
        trackList,
      );
      if (!diagnostics.trackListReadable) {
        logger.warn('[mpv-subtitles] mpv reported an unreadable subtitle track list', diagnostics);
      } else if (diagnostics.subtitleTrackCount === 0) {
        logger.warn('[mpv-subtitles] mpv reported no subtitle tracks', diagnostics);
      } else {
        logger.info('[mpv-subtitles] subtitle track list updated', diagnostics);
      }
      managedLocalSubtitleSelectionRuntime.handleSubtitleTrackListChange(trackList);
      secondarySubtitleTrackController.scheduleRefresh(0);
      autoplaySubtitlePrimingRuntime.scheduleSubtitlePrefetchRefresh();
      youtube.primarySubtitleNotification.handleSubtitleTrackListChange(trackList);
    },
    updateSubtitleRenderMetrics: (patch) => {
      updateMpvSubtitleRenderMetricsHandler(patch as Partial<MpvSubtitleRenderMetrics>);
    },
    syncOverlayMpvSubtitleSuppression: () => {
      syncOverlayMpvSubtitleSuppression();
    },
  },
  mpvClientRuntimeServiceFactoryMainDeps: {
    createClient: MpvIpcClient,
    getSocketPath: () => appState.mpvSocketPath,
    getResolvedConfig: () => configService.getConfig(),
    isAutoStartOverlayEnabled: () => appState.autoStartOverlay,
    setOverlayVisible: (visible: boolean) => setOverlayVisible(visible),
    isVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    getReconnectTimer: () => appState.reconnectTimer,
    setReconnectTimer: (timer: ReturnType<typeof setTimeout> | null) => {
      appState.reconnectTimer = timer;
    },
    shouldAutoLoadSecondarySubTrack: (path: string) =>
      shouldAutoLoadSecondarySubTrackForJellyfinPlayback(jellyfin.getActivePlayback(), path),
    shouldQuitOnMpvShutdown: () =>
      shouldQuitOnMpvShutdownForTrayState({
        managedPlayback: appState.initialArgs?.managedPlayback === true,
        backgroundMode: appState.backgroundMode,
        hasTray: Boolean(appTray),
      }),
    requestAppQuit: () => requestAppQuit(),
  },
  updateMpvSubtitleRenderMetricsMainDeps: {
    getCurrentMetrics: () => appState.mpvSubtitleRenderMetrics,
    setCurrentMetrics: (metrics) => {
      appState.mpvSubtitleRenderMetrics = metrics;
    },
    applyPatch: (current, patch) => applyMpvSubtitleRenderMetricsPatch(current, patch),
    broadcastMetrics: () => {
      // no renderer consumer for subtitle render metrics updates at present
    },
  },
  tokenizer: {
    buildTokenizerDepsMainDeps: {
      getYomitanExt: () => appState.yomitanExt,
      getYomitanSession: () => appState.yomitanSession,
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
      isKnownWord: (text, reading, options) =>
        Boolean(appState.ankiIntegration?.isKnownWord(text, reading, options)),
      getKnownWordTier: (text, reading, options) =>
        appState.ankiIntegration?.getKnownWordTier(text, reading, options) ?? null,
      recordLookup: (hit) => {
        ensureImmersionTrackerStarted();
        appState.immersionTracker?.recordLookup(hit);
      },
      getKnownWordMatchMode: () =>
        appState.ankiIntegration?.getKnownWordMatchMode() ??
        configService.getConfig().ankiConnect.knownWords.matchMode,
      getKnownWordsEnabled: () =>
        getRuntimeBooleanOption(
          'subtitle.annotation.knownWords.highlightEnabled',
          configService.getConfig().ankiConnect.knownWords.highlightEnabled,
        ),
      getNPlusOneEnabled: () =>
        getRuntimeBooleanOption(
          'subtitle.annotation.nPlusOne',
          configService.getConfig().ankiConnect.nPlusOne.enabled,
        ),
      getMinSentenceWordsForNPlusOne: () =>
        configService.getConfig().ankiConnect.nPlusOne.minSentenceWords,
      getJlptLevel: (text) => appState.jlptLevelLookup(text),
      getJlptEnabled: () =>
        getRuntimeBooleanOption(
          'subtitle.annotation.jlpt',
          configService.getConfig().subtitleStyle.enableJlpt,
        ),
      getCharacterDictionaryEnabled: () =>
        configService.getConfig().subtitleStyle.nameMatchEnabled &&
        yomitanProfilePolicy.isCharacterDictionaryEnabled() &&
        !youtube.isPlaybackActiveNow(),
      getNameMatchEnabled: () => configService.getConfig().subtitleStyle.nameMatchEnabled,
      getNameMatchImagesEnabled: () =>
        configService.getConfig().subtitleStyle.nameMatchImagesEnabled,
      getCharacterNameImage: (term) => characterDictionary.imageLookup.get(term),
      getCurrentCharacterDictionaryMediaId: () => characterDictionary.autoSync.getCurrentMediaId(),
      getCharacterNameCandidates: () => characterDictionary.nameCandidates.get(),
      getFrequencyDictionaryEnabled: () =>
        getRuntimeBooleanOption(
          'subtitle.annotation.frequency',
          configService.getConfig().subtitleStyle.frequencyDictionary.enabled,
        ),
      getFrequencyDictionaryMatchMode: () =>
        configService.getConfig().subtitleStyle.frequencyDictionary.matchMode,
      getFrequencyRank: (text) => appState.frequencyRankLookup(text),
      getYomitanGroupDebugEnabled: () => appState.overlayDebugVisualizationEnabled,
      getMecabTokenizer: () => appState.mecabTokenizer,
      onTokenizationReady: () => {
        currentMediaTokenizationGate.markReady(
          appState.currentMediaPath?.trim() || appState.mpvClient?.currentVideoPath?.trim() || null,
        );
        startupOsdSequencer.markTokenizationReady();
      },
    },
    createTokenizerRuntimeDeps: (deps) => createTokenizerDepsRuntime(deps),
    tokenizeSubtitle: (text, deps) => tokenizeSubtitleCore(text, deps),
    createMecabTokenizerAndCheckMainDeps: {
      getMecabTokenizer: () => appState.mecabTokenizer,
      setMecabTokenizer: (tokenizer) => {
        appState.mecabTokenizer = tokenizer as MecabTokenizer | null;
      },
      createMecabTokenizer: () => new MecabTokenizer(),
      checkAvailability: async (tokenizer) => (tokenizer as MecabTokenizer).checkAvailability(),
    },
    prewarmSubtitleDictionariesMainDeps: {
      ensureJlptDictionaryLookup: () => dictionaries.jlpt.ensureJlptDictionaryLookup(),
      ensureFrequencyDictionaryLookup: () =>
        dictionaries.frequency.ensureFrequencyDictionaryLookup(),
      showMpvOsd: (message: string) =>
        overlayNotificationsRuntime.showConfiguredStatusNotification(message),
      showLoadingOsd: (message: string) => startupOsdSequencer.showAnnotationLoading(message),
      showLoadedOsd: (message: string) =>
        startupOsdSequencer.markAnnotationLoadingComplete(message),
      shouldShowOsdNotification: () => {
        const type = overlayNotificationsRuntime.getConfiguredStatusNotificationType();
        return type === 'osd' || type === 'osd-system';
      },
    },
  },
  warmups: {
    launchBackgroundWarmupTaskMainDeps: {
      now: () => Date.now(),
      logDebug: (message) => logger.debug(message),
      logWarn: (message) => logger.warn(message),
    },
    startBackgroundWarmupsMainDeps: {
      getStarted: () => backgroundWarmupsStarted,
      setStarted: (started) => {
        backgroundWarmupsStarted = started;
      },
      isTexthookerOnlyMode: () => appState.texthookerOnlyMode,
      ensureYomitanExtensionLoaded: () => ensureYomitanExtensionLoaded().then(() => {}),
      shouldWarmupMecab: () => {
        const startupWarmups = configService.getConfig().startupWarmups;
        if (startupWarmups.lowPowerMode) {
          return false;
        }
        if (!startupWarmups.mecab) {
          return false;
        }
        return shouldInitializeMecabForAnnotations();
      },
      shouldWarmupYomitanExtension: () => configService.getConfig().startupWarmups.yomitanExtension,
      shouldWarmupSubtitleDictionaries: () => {
        const startupWarmups = configService.getConfig().startupWarmups;
        if (startupWarmups.lowPowerMode) {
          return false;
        }
        return startupWarmups.subtitleDictionaries;
      },
      shouldWarmupJellyfinRemoteSession: () => {
        const startupWarmups = configService.getConfig().startupWarmups;
        if (startupWarmups.lowPowerMode) {
          return false;
        }
        return startupWarmups.jellyfinRemoteSession;
      },
      shouldAutoConnectJellyfinRemote: () => {
        const jellyfin = configService.getConfig().jellyfin;
        return (
          jellyfin.enabled && jellyfin.remoteControlEnabled && jellyfin.remoteControlAutoConnect
        );
      },
      startJellyfinRemoteSession: () => jellyfin.startJellyfinRemoteSession(),
      logDebug: (message) => logger.debug(message),
    },
  },
});

async function tokenizeSubtitle(text: string): Promise<SubtitleData> {
  if (!isTokenizationWarmupReady()) {
    startupOsdSequencer.showTokenizationLoading('Loading subtitle tokenization...');
  }
  return await tokenizeSubtitleRuntime(text);
}

signalAutoplayReadyFromWarmTokenization = createAutoplayTokenizationWarmRelease({
  isTokenizationWarmupReady: () => isTokenizationWarmupReady(),
  startTokenizationWarmups: async () => {
    await startTokenizationWarmups();
  },
  getCurrentMediaPath: () =>
    appState.currentMediaPath?.trim() || appState.mpvClient?.currentVideoPath?.trim() || null,
  primeCurrentSubtitle: (mediaPath) =>
    autoplaySubtitlePrimingRuntime.primeCurrentSubtitleForAutoplay(mediaPath),
  signalAutoplayReady: () => signalCurrentSubtitleAutoplayReady(),
  warn: (message, error) => logger.warn(message, error),
});
tokenizeSubtitleDeferred = tokenizeSubtitle;

const aniSkipRuntime = createAniSkipRuntime({
  getAniSkipConfig: () => ({
    aniskipEnabled: configService.getConfig().mpv.aniskipEnabled,
    aniskipButtonKey: configService.getConfig().mpv.aniskipButtonKey,
  }),
  resolveMetadataForFile: (mediaPath) => resolveAniSkipMetadataForFile(mediaPath),
  sendMpvCommand: (command) => {
    appState.mpvClient?.send({ command });
  },
  requestMpvProperty: (name) => {
    if (!appState.mpvClient) {
      return Promise.reject(new Error('MPV not connected'));
    }
    return appState.mpvClient.requestProperty(name);
  },
  isMpvConnected: () => appState.mpvClient?.connected === true,
  getCurrentTimePos: () => appState.mpvClient?.currentTimePos ?? Number.NaN,
  showMpvOsd: (text, durationMs) => {
    appState.mpvClient?.send({ command: ['show-text', text, durationMs] });
  },
  showPlaybackFeedback: (text) => showConfiguredPlaybackFeedback(text),
  logInfo: (message) => logger.info(message),
  logWarn: (message, error) => logger.warn(message, error),
  logDebug: (message) => logger.debug(message),
});

function createMpvClientRuntimeService(): MpvIpcClient {
  const client = createMpvClientRuntimeServiceHandler() as MpvIpcClient;
  client.on('connection-change', ({ connected }) => {
    if (connected) {
      return;
    }
    youtube.browserPlayback.handleMpvDisconnected();
    youtube.flow.cancelWhisperGeneration();
    subtitleGenerationRuntime.handleMediaPathChange(null);
    if (youtube.shouldQuitWhenMpvCloses()) {
      requestAppQuit();
      return;
    }
    if (!youtube.flow.hasActiveSession()) {
      return;
    }
    youtube.flow.cancelActivePicker();
    overlayManager.broadcastToOverlayWindows(IPC_CHANNELS.event.youtubePickerCancel, null);
    overlayModalRuntime.handleOverlayModalClosed('youtube-track-picker');
  });
  client.on('connection-change', aniSkipRuntime.handleConnectionChange);
  client.on('media-path-change', aniSkipRuntime.handleMediaPathChange);
  client.on('time-pos-change', aniSkipRuntime.handleTimePosChange);
  client.on('client-message', aniSkipRuntime.handleClientMessage);
  pluginHost.attachMpvClient(client);
  return client;
}

function resetSubtitleSidebarEmbeddedLayoutRuntime(): void {
  sendMpvCommandRuntime(appState.mpvClient, ['set_property', 'video-margin-ratio-right', 0]);
  sendMpvCommandRuntime(appState.mpvClient, ['set_property', 'video-pan-x', 0]);
}

const overlayGeometryRuntime = createOverlayGeometryRuntime({
  overlayManager: {
    getMainWindow: () => overlayManager.getMainWindow(),
    getModalWindow: () => overlayManager.getModalWindow(),
    getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    setOverlayWindowBounds: (geometry) => overlayManager.setOverlayWindowBounds(geometry),
    setModalWindowBounds: (geometry) => overlayManager.setModalWindowBounds(geometry),
  },
  getTrackedWindowGeometry: () => appState.windowTracker?.getGeometry() ?? null,
  getTrackedWindowMediaSourceId: () => appState.windowTracker?.getTargetWindowMediaSourceId?.(),
  getTrackedWindowNativeId: () => appState.windowTracker?.getTargetWindowNativeId?.(),
  getStatsOverlayVisible: () => appState.statsOverlayVisible,
  getOverlayForegroundSeparateWindows: () => getOverlayForegroundSeparateWindows(),
  getLinuxVisibleOverlayWindowMode: () => linuxOverlayModeRuntime.mode,
  getLinuxTrackedMpvFullscreen: () => linuxOverlayModeRuntime.fullscreen,
  getLinuxTrackedMpvFullscreenChangedAtMs: () => linuxOverlayModeRuntime.fullscreenChangedAtMs,
  syncLinuxVisibleOverlayMpvFullscreenMode: (fullscreen) =>
    linuxOverlayModeRuntime.sync(fullscreen),
  getLinuxVisibleOverlayOwnerBindingKey: () => linuxOverlayModeRuntime.ownerBindingKey,
  setLinuxVisibleOverlayOwnerBindingKey: (key) => {
    linuxOverlayModeRuntime.ownerBindingKey = key;
  },
  clearVisibleOverlayX11OwnerBinding: (window) =>
    visibleOverlayInteractionRuntime.clearVisibleOverlayX11OwnerBinding(window),
  getNativeWindowHandleDecimal: (window) =>
    visibleOverlayInteractionRuntime.getNativeWindowHandleDecimal(window),
  enqueueVisibleOverlayX11OwnerBindingOperation: (window, args, onError) =>
    visibleOverlayInteractionRuntime.enqueueVisibleOverlayX11OwnerBindingOperation(
      window,
      args,
      onError,
    ),
  scheduleWindowsVisibleOverlayZOrderSyncBurst: () =>
    visibleOverlayInteractionRuntime.scheduleWindowsVisibleOverlayZOrderSyncBurst(),
  logDebug: (message, ...args) => logger.debug(message, ...args),
});

async function loadYomitanExtension(): Promise<Extension | null> {
  const extension = await yomitanExtensionRuntime.loadYomitanExtension();
  if (extension && !yomitanProfilePolicy.isExternalReadOnlyMode()) {
    await syncYomitanDefaultProfileAnkiServer();
  }
  return extension;
}

async function ensureYomitanExtensionLoaded(): Promise<Extension | null> {
  const extension = await yomitanExtensionRuntime.ensureYomitanExtensionLoaded();
  if (extension && !yomitanProfilePolicy.isExternalReadOnlyMode()) {
    await syncYomitanDefaultProfileAnkiServer();
  }
  return extension;
}

const { syncYomitanDefaultProfileAnkiServer } = createYomitanAnkiServerSyncRuntime({
  isExternalReadOnlyMode: () => yomitanProfilePolicy.isExternalReadOnlyMode(),
  getResolvedConfig: () => configService.getConfig(),
  getYomitanParserRuntimeDeps: () => getYomitanParserRuntimeDeps(),
  logError: (message, ...args) => {
    logger.error(message, ...args);
  },
  logInfo: (message, ...args) => {
    logger.info(message, ...args);
  },
});

function getYomitanParserRuntimeDeps() {
  return {
    getYomitanExt: () => appState.yomitanExt,
    getYomitanSession: () => appState.yomitanSession,
    getYomitanParserWindow: () => appState.yomitanParserWindow,
    setYomitanParserWindow: (window: BrowserWindow | null) => {
      appState.yomitanParserWindow = window;
    },
    getYomitanParserReadyPromise: () => appState.yomitanParserReadyPromise,
    setYomitanParserReadyPromise: (promise: Promise<void> | null) => {
      appState.yomitanParserReadyPromise = promise;
    },
    getYomitanParserInitPromise: () => appState.yomitanParserInitPromise,
    setYomitanParserInitPromise: (promise: Promise<boolean> | null) => {
      appState.yomitanParserInitPromise = promise;
    },
  };
}

function createModalWindow(): BrowserWindow {
  const existingWindow = overlayManager.getModalWindow();
  if (existingWindow && !existingWindow.isDestroyed()) {
    return existingWindow;
  }
  const window = createModalWindowHandler();
  overlayManager.setModalWindowBounds(overlayGeometryRuntime.getCurrentOverlayGeometry());
  return window;
}

function createMainWindow(): BrowserWindow {
  const window = createMainWindowHandler();
  if (process.platform === 'win32') {
    const overlayHwnd = visibleOverlayInteractionRuntime.getWindowsNativeWindowHandleNumber(window);
    if (!ensureWindowsOverlayTransparency(overlayHwnd)) {
      logger.warn('Failed to eagerly extend Windows overlay transparency via koffi');
    }
  }
  return window;
}

function generateMiningSentenceFurigana(
  text: string,
  highlightedText?: string,
): Promise<string | null> {
  return generateSentenceFurigana(text, highlightedText, getYomitanParserRuntimeDeps(), logger);
}

function initializeOverlayRuntime(): void {
  initializeOverlayRuntimeHandler();
  if (!(appState.initialArgs && isHeadlessInitialCommand(appState.initialArgs))) {
    overlayModalRuntime.primeModalWindow();
  }
  appState.ankiIntegration?.setRecordCardsMinedCallback(recordTrackedCardsMined);
  appState.ankiIntegration?.setMiningProvenanceProvider(() => ({
    minedAt: Date.now(),
    media: pluginHost.getMediaSnapshot(),
    subtitle: getPluginCurrentSubtitle(),
  }));
  appState.ankiIntegration?.setCardsReadyCallback((noteIds, notes) =>
    pluginHost.emitAnkiEvent('card-mined', { noteIds, notes }),
  );
  appState.ankiIntegration?.setKnownWordCacheUpdatedCallback(refreshCurrentSubtitleAnnotations);
  appState.ankiIntegration?.setSubtitleMiningContextConsumer(consumePendingSubtitleMiningContext);
  appState.ankiIntegration?.setMediaTimingReviewCallback(mediaTimingReviewRuntime.requestReview);
  appState.ankiIntegration?.setSentenceFuriganaGenerator(generateMiningSentenceFurigana);
  syncOverlayMpvSubtitleSuppression();
}

function openYomitanSettings(): boolean {
  if (activeDictionaryBackend === 'hachidori') {
    if (configService.getConfig().yomitan.externalProfilePath.trim()) {
      logger.warn('Yomitan settings unavailable while using read-only external-profile mode.');
      return false;
    }
    inactiveYomitanLoad ??= inactiveYomitanExtension
      ? Promise.resolve(inactiveYomitanExtension)
      : loadYomitanExtensionCore({
          userDataPath: USER_DATA_PATH,
          getYomitanParserWindow: () => null,
          setYomitanParserWindow: () => {},
          setYomitanParserReadyPromise: () => {},
          setYomitanParserInitPromise: () => {},
          setYomitanExtension: (extension) => {
            inactiveYomitanExtension = extension;
          },
          setYomitanSession: () => {},
        }).finally(() => {
          inactiveYomitanLoad = null;
        });
    void (
      inactiveYomitanExtension ? Promise.resolve(inactiveYomitanExtension) : inactiveYomitanLoad
    )
      .then((extension) =>
        openYomitanSettingsWindow({
          yomitanExt: extension,
          getExistingWindow: () => inactiveYomitanSettingsWindow,
          setWindow: (window) => {
            inactiveYomitanSettingsWindow = window;
          },
        }),
      )
      .catch((error: unknown) => logger.error('Failed to open Yomitan settings', error));
    return true;
  }
  if (yomitanProfilePolicy.isExternalReadOnlyMode()) {
    const message = 'Yomitan settings unavailable while using read-only external-profile mode.';
    logger.warn(
      'Yomitan settings window disabled while yomitan.externalProfilePath is configured because external profile mode is read-only.',
    );
    overlayNotificationsRuntime.showConfiguredStatusNotification(message, { variant: 'warning' });
    return false;
  }
  openYomitanSettingsHandler();
  return true;
}

function openHachidoriSettings(): void {
  void (
    activeDictionaryBackend === 'hachidori'
      ? ensureYomitanExtensionLoaded()
      : hachidoriExtensionRuntime.ensureLoaded()
  )
    .then((extension) =>
      openYomitanSettingsWindow({
        backend: 'hachidori',
        yomitanExt: extension,
        yomitanSession: getHachidoriSession(),
        getExistingWindow: () => hachidoriSettingsWindow,
        setWindow: (window) => {
          hachidoriSettingsWindow = window;
        },
        onWindowClosed: () => {
          if (activeDictionaryBackend === 'hachidori' && appState.yomitanParserWindow) {
            clearYomitanParserCachesForWindow(appState.yomitanParserWindow);
          }
        },
      }),
    )
    .catch((error: unknown) => logger.error('Failed to open Hachidori settings', error));
}

function openDictionarySettings(): boolean {
  if (activeDictionaryBackend === 'hachidori') {
    openHachidoriSettings();
    return true;
  }
  return openYomitanSettings();
}

const { exportLogsFromTray } = createLogExportTrayRuntime({
  flushMpvLog: () => flushMpvLog(),
  logInfo: (message) => logger.info(message),
  logWarn: (message, details) => logger.warn(message, details),
});

const {
  getConfiguredShortcuts,
  registerGlobalShortcuts,
  refreshGlobalAndOverlayShortcuts,
  cancelPendingMultiCopy,
  startPendingMultiCopy,
  cancelPendingMineSentenceMultiple,
  startPendingMineSentenceMultiple,
  syncOverlayShortcuts,
  refreshOverlayShortcuts,
} = composeShortcutRuntimes({
  globalShortcuts: {
    getConfiguredShortcutsMainDeps: {
      getResolvedConfig: () => configService.getConfig(),
      defaultConfig: DEFAULT_CONFIG,
      resolveConfiguredShortcuts,
    },
    buildRegisterGlobalShortcutsMainDeps: (getConfiguredShortcutsHandler) => ({
      getConfiguredShortcuts: () => getConfiguredShortcutsHandler(),
      registerGlobalShortcutsCore,
      onToggleVisibleOverlay: () => toggleVisibleOverlay(),
      onOpenYomitanSettings: () => openDictionarySettings(),
      isDev,
      getMainWindow: () => overlayManager.getMainWindow(),
    }),
    buildRefreshGlobalAndOverlayShortcutsMainDeps: (registerGlobalShortcutsHandler) => ({
      unregisterAllGlobalShortcuts: () => globalShortcut.unregisterAll(),
      registerGlobalShortcuts: () => registerGlobalShortcutsHandler(),
      syncOverlayShortcuts: () => syncOverlayShortcuts(),
    }),
  },
  numericShortcutRuntimeMainDeps: {
    globalShortcut,
    showMpvOsd: (text) => overlayNotificationsRuntime.showConfiguredStatusNotification(text),
    setTimer: (handler, timeoutMs) => setTimeout(handler, timeoutMs),
    clearTimer: (timer) => clearTimeout(timer),
  },
  numericSessions: {
    onMultiCopyDigit: (count) => handleMultiCopyDigitHandler(count),
    onMineSentenceDigit: (count) => handleMineSentenceDigitHandler(count),
    tryBeginMultiCopyOverlaySelection: (timeoutMs) =>
      tryBeginVisibleOverlayNumericSelection({
        actionId: 'copySubtitleMultiple',
        timeoutMs,
        getMainWindow: () => overlayManager.getMainWindow(),
        getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
      }),
    tryBeginMineSentenceOverlaySelection: (timeoutMs) =>
      tryBeginVisibleOverlayNumericSelection({
        actionId: 'mineSentenceMultiple',
        timeoutMs,
        getMainWindow: () => overlayManager.getMainWindow(),
        getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
      }),
  },
  overlayShortcutsRuntimeMainDeps: {
    overlayShortcutsRuntime,
  },
});

const { persistSessionBindings, refreshCurrentSessionBindings, refreshMpvSessionBindings } =
  createSessionBindingsRuntime({
    configDir: CONFIG_DIR,
    getKeybindings: () => appState.keybindings,
    getConfiguredShortcuts: () => getConfiguredShortcuts(),
    getResolvedConfig: () => configService.getConfig(),
    getMpvClient: () => appState.mpvClient,
    setSessionBindings: (bindings) => {
      appState.sessionBindings = bindings;
    },
    setSessionBindingsInitialized: (initialized) => {
      appState.sessionBindingsInitialized = initialized;
    },
    logWarn: (message) => logger.warn(message),
    onBindingsChanged: (bindings) =>
      overlayManager.broadcastToOverlayWindows(IPC_CHANNELS.event.sessionBindingsChanged, bindings),
    onMpvInputBindingsChanged: () =>
      overlayManager.broadcastToOverlayWindows(IPC_CHANNELS.event.mpvInputBindingsChanged),
    onWarning: (warning) => {
      if (warning.kind !== 'conflict') return;
      overlayNotificationsRuntime.showOverlayNotification({
        id: `session-binding-conflict:${warning.path}`,
        title: 'Shortcut conflict',
        body: warning.message,
        variant: 'warning',
      });
    },
  });

const { flushMpvLog, showMpvOsd } = createMpvOsdRuntimeHandlers({
  appendToMpvLogMainDeps: {
    getLogPath: () => (isLogFileEnabled('mpv') ? DEFAULT_MPV_LOG_PATH : ''),
    dirname: (targetPath) => path.dirname(targetPath),
    mkdir: async (targetPath, options) => {
      await fs.promises.mkdir(targetPath, options);
    },
    appendFile: async (targetPath, data, options) => {
      await fs.promises.appendFile(targetPath, data, options);
    },
    now: () => new Date(),
  },
  buildShowMpvOsdMainDeps: (appendToMpvLogHandler) => ({
    appendToMpvLog: (message) => appendToMpvLogHandler(message),
    showMpvOsdRuntime: (mpvClient, text, fallbackLog) =>
      showMpvOsdRuntime(mpvClient, text, fallbackLog),
    getMpvClient: () => appState.mpvClient,
    logInfo: (line) => logger.info(line),
  }),
});
flushPendingMpvLogWrites = () => {
  void flushMpvLog();
};

const { getChangelogSnapshot } = createChangelogRuntime({
  getInstalledVersion: () => app.getVersion(),
  getUpdateChannel: () => configService.getConfig().updates.channel,
  resourcesPath: process.resourcesPath,
  appPath: app.getAppPath(),
  dirname: __dirname,
  joinPath: (...parts) => path.join(...parts),
  fileExists: (candidate) => fs.existsSync(candidate),
  readFile: (candidate) => fs.readFileSync(candidate, 'utf8'),
  logWarn: (message) => logger.warn(message),
});

const { getUpdateService, takePendingLauncherMigrationPath } = createUpdateServiceRuntime({
  userDataPath: USER_DATA_PATH,
  getUpdatesConfig: () => configService.getConfig().updates,
  logInfo: (message) => logger.info(message),
  logWarn: (message, details) => logger.warn(message, details),
  showOverlayNotification: (payload) =>
    overlayNotificationsRuntime.showOverlayNotification(payload),
  showDesktopNotification: (title, options) => showDesktopNotification(title, options),
  showMpvOsd: (message) => {
    showMpvOsd(message);
  },
  withStatsWindowLayerSuspendedForNativeDialog: (showDialog) =>
    withStatsWindowLayerSuspendedForNativeDialog(showDialog),
});

const cycleSecondarySubMode = createCycleSecondarySubModeRuntimeHandler({
  cycleSecondarySubModeMainDeps: {
    getSecondarySubMode: () => appState.secondarySubMode,
    setSecondarySubMode: (mode: SecondarySubMode) => {
      setSecondarySubMode(mode);
    },
    getLastSecondarySubToggleAtMs: () => appState.lastSecondarySubToggleAtMs,
    setLastSecondarySubToggleAtMs: (timestampMs: number) => {
      appState.lastSecondarySubToggleAtMs = timestampMs;
    },
    broadcastSecondarySubMode: (mode) => {
      overlayManager.broadcastToOverlayWindows('secondary-subtitle:mode', mode);
    },
    showMpvOsd: (text: string) => showConfiguredPlaybackFeedback(text),
  },
  cycleSecondarySubMode: (deps) => cycleSecondarySubModeCore(deps),
});

function setSecondarySubMode(mode: SecondarySubMode): void {
  appState.secondarySubMode = mode;
}

const updateLastCardFromClipboardHandler = createUpdateLastCardFromClipboardHandler({
  getAnkiIntegration: () => appState.ankiIntegration,
  readClipboardText: () => clipboard.readText(),
  showMpvOsd: (text) => overlayNotificationsRuntime.showConfiguredStatusNotification(text),
  updateLastCardFromClipboardCore,
});

const refreshKnownWordCacheHandler = createRefreshKnownWordCacheHandler({
  getAnkiIntegration: () => appState.ankiIntegration,
  missingIntegrationMessage: 'AnkiConnect integration not enabled',
});

const triggerFieldGroupingHandler = createTriggerFieldGroupingHandler({
  getAnkiIntegration: () => appState.ankiIntegration,
  showMpvOsd: (text) => overlayNotificationsRuntime.showConfiguredStatusNotification(text),
  triggerFieldGroupingCore,
});

const markLastCardAsAudioCardHandler = createMarkLastCardAsAudioCardHandler({
  getAnkiIntegration: () => appState.ankiIntegration,
  showMpvOsd: (text) => overlayNotificationsRuntime.showConfiguredStatusNotification(text),
  markLastCardAsAudioCardCore,
});

const mineSentenceCardHandler = createMineSentenceCardHandler({
  getAnkiIntegration: () => appState.ankiIntegration,
  getMpvClient: () => appState.mpvClient,
  getPrimarySubtitle: () => captureCurrentPrimarySubtitleMiningContext(),
  showMpvOsd: (text) => overlayNotificationsRuntime.showConfiguredStatusNotification(text),
  mineSentenceCardCore,
});

const handleMultiCopyDigitHandler = createHandleMultiCopyDigitHandler({
  getSubtitleTimingTracker: () => appState.subtitleTimingTracker,
  writeClipboardText: (text) => clipboard.writeText(text),
  showMpvOsd: (text) => showConfiguredPlaybackFeedback(text),
  handleMultiCopyDigitCore,
});

const copyCurrentSubtitleHandler = createCopyCurrentSubtitleHandler({
  getSubtitleTimingTracker: () => appState.subtitleTimingTracker,
  writeClipboardText: (text) => clipboard.writeText(text),
  showMpvOsd: (text) => overlayNotificationsRuntime.showConfiguredStatusNotification(text),
  copyCurrentSubtitleCore,
});

const handleMineSentenceDigitHandler = createHandleMineSentenceDigitHandler({
  getSubtitleTimingTracker: () => appState.subtitleTimingTracker,
  getAnkiIntegration: () => appState.ankiIntegration,
  getCurrentSecondarySubText: () => appState.mpvClient?.currentSecondarySubText || undefined,
  showMpvOsd: (text) => overlayNotificationsRuntime.showConfiguredStatusNotification(text),
  logError: (message, err) => {
    logger.error(message, err);
  },
  handleMineSentenceDigitCore,
});
const {
  setVisibleOverlayVisible: setVisibleOverlayVisibleHandler,
  toggleVisibleOverlay: toggleVisibleOverlayHandler,
  setOverlayVisible: setOverlayVisibleHandler,
} = createOverlayVisibilityRuntime({
  setVisibleOverlayVisibleDeps: {
    setVisibleOverlayVisibleCore,
    setVisibleOverlayVisibleState: (nextVisible) => {
      overlayManager.setVisibleOverlayVisible(nextVisible);
    },
    updateVisibleOverlayVisibility: () => overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
  },
  getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
});

const handleOverlayModalClosedHandler = createHandleOverlayModalClosedHandler({
  handleOverlayModalClosedRuntime: (modal) => overlayModalRuntime.handleOverlayModalClosed(modal),
});

const appendClipboardVideoToQueueHandler = createAppendClipboardVideoToQueueHandler({
  appendClipboardVideoToQueueRuntime,
  getMpvClient: () => appState.mpvClient,
  readClipboardText: () => clipboard.readText(),
  showMpvOsd: (text) => overlayNotificationsRuntime.showConfiguredStatusNotification(text),
  sendMpvCommand: (command) => {
    sendMpvCommandRuntime(appState.mpvClient, command);
  },
});

async function dispatchSessionAction(request: SessionActionDispatchRequest): Promise<void> {
  await dispatchSessionActionCore(request, {
    toggleStatsOverlay: async () =>
      await toggleStatsOverlayWindow({
        preloadPath: statsPreloadPath,
        getApiBaseUrl: async () => (await ensureStatsServerStarted()).url,
        onStartupError: (error) =>
          overlayNotificationsRuntime.showConfiguredStatusNotification(
            `Stats server startup failed: ${error instanceof Error ? error.message : String(error)}`,
            { title: 'Stats' },
          ),
        getToggleKey: () => configService.getConfig().stats.toggleKey,
        resolveBounds: () => overlayGeometryRuntime.getCurrentOverlayGeometry(),
        onVisibilityChanged: (visible) => {
          visibleOverlayInteractionRuntime.handleStatsOverlayVisibilityChanged(visible);
        },
      }),
    toggleVisibleOverlay: () => toggleVisibleOverlay(),
    copyCurrentSubtitle: () => copyCurrentSubtitleHandler(),
    copySubtitleCount: (count) => handleMultiCopyDigitHandler(count),
    updateLastCardFromClipboard: () => updateLastCardFromClipboardHandler(),
    triggerFieldGrouping: () => triggerFieldGroupingHandler(),
    triggerSubsyncFromConfig: () => subsyncRuntime.triggerFromConfig(),
    mineSentenceCard: () => mineSentenceCardHandler(),
    mineSentenceCount: (count) => handleMineSentenceDigitHandler(count),
    toggleSecondarySub: () => cycleSecondarySubMode(),
    toggleSubtitleSidebar: () =>
      overlayManager.broadcastToOverlayWindows(IPC_CHANNELS.event.subtitleSidebarToggle),
    toggleNotificationHistory: () => toggleNotificationHistoryPanel(),
    appendClipboardVideoToQueue: () => {
      appendClipboardVideoToQueueHandler();
    },
    markLastCardAsAudioCard: () => markLastCardAsAudioCardHandler(),
    markActiveVideoWatched: async () => {
      ensureImmersionTrackerStarted();
      const marked = (await appState.immersionTracker?.markActiveVideoWatched()) ?? false;
      if (marked) {
        try {
          await anilist.maybeRunAnilistPostWatchUpdate({ force: true });
        } catch (error) {
          logger.warn('Failed to run AniList post-watch update after manual watched mark:', error);
        }
      }
      return marked;
    },
    openRuntimeOptionsPalette: () => overlayModals.openRuntimeOptionsPalette(),
    openJimaku: () => overlayModals.openJimakuOverlay(),
    openTsukihime: () => overlayModals.openTsukihimeOverlay(),
    openSessionHelp: () => overlayModals.openSessionHelpOverlay(),
    openSubtitleSelection: () => {
      if (!configService.getConfig().subtitleSelection.enabled) return;
      overlayModals.openOverlayHostedModalWithOsd(
        openSubtitleSelectionModal,
        'Subtitle selection overlay unavailable.',
        'Failed to open subtitle selection overlay.',
      );
    },
    openSubtitleGeneration: () => overlayModals.openSubtitleGenerationOverlay(),
    openCharacterDictionaryManager: () => overlayModals.openCharacterDictionaryManagerOverlay(),
    openControllerSelect: () => overlayModals.openControllerSelectOverlay(),
    openControllerDebug: () => overlayModals.openControllerDebugOverlay(),
    openYoutubeTrackPicker: () => youtube.openTrackPickerFromPlayback(),
    openPlaylistBrowser: () => overlayModals.openPlaylistBrowser(),
    replayCurrentSubtitle: () => replayCurrentSubtitleRuntime(appState.mpvClient),
    playNextSubtitle: () => playNextSubtitleRuntime(appState.mpvClient),
    cycleRuntimeOption: (id, direction) => {
      if (!appState.runtimeOptionsManager) {
        return { ok: false, error: 'Runtime options manager unavailable' };
      }
      return applyRuntimeOptionResultRuntime(
        appState.runtimeOptionsManager.cycleOption(id, direction),
        (text) => showConfiguredPlaybackFeedback(text),
      );
    },
    playNextPlaylistItem: () =>
      sendMpvCommandRuntime(appState.mpvClient, ['playlist-next', 'force']),
    runPluginAction: async (pluginActionId) => {
      const ran = await pluginHost.runAction(pluginActionId);
      if (!ran) logger.warn(`Plugin action is not registered: ${pluginActionId}`);
      return ran;
    },
    showMpvOsd: (text) => showConfiguredPlaybackFeedback(text),
  });
}

const { playlistBrowserMainDeps } = createPlaylistBrowserIpcRuntime(() => appState.mpvClient, {
  getPrimarySubtitleLanguages: () => configService.getConfig().youtube.primarySubLanguages,
  getSecondarySubtitleLanguages: () => configService.getConfig().secondarySub.secondarySubLanguages,
});

const { registerIpcRuntimeHandlers } = composeIpcRuntimeHandlers({
  mpvCommandMainDeps: {
    triggerSubsyncFromConfig: () => subsyncRuntime.triggerFromConfig(),
    openRuntimeOptionsPalette: () => overlayModals.openRuntimeOptionsPalette(),
    openJimaku: () => overlayModals.openJimakuOverlay(),
    openTsukihime: () => overlayModals.openTsukihimeOverlay(),
    openYoutubeTrackPicker: () => youtube.openTrackPickerFromPlayback(),
    openPlaylistBrowser: () => overlayModals.openPlaylistBrowser(),
    cycleRuntimeOption: (id, direction) => {
      if (!appState.runtimeOptionsManager) {
        return { ok: false, error: 'Runtime options manager unavailable' };
      }
      return applyRuntimeOptionResultRuntime(
        appState.runtimeOptionsManager.cycleOption(id, direction),
        (text) => showConfiguredPlaybackFeedback(text),
      );
    },
    showMpvOsd: (text: string) =>
      overlayNotificationsRuntime.showConfiguredStatusNotification(text),
    showRawMpvOsd: (text: string) => showMpvOsd(text),
    showPlaybackFeedback: (text: string) => showConfiguredPlaybackFeedback(text),
    replayCurrentSubtitle: () => replayCurrentSubtitleRuntime(appState.mpvClient),
    playNextSubtitle: () => playNextSubtitleRuntime(appState.mpvClient),
    sendMpvCommand: (rawCommand: (string | number)[]) => sendRendererMpvCommand(rawCommand),
    getMpvClient: () => appState.mpvClient,
    isMpvConnected: () => Boolean(appState.mpvClient && appState.mpvClient.connected),
    hasRuntimeOptionsManager: () => appState.runtimeOptionsManager !== null,
  },
  handleMpvCommandFromIpcRuntime,
  runSubsyncManualFromIpc: (request) => subsyncRuntime.runManualFromIpc(request),
  registration: {
    runtimeOptions: {
      getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
      showMpvOsd: (text: string) => showConfiguredPlaybackFeedback(text),
    },
    mainDeps: {
      previewMediaTimingReview: (request) => mediaTimingReviewRuntime.previewRange(request),
      getMediaTimingReviewFrame: (request) => mediaTimingReviewRuntime.getFrame(request),
      getMediaTimingReviewWaveform: (request) => mediaTimingReviewRuntime.getWaveform(request),
      stopMediaTimingReviewPreview: (reviewId) => mediaTimingReviewRuntime.stopPreview(reviewId),
      resolveMediaTimingReview: (request) => mediaTimingReviewRuntime.resolveReview(request),
      getMainWindow: () => overlayManager.getMainWindow(),
      focusMainWindow: () => {
        const mainWindow = overlayManager.getMainWindow();
        if (!mainWindow || mainWindow.isDestroyed()) return;
        if (process.platform === 'darwin') {
          focusMacOSOverlayWindow({
            platform: process.platform,
            getOverlayWindow: () => mainWindow,
            stealAppFocus: () => app.focus({ steal: true }),
            warn: (message, details) => logger.warn(message, details),
          });
          return;
        }
        if (!mainWindow.isFocused()) {
          mainWindow.focus();
        }
      },
      activatePlaybackWindowForOverlayInteraction: async () => {
        try {
          const raised = (await appState.windowTracker?.raiseTargetWindow?.()) ?? false;
          overlayGeometryRuntime.enforceOverlayLayerOrder();
          return raised;
        } catch (error) {
          logger.debug(
            'Failed to raise tracked mpv window for overlay interaction:',
            error instanceof Error ? error.message : String(error),
          );
          overlayGeometryRuntime.enforceOverlayLayerOrder();
          return false;
        }
      },
      onOverlayModalClosed: (modal, senderWindow, instanceId) => {
        // A late close for a replaced plugin panel must not hide the panel now showing.
        if (
          modal === 'plugin-panel' &&
          instanceId !== undefined &&
          !pluginPanelRuntime.handleModalClosed(instanceId)
        ) {
          return;
        }
        if (modal === 'media-timing-review') {
          void mediaTimingReviewRuntime.dispose();
        }
        if (modal === 'subtitle-sidebar' && senderWindow === overlayManager.getMainWindow()) {
          subtitleSidebarRequestedOpen = false;
        }
        const modalWindow = overlayManager.getModalWindow();
        if (
          senderWindow &&
          modalWindow &&
          senderWindow === modalWindow &&
          !senderWindow.isDestroyed()
        ) {
          applyOverlayClickThrough(senderWindow);
          senderWindow.hide();
        }
        handleOverlayModalClosedHandler(modal);
      },
      onOverlayModalOpened: (modal, senderWindow) => {
        if (modal === 'subtitle-sidebar' && senderWindow === overlayManager.getMainWindow()) {
          subtitleSidebarRequestedOpen = true;
        }
        overlayModalRuntime.notifyOverlayModalOpened(modal);
      },
      onOverlayMouseInteractionChanged: (active, senderWindow) => {
        const mainWindow = overlayManager.getMainWindow();
        if (!mainWindow || senderWindow !== mainWindow) {
          return;
        }
        const previousActive =
          visibleOverlayInteractionRuntime.getVisibleOverlayInteractionActive();
        visibleOverlayInteractionRuntime.setVisibleOverlayInteractionActive(active);
        if (previousActive === active) {
          if (active && process.platform === 'darwin' && !mainWindow.isFocused()) {
            overlayVisibilityRuntime.updateVisibleOverlayVisibility();
          }
          return;
        }
        overlayVisibilityRuntime.updateVisibleOverlayVisibility();
      },
      onOverlayInteractiveHint: (interactive, senderWindow) => {
        const mainWindow = overlayManager.getMainWindow();
        if (!mainWindow || senderWindow !== mainWindow) {
          return;
        }
        visibleOverlayInteractionRuntime.setLinuxOverlayInteractiveHint(interactive);
        visibleOverlayInteractionRuntime.applyLinuxOverlayInputShapeFromLatestMeasurement();
      },
      handleOverlayNotificationAction: (notificationId, actionId, noteId) => {
        if (
          notificationId === UPDATE_AVAILABLE_NOTIFICATION_ID &&
          actionId === INSTALL_UPDATE_ACTION_ID
        ) {
          void getUpdateService()
            .checkForUpdates({
              source: 'manual',
              installWhenAvailable: true,
            })
            .catch((error) => {
              logger.warn('Failed to install update from overlay notification action:', error);
            });
        }
        if (
          notificationId === UPDATE_AVAILABLE_NOTIFICATION_ID &&
          actionId === VIEW_CHANGELOG_ACTION_ID
        ) {
          overlayModals.openChangelogOverlay();
        }
        if (actionId === OPEN_ANKI_CARD_ACTION_ID && noteId !== undefined) {
          void openAnkiCardFromNotification(noteId).catch((error) => {
            logger.warn('Failed to open Anki card from overlay notification action:', error);
            overlayNotificationsRuntime.showConfiguredStatusNotification(
              'Failed to open Anki card in Anki.',
              {
                id: 'open-anki-card-failed',
                variant: 'error',
              },
            );
          });
        }
      },
      onYoutubePickerResolve: (request) => youtube.flow.resolveActivePicker(request),
      // Overlay lookups carry no cue context of their own; fall back to snapshotting the
      // live mpv sub timings at lookup time so media generation clips the mined line even
      // when extraction finishes long after playback has moved on.
      recordSubtitleMiningContext: (context) =>
        recordSubtitleMiningContext(
          context
            ? shiftSubtitleRangeToVideoClock(context, appState.mpvClient?.currentSubDelay ?? 0)
            : captureCurrentPrimarySubtitleMiningContext(),
        ),
      tokenizeCurrentSubtitle: async () => {
        const tokenizeSubtitleForCurrent = tokenizeSubtitleDeferred;
        return resolveCurrentSubtitleForRenderer({
          currentSubText: appState.currentSubText,
          currentSubtitleData: appState.currentSubtitleData,
          withCurrentSubtitleTiming: (payload) => withCurrentSubtitleTiming(payload),
          tokenizeUncached: false,
          tokenizeSubtitle: tokenizeSubtitleForCurrent
            ? (text) => tokenizeSubtitleForCurrent(text)
            : undefined,
          onResolvedSubtitle: (payload) => {
            appState.currentSubtitleData = payload;
            autoplayReadyGate.maybeSignalPluginAutoplayReady(payload, { forceWhilePaused: true });
          },
        });
      },
      getCurrentSubtitleRaw: () => appState.currentSubText,
      getSubtitleSidebarOpen: () => subtitleSidebarRequestedOpen,
      getSubtitleSidebarPlaybackState: () => getSubtitleSidebarPlaybackState(),
      // Full snapshot: resolves the active source over mpv IPC and re-parses it when it changed.
      // The renderer fetches it on open, on source/cue changes, and on a slow fallback cadence.
      getSubtitleSidebarSnapshot: async () => {
        const config = configService.getConfig().subtitleSidebar;
        const cachedSnapshot = () => ({
          ...getSubtitleSidebarPlaybackState(),
          cues: appState.activeParsedSubtitleCues,
          config,
        });
        const client = appState.mpvClient;
        if (!client?.connected) {
          return cachedSnapshot();
        }

        try {
          const [currentExternalFilenameRaw, currentTrackRaw, trackListRaw, sidRaw, videoPathRaw] =
            await Promise.all([
              client.requestProperty('current-tracks/sub/external-filename').catch(() => null),
              client.requestProperty('current-tracks/sub').catch(() => null),
              client.requestProperty('track-list'),
              client.requestProperty('sid'),
              client.requestProperty('path'),
            ]);
          const videoPath = typeof videoPathRaw === 'string' ? videoPathRaw : '';
          if (!videoPath) {
            return cachedSnapshot();
          }
          if (
            shouldUseCachedYoutubeParsedCues({
              videoPath,
              cachedMediaPath: appState.activeParsedSubtitleMediaPath,
              cachedCueCount: appState.activeParsedSubtitleCues.length,
            })
          ) {
            return cachedSnapshot();
          }

          const resolvedSource = await resolveActiveSubtitleSidebarSourceHandler({
            currentExternalFilenameRaw,
            currentTrackRaw,
            trackListRaw,
            sidRaw,
            videoPath,
          });
          if (!resolvedSource) {
            return cachedSnapshot();
          }

          try {
            if (appState.activeParsedSubtitleSource === resolvedSource.sourceKey) {
              return cachedSnapshot();
            }

            const content = await loadSubtitleSourceText(resolvedSource.path);
            const cues = parseSubtitleCues(content, resolvedSource.path);
            appState.activeParsedSubtitleCues = cues;
            appState.activeParsedSubtitleSource = resolvedSource.sourceKey;
            appState.activeParsedSubtitleMediaPath = videoPath || null;
            return cachedSnapshot();
          } finally {
            await resolvedSource.cleanup?.();
          }
        } catch {
          return cachedSnapshot();
        }
      },
      getPlaybackPaused: () =>
        resolveFreshPlaybackPaused({
          getCachedPlaybackPaused: () => appState.playbackPaused,
          getMpvClient: () => appState.mpvClient,
        }),
      getSubtitlePosition: () => loadSubtitlePosition(),
      getSubtitleStyle: () => {
        const resolvedConfig = configService.getConfig();
        return resolveSubtitleStyleForRenderer(resolvedConfig);
      },
      saveSubtitlePosition: (position) => saveSubtitlePosition(position),
      getMpvInputBindings: async () => {
        await refreshMpvSessionBindings();
        return readMpvInputBindings({
          getMpvClient: () => appState.mpvClient,
          getConfiguredKeybindings: () => configService.getConfig().keybindings ?? [],
          platform:
            process.platform === 'darwin'
              ? 'darwin'
              : process.platform === 'win32'
                ? 'win32'
                : 'linux',
        });
      },
      getSessionBindings: async () => {
        await refreshMpvSessionBindings();
        return appState.sessionBindings;
      },
      getConfiguredShortcuts: () => getConfiguredShortcuts(),
      dispatchSessionAction: (request) => dispatchSessionAction(request),
      getStatsToggleKey: () => configService.getConfig().stats.toggleKey,
      getMarkWatchedKey: () => configService.getConfig().stats.markWatchedKey,
      getOverlayNotificationPosition: () => configService.getConfig().notifications.overlayPosition,
      getControllerConfig: () => configService.getConfig().controller,
      saveControllerConfig: (update) => {
        const currentRawConfig = configService.getRawConfig();
        configService.patchRawConfig({
          controller: applyControllerConfigUpdate(currentRawConfig.controller, update),
        });
      },
      getSecondarySubMode: () => appState.secondarySubMode,
      getMpvClient: () => appState.mpvClient,
      getRuntimeOptions: () => getRuntimeOptionsStateHandler(),
      reportOverlayContentBounds: (payload: unknown) => {
        if (overlayContentMeasurementStore.report(payload)) {
          visibleOverlayInteractionRuntime.tickLinuxOverlayPointerInteractionNow();
          visibleOverlayInteractionRuntime.tickWindowsOverlayPointerInteractionNow();
          visibleOverlayInteractionRuntime.primeLinuxOverlayPointerInteractionAfterFirstMeasurement();
          autoplayReadyGate.flushPendingAutoplayReadySignal();
          autoplaySubtitlePrimingRuntime.scheduleVisibleOverlaySubtitleRefreshAfterFirstPaint();
        }
      },
      getAnilistStatus: () => anilist.stateRuntime.getStatusSnapshot(),
      clearAnilistToken: () => anilist.stateRuntime.clearTokenState(),
      openAnilistSetup: () => anilist.openSetupWindow(),
      getAnilistQueueStatus: () => anilist.stateRuntime.getQueueStatusSnapshot(),
      retryAnilistQueueNow: () => anilist.processNextAnilistRetryUpdate(),
      runAnilistPostWatchUpdateOnManualMark: () =>
        anilist.maybeRunAnilistPostWatchUpdate({ force: true }),
      getCharacterDictionarySelection: (searchTitle?: string) =>
        characterDictionary.runtime.getManualSelectionSnapshot(undefined, searchTitle),
      setCharacterDictionarySelection: async (
        mediaId: number,
        replaceManagedMediaId?: number,
        mediaTitle?: string,
      ) => {
        if (replaceManagedMediaId !== undefined && mediaTitle) {
          const result = replaceCharacterDictionaryManagedEntry(
            USER_DATA_PATH,
            replaceManagedMediaId,
            {
              mediaId,
              mediaTitle,
            },
          );
          if (result.ok && result.rebuildRequired) {
            try {
              await characterDictionary.autoSync.runSyncNow();
            } catch (error) {
              logger.warn('Failed to rebuild character dictionary after manager override:', error);
            }
          }
          return result;
        }
        return await applyCharacterDictionarySelection(
          { mediaId },
          {
            setManualSelection: (request) =>
              characterDictionary.runtime.setManualSelection(request),
            resetAnilistMediaGuessState: anilist.resetAnilistMediaGuessState,
            runSyncNow: () => characterDictionary.autoSync.runSyncNow(),
            warn: (message, error) => logger.warn(message, error),
          },
        );
      },
      getCharacterDictionaryManagerSnapshot: async () =>
        getCharacterDictionaryManagerSnapshot(
          USER_DATA_PATH,
          characterDictionary.autoSync.getCurrentMediaId(),
        ),
      removeCharacterDictionaryManagedEntry: async (mediaId: number) => {
        const result = removeCharacterDictionaryManagedEntry(
          USER_DATA_PATH,
          mediaId,
          characterDictionary.autoSync.getCurrentMediaId(),
        );
        if (result.ok && result.rebuildRequired) {
          try {
            await characterDictionary.autoSync.runSyncNow();
          } catch (error) {
            logger.warn('Failed to rebuild character dictionary after manager removal:', error);
          }
        }
        return result;
      },
      moveCharacterDictionaryManagedEntry: async (mediaId: number, direction: 1 | -1) => {
        const result = moveCharacterDictionaryManagedEntry(
          USER_DATA_PATH,
          mediaId,
          direction,
          characterDictionary.autoSync.getCurrentMediaId(),
        );
        if (result.ok && result.rebuildRequired) {
          try {
            await characterDictionary.autoSync.runSyncNow();
          } catch (error) {
            logger.warn('Failed to rebuild character dictionary after manager reorder:', error);
          }
        }
        return result;
      },
      getChangelogSnapshot: (options) => getChangelogSnapshot(options),
      ...playlistBrowserMainDeps,
      getImmersionTracker: () => appState.immersionTracker,
    },
    ankiJimakuDeps: createAnkiJimakuIpcRuntimeServiceDeps({
      getResolvedConfig: () => configService.getConfig(),
      getMpvClient: () => appState.mpvClient,
      getAnkiIntegration: () => appState.ankiIntegration,
      getFieldGroupingResolver: () => getFieldGroupingResolverHandler(),
      setFieldGroupingResolver: (resolver: ((choice: KikuFieldGroupingChoice) => void) | null) =>
        setFieldGroupingResolverHandler(resolver),
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
      downloadToFile: (
        url: string,
        destPath: string,
        headers: Record<string, string>,
        downloadOptions?: { isAllowedRedirect?: (url: URL) => boolean },
      ) => downloadToFile(url, destPath, headers, downloadOptions),
    }),
    registerIpcRuntimeServices,
  },
});
const { handleCliCommand, handleInitialArgs } = composeCliStartupHandlers({
  cliCommandContextMainDeps: {
    appState,
    setLogLevel: (level) => setLogLevel(level, 'cli'),
    onMpvSocketPathChanged: (nextSocketPath, previousSocketPath) =>
      visibleOverlayInteractionRuntime.retargetOverlayWindowTrackerForMpvSocket(
        nextSocketPath,
        previousSocketPath,
      ),
    texthookerService,
    getResolvedConfig: () => configService.getConfig(),
    defaultWebsocketPort: DEFAULT_CONFIG.websocket.port,
    defaultAnnotationWebsocketPort: DEFAULT_CONFIG.annotationWebsocket.port,
    hasMpvWebsocketPlugin: () => hasMpvWebsocketPlugin(),
    openExternal: (url: string) => shell.openExternal(url),
    logBrowserOpenError: (url: string, error: unknown) =>
      logger.error(`Failed to open browser for texthooker URL: ${url}`, error),
    showMpvOsd: (text: string) =>
      overlayNotificationsRuntime.showConfiguredStatusNotification(text),
    showPlaybackFeedback: (text: string) => showConfiguredPlaybackFeedback(text),
    initializeOverlayRuntime: () => initializeOverlayRuntime(),
    toggleVisibleOverlay: () => toggleVisibleOverlay(),
    togglePrimarySubtitleBar: () =>
      overlayManager.broadcastToOverlayWindows(IPC_CHANNELS.event.primarySubtitleBarToggle),
    openFirstRunSetupWindow: (force?: boolean) => firstRunSetup.openWindow(force),
    setVisibleOverlayVisible: (visible: boolean) => setVisibleOverlayVisible(visible),
    copyCurrentSubtitle: () => copyCurrentSubtitleHandler(),
    startPendingMultiCopy: (timeoutMs: number) => startPendingMultiCopy(timeoutMs),
    mineSentenceCard: () => mineSentenceCardHandler(),
    startPendingMineSentenceMultiple: (timeoutMs: number) =>
      startPendingMineSentenceMultiple(timeoutMs),
    updateLastCardFromClipboard: () => updateLastCardFromClipboardHandler(),
    refreshKnownWordCache: () => refreshKnownWordCacheHandler(),
    triggerFieldGrouping: () => triggerFieldGroupingHandler(),
    triggerSubsyncFromConfig: () => subsyncRuntime.triggerFromConfig(),
    markLastCardAsAudioCard: () => markLastCardAsAudioCardHandler(),
    getAnilistStatus: () => anilist.stateRuntime.getStatusSnapshot(),
    clearAnilistToken: () => anilist.stateRuntime.clearTokenState(),
    openAnilistSetupWindow: () => anilist.openSetupWindow(),
    openJellyfinSetupWindow: () => jellyfin.openJellyfinSetupWindow(),
    getAnilistQueueStatus: () => anilist.stateRuntime.getQueueStatusSnapshot(),
    processNextAnilistRetryUpdate: () => anilist.processNextAnilistRetryUpdate(),
    generateCharacterDictionary: async (targetPath?: string) => {
      const disabledReason = yomitanProfilePolicy.getCharacterDictionaryDisabledReason();
      if (disabledReason) {
        throw new Error(disabledReason);
      }
      return await characterDictionary.runtime.generateForCurrentMedia(targetPath);
    },
    getCharacterDictionarySelection: async (targetPath?: string) =>
      characterDictionary.runtime.getManualSelectionSnapshot(targetPath),
    setCharacterDictionarySelection: async (request) =>
      applyCharacterDictionarySelection(request, {
        setManualSelection: (selectionRequest) =>
          characterDictionary.runtime.setManualSelection(selectionRequest),
        resetAnilistMediaGuessState: anilist.resetAnilistMediaGuessState,
        runSyncNow: () => characterDictionary.autoSync.runSyncNow(),
        warn: (message, error) => logger.warn(message, error),
      }),
    runJellyfinCommand: (argsFromCommand: CliArgs) => jellyfin.runJellyfinCommand(argsFromCommand),
    runStatsCommand: (argsFromCommand: CliArgs, source: CliCommandSource) =>
      runStatsCliCommand(argsFromCommand, source),
    runUpdateCommand: async (argsFromCommand: CliArgs, source: CliCommandSource) => {
      await runUpdateCliCommand(argsFromCommand, source, {
        checkForUpdates: (request) => getUpdateService().checkForUpdates(request),
        writeResponse: (responsePath, payload) =>
          writeUpdateCliCommandResponse(responsePath, payload),
        logWarn: (message, error) => logger.warn(message, error),
      });
    },
    runEnsureLinuxRuntimePluginAssetsCommand: async (
      argsFromCommand: CliArgs,
      source: CliCommandSource,
    ) => {
      await runEnsureLinuxRuntimePluginAssetsCliCommand(
        argsFromCommand,
        {
          ensureLinuxRuntimePluginAssets: () => ensureLinuxRuntimePluginAssets(),
          writeResponse: (responsePath, payload) =>
            writeEnsureLinuxRuntimePluginAssetsCliCommandResponse(responsePath, payload),
          logWarn: (message, error) => logger.warn(message, error),
        },
        source,
      );
    },
    runYoutubePlaybackFlow: (request) => youtube.playback.runYoutubePlaybackFlow(request),
    ensureBackgroundStatsServer: () => ensureBackgroundStatsServer(),
    openHachidoriSettings: () => openHachidoriSettings(),
    openYomitanSettings: () => openYomitanSettings(),
    openConfigSettingsWindow: () => settings.window.openWindow(),
    openSyncUiWindow: () => sync.openWindow(),
    openYoutubeBrowserWindow: () => youtube.browserWindow.open(),
    cycleSecondarySubMode: () => cycleSecondarySubMode(),
    openRuntimeOptionsPalette: () => overlayModals.openRuntimeOptionsPalette(),
    printHelp: () => printHelp(DEFAULT_TEXTHOOKER_PORT),
    stopApp: () => requestAppQuit(),
    hasMainWindow: () => Boolean(overlayManager.getMainWindow()),
    dispatchSessionAction: (request: SessionActionDispatchRequest) =>
      dispatchSessionAction(request),
    getMultiCopyTimeoutMs: () => getConfiguredShortcuts().multiCopyTimeoutMs,
    schedule: (fn: () => void, delayMs: number) => setTimeout(fn, delayMs),
    logInfo: (message: string) => logger.info(message),
    logDebug: (message: string) => logger.debug(message),
    logWarn: (message: string) => logger.warn(message),
    logError: (message: string, err: unknown) => logger.error(message, err),
  },
  cliCommandRuntimeHandlerMainDeps: {
    handleTexthookerOnlyModeTransitionMainDeps: {
      isTexthookerOnlyMode: () => appState.texthookerOnlyMode,
      ensureOverlayStartupPrereqs: () => ensureOverlayStartupPrereqs(),
      setTexthookerOnlyMode: (enabled) => {
        appState.texthookerOnlyMode = enabled;
      },
      commandNeedsOverlayStartupPrereqs: (inputArgs) =>
        commandNeedsOverlayStartupPrereqs(inputArgs),
      startBackgroundWarmups: () => startBackgroundWarmups(),
      logInfo: (message: string) => logger.info(message),
    },
    ensureTrayForCommand: (args) => {
      if (args.background || args.managedPlayback) {
        ensureTrayHandler();
      }
    },
    handleCliCommandRuntimeServiceWithContext: (args, source, cliContext) =>
      handleCliCommandRuntimeServiceWithContext(args, source, cliContext),
  },
  initialArgsRuntimeHandlerMainDeps: {
    getInitialArgs: () => appState.initialArgs,
    isBackgroundMode: () => appState.backgroundMode,
    shouldEnsureTrayOnStartup: () =>
      shouldEnsureTrayOnStartupForInitialArgs(process.platform, appState.initialArgs),
    shouldRunHeadlessInitialCommand: (args) => isHeadlessInitialCommand(args),
    ensureTray: () => ensureTrayHandler(),
    isTexthookerOnlyMode: () => appState.texthookerOnlyMode,
    hasImmersionTracker: () => Boolean(appState.immersionTracker),
    getMpvClient: () => appState.mpvClient,
    commandNeedsOverlayStartupPrereqs: (args) => commandNeedsOverlayStartupPrereqs(args),
    commandNeedsOverlayRuntime: (args) => commandNeedsOverlayRuntime(args),
    ensureOverlayStartupPrereqs: () => ensureOverlayStartupPrereqs(),
    isOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
    initializeOverlayRuntime: () => initializeOverlayRuntime(),
    logInfo: (message) => logger.info(message),
  },
});
const { runAndApplyStartupState } = composeHeadlessStartupHandlers<
  CliArgs,
  StartupState,
  ReturnType<typeof createStartupBootstrapRuntimeDeps>
>({
  startupRuntimeHandlersDeps: {
    appLifecycleRuntimeRunnerMainDeps: {
      app: appLifecycleApp,
      platform: process.platform,
      shouldStartApp: (nextArgs: CliArgs) => shouldStartApp(nextArgs),
      parseArgs: (argv: string[]) => parseArgs(argv),
      handleCliCommand: (nextArgs: CliArgs, source: CliCommandSource) =>
        handleCliCommand(nextArgs, source),
      printHelp: () => printHelp(DEFAULT_TEXTHOOKER_PORT),
      logNoRunningInstance: () => appLogger.logNoRunningInstance(),
      startControlServer: (handleArgv: (argv: string[]) => void) => {
        const server = startAppControlServer({
          socketPath: getAppControlSocketPath({ configDir: CONFIG_DIR }),
          platform: process.platform,
          handleArgv,
          logDebug: (message) => logger.debug(message),
          logWarn: (message, error) => logger.warn(message, error),
        });
        return () => server.close();
      },
      onReady: runAppReadyRuntimeWithFatalReporting,
      onWillQuitCleanup: () => onWillQuitCleanupHandler(),
      shouldRestoreWindowsOnActivate: () => shouldRestoreWindowsOnActivateHandler(),
      restoreWindowsOnActivate: () => restoreWindowsOnActivateHandler(),
      shouldQuitOnWindowAllClosed: () =>
        shouldQuitOnWindowAllClosedForTrayState({
          backgroundMode: appState.backgroundMode,
          hasTray: Boolean(appTray),
        }),
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
      getDefaultSocketPath: () =>
        configService.getConfig().mpv.socketPath || getDefaultSocketPathHandler(),
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
      quitApp: () => requestAppQuit(),
      logGenerateConfigError: (message) => logger.error(message),
      startAppLifecycle,
    }),
    createStartupBootstrapRuntimeDeps: (deps) => createStartupBootstrapRuntimeDeps(deps),
    runStartupBootstrapRuntime,
    applyStartupState: (startupState) => applyStartupState(appState, startupState),
  },
});

runAndApplyStartupState();
void app.whenReady().then(() => {
  void takePendingLauncherMigrationPath(async (pendingLauncherPath) => {
    const acknowledgedPaths = await refreshManagedCommandLineLauncher({
      ...(await createCommandLineLauncherRuntimeOptions()),
      additionalLauncherPaths: pendingLauncherPath ? [pendingLauncherPath] : [],
    });
    return pendingLauncherPath !== undefined && acknowledgedPaths.includes(pendingLauncherPath);
  }).catch((error) => {
    logger.warn('Failed to refresh the installed command-line launcher', error);
  });
  if (!shouldStartAutomaticUpdateChecks(appState.initialArgs)) {
    return;
  }
  getUpdateService().startAutomaticChecks();
});
const startupModeFlags = getStartupModeFlags(appState.initialArgs);
const shouldUseMinimalStartup = startupModeFlags.shouldUseMinimalStartup;
const shouldSkipHeavyStartup = startupModeFlags.shouldSkipHeavyStartup;
if (!appState.initialArgs || (!shouldUseMinimalStartup && !shouldSkipHeavyStartup)) {
  if (isAnilistTrackingEnabled(configService.getConfig())) {
    void anilist
      .refreshClientSecretStateIfEnabled({
        force: true,
        allowSetupPrompt: false,
      })
      .catch((error) => {
        logger.error('Failed to refresh AniList client secret state during startup', error);
      });
    anilist.stateRuntime.refreshRetryQueueState();
  }
  void initializeDiscordPresenceService().catch((error) => {
    logger.error('Failed to initialize Discord presence service during startup', error);
  });
}
const { createMainWindow: createMainWindowHandler, createModalWindow: createModalWindowHandler } =
  createOverlayWindowRuntimeHandlers<BrowserWindow>({
    createOverlayWindowDeps: {
      createOverlayWindowCore: (kind, options) => createOverlayWindowCore(kind, options),
      isDev,
      ensureOverlayWindowLevel: (window) => overlayGeometryRuntime.ensureOverlayWindowLevel(window),
      onRuntimeOptionsChanged: () => overlayVisibilityComposer.broadcastRuntimeOptionsChanged(),
      setOverlayDebugVisualizationEnabled: (enabled) =>
        overlayVisibilityComposer.setOverlayDebugVisualizationEnabled(enabled),
      isOverlayVisible: (windowKind) =>
        windowKind === 'visible' ? overlayManager.getVisibleOverlayVisible() : false,
      getYomitanSession: () =>
        activeDictionaryBackend === 'hachidori' ? getHachidoriSession() : appState.yomitanSession,
      tryHandleOverlayShortcutLocalFallback: (input) =>
        overlayShortcutsRuntime.tryHandleOverlayShortcutLocalFallback(input),
      forwardTabToMpv: () => sendMpvCommandRuntime(appState.mpvClient, ['keypress', 'TAB']),
      getLinuxX11FullscreenOverlay: () =>
        shouldRunLinuxOverlayZOrderKeepAlive() &&
        linuxOverlayModeRuntime.fullscreen &&
        linuxOverlayModeRuntime.mode === 'fullscreen-override',
      onVisibleWindowBlurred: () =>
        visibleOverlayInteractionRuntime.scheduleVisibleOverlayBlurRefresh(),
      onVisibleWindowFocused: () =>
        visibleOverlayInteractionRuntime.requestLinuxOverlayZOrderFollow(),
      onWindowDidFinishLoad: () => {
        flushQueuedOverlayNotifications();
      },
      onWindowContentReady: () => {
        overlayNotificationsRuntime.dismissOverlayLoadingStatusNotification();
        flushQueuedOverlayNotifications();
        overlayVisibilityRuntime.updateVisibleOverlayVisibility();
        visibleOverlayInteractionRuntime.primeLinuxOverlayPointerInteractionAfterFirstMeasurement();
        autoplayReadyGate.flushPendingAutoplayReadySignal();
      },
      onWindowClosed: (windowKind, window) => {
        if (windowKind === 'visible') {
          if (overlayManager.getMainWindow() !== window) {
            return;
          }
          cancelPendingLinuxMpvFullscreenOverlayRefreshBurst();
          overlayManager.setMainWindow(null);
        } else {
          if (overlayManager.getModalWindow() !== window) {
            return;
          }
          void mediaTimingReviewRuntime.dispose();
          overlayManager.setModalWindow(null);
        }
      },
    },
    getMainWindow: () => overlayManager.getMainWindow(),
    isWindowDestroyed: (window) => window.isDestroyed(),
    setMainWindow: (window) => overlayManager.setMainWindow(window),
    setModalWindow: (window) => overlayManager.setModalWindow(window),
  });

function refreshTrayMenuIfPresent(): void {
  if (appTray) {
    ensureTrayHandler();
  }
}

const { ensureTray: ensureTrayHandler, destroyTray: destroyTrayHandler } =
  createTrayRuntimeHandlers({
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
      platform: process.platform,
      initializeOverlayRuntime: () => initializeOverlayRuntime(),
      isOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
      openSessionHelpModal: () => overlayModals.openSessionHelpOverlay(),
      openChangelogModal: () => overlayModals.openChangelogOverlay(),
      openTexthookerInBrowser: () =>
        handleCliCommand(parseArgs(['--texthooker', '--open-browser'])),
      showTexthookerPage: () => shouldShowTexthookerTrayEntry(configService.getConfig()),
      showFirstRunSetup: () => !firstRunSetup.service.isSetupCompleted(),
      openFirstRunSetupWindow: (force?: boolean) => firstRunSetup.openWindow(force),
      showWindowsMpvLauncherSetup: () => process.platform === 'win32',
      getDictionaryBackend: () => activeDictionaryBackend,
      openHachidoriSettings: () => openHachidoriSettings(),
      openYomitanSettings: () => openYomitanSettings(),
      openConfigSettingsWindow: () => settings.window.openWindow(),
      openSyncUiWindow: () => sync.openWindow(),
      openYoutubeBrowserWindow: () => youtube.browserWindow.open(),
      exportLogs: () => {
        void exportLogsFromTray();
      },
      openJellyfinSetupWindow: () => jellyfin.openJellyfinSetupWindow(),
      isJellyfinConfigured: () =>
        isJellyfinConfiguredForTrayRuntime(jellyfin.getTrayDiscoveryDeps()),
      isJellyfinDiscoveryActive: () => Boolean(appState.jellyfinRemoteSession),
      toggleJellyfinDiscovery: (checked: boolean) =>
        toggleJellyfinDiscoveryFromTrayRuntime(jellyfin.getTrayDiscoveryDeps(), {
          desiredActive: checked,
        }),
      openAnilistSetupWindow: () => anilist.openSetupWindow(),
      checkForUpdates: () => {
        void getUpdateService().checkForUpdates({ source: 'manual' });
      },
      quitApp: () => requestAppQuit(),
    },
    ensureTrayDeps: {
      getTray: () => appTray,
      setTray: (tray) => {
        appTray = tray as Tray | null;
      },
      createImageFromPath: (iconPath) => nativeImage.createFromPath(iconPath),
      createEmptyImage: () => nativeImage.createEmpty(),
      createTray: (icon) => new Tray(icon as ConstructorParameters<typeof Tray>[0]),
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
const yomitanProfilePolicy = createYomitanProfilePolicy({
  externalProfilePath:
    activeDictionaryBackend === 'yomitan'
      ? configService.getConfig().yomitan.externalProfilePath
      : '',
  logInfo: (message) => logger.info(message),
});
const configuredExternalYomitanProfilePath = yomitanProfilePolicy.externalProfilePath;
const yomitanExtensionRuntime = createYomitanExtensionRuntime({
  loadYomitanExtensionCore: async (deps) => {
    if (activeDictionaryBackend === 'yomitan') return loadYomitanExtensionCore(deps);
    try {
      const extension = await hachidoriExtensionRuntime.ensureLoaded();
      deps.setYomitanExtension(extension);
      deps.setYomitanSession(getHachidoriSession());
      return extension;
    } catch (error) {
      logger.error('Failed to load Hachidori extension:', error);
      deps.setYomitanExtension(null);
      deps.setYomitanSession(null);
      return null;
    }
  },
  userDataPath: USER_DATA_PATH,
  externalProfilePath: configuredExternalYomitanProfilePath,
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
  setYomitanSession: (nextSession) => {
    appState.yomitanSession = nextSession;
  },
  getYomitanExtension: () => appState.yomitanExt,
  getLoadInFlight: () => yomitanLoadInFlight,
  setLoadInFlight: (promise) => {
    yomitanLoadInFlight = promise;
  },
  onYomitanExtensionLoaded: () => {
    const reloaded = reloadOverlayWindowsForYomitanContentScripts(
      overlayManager.getOverlayWindows(),
      (message, error) => logger.warn(message, error),
    );
    if (reloaded > 0) {
      logger.debug(`Reloaded ${reloaded} overlay window(s) after Yomitan extension load.`);
    }
  },
});
const { initializeOverlayRuntime: initializeOverlayRuntimeHandler } =
  runtimeRegistry.overlay.createOverlayRuntimeBootstrapHandlers({
    initializeOverlayRuntimeMainDeps: {
      appState,
      overlayManager: {
        getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
      },
      overlayVisibilityRuntime: {
        updateVisibleOverlayVisibility: () =>
          overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
      },
      refreshCurrentSubtitle: () => {
        subtitleProcessingController.refreshCurrentSubtitle(appState.currentSubText);
      },
      overlayShortcutsRuntime: {
        syncOverlayShortcuts: () => overlayShortcutsRuntime.syncOverlayShortcuts(),
      },
      createMainWindow: () => {
        if (appState.initialArgs && isHeadlessInitialCommand(appState.initialArgs)) {
          return;
        }
        createMainWindow();
      },
      registerGlobalShortcuts: () => {
        if (appState.initialArgs && isHeadlessInitialCommand(appState.initialArgs)) {
          return;
        }
        registerGlobalShortcuts();
      },
      createWindowTracker: (override, targetMpvSocketPath) =>
        visibleOverlayInteractionRuntime.createOverlayWindowTracker(override, targetMpvSocketPath),
      updateVisibleOverlayBounds: (geometry: WindowGeometry) =>
        overlayGeometryRuntime.updateVisibleOverlayBounds(geometry),
      bindOverlayOwner: () => visibleOverlayInteractionRuntime.bindVisibleOverlayOwner(),
      releaseOverlayOwner: () => visibleOverlayInteractionRuntime.releaseVisibleOverlayOwner(),
      getOverlayWindows: () => overlayManager.getOverlayWindows(),
      getResolvedConfig: () => configService.getConfig(),
      showDesktopNotification,
      showOverlayNotification: (payload) =>
        overlayNotificationsRuntime.showOverlayNotification(payload),
      dismissOverlayNotification: (id) =>
        overlayNotificationsRuntime.dismissOverlayNotification(id),
      createFieldGroupingCallback: () => createFieldGroupingCallback(),
      getKnownWordCacheStatePath: () => path.join(USER_DATA_PATH, 'known-words-cache.json'),
      getCachedMediaPath: (currentVideoPath, kind) =>
        youtube.getCachedMediaPathForCurrentPlayback(currentVideoPath, kind),
      shouldRequireRemoteMediaCache: () => youtube.shouldRequireMediaCacheForCurrentPlayback(),
      getYoutubeMediaSourceUrl: () => youtube.getCurrentMediaCacheSourceUrl(),
      shouldStartAnkiIntegration: () =>
        !(appState.initialArgs && isHeadlessInitialCommand(appState.initialArgs)),
    },
    initializeOverlayRuntimeBootstrapDeps: {
      isOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
      initializeOverlayRuntimeCore,
      setOverlayRuntimeInitialized: (initialized) => {
        appState.overlayRuntimeInitialized = initialized;
      },
      startBackgroundWarmups: () => {
        if (appState.initialArgs && isHeadlessInitialCommand(appState.initialArgs)) {
          return;
        }
        startBackgroundWarmups();
      },
    },
  });
const { openYomitanSettings: openYomitanSettingsHandler } = createYomitanSettingsRuntime({
  ensureYomitanExtensionLoaded: () => ensureYomitanExtensionLoaded(),
  getYomitanExtension: () => appState.yomitanExt,
  getYomitanExtensionLoadInFlight: () => yomitanLoadInFlight,
  getYomitanSession: () => appState.yomitanSession,
  openYomitanSettingsWindow: ({ yomitanExt, getExistingWindow, setWindow, yomitanSession }) => {
    openYomitanSettingsWindow({
      yomitanExt: yomitanExt as Extension,
      getExistingWindow: () => getExistingWindow() as BrowserWindow | null,
      setWindow: (window) => setWindow(window as BrowserWindow | null),
      yomitanSession: (yomitanSession as Session | null | undefined) ?? appState.yomitanSession,
      onWindowClosed: () => {
        if (appState.yomitanParserWindow) {
          clearYomitanParserCachesForWindow(appState.yomitanParserWindow);
        }
      },
    });
  },
  getExistingWindow: () => appState.yomitanSettingsWindow,
  setWindow: (window) => {
    appState.yomitanSettingsWindow = window as BrowserWindow | null;
  },
  logWarn: (message) => logger.warn(message),
  logError: (message, error) => logger.error(message, error),
});

function ensureOverlayWindowsReadyForVisibilityActions(): void {
  if (!appState.overlayRuntimeInitialized) {
    initializeOverlayRuntime();
    return;
  }

  const mainWindow = overlayManager.getMainWindow();
  if (!mainWindow || mainWindow.isDestroyed()) {
    createMainWindow();
  }
}

function notifyMpvPluginVisibleOverlayVisibility(visible: boolean): void {
  sendMpvCommandRuntime(appState.mpvClient, [
    'script-message',
    visible ? 'subminer-visible-overlay-shown' : 'subminer-visible-overlay-hidden',
  ]);
}

function setVisibleOverlayVisible(visible: boolean): void {
  ensureOverlayWindowsReadyForVisibilityActions();
  if (!visible) {
    overlayNotificationsRuntime.dismissOverlayLoadingStatusNotification();
    autoplayReadyGate.markCurrentMediaAutoplayReady();
    autoplaySubtitlePrimingRuntime.cancelVisibleOverlaySubtitleRefreshAfterFirstPaint();
    cancelPendingLinuxMpvFullscreenOverlayRefreshBurst();
    visibleOverlayInteractionRuntime.resetVisibleOverlayInputState();
  }
  if (visible) {
    maybeStartOverlayLoadingOsd();
    visibleOverlayInteractionRuntime.resetLinuxVisibleOverlayStartupInputPrimer();
    visibleOverlayInteractionRuntime.startLinuxVisibleOverlayStartupInputGrace();
    visibleOverlayInteractionRuntime.restoreVisibleOverlayWindowShapeForShow();
    void ensureOverlayMpvSubtitlesHidden();
    void autoplaySubtitlePrimingRuntime.primeCurrentSubtitleForVisibleOverlay();
  }
  setVisibleOverlayVisibleHandler(visible);
  notifyMpvPluginVisibleOverlayVisibility(visible);
  syncOverlayMpvSubtitleSuppression();
}

function toggleVisibleOverlay(): void {
  ensureOverlayWindowsReadyForVisibilityActions();
  const nextVisible = !overlayManager.getVisibleOverlayVisible();
  if (!nextVisible) {
    overlayNotificationsRuntime.dismissOverlayLoadingStatusNotification();
    autoplayReadyGate.markCurrentMediaAutoplayReady();
    autoplaySubtitlePrimingRuntime.cancelVisibleOverlaySubtitleRefreshAfterFirstPaint();
    cancelPendingLinuxMpvFullscreenOverlayRefreshBurst();
    visibleOverlayInteractionRuntime.resetVisibleOverlayInputState();
  } else {
    maybeStartOverlayLoadingOsd();
    visibleOverlayInteractionRuntime.resetLinuxVisibleOverlayStartupInputPrimer();
    visibleOverlayInteractionRuntime.startLinuxVisibleOverlayStartupInputGrace();
    visibleOverlayInteractionRuntime.restoreVisibleOverlayWindowShapeForShow();
    void ensureOverlayMpvSubtitlesHidden();
    void autoplaySubtitlePrimingRuntime.primeCurrentSubtitleForVisibleOverlay();
  }
  toggleVisibleOverlayHandler();
  notifyMpvPluginVisibleOverlayVisibility(nextVisible);
  syncOverlayMpvSubtitleSuppression();
}
function setOverlayVisible(visible: boolean): void {
  if (!visible) {
    overlayNotificationsRuntime.dismissOverlayLoadingStatusNotification();
    autoplaySubtitlePrimingRuntime.cancelVisibleOverlaySubtitleRefreshAfterFirstPaint();
    visibleOverlayInteractionRuntime.resetVisibleOverlayInputState();
    autoplayReadyGate.markCurrentMediaAutoplayReady();
    cancelPendingLinuxMpvFullscreenOverlayRefreshBurst();
  }
  if (visible) {
    maybeStartOverlayLoadingOsd();
    visibleOverlayInteractionRuntime.resetLinuxVisibleOverlayStartupInputPrimer();
    visibleOverlayInteractionRuntime.startLinuxVisibleOverlayStartupInputGrace();
    visibleOverlayInteractionRuntime.restoreVisibleOverlayWindowShapeForShow();
    void ensureOverlayMpvSubtitlesHidden();
    void autoplaySubtitlePrimingRuntime.primeCurrentSubtitleForVisibleOverlay();
  }
  setOverlayVisibleHandler(visible);
  notifyMpvPluginVisibleOverlayVisibility(visible);
  syncOverlayMpvSubtitleSuppression();
}

registerIpcRuntimeHandlers();
registerSubtitleSelectionIpc({
  ipc: ipcMain,
  isAllowedSender: (sender) =>
    [overlayManager.getMainWindow(), overlayManager.getModalWindow()].some(
      (window) => window && !window.isDestroyed() && window.webContents === sender,
    ),
  runtime: createSubtitleSelectionRuntime({
    isEnabled: () => configService.getConfig().subtitleSelection.enabled,
    getMpvClient: () => appState.mpvClient,
  }),
});
const subtitleGenerationRuntime = createSubtitleGenerationRuntime({
  getConfig: () => configService.getConfig().subtitleGeneration,
  getModelDirectory: getSubtitleGenerationModelDirectory,
  generateYoutube: async ({ url, config, signal, onProgress }) =>
    youtube.whisperSubtitles.generate({
      url,
      config,
      signal,
      onProgress,
      outputPath: path.join(await youtube.flow.createSubtitleTempDir(), 'youtube-whisper.ja.srt'),
    }),
  resolveYoutubeSourceUrl: (mediaPath) => youtube.playback.getYoutubeSourceUrlForStream(mediaPath),
  resolveJellyfinMedia: (mediaPath) => {
    const playback = jellyfin.getActivePlayback();
    const sourceUrl = playback && toJellyfinOriginalStreamUrl(mediaPath, playback.itemId);
    return sourceUrl
      ? {
          sourceUrl,
          transcoded: playback.playMethod === 'Transcode',
          audioStreamIndex: playback.audioStreamIndex ?? null,
        }
      : null;
  },
  getMpvClient: () => appState.mpvClient,
  onProgress: (progress) => {
    for (const window of [overlayManager.getMainWindow(), overlayManager.getModalWindow()]) {
      if (window && !window.isDestroyed())
        window.webContents.send(IPC_CHANNELS.event.subtitleGenerationProgress, progress);
    }
  },
});
registerSubtitleGenerationIpc({
  ipc: ipcMain,
  isAllowedSender: (sender) =>
    [overlayManager.getMainWindow(), overlayManager.getModalWindow()].some(
      (window) => window && !window.isDestroyed() && window.webContents === sender,
    ),
  runtime: subtitleGenerationRuntime,
  openModal: () => openSubtitleGenerationModal(overlayModals.createOverlayHostedModalOpenDeps()),
});
app.on('before-quit', () => subtitleGenerationRuntime.cancel());
