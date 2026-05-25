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
import { openPlaylistBrowser as openPlaylistBrowserRuntime } from './main/runtime/playlist-browser-open';
import { createDiscordRpcClient } from './main/runtime/discord-rpc-client.js';
import { startAppControlServer } from './main/runtime/app-control-server';
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
import { mergeAiConfig } from './ai/config';

function getPasswordStoreArg(argv: string[]): string | null {
  let resolved: string | null = null;
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (!arg?.startsWith('--password-store')) {
      continue;
    }

    if (arg === '--password-store') {
      const value = argv[i + 1];
      if (value && !value.startsWith('--')) {
        resolved = value.trim();
        i += 1;
      }
      continue;
    }

    const [prefix, value] = arg.split('=', 2);
    if (prefix === '--password-store' && value && value.trim().length > 0) {
      resolved = value.trim();
    }
  }
  return resolved;
}

function normalizePasswordStoreArg(value: string): string {
  const normalized = value.trim();
  if (normalized.toLowerCase() === 'gnome') {
    return 'gnome-libsecret';
  }
  return normalized;
}

function getDefaultPasswordStore(): string {
  return 'gnome-libsecret';
}

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
import { execFile, spawn } from 'node:child_process';
import * as os from 'os';
import * as path from 'path';
import { MecabTokenizer } from './mecab-tokenizer';
import type {
  CompiledSessionBinding,
  JimakuApiResponse,
  KikuFieldGroupingChoice,
  MpvSubtitleRenderMetrics,
  ResolvedConfig,
  RuntimeOptionState,
  SessionActionDispatchRequest,
  SecondarySubMode,
  SubtitleCue,
  SubtitleData,
  SubtitlePosition,
  UpdateChannel,
  WindowGeometry,
} from './types';
import { AnkiIntegration } from './anki-integration';
import { SubtitleTimingTracker } from './subtitle-timing-tracker';
import { RuntimeOptionsManager } from './runtime-options';
import { downloadToFile, isRemoteMediaPath, parseMediaInfo } from './jimaku/utils';
import {
  createLogger,
  setLogLevel,
  resolveDefaultLogFilePath,
  type LogLevelSource,
} from './logger';
import { createFatalErrorReporter } from './main/fatal-error';
import { createWindowTracker as createWindowTrackerCore } from './window-trackers';
import {
  bindWindowsOverlayAboveMpv,
  clearWindowsOverlayOwner,
  ensureWindowsOverlayTransparency,
  findWindowsMpvTargetWindowHandle,
  getWindowsForegroundProcessName,
  setWindowsOverlayOwner,
} from './window-trackers/windows-helper';
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
import { AnkiConnectClient } from './anki-connect';
import {
  getStartupModeFlags,
  shouldRefreshAnilistOnConfigReload,
  shouldStartAutomaticUpdateChecks,
} from './main/runtime/startup-mode-flags';
import {
  buildConfigParseErrorDetails,
  buildConfigWarningDialogDetails,
  buildConfigWarningNotificationBody,
  failStartupFromConfig,
} from './main/config-validation';
import {
  buildAnilistSetupUrl,
  consumeAnilistSetupCallbackUrl,
  createAnilistStateRuntime,
  createBuildOpenAnilistSetupWindowMainDepsHandler,
  createMaybeFocusExistingAnilistSetupWindowHandler,
  createOpenAnilistSetupWindowHandler,
  findAnilistSetupDeepLinkArgvUrl,
  isAnilistTrackingEnabled,
  loadAnilistManualTokenEntry,
  openAnilistSetupInBrowser,
  rememberAnilistAttemptedUpdateKey,
} from './main/runtime/domains/anilist';
import { DEFAULT_MIN_WATCH_RATIO } from './shared/watch-threshold';
import { shouldShowTexthookerTrayEntry } from './main/runtime/tray-main-actions';
import {
  createApplyJellyfinMpvDefaultsHandler,
  createBuildApplyJellyfinMpvDefaultsMainDepsHandler,
  createBuildGetDefaultSocketPathMainDepsHandler,
  createGetDefaultSocketPathHandler,
  buildJellyfinSetupFormHtml,
  parseJellyfinSetupSubmissionUrl,
  getConfiguredJellyfinSession,
  mergeJellyfinRecentServers,
  persistJellyfinAuthSession,
  type ActiveJellyfinRemotePlaybackState,
} from './main/runtime/domains/jellyfin';
import {
  createBuildConfigHotReloadMessageMainDepsHandler,
  createBuildConfigHotReloadAppliedMainDepsHandler,
  createBuildConfigHotReloadRuntimeMainDepsHandler,
  createBuildWatchConfigPathMainDepsHandler,
  createWatchConfigPathHandler,
  createBuildOverlayContentMeasurementStoreMainDepsHandler,
  createBuildOverlayModalRuntimeMainDepsHandler,
  createBuildAppendClipboardVideoToQueueMainDepsHandler,
  createBuildHandleOverlayModalClosedMainDepsHandler,
  createBuildLoadSubtitlePositionMainDepsHandler,
  createBuildSaveSubtitlePositionMainDepsHandler,
  createBuildFieldGroupingOverlayMainDepsHandler,
  createBuildGetFieldGroupingResolverMainDepsHandler,
  createBuildSetFieldGroupingResolverMainDepsHandler,
  createBuildOverlayVisibilityRuntimeMainDepsHandler,
  createBuildBroadcastRuntimeOptionsChangedMainDepsHandler,
  createBuildGetRuntimeOptionsStateMainDepsHandler,
  createBuildOpenRuntimeOptionsPaletteMainDepsHandler,
  createBuildRestorePreviousSecondarySubVisibilityMainDepsHandler,
  createBuildSendToActiveOverlayWindowMainDepsHandler,
  createBuildSetOverlayDebugVisualizationEnabledMainDepsHandler,
  createBuildEnforceOverlayLayerOrderMainDepsHandler,
  createBuildEnsureOverlayWindowLevelMainDepsHandler,
  createBuildUpdateVisibleOverlayBoundsMainDepsHandler,
  createTrayRuntimeHandlers,
  createOverlayVisibilityRuntime,
  createBroadcastRuntimeOptionsChangedHandler,
  createGetRuntimeOptionsStateHandler,
  createGetFieldGroupingResolverHandler,
  createSetFieldGroupingResolverHandler,
  createOpenRuntimeOptionsPaletteHandler,
  createRestorePreviousSecondarySubVisibilityHandler,
  createSendToActiveOverlayWindowHandler,
  createSetOverlayDebugVisualizationEnabledHandler,
  createEnforceOverlayLayerOrderHandler,
  createEnsureOverlayWindowLevelHandler,
  createUpdateVisibleOverlayBoundsHandler,
  createLoadSubtitlePositionHandler,
  createSaveSubtitlePositionHandler,
  createAppendClipboardVideoToQueueHandler,
  createHandleOverlayModalClosedHandler,
  createConfigHotReloadMessageHandler,
  createConfigHotReloadAppliedHandler,
  buildTrayMenuTemplateRuntime,
  resolveTrayIconPathRuntime,
  createYomitanExtensionRuntime,
  createYomitanSettingsRuntime,
  buildRestartRequiredConfigMessage,
  resolveSubtitleStyleForRenderer,
} from './main/runtime/domains/overlay';
import {
  createBuildAnilistStateRuntimeMainDepsHandler,
  createBuildConfigDerivedRuntimeMainDepsHandler,
  createBuildImmersionMediaRuntimeMainDepsHandler,
  createBuildMainSubsyncRuntimeMainDepsHandler,
  createBuildSubtitleProcessingControllerMainDepsHandler,
  createBuildMediaRuntimeMainDepsHandler,
  createBuildDictionaryRootsMainHandler,
  createBuildFrequencyDictionaryRootsMainHandler,
  createBuildFrequencyDictionaryRuntimeMainDepsHandler,
  createBuildJlptDictionaryRuntimeMainDepsHandler,
  createImmersionMediaRuntime,
  createConfigDerivedRuntime,
  appendClipboardVideoToQueueRuntime,
  createMainSubsyncRuntime,
} from './main/runtime/domains/startup';
import {
  createMpvOsdRuntimeHandlers,
  createCycleSecondarySubModeRuntimeHandler,
} from './main/runtime/domains/mpv';
import {
  createBuildCopyCurrentSubtitleMainDepsHandler,
  createBuildHandleMineSentenceDigitMainDepsHandler,
  createBuildHandleMultiCopyDigitMainDepsHandler,
  createBuildMarkLastCardAsAudioCardMainDepsHandler,
  createBuildMineSentenceCardMainDepsHandler,
  createBuildRefreshKnownWordCacheMainDepsHandler,
  createBuildTriggerFieldGroupingMainDepsHandler,
  createBuildUpdateLastCardFromClipboardMainDepsHandler,
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
import {
  ensureDefaultConfigBootstrap,
  getDefaultConfigFilePaths,
  resolveDefaultMpvInstallPaths,
} from './shared/setup-state';
import {
  ImmersionTrackerService,
  JellyfinRemoteSessionService,
  MpvIpcClient,
  SubtitleWebSocket,
  Texthooker,
  applyMpvSubtitleRenderMetricsPatch,
  authenticateWithPasswordRuntime,
  broadcastRuntimeOptionsChangedRuntime,
  copyCurrentSubtitle as copyCurrentSubtitleCore,
  createConfigHotReloadRuntime,
  createDiscordPresenceService,
  createShiftSubtitleDelayToAdjacentCueHandler,
  createFieldGroupingOverlayRuntime,
  createOverlayContentMeasurementStore,
  createOverlayManager,
  createOverlayWindow as createOverlayWindowCore,
  createSubtitleProcessingController,
  createTokenizerDepsRuntime,
  cycleSecondarySubMode as cycleSecondarySubModeCore,
  deleteYomitanDictionaryByTitle,
  destroyYomitanSettingsWindow,
  enforceOverlayLayerOrder as enforceOverlayLayerOrderCore,
  ensureOverlayWindowLevel as ensureOverlayWindowLevelCore,
  getYomitanDictionaryInfo,
  handleMineSentenceDigit as handleMineSentenceDigitCore,
  handleMultiCopyDigit as handleMultiCopyDigitCore,
  hasMpvWebsocketPlugin,
  importYomitanDictionaryFromZip,
  initializeOverlayAnkiIntegration as initializeOverlayAnkiIntegrationCore,
  initializeOverlayRuntime as initializeOverlayRuntimeCore,
  isOverlayWindowContentReady,
  jellyfinTicksToSecondsRuntime,
  listJellyfinItemsRuntime,
  listJellyfinLibrariesRuntime,
  listJellyfinSubtitleTracksRuntime,
  loadJellyfinSubtitleDelay,
  loadSubtitlePosition as loadSubtitlePositionCore,
  loadYomitanExtension as loadYomitanExtensionCore,
  markLastCardAsAudioCard as markLastCardAsAudioCardCore,
  mineSentenceCard as mineSentenceCardCore,
  openYomitanSettingsWindow,
  playNextSubtitleRuntime,
  registerGlobalShortcuts as registerGlobalShortcutsCore,
  replayCurrentSubtitleRuntime,
  resolveJellyfinPlaybackPlanRuntime,
  runStartupBootstrapRuntime,
  saveJellyfinSubtitleDelay,
  saveSubtitlePosition as saveSubtitlePositionCore,
  addYomitanNoteViaSearch,
  clearYomitanParserCachesForWindow,
  syncYomitanDefaultAnkiServer as syncYomitanDefaultAnkiServerCore,
  sendMpvCommandRuntime,
  setMpvSubVisibilityRuntime,
  setOverlayDebugVisualizationEnabledRuntime,
  syncOverlayWindowLayer,
  setVisibleOverlayVisible as setVisibleOverlayVisibleCore,
  showMpvOsdRuntime,
  startOverlayWindowTracker as startOverlayWindowTrackerCore,
  tokenizeSubtitle as tokenizeSubtitleCore,
  triggerFieldGrouping as triggerFieldGroupingCore,
  upsertYomitanDictionarySettings,
  updateLastCardFromClipboard as updateLastCardFromClipboardCore,
} from './core/services';
import {
  acquireYoutubeSubtitleTrack,
  acquireYoutubeSubtitleTracks,
} from './core/services/youtube/generate';
import { resolveYoutubePlaybackUrl } from './core/services/youtube/playback-resolve';
import { probeYoutubeTracks } from './core/services/youtube/track-probe';
import { startStatsServer } from './core/services/stats-server';
import {
  destroyStatsWindow,
  promoteStatsOverlayAbovePlayback,
  registerStatsOverlayToggle,
  toggleStatsOverlay as toggleStatsOverlayWindow,
  withStatsWindowLayerSuspendedForNativeDialog,
} from './core/services/stats-window.js';
import {
  createFirstRunSetupService,
  getFirstRunSetupCompletionMessage,
  isStandaloneFirstRunSetupCommand,
  shouldAutoOpenFirstRunSetup,
} from './main/runtime/first-run-setup-service';
import { createYoutubeFlowRuntime } from './main/runtime/youtube-flow';
import { createYoutubePlaybackRuntime } from './main/runtime/youtube-playback-runtime';
import {
  clearYoutubePrimarySubtitleNotificationTimer,
  createYoutubePrimarySubtitleNotificationRuntime,
} from './main/runtime/youtube-primary-subtitle-notification';
import { createAutoplayReadyGate } from './main/runtime/autoplay-ready-gate';
import { selectAutoplayStartupCue } from './main/runtime/autoplay-subtitle-primer';
import { createAutoplayTokenizationWarmRelease } from './main/runtime/autoplay-tokenization-warm-release';
import { createManagedLocalSubtitleSelectionRuntime } from './main/runtime/local-subtitle-selection';
import {
  buildFirstRunSetupHtml,
  createMaybeFocusExistingFirstRunSetupWindowHandler,
  createOpenFirstRunSetupWindowHandler,
  parseFirstRunSetupSubmissionUrl,
  type FirstRunSetupSubmission,
} from './main/runtime/first-run-setup-window';
import {
  detectInstalledFirstRunPlugin,
  detectInstalledFirstRunPluginCandidates,
  detectInstalledMpvPlugin,
  removeLegacyMpvPluginCandidates,
  resolvePackagedRuntimePluginPath,
} from './main/runtime/first-run-setup-plugin';
import {
  applyWindowsMpvShortcuts,
  detectWindowsMpvShortcuts,
  resolveWindowsMpvShortcutPaths,
} from './main/runtime/windows-mpv-shortcuts';
import {
  detectCommandLineLauncher,
  installBun as installCommandLineBun,
  installLauncher as installCommandLineLauncher,
} from './main/runtime/command-line-launcher';
import {
  createWindowsMpvLaunchDeps,
  getConfiguredWindowsMpvPathStatus,
  launchWindowsMpv,
} from './main/runtime/windows-mpv-launch';
import { createWaitForMpvConnectedHandler } from './main/runtime/jellyfin-remote-connection';
import {
  DEFAULT_JELLYFIN_CLIENT_NAME,
  DEFAULT_JELLYFIN_CLIENT_VERSION,
  createHostDerivedJellyfinDeviceId,
} from './main/runtime/jellyfin-device-identity';
import {
  clearJellyfinAuthSessionAndRefreshTray as clearJellyfinAuthSessionAndRefreshTrayRuntime,
  isJellyfinConfiguredForTray as isJellyfinConfiguredForTrayRuntime,
  toggleJellyfinDiscoveryFromTray as toggleJellyfinDiscoveryFromTrayRuntime,
} from './main/runtime/jellyfin-tray-discovery';
import { createPrepareYoutubePlaybackInMpvHandler } from './main/runtime/youtube-playback-launch';
import {
  shouldEnsureTrayOnStartupForInitialArgs,
  shouldQuitOnMpvShutdownForTrayState,
  shouldQuitOnWindowAllClosedForTrayState,
} from './main/runtime/startup-tray-policy';
import { createImmersionTrackerStartupHandler } from './main/runtime/immersion-startup';
import { createBuildImmersionTrackerStartupMainDepsHandler } from './main/runtime/immersion-startup-main-deps';
import {
  createRunStatsCliCommandHandler,
  writeStatsCliCommandResponse,
} from './main/runtime/stats-cli-command';
import {
  isBackgroundStatsServerProcessAlive,
  readBackgroundStatsServerState,
  removeBackgroundStatsServerState,
  resolveBackgroundStatsServerUrl,
  writeBackgroundStatsServerState,
} from './main/runtime/stats-daemon';
import { createEnsureStatsServerUrlHandler } from './main/runtime/stats-server-routing';
import { resolveLegacyVocabularyPosFromTokens } from './core/services/immersion-tracker/legacy-vocabulary-pos';
import { createAnilistUpdateQueue } from './core/services/anilist/anilist-update-queue';
import {
  guessAnilistMediaInfo,
  updateAnilistPostWatchProgress,
} from './core/services/anilist/anilist-updater';
import { createCoverArtFetcher } from './core/services/anilist/cover-art-fetcher';
import { createAnilistRateLimiter } from './core/services/anilist/rate-limiter';
import { createJellyfinTokenStore } from './core/services/jellyfin-token-store';
import { applyRuntimeOptionResultRuntime } from './core/services/runtime-options-ipc';
import { createAnilistTokenStore } from './core/services/anilist/anilist-token-store';
import {
  buildPluginSessionBindingsArtifact,
  compileSessionBindings,
} from './core/services/session-bindings';
import { dispatchSessionAction as dispatchSessionActionCore } from './core/services/session-actions';
import { createBuildOverlayShortcutsRuntimeMainDepsHandler } from './main/runtime/domains/shortcuts';
import { createMainRuntimeRegistry } from './main/runtime/registry';
import {
  createEnsureOverlayMpvSubtitlesHiddenHandler,
  createRestoreOverlayMpvSubtitlesHandler,
} from './main/runtime/overlay-mpv-sub-visibility';
import {
  composeAnilistSetupHandlers,
  composeAnilistTrackingHandlers,
  composeAppReadyRuntime,
  composeCliStartupHandlers,
  composeHeadlessStartupHandlers,
  composeIpcRuntimeHandlers,
  composeJellyfinRuntimeHandlers,
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
import { createAnkiJimakuIpcRuntimeServiceDeps } from './main/dependencies';
import { createMainBootServices, type MainBootServicesResult } from './main/boot/services';
import { handleCliCommandRuntimeServiceWithContext } from './main/cli-runtime';
import { createOverlayModalRuntimeService } from './main/overlay-runtime';
import { createOverlayModalInputState } from './main/runtime/overlay-modal-input-state';
import { openYoutubeTrackPicker } from './main/runtime/youtube-picker-open';
import { openRuntimeOptionsModal as openRuntimeOptionsModalRuntime } from './main/runtime/runtime-options-open';
import { openJimakuModal as openJimakuModalRuntime } from './main/runtime/jimaku-open';
import { openSubsyncManualModal as openSubsyncManualModalRuntime } from './main/runtime/subsync-open';
import { openSessionHelpModal as openSessionHelpModalRuntime } from './main/runtime/session-help-open';
import { openCharacterDictionaryModal as openCharacterDictionaryModalRuntime } from './main/runtime/character-dictionary-open';
import { openControllerSelectModal as openControllerSelectModalRuntime } from './main/runtime/controller-select-open';
import { openControllerDebugModal as openControllerDebugModalRuntime } from './main/runtime/controller-debug-open';
import { createPlaylistBrowserIpcRuntime } from './main/runtime/playlist-browser-ipc';
import { writeSessionBindingsArtifact } from './main/runtime/session-bindings-artifact';
import { createOverlayShortcutsRuntimeService } from './main/overlay-shortcuts-runtime';
import {
  createFrequencyDictionaryRuntimeService,
  getFrequencyDictionarySearchPaths,
} from './main/frequency-dictionary-runtime';
import {
  createJlptDictionaryRuntimeService,
  getJlptDictionarySearchPaths,
} from './main/jlpt-runtime';
import { createMediaRuntimeService } from './main/media-runtime';
import { createOverlayVisibilityRuntimeService } from './main/overlay-visibility-runtime';
import { createStatsOverlayVisibilityChangeHandler } from './main/runtime/stats-overlay-visibility';
import { createDiscordPresenceRuntime } from './main/runtime/discord-presence-runtime';
import { createCharacterDictionaryRuntimeService } from './main/character-dictionary-runtime';
import { createCharacterDictionaryAutoSyncRuntimeService } from './main/runtime/character-dictionary-auto-sync';
import { handleCharacterDictionaryAutoSyncComplete } from './main/runtime/character-dictionary-auto-sync-completion';
import { notifyCharacterDictionaryAutoSyncStatus } from './main/runtime/character-dictionary-auto-sync-notifications';
import { createCurrentMediaTokenizationGate } from './main/runtime/current-media-tokenization-gate';
import { resolveCurrentSubtitleForRenderer } from './main/runtime/current-subtitle-snapshot';
import { createJellyfinSubtitleCacheIo } from './main/runtime/jellyfin-subtitle-cache-io';
import { createStartupOsdSequencer } from './main/runtime/startup-osd-sequencer';
import {
  createElectronAppUpdater,
  isNativeUpdaterSupported,
} from './main/runtime/update/app-updater';
import { createCurlFetch, createGlobalFetch } from './main/runtime/update/fetch-adapter';
import { createCurlHttpExecutor } from './main/runtime/update/curl-http-executor';
import { createFetchHttpExecutor } from './main/runtime/update/fetch-http-executor';
import {
  fetchLatestStableRelease,
  fetchReleaseAssetBuffer,
  fetchReleaseAssetText,
  findReleaseAsset,
  parseSha256Sums,
  type GitHubRelease,
} from './main/runtime/update/release-assets';
import { shouldFetchReleaseMetadataForPlatform } from './main/runtime/update/release-metadata-policy';
import { updateLauncherFromRelease } from './main/runtime/update/launcher-updater';
import { notifyUpdateAvailable } from './main/runtime/update/update-notifications';
import { createUpdateDialogPresenter } from './main/runtime/update/update-dialogs';
import {
  runUpdateCliCommand,
  writeUpdateCliCommandResponse,
} from './main/runtime/update/update-cli-command';
import {
  createFileUpdateStateStore,
  createUpdateService,
} from './main/runtime/update/update-service';
import { updateSupportAssetsFromRelease } from './main/runtime/update/support-assets';
import {
  createRefreshSubtitlePrefetchFromActiveTrackHandler,
  createResolveActiveSubtitleSidebarSourceHandler,
} from './main/runtime/subtitle-prefetch-runtime';
import {
  createCreateAnilistSetupWindowHandler,
  createCreateConfigSettingsWindowHandler,
  createCreateFirstRunSetupWindowHandler,
  createCreateJellyfinSetupWindowHandler,
} from './main/runtime/setup-window-factory';
import { createConfigSettingsRuntime } from './main/runtime/config-settings-runtime';
import {
  isSameYoutubeMediaPath,
  isYoutubeMediaPath,
  isYoutubePlaybackActive,
  shouldUseCachedYoutubeParsedCues,
} from './main/runtime/youtube-playback';
import { createYomitanProfilePolicy } from './main/runtime/yomitan-profile-policy';
import { reloadOverlayWindowsForYomitanContentScripts } from './main/runtime/yomitan-extension-overlay-reload';
import { formatSkippedYomitanWriteAction } from './main/runtime/yomitan-read-only-log';
import {
  getPreferredYomitanAnkiServerUrl as getPreferredYomitanAnkiServerUrlRuntime,
  shouldForceOverrideYomitanAnkiServer,
} from './main/runtime/yomitan-anki-server';
import {
  type AnilistMediaGuessRuntimeState,
  type StartupState,
  applyStartupState,
  createAppState,
  createInitialAnilistMediaGuessRuntimeState,
  createInitialAnilistUpdateInFlightState,
  transitionAnilistClientSecretState,
  transitionAnilistMediaGuessRuntimeState,
  transitionAnilistRetryQueueLastAttemptAt,
  transitionAnilistRetryQueueLastError,
  transitionAnilistRetryQueueState,
  transitionAnilistUpdateInFlightState,
} from './main/state';
import {
  isAllowedAnilistExternalUrl,
  isAllowedAnilistSetupNavigationUrl,
} from './main/anilist-url-guard';
import {
  ConfigService,
  ConfigStartupParseError,
  DEFAULT_CONFIG,
  DEFAULT_KEYBINDINGS,
  generateConfigTemplate,
} from './config';
import { resolveConfigDir } from './config/path-resolution';
import { buildConfigSettingsRegistry } from './config/settings/registry';
import { parseSubtitleCues } from './core/services/subtitle-cue-parser';
import {
  createSubtitlePrefetchService,
  type SubtitlePrefetchService,
} from './core/services/subtitle-prefetch';
import {
  buildSubtitleSidebarSourceKey,
  resolveSubtitleSourcePath,
} from './main/runtime/subtitle-prefetch-source';
import { createSubtitlePrefetchInitController } from './main/runtime/subtitle-prefetch-init';
import { applyCharacterDictionarySelection } from './main/character-dictionary-selection';
import { codecToExtension, getSubsyncConfig } from './subsync/utils';

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
const DEFAULT_MPV_LOG_FILE = resolveDefaultLogFilePath({
  platform: process.platform,
  homeDir: os.homedir(),
  appDataDir: process.env.APPDATA,
});
const ANILIST_SETUP_CLIENT_ID_URL = 'https://anilist.co/api/v2/oauth/authorize';
const jellyfinSubtitleCacheIo = createJellyfinSubtitleCacheIo({
  tmpDir: () => os.tmpdir(),
  makeTempDir: (prefix) => fs.promises.mkdtemp(prefix),
  writeFile: (filePath, bytes) => fs.promises.writeFile(filePath, bytes),
  removeDir: (dir, options) => {
    fs.rmSync(dir, options);
  },
  fetch: (url) => fetch(url),
});
const ANILIST_SETUP_RESPONSE_TYPE = 'token';
const ANILIST_DEFAULT_CLIENT_ID = '36084';
const ANILIST_REDIRECT_URI = 'https://anilist.subminer.moe/';
const ANILIST_DEVELOPER_SETTINGS_URL = 'https://anilist.co/settings/developer';
const ANILIST_UPDATE_MIN_WATCH_SECONDS = 10 * 60;
const ANILIST_DURATION_RETRY_INTERVAL_MS = 15_000;
const ANILIST_MAX_ATTEMPTED_UPDATE_KEYS = 1000;
const TRAY_TOOLTIP = 'SubMiner';
const SUBMINER_BUNDLE_ID = 'com.sudacode.SubMiner';
const JELLYFIN_SETUP_PRELOAD_PATH = path.join(__dirname, 'preload-jellyfin-setup.js');

let anilistMediaGuessRuntimeState: AnilistMediaGuessRuntimeState =
  createInitialAnilistMediaGuessRuntimeState();
let anilistUpdateInFlightState = createInitialAnilistUpdateInFlightState();
const anilistAttemptedUpdateKeys = new Set<string>();
let anilistCachedAccessToken: string | null = null;
let jellyfinPlayQuitOnDisconnectArmed = false;
const JELLYFIN_LANG_PREF = 'ja,jp,jpn,japanese,en,eng,english,enUS,en-US';
const JELLYFIN_TICKS_PER_SECOND = 10_000_000;
const JELLYFIN_REMOTE_PROGRESS_INTERVAL_MS = 3000;
const JELLYFIN_REMOTE_STARTUP_STOP_GRACE_MS = 10_000;
const DISCORD_PRESENCE_APP_ID = '1475264834730856619';
const JELLYFIN_MPV_CONNECT_TIMEOUT_MS = 3000;
const JELLYFIN_MPV_AUTO_LAUNCH_TIMEOUT_MS = 20000;
const YOUTUBE_MPV_CONNECT_TIMEOUT_MS = 3000;
const YOUTUBE_MPV_AUTO_LAUNCH_TIMEOUT_MS = 10000;
const YOUTUBE_MPV_YTDL_FORMAT = 'bestvideo*+bestaudio/best';
const YOUTUBE_DIRECT_PLAYBACK_FORMAT = 'b';
const MPV_JELLYFIN_DEFAULT_ARGS = [
  '--sub-auto=no',
  '--sub-file-paths=.;subs;subtitles',
  '--sid=no',
  '--secondary-sid=no',
  '--sub-visibility=no',
  '--secondary-sub-visibility=no',
  '--alang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
  '--slang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
] as const;

let activeJellyfinRemotePlayback: ActiveJellyfinRemotePlaybackState | null = null;
let activeJellyfinSubtitleDelayKey: { itemId: string; streamIndex: number } | null = null;
let jellyfinRemoteLastProgressAtMs = 0;
let jellyfinMpvAutoLaunchInFlight: Promise<boolean> | null = null;
let backgroundWarmupsStarted = false;
let yomitanLoadInFlight: Promise<Extension | null> | null = null;
let notifyAnilistTokenStoreWarning: (message: string) => void = () => {};

const buildApplyJellyfinMpvDefaultsMainDepsHandler =
  createBuildApplyJellyfinMpvDefaultsMainDepsHandler({
    sendMpvCommandRuntime: (client, command) => sendMpvCommandRuntime(client, command),
    jellyfinLangPref: JELLYFIN_LANG_PREF,
  });
const applyJellyfinMpvDefaultsMainDeps = buildApplyJellyfinMpvDefaultsMainDepsHandler();
const applyJellyfinMpvDefaultsHandler = createApplyJellyfinMpvDefaultsHandler(
  applyJellyfinMpvDefaultsMainDeps,
);

function applyJellyfinMpvDefaults(
  client: Parameters<typeof applyJellyfinMpvDefaultsHandler>[0],
): void {
  applyJellyfinMpvDefaultsHandler(client);
}

const isDev = process.argv.includes('--dev') || process.argv.includes('--debug');
const texthookerService = new Texthooker(() => {
  const config = getResolvedConfig();
  const characterDictionaryEnabled =
    config.anilist.characterDictionary.enabled &&
    yomitanProfilePolicy.isCharacterDictionaryEnabled();
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
const buildGetDefaultSocketPathMainDepsHandler = createBuildGetDefaultSocketPathMainDepsHandler({
  platform: process.platform,
});
const getDefaultSocketPathMainDeps = buildGetDefaultSocketPathMainDepsHandler();
const getDefaultSocketPathHandler = createGetDefaultSocketPathHandler(getDefaultSocketPathMainDeps);

function getDefaultSocketPath(): string {
  return getDefaultSocketPathHandler();
}

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
  appDataDir: process.env.APPDATA,
  xdgConfigHome: process.env.XDG_CONFIG_HOME,
  homeDir: os.homedir(),
  defaultMpvLogFile: DEFAULT_MPV_LOG_FILE,
  envMpvLog: process.env.SUBMINER_MPV_LOG,
  defaultTexthookerPort: DEFAULT_TEXTHOOKER_PORT,
  getDefaultSocketPath: () => getDefaultSocketPath(),
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
  createOverlayManager,
  createOverlayModalInputState,
  createOverlayContentMeasurementStore: ({ logger }) => {
    const buildHandler = createBuildOverlayContentMeasurementStoreMainDepsHandler({
      now: () => Date.now(),
      warn: (message: string) => logger.warn(message),
    });
    return createOverlayContentMeasurementStore(buildHandler());
  },
  getSyncOverlayShortcutsForModal: () => syncOverlayShortcutsForModal,
  getSyncOverlayVisibilityForModal: () => syncOverlayVisibilityForModal,
  createOverlayModalRuntime: ({ overlayManager, overlayModalInputState }) => {
    const buildHandler = createBuildOverlayModalRuntimeMainDepsHandler({
      getMainWindow: () => overlayManager.getMainWindow(),
      getModalWindow: () => overlayManager.getModalWindow(),
      createModalWindow: () => createModalWindow(),
      getModalGeometry: () => getCurrentOverlayGeometry(),
      setModalWindowBounds: (geometry) => overlayManager.setModalWindowBounds(geometry),
    });
    return createOverlayModalRuntimeService(buildHandler(), {
      onModalStateChange: (isActive: boolean) =>
        overlayModalInputState.handleModalInputStateChange(isActive),
    });
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
const configSettingsFields = buildConfigSettingsRegistry(DEFAULT_CONFIG);
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
let statsServer: ReturnType<typeof startStatsServer> | null = null;
const statsDaemonStatePath = path.join(USER_DATA_PATH, 'stats-daemon.json');

function readLiveBackgroundStatsDaemonState(): {
  pid: number;
  port: number;
  startedAtMs: number;
} | null {
  const state = readBackgroundStatsServerState(statsDaemonStatePath);
  if (!state) {
    removeBackgroundStatsServerState(statsDaemonStatePath);
    return null;
  }
  if (state.pid === process.pid && !statsServer) {
    removeBackgroundStatsServerState(statsDaemonStatePath);
    return null;
  }
  if (!isBackgroundStatsServerProcessAlive(state.pid)) {
    removeBackgroundStatsServerState(statsDaemonStatePath);
    return null;
  }
  return state;
}

function clearOwnedBackgroundStatsDaemonState(): void {
  const state = readBackgroundStatsServerState(statsDaemonStatePath);
  if (state?.pid === process.pid) {
    removeBackgroundStatsServerState(statsDaemonStatePath);
  }
}

function stopStatsServer(): void {
  if (!statsServer) {
    return;
  }
  statsServer.close();
  statsServer = null;
  clearOwnedBackgroundStatsDaemonState();
}

function requestAppQuit(): void {
  destroyYomitanSettingsWindow(appState.yomitanSettingsWindow);
  appState.yomitanSettingsWindow = null;
  destroyStatsWindow();
  stopStatsServer();
  if (!forceQuitTimer) {
    forceQuitTimer = setTimeout(() => {
      logger.warn('App quit timed out; forcing process exit.');
      app.exit(0);
    }, 2000);
  }
  app.quit();
}

process.on('SIGINT', () => {
  requestAppQuit();
});
process.on('SIGTERM', () => {
  requestAppQuit();
});

const startBackgroundWarmupsIfAllowed = (): void => {
  startBackgroundWarmups();
};
const youtubeFlowRuntime = createYoutubeFlowRuntime({
  probeYoutubeTracks: (url: string) => probeYoutubeTracks(url),
  acquireYoutubeSubtitleTrack: (input) => acquireYoutubeSubtitleTrack(input),
  acquireYoutubeSubtitleTracks: (input) => acquireYoutubeSubtitleTracks(input),
  openPicker: async (payload) => {
    return await openYoutubeTrackPicker(
      {
        sendToActiveOverlayWindow: (channel, nextPayload, runtimeOptions) =>
          overlayModalRuntime.sendToActiveOverlayWindow(channel, nextPayload, runtimeOptions),
        waitForModalOpen: (modal, timeoutMs) =>
          overlayModalRuntime.waitForModalOpen(modal, timeoutMs),
        logWarn: (message) => logger.warn(message),
      },
      payload,
    );
  },
  pauseMpv: () => {
    sendMpvCommandRuntime(appState.mpvClient, ['set_property', 'pause', 'yes']);
  },
  resumeMpv: () => {
    sendMpvCommandRuntime(appState.mpvClient, ['set_property', 'pause', 'no']);
  },
  sendMpvCommand: (command) => {
    sendMpvCommandRuntime(appState.mpvClient, command);
  },
  requestMpvProperty: async (name: string) => {
    const client = appState.mpvClient;
    if (!client) return null;
    return await client.requestProperty(name);
  },
  refreshCurrentSubtitle: (text: string) => {
    subtitleProcessingController.refreshCurrentSubtitle(text);
  },
  refreshSubtitleSidebarSource: async (sourcePath: string, mediaPath?: string) => {
    await subtitlePrefetchRuntime.refreshSubtitleSidebarFromSource(sourcePath, mediaPath);
  },
  startTokenizationWarmups: async () => {
    await startTokenizationWarmups();
  },
  waitForTokenizationReady: async () => {
    await currentMediaTokenizationGate.waitUntilReady(
      appState.currentMediaPath?.trim() || appState.mpvClient?.currentVideoPath?.trim() || null,
    );
  },
  waitForAnkiReady: async () => {
    const integration = appState.ankiIntegration;
    if (!integration) {
      return;
    }
    try {
      await Promise.race([
        integration.waitUntilReady(),
        new Promise<never>((_, reject) => {
          setTimeout(
            () => reject(new Error('Timed out waiting for AnkiConnect integration')),
            2500,
          );
        }),
      ]);
    } catch (error) {
      logger.warn(
        'Continuing YouTube playback before AnkiConnect integration reported ready:',
        error instanceof Error ? error.message : String(error),
      );
    }
  },
  wait: (ms: number) => new Promise((resolve) => setTimeout(resolve, ms)),
  waitForPlaybackWindowReady: async () => {
    const deadline = Date.now() + 4000;
    let stableGeometry: WindowGeometry | null = null;
    let stableSinceMs = 0;
    while (Date.now() < deadline) {
      const tracker = appState.windowTracker;
      const trackerGeometry = tracker?.getGeometry() ?? null;
      const mediaPath =
        appState.currentMediaPath?.trim() || appState.mpvClient?.currentVideoPath?.trim() || '';
      const trackerFocused = tracker?.isTargetWindowFocused() ?? false;
      if (tracker && tracker.isTracking() && trackerGeometry && trackerFocused && mediaPath) {
        if (!geometryMatches(stableGeometry, trackerGeometry)) {
          stableGeometry = trackerGeometry;
          stableSinceMs = Date.now();
        } else if (Date.now() - stableSinceMs >= 200) {
          return;
        }
      } else {
        stableGeometry = null;
        stableSinceMs = 0;
      }
      await new Promise((resolve) => setTimeout(resolve, 100));
    }
    logger.warn(
      'Timed out waiting for tracked playback window focus/media readiness before opening YouTube subtitle picker.',
    );
  },
  waitForOverlayGeometryReady: async () => {
    const deadline = Date.now() + 4000;
    while (Date.now() < deadline) {
      const tracker = appState.windowTracker;
      const trackerGeometry = tracker?.getGeometry() ?? null;
      if (trackerGeometry && geometryMatches(lastOverlayWindowGeometry, trackerGeometry)) {
        return;
      }
      await new Promise((resolve) => setTimeout(resolve, 50));
    }
    logger.warn('Timed out waiting for overlay geometry to match tracked playback window.');
  },
  focusOverlayWindow: () => {
    const mainWindow = overlayManager.getMainWindow();
    if (!mainWindow || mainWindow.isDestroyed()) {
      return;
    }
    mainWindow.setIgnoreMouseEvents(false);
    if (!mainWindow.isFocused()) {
      mainWindow.focus();
    }
    if (!mainWindow.webContents.isFocused()) {
      mainWindow.webContents.focus();
    }
  },
  showMpvOsd: (text: string) => showMpvOsd(text),
  reportSubtitleFailure: (message: string) => reportYoutubeSubtitleFailure(message),
  notifyPrimarySubtitleLoaded: () =>
    youtubePrimarySubtitleNotificationRuntime.markCurrentMediaPrimarySubtitleLoaded(),
  warn: (message: string) => logger.warn(message),
  log: (message: string) => logger.info(message),
  getYoutubeOutputDir: () => path.join(os.homedir(), '.cache', 'subminer', 'youtube-subs'),
  createSubtitleTempDir: () =>
    fs.promises.mkdtemp(path.join(os.tmpdir(), 'subminer-youtube-subtitles-')),
  cleanupSubtitleTempDirs: (dirs) => {
    for (const dir of dirs) {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  },
});
const prepareYoutubePlaybackInMpv = createPrepareYoutubePlaybackInMpvHandler({
  requestPath: async () => {
    const client = appState.mpvClient;
    if (!client) return null;
    const value = await client.requestProperty('path').catch(() => null);
    return typeof value === 'string' ? value : null;
  },
  requestProperty: async (name) => {
    const client = appState.mpvClient;
    if (!client) return null;
    return await client.requestProperty(name);
  },
  sendMpvCommand: (command) => {
    sendMpvCommandRuntime(appState.mpvClient, command);
  },
  wait: (ms) => new Promise((resolve) => setTimeout(resolve, ms)),
});
const waitForYoutubeMpvConnected = createWaitForMpvConnectedHandler({
  getMpvClient: () => appState.mpvClient,
  now: () => Date.now(),
  sleep: (delayMs) => new Promise((resolve) => setTimeout(resolve, delayMs)),
});
const autoplayReadyGate = createAutoplayReadyGate({
  isAppOwnedFlowInFlight: () => youtubePrimarySubtitleNotificationRuntime.isAppOwnedFlowInFlight(),
  getCurrentMediaPath: () => appState.currentMediaPath,
  getCurrentVideoPath: () => appState.mpvClient?.currentVideoPath ?? null,
  getPlaybackPaused: () => appState.playbackPaused,
  getMpvClient: () => appState.mpvClient,
  signalPluginAutoplayReady: () => {
    sendMpvCommandRuntime(appState.mpvClient, ['script-message', 'subminer-autoplay-ready']);
  },
  isSignalTargetReady: () => {
    if (!overlayManager.getVisibleOverlayVisible()) {
      return true;
    }
    const overlayWindow = overlayManager.getMainWindow();
    return Boolean(overlayWindow && isOverlayWindowContentReady(overlayWindow));
  },
  schedule: (callback, delayMs) => setTimeout(callback, delayMs),
  logDebug: (message) => logger.debug(message),
});
const managedLocalSubtitleSelectionRuntime = createManagedLocalSubtitleSelectionRuntime({
  getCurrentMediaPath: () => appState.currentMediaPath,
  getMpvClient: () => appState.mpvClient,
  getPrimarySubtitleLanguages: () => getResolvedConfig().youtube.primarySubLanguages,
  getSecondarySubtitleLanguages: () => getResolvedConfig().secondarySub.secondarySubLanguages,
  sendMpvCommand: (command) => {
    sendMpvCommandRuntime(appState.mpvClient, command);
  },
  schedule: (callback, delayMs) => setTimeout(callback, delayMs),
  clearScheduled: (timer) => clearTimeout(timer),
});

function resolveBundledMpvRuntimePluginEntrypoint(): string | undefined {
  return (
    resolvePackagedRuntimePluginPath({
      dirname: __dirname,
      appPath: app.getAppPath(),
      resourcesPath: process.resourcesPath,
    }) ?? undefined
  );
}

function detectWindowsInstalledMpvPlugin(mpvExecutablePath: string) {
  return detectInstalledMpvPlugin({
    platform: 'win32',
    homeDir: os.homedir(),
    appDataDir: app.getPath('appData'),
    mpvExecutablePath,
  });
}

function logInstalledMpvPluginDetected(detection: { path: string | null; version: string | null }) {
  if (!detection.path) return;
  logger.warn(
    `SubMiner detected an installed mpv plugin at ${detection.path}. This mpv session will use the installed plugin. Remove it to use the bundled runtime plugin automatically. Detected plugin version: ${detection.version ?? 'unknown or legacy'}.`,
  );
}

async function promptForLegacyMpvPluginRemovalBeforeWindowsLaunch(
  mpvPath: string,
  detection: { path: string | null; version: string | null },
): Promise<'removed' | 'continue' | 'cancel'> {
  const response = await dialog.showMessageBox({
    type: 'warning',
    title: 'SubMiner mpv plugin detected',
    message: [
      'SubMiner detected an installed mpv plugin at:',
      detection.path ?? 'unknown path',
      '',
      "This mpv session will use the installed plugin unless it is removed. Remove it now to use SubMiner's bundled runtime plugin automatically.",
      `Detected plugin version: ${detection.version ?? 'unknown or legacy'}`,
    ].join('\n'),
    detail:
      'Remove the legacy SubMiner mpv plugin files from mpv before launching this video? This moves the files to the OS trash.',
    buttons: ['Remove legacy plugin', 'Continue with installed plugin', 'Cancel'],
    defaultId: 0,
    cancelId: 2,
  });

  if (response.response === 2) {
    return 'cancel';
  }
  if (response.response === 1) {
    return 'continue';
  }

  const result = await removeLegacyMpvPluginCandidates({
    candidates: detectInstalledFirstRunPluginCandidates({
      platform: 'win32',
      homeDir: os.homedir(),
      appDataDir: app.getPath('appData'),
      mpvExecutablePath: mpvPath,
    }),
    trashItem: (candidatePath) => shell.trashItem(candidatePath),
  });
  if (result.ok) {
    await dialog.showMessageBox({
      type: 'info',
      title: 'Legacy mpv plugin removed',
      message:
        'Legacy mpv plugin removed. SubMiner-managed playback will use the bundled runtime plugin.',
    });
    return 'removed';
  }

  await dialog.showMessageBox({
    type: 'error',
    title: 'Could not remove legacy mpv plugin',
    message: 'Some legacy SubMiner mpv plugin files could not be moved to the trash.',
    detail: result.failedPaths.map((failure) => `${failure.path}: ${failure.message}`).join('\n'),
  });
  return 'cancel';
}

const youtubePlaybackRuntime = createYoutubePlaybackRuntime({
  platform: process.platform,
  directPlaybackFormat: YOUTUBE_DIRECT_PLAYBACK_FORMAT,
  mpvYtdlFormat: YOUTUBE_MPV_YTDL_FORMAT,
  autoLaunchTimeoutMs: YOUTUBE_MPV_AUTO_LAUNCH_TIMEOUT_MS,
  connectTimeoutMs: YOUTUBE_MPV_CONNECT_TIMEOUT_MS,
  getSocketPath: () => appState.mpvSocketPath,
  getMpvConnected: () => Boolean(appState.mpvClient?.connected),
  invalidatePendingAutoplayReadyFallbacks: () =>
    autoplayReadyGate.invalidatePendingAutoplayReadyFallbacks(),
  setAppOwnedFlowInFlight: (next) => {
    youtubePrimarySubtitleNotificationRuntime.setAppOwnedFlowInFlight(next);
  },
  ensureYoutubePlaybackRuntimeReady: async () => {
    await ensureYoutubePlaybackRuntimeReady();
  },
  resolveYoutubePlaybackUrl: (url, format) => resolveYoutubePlaybackUrl(url, format),
  launchWindowsMpv: (playbackUrl, args) =>
    launchWindowsMpv(
      [playbackUrl],
      createWindowsMpvLaunchDeps({
        showError: (title, content) => dialog.showErrorBox(title, content),
      }),
      [...args, `--log-file=${DEFAULT_MPV_LOG_PATH}`],
      process.execPath,
      resolveBundledMpvRuntimePluginEntrypoint(),
      getResolvedConfig().mpv.executablePath,
      getResolvedConfig().mpv.launchMode,
      {
        detectInstalledMpvPlugin: detectWindowsInstalledMpvPlugin,
        notifyInstalledPluginDetected: logInstalledMpvPluginDetected,
        resolveInstalledPluginBeforeLaunch: (detection, mpvPath) =>
          promptForLegacyMpvPluginRemovalBeforeWindowsLaunch(mpvPath, detection),
      },
      getMpvPluginRuntimeConfig(),
    ),
  waitForYoutubeMpvConnected: (timeoutMs) => waitForYoutubeMpvConnected(timeoutMs),
  prepareYoutubePlaybackInMpv: (request) => prepareYoutubePlaybackInMpv(request),
  runYoutubePlaybackFlow: (request) => youtubeFlowRuntime.runYoutubePlaybackFlow(request),
  logInfo: (message) => logger.info(message),
  logWarn: (message) => logger.warn(message),
  schedule: (callback, delayMs) => setTimeout(callback, delayMs),
  clearScheduled: (timer) => clearTimeout(timer),
});

function getMpvPluginRuntimeConfig() {
  const config = getResolvedConfig();
  return {
    socketPath: appState.mpvSocketPath,
    binaryPath: config.mpv.subminerBinaryPath,
    backend: config.mpv.backend,
    autoStart: config.mpv.autoStartSubMiner,
    autoStartVisibleOverlay: config.auto_start_overlay,
    autoStartPauseUntilReady: config.mpv.pauseUntilOverlayReady,
    texthookerEnabled: config.texthooker.launchAtStartup,
    aniskipEnabled: config.mpv.aniskipEnabled,
    aniskipButtonKey: config.mpv.aniskipButtonKey,
  };
}

let firstRunSetupMessage: string | null = null;
const resolveWindowsMpvShortcutRuntimePaths = () =>
  resolveWindowsMpvShortcutPaths({
    appDataDir: app.getPath('appData'),
    desktopDir: app.getPath('desktop'),
  });
const createCommandLineLauncherRuntimeOptions = () => ({
  platform: process.platform,
  env: process.env,
  homeDir: os.homedir(),
  localAppData: process.env.LOCALAPPDATA,
  userProfile: process.env.USERPROFILE,
  cwd: process.cwd(),
  resourcesPath: process.resourcesPath,
  appExePath: process.execPath,
});
const firstRunSetupService = createFirstRunSetupService({
  platform: process.platform,
  configDir: CONFIG_DIR,
  getYomitanDictionaryCount: async () => {
    await ensureYomitanExtensionLoaded();
    const dictionaries = await getYomitanDictionaryInfo(getYomitanParserRuntimeDeps(), {
      error: (message, ...args) => logger.error(message, ...args),
      info: (message, ...args) => logger.info(message, ...args),
    });
    return dictionaries.length;
  },
  isExternalYomitanConfigured: () =>
    getResolvedConfig().yomitan.externalProfilePath.trim().length > 0,
  detectPluginInstalled: () => {
    const candidates = detectInstalledFirstRunPluginCandidates({
      platform: process.platform,
      homeDir: os.homedir(),
      xdgConfigHome: process.env.XDG_CONFIG_HOME,
      appDataDir: app.getPath('appData'),
      mpvExecutablePath: getResolvedConfig().mpv.executablePath,
    });
    if (candidates.length > 0) {
      return true;
    }
    const installPaths = resolveDefaultMpvInstallPaths(
      process.platform,
      os.homedir(),
      process.env.XDG_CONFIG_HOME,
    );
    return detectInstalledFirstRunPlugin(installPaths);
  },
  detectLegacyMpvPluginCandidates: () =>
    detectInstalledFirstRunPluginCandidates({
      platform: process.platform,
      homeDir: os.homedir(),
      xdgConfigHome: process.env.XDG_CONFIG_HOME,
      appDataDir: app.getPath('appData'),
      mpvExecutablePath: getResolvedConfig().mpv.executablePath,
    }),
  removeLegacyMpvPlugins: (candidates) =>
    removeLegacyMpvPluginCandidates({
      candidates,
      trashItem: (candidatePath) => shell.trashItem(candidatePath),
    }),
  detectWindowsMpvShortcuts: () => {
    if (process.platform !== 'win32') {
      return {
        startMenuInstalled: false,
        desktopInstalled: false,
      };
    }
    return detectWindowsMpvShortcuts(resolveWindowsMpvShortcutRuntimePaths());
  },
  applyWindowsMpvShortcuts: async (preferences) => {
    if (process.platform !== 'win32') {
      return {
        ok: true,
        status: 'unknown' as const,
        message: '',
      };
    }
    return applyWindowsMpvShortcuts({
      preferences,
      paths: resolveWindowsMpvShortcutRuntimePaths(),
      exePath: process.execPath,
      writeShortcutLink: (shortcutPath, operation, details) =>
        shell.writeShortcutLink(shortcutPath, operation, details),
    });
  },
  detectCommandLineLauncher: () =>
    detectCommandLineLauncher(createCommandLineLauncherRuntimeOptions()),
  installBun: async () => {
    const snapshot = await installCommandLineBun(createCommandLineLauncherRuntimeOptions());
    return {
      ok: snapshot.status === 'ready',
      message:
        snapshot.message ??
        (snapshot.status === 'ready'
          ? 'Bun is ready. Open a new terminal.'
          : 'Bun installation failed.'),
    };
  },
  installCommandLineLauncher: async () => {
    const snapshot = await installCommandLineLauncher(createCommandLineLauncherRuntimeOptions());
    const ok = snapshot.status === 'ready' || snapshot.status === 'installed_bun_missing';
    return {
      ok,
      installPath: snapshot.installPath,
      message:
        snapshot.message ??
        (ok
          ? 'Command-line launcher installed. Open a new terminal.'
          : 'Command-line launcher installation failed.'),
    };
  },
  onStateChanged: (state) => {
    appState.firstRunSetupCompleted = state.status === 'completed';
    if (appTray) {
      ensureTray();
    }
  },
});
const discordPresenceSessionStartedAtMs = Date.now();
let discordPresenceMediaDurationSec: number | null = null;
const discordPresenceRuntime = createDiscordPresenceRuntime({
  getDiscordPresenceService: () => appState.discordPresenceService,
  isDiscordPresenceEnabled: () => getResolvedConfig().discordPresence.enabled === true,
  getMpvClient: () => appState.mpvClient,
  getCurrentMediaTitle: () => appState.currentMediaTitle,
  getCurrentMediaPath: () => appState.currentMediaPath,
  getCurrentSubtitleText: () => appState.currentSubText,
  getPlaybackPaused: () => appState.playbackPaused,
  getFallbackMediaDurationSec: () => anilistMediaGuessRuntimeState.mediaDurationSec,
  getSessionStartedAtMs: () => discordPresenceSessionStartedAtMs,
  getMediaDurationSec: () => discordPresenceMediaDurationSec,
  setMediaDurationSec: (next) => {
    discordPresenceMediaDurationSec = next;
  },
});

async function initializeDiscordPresenceService(): Promise<void> {
  if (getResolvedConfig().discordPresence.enabled !== true) {
    appState.discordPresenceService = null;
    return;
  }

  appState.discordPresenceService = createDiscordPresenceService({
    config: getResolvedConfig().discordPresence,
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
  shouldKeepSuppressedFromVisibleOverlayBinding: () => shouldSuppressMpvSubtitlesForOverlay(),
  setMpvSubVisibility: (visible) => {
    setMpvSubVisibilityRuntime(appState.mpvClient, visible);
  },
});

function shouldSuppressMpvSubtitlesForOverlay(): boolean {
  return overlayManager.getVisibleOverlayVisible();
}

function syncOverlayMpvSubtitleSuppression(): void {
  if (shouldSuppressMpvSubtitlesForOverlay()) {
    void ensureOverlayMpvSubtitlesHidden();
    return;
  }

  restoreOverlayMpvSubtitles();
}

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
    appState.anilistClientSecretState = transitionAnilistClientSecretState(
      appState.anilistClientSecretState,
      next,
    );
  },
  getRetryQueueState: () => appState.anilistRetryQueueState,
  setRetryQueueState: (next) => {
    appState.anilistRetryQueueState = transitionAnilistRetryQueueState(
      appState.anilistRetryQueueState,
      next,
    );
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
    openOverlayHostedModalWithOsd(
      (deps) => openSubsyncManualModalRuntime(deps, payload),
      'Subsync overlay unavailable.',
      'Failed to open subsync overlay.',
    );
  },
});
const immersionMediaRuntime = createImmersionMediaRuntime(
  buildImmersionMediaRuntimeMainDepsHandler(),
);
const anilistRateLimiter = createAnilistRateLimiter();
const statsCoverArtFetcher = createCoverArtFetcher(
  anilistRateLimiter,
  createLogger('main:stats-cover-art'),
);
const anilistStateRuntime = createAnilistStateRuntime(buildAnilistStateRuntimeMainDepsHandler());
const configDerivedRuntime = createConfigDerivedRuntime(buildConfigDerivedRuntimeMainDepsHandler());
const subsyncRuntime = createMainSubsyncRuntime(buildMainSubsyncRuntimeMainDepsHandler());
const currentMediaTokenizationGate = createCurrentMediaTokenizationGate();
const startupOsdSequencer = createStartupOsdSequencer({
  showOsd: (message) => showMpvOsd(message),
});
const youtubePrimarySubtitleNotificationRuntime = createYoutubePrimarySubtitleNotificationRuntime({
  getPrimarySubtitleLanguages: () => getResolvedConfig().youtube.primarySubLanguages,
  notifyFailure: (message) => reportYoutubeSubtitleFailure(message),
  schedule: (fn, delayMs) => setTimeout(fn, delayMs),
  clearSchedule: clearYoutubePrimarySubtitleNotificationTimer,
  getCurrentSubtitleState: async () => {
    const client = appState.mpvClient;
    if (!client?.connected) {
      return null;
    }
    const [sid, trackList] = await Promise.all([
      client.requestProperty('sid').catch(() => null),
      client.requestProperty('track-list').catch(() => null),
    ]);
    return {
      sid,
      trackList: Array.isArray(trackList) ? trackList : null,
    };
  },
});

function isYoutubePlaybackActiveNow(): boolean {
  return isYoutubePlaybackActive(
    appState.currentMediaPath,
    appState.mpvClient?.currentVideoPath ?? null,
  );
}

function reportYoutubeSubtitleFailure(message: string): void {
  const type = getResolvedConfig().ankiConnect.behavior.notificationType;
  if (type === 'osd' || type === 'both') {
    showMpvOsd(message);
  }
  if (type === 'system' || type === 'both') {
    try {
      showDesktopNotification('SubMiner', { body: message });
    } catch {
      logger.warn(`Unable to show desktop notification: ${message}`);
    }
  }
}

async function openYoutubeTrackPickerFromPlayback(): Promise<void> {
  if (youtubeFlowRuntime.hasActiveSession()) {
    showMpvOsd('YouTube subtitle flow already in progress.');
    return;
  }
  const currentMediaPath =
    appState.currentMediaPath?.trim() || appState.mpvClient?.currentVideoPath?.trim() || '';
  if (!isYoutubePlaybackActiveNow() || !currentMediaPath) {
    showMpvOsd('YouTube subtitle picker is only available during YouTube playback.');
    return;
  }
  await youtubeFlowRuntime.openManualPicker({
    url: currentMediaPath,
  });
}

let appTray: Tray | null = null;
let tokenizeSubtitleDeferred: ((text: string) => Promise<SubtitleData>) | null = null;
function withCurrentSubtitleTiming(payload: SubtitleData): SubtitleData {
  return {
    ...payload,
    startTime: appState.mpvClient?.currentSubStart ?? null,
    endTime: appState.mpvClient?.currentSubEnd ?? null,
  };
}
function emitSubtitlePayload(payload: SubtitleData): void {
  const timedPayload = withCurrentSubtitleTiming(payload);
  appState.currentSubtitleData = timedPayload;
  broadcastToOverlayWindows('subtitle:set', timedPayload);
  subtitleWsService.broadcast(timedPayload, {
    enabled: getResolvedConfig().subtitleStyle.frequencyDictionary.enabled,
    topX: getResolvedConfig().subtitleStyle.frequencyDictionary.topX,
    mode: getResolvedConfig().subtitleStyle.frequencyDictionary.mode,
  });
  annotationSubtitleWsService.broadcast(timedPayload, {
    enabled: getResolvedConfig().subtitleStyle.frequencyDictionary.enabled,
    topX: getResolvedConfig().subtitleStyle.frequencyDictionary.topX,
    mode: getResolvedConfig().subtitleStyle.frequencyDictionary.mode,
  });
  autoplayReadyGate.maybeSignalPluginAutoplayReady(timedPayload, { forceWhilePaused: true });
  subtitlePrefetchService?.resume();
}
const buildSubtitleProcessingControllerMainDepsHandler =
  createBuildSubtitleProcessingControllerMainDepsHandler({
    tokenizeSubtitle: async (text: string) =>
      tokenizeSubtitleDeferred ? await tokenizeSubtitleDeferred(text) : { text, tokens: null },
    emitSubtitle: (payload) => emitSubtitlePayload(payload),
    logDebug: (message) => {
      logger.debug(`[subtitle-processing] ${message}`);
    },
    now: () => Date.now(),
  });
const subtitleProcessingControllerMainDeps = buildSubtitleProcessingControllerMainDepsHandler();
const subtitleProcessingController = createSubtitleProcessingController(
  subtitleProcessingControllerMainDeps,
);

let subtitlePrefetchService: SubtitlePrefetchService | null = null;
let subtitlePrefetchRefreshTimer: ReturnType<typeof setTimeout> | null = null;
let lastObservedTimePos = 0;
let cancelLinuxMpvFullscreenOverlayRefreshBurst: CancelLinuxMpvFullscreenOverlayRefreshBurst | null =
  null;
const SEEK_THRESHOLD_SECONDS = 3;
const AUTOPLAY_SUBTITLE_PRIME_LOOKAHEAD_SECONDS = 2;
let autoplaySubtitlePrimedMediaPath: string | null = null;

function getCurrentAutoplayMediaPath(): string | null {
  return appState.currentMediaPath?.trim() || appState.mpvClient?.currentVideoPath?.trim() || null;
}

function isCurrentAutoplayMediaPath(mediaPath: string): boolean {
  return getCurrentAutoplayMediaPath() === mediaPath;
}

function markAutoplaySubtitlePrimeConsumed(mediaPath: string): boolean {
  if (autoplaySubtitlePrimedMediaPath === mediaPath) {
    return false;
  }
  autoplaySubtitlePrimedMediaPath = mediaPath;
  return true;
}

function emitAutoplayPrimedSubtitle(mediaPath: string, text: string): boolean {
  if (!text.trim() || !isCurrentAutoplayMediaPath(mediaPath)) {
    return false;
  }
  if (!markAutoplaySubtitlePrimeConsumed(mediaPath)) {
    return false;
  }

  appState.currentSubText = text;
  const rawPayload = withCurrentSubtitleTiming({ text, tokens: null });
  appState.currentSubtitleData = rawPayload;
  broadcastToOverlayWindows('subtitle:set', rawPayload);
  subtitlePrefetchService?.pause();
  subtitleProcessingController.onSubtitleChange(text);
  return true;
}

async function primeCurrentSubtitleForAutoplay(mediaPath: string): Promise<void> {
  const client = appState.mpvClient;
  if (!client?.connected || !isCurrentAutoplayMediaPath(mediaPath)) {
    return;
  }

  const subTextRaw = await client.requestProperty('sub-text').catch((error) => {
    logger.debug(
      `[autoplay-subtitle-prime] failed to read sub-text: ${
        error instanceof Error ? error.message : String(error)
      }`,
    );
    return null;
  });
  const text = typeof subTextRaw === 'string' ? subTextRaw : '';
  emitAutoplayPrimedSubtitle(mediaPath, text);
}

async function primeAutoplaySubtitleFromParsedCues(
  mediaPath: string,
  cues: SubtitleCue[],
): Promise<void> {
  if (
    cues.length === 0 ||
    autoplaySubtitlePrimedMediaPath === mediaPath ||
    !isCurrentAutoplayMediaPath(mediaPath)
  ) {
    return;
  }

  const client = appState.mpvClient;
  const timePosRaw = await client?.requestProperty('time-pos').catch(() => null);
  const currentTimeSeconds = Number(
    timePosRaw ?? client?.currentTimePos ?? lastObservedTimePos ?? 0,
  );
  const cue = selectAutoplayStartupCue(
    cues,
    Number.isFinite(currentTimeSeconds) ? currentTimeSeconds : 0,
    AUTOPLAY_SUBTITLE_PRIME_LOOKAHEAD_SECONDS,
  );
  if (!cue) {
    return;
  }

  emitAutoplayPrimedSubtitle(mediaPath, cue.text);
}

function clearScheduledSubtitlePrefetchRefresh(): void {
  if (subtitlePrefetchRefreshTimer) {
    clearTimeout(subtitlePrefetchRefreshTimer);
    subtitlePrefetchRefreshTimer = null;
  }
}

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
  isCacheFull: () => subtitleProcessingController.isCacheFull(),
  logInfo: (message) => logger.info(message),
  logWarn: (message) => logger.warn(message),
  onParsedSubtitleCuesChanged: (cues, sourceKey) => {
    appState.activeParsedSubtitleCues = cues ?? [];
    appState.activeParsedSubtitleSource = sourceKey;
    if (!cues?.length) {
      appState.activeParsedSubtitleMediaPath = null;
    }
    const mediaPath = getCurrentAutoplayMediaPath();
    if (mediaPath && cues?.length) {
      void primeAutoplaySubtitleFromParsedCues(mediaPath, cues).catch((error) => {
        logger.debug(
          `[autoplay-subtitle-prime] failed to prime from parsed cues: ${
            error instanceof Error ? error.message : String(error)
          }`,
        );
      });
    }
  },
});
const resolveActiveSubtitleSidebarSourceHandler = createResolveActiveSubtitleSidebarSourceHandler({
  getFfmpegPath: () => getResolvedConfig().subsync.ffmpeg_path.trim() || 'ffmpeg',
  extractInternalSubtitleTrack: (ffmpegPath, videoPath, track) =>
    extractInternalSubtitleTrackToTempFile(ffmpegPath, videoPath, track),
});

async function refreshSubtitleSidebarFromSource(
  sourcePath: string,
  mediaPath?: string,
): Promise<void> {
  const normalizedSourcePath = resolveSubtitleSourcePath(sourcePath.trim());
  if (!normalizedSourcePath) {
    return;
  }
  const nextMediaPath = mediaPath?.trim() || getCurrentAutoplayMediaPath();
  await subtitlePrefetchInitController.initSubtitlePrefetch(
    normalizedSourcePath,
    lastObservedTimePos,
    normalizedSourcePath,
  );
  appState.activeParsedSubtitleMediaPath = nextMediaPath;
}
const refreshSubtitlePrefetchFromActiveTrackHandler =
  createRefreshSubtitlePrefetchFromActiveTrackHandler({
    getMpvClient: () => appState.mpvClient,
    getLastObservedTimePos: () => lastObservedTimePos,
    shouldKeepExistingCuesOnMissingSource: (videoPath) => isYoutubeMediaPath(videoPath),
    subtitlePrefetchInitController,
    resolveActiveSubtitleSidebarSource: (input) => resolveActiveSubtitleSidebarSourceHandler(input),
  });

function scheduleSubtitlePrefetchRefresh(delayMs = 0): void {
  clearScheduledSubtitlePrefetchRefresh();
  subtitlePrefetchRefreshTimer = setTimeout(() => {
    subtitlePrefetchRefreshTimer = null;
    void refreshSubtitlePrefetchFromActiveTrackHandler();
  }, delayMs);
}
const subtitlePrefetchRuntime = {
  cancelPendingInit: () => subtitlePrefetchInitController.cancelPendingInit(),
  initSubtitlePrefetch: subtitlePrefetchInitController.initSubtitlePrefetch,
  refreshSubtitleSidebarFromSource: (sourcePath: string, mediaPath?: string) =>
    refreshSubtitleSidebarFromSource(sourcePath, mediaPath),
  refreshSubtitlePrefetchFromActiveTrack: () => refreshSubtitlePrefetchFromActiveTrackHandler(),
  scheduleSubtitlePrefetchRefresh: (delayMs?: number) => scheduleSubtitlePrefetchRefresh(delayMs),
  clearScheduledSubtitlePrefetchRefresh: () => clearScheduledSubtitlePrefetchRefresh(),
} as const;

const overlayShortcutsRuntime = createOverlayShortcutsRuntimeService(
  createBuildOverlayShortcutsRuntimeMainDepsHandler({
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
    showMpvOsd: (text: string) => showMpvOsd(text),
    openRuntimeOptionsPalette: () => {
      openRuntimeOptionsPalette();
    },
    openCharacterDictionary: () => {
      openCharacterDictionaryOverlay();
    },
    openJimaku: () => {
      openJimakuOverlay();
    },
    markAudioCard: () => markLastCardAsAudioCard(),
    copySubtitleMultiple: (timeoutMs: number) => {
      startPendingMultiCopy(timeoutMs);
    },
    copySubtitle: () => {
      copyCurrentSubtitle();
    },
    toggleSecondarySubMode: () => handleCycleSecondarySubMode(),
    updateLastCardFromClipboard: () => updateLastCardFromClipboard(),
    triggerFieldGrouping: () => triggerFieldGrouping(),
    triggerSubsyncFromConfig: () => triggerSubsyncFromConfig(),
    mineSentenceCard: () => mineSentenceCard(),
    mineSentenceMultiple: (timeoutMs: number) => {
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
syncOverlayShortcutsForModal = (isActive: boolean): void => {
  if (isActive) {
    overlayShortcutsRuntime.unregisterOverlayShortcuts();
  } else {
    overlayShortcutsRuntime.syncOverlayShortcuts();
  }
};

const buildConfigHotReloadMessageMainDepsHandler = createBuildConfigHotReloadMessageMainDepsHandler(
  {
    showMpvOsd: (message) => showMpvOsd(message),
    showDesktopNotification: (title, options) => showDesktopNotification(title, options),
  },
);
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
const buildConfigHotReloadAppliedMainDepsHandler = createBuildConfigHotReloadAppliedMainDepsHandler(
  {
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
      broadcastToOverlayWindows(channel, payload);
    },
    applyAnkiRuntimeConfigPatch: (patch) => {
      if (appState.ankiIntegration) {
        appState.ankiIntegration.applyRuntimeConfigPatch(patch);
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
  },
);
const applyConfigHotReloadDiff = createConfigHotReloadAppliedHandler(
  buildConfigHotReloadAppliedMainDepsHandler(),
);
const buildConfigHotReloadRuntimeMainDepsHandler = createBuildConfigHotReloadRuntimeMainDepsHandler(
  {
    getCurrentConfig: () => getResolvedConfig(),
    reloadConfigStrict: () => configService.reloadConfigStrict(),
    watchConfigPath: (configPath, onChange) => watchConfigPathHandler(configPath, onChange),
    setTimeout: (callback, delayMs) => setTimeout(callback, delayMs),
    clearTimeout: (timeout) => clearTimeout(timeout),
    debounceMs: 250,
    onHotReloadApplied: applyConfigHotReloadDiff,
    onRestartRequired: (fields) =>
      notifyConfigHotReloadMessage(buildRestartRequiredConfigMessage(fields)),
    onInvalidConfig: notifyConfigHotReloadMessage,
    onValidationWarnings: (configPath, warnings) => {
      showDesktopNotification('SubMiner', {
        body: buildConfigWarningNotificationBody(configPath, warnings),
      });
      if (process.platform === 'darwin') {
        dialog.showErrorBox(
          'SubMiner config validation warning',
          buildConfigWarningDialogDetails(configPath, warnings),
        );
      }
    },
  },
);
const configHotReloadRuntime = createConfigHotReloadRuntime(
  buildConfigHotReloadRuntimeMainDepsHandler(),
);

const configSettingsRuntime = createConfigSettingsRuntime({
  fields: configSettingsFields,
  getConfigPath: () => configService.getConfigPath(),
  getRawConfig: () => configService.getRawConfig(),
  getConfig: () => configService.getConfig(),
  getWarnings: () => configService.getWarnings(),
  reloadConfigStrict: () => configService.reloadConfigStrict(),
  onHotReloadApplied: applyConfigHotReloadDiff,
  defaultAnkiConnectUrl: DEFAULT_CONFIG.ankiConnect.url,
  createAnkiClient: (url) => new AnkiConnectClient(url),
  getSettingsWindow: () => appState.configSettingsWindow,
  setSettingsWindow: (window) => {
    appState.configSettingsWindow = window as BrowserWindow | null;
  },
  createSettingsWindow: createCreateConfigSettingsWindowHandler({
    createBrowserWindow: (options) => new BrowserWindow(options),
    preloadPath: path.join(__dirname, 'preload-settings.js'),
  }),
  settingsHtmlPath: path.join(__dirname, 'settings', 'index.html'),
  openPath: (targetPath) => shell.openPath(targetPath),
  ipcMain,
  ipcChannels: IPC_CHANNELS.request,
  log: (message) => logger.error(message),
});

configSettingsRuntime.registerHandlers();
const openConfigSettingsWindow = () => configSettingsRuntime.openWindow();

const buildDictionaryRootsHandler = createBuildDictionaryRootsMainHandler({
  platform: process.platform,
  dirname: __dirname,
  appPath: app.getAppPath(),
  resourcesPath: process.resourcesPath,
  userDataPath: USER_DATA_PATH,
  appUserDataPath: app.getPath('userData'),
  homeDir: os.homedir(),
  appDataDir: process.env.APPDATA,
  cwd: process.cwd(),
  joinPath: (...parts) => path.join(...parts),
});
const buildFrequencyDictionaryRootsHandler = createBuildFrequencyDictionaryRootsMainHandler({
  platform: process.platform,
  dirname: __dirname,
  appPath: app.getAppPath(),
  resourcesPath: process.resourcesPath,
  userDataPath: USER_DATA_PATH,
  appUserDataPath: app.getPath('userData'),
  homeDir: os.homedir(),
  appDataDir: process.env.APPDATA,
  cwd: process.cwd(),
  joinPath: (...parts) => path.join(...parts),
});

const jlptDictionaryRuntime = createJlptDictionaryRuntimeService(
  createBuildJlptDictionaryRuntimeMainDepsHandler({
    isJlptEnabled: () => getResolvedConfig().subtitleStyle.enableJlpt,
    getDictionaryRoots: () => buildDictionaryRootsHandler(),
    getJlptDictionarySearchPaths,
    setJlptLevelLookup: (lookup) => {
      appState.jlptLevelLookup = lookup;
    },
    logInfo: (message) => logger.info(message),
  })(),
);

const frequencyDictionaryRuntime = createFrequencyDictionaryRuntimeService(
  createBuildFrequencyDictionaryRuntimeMainDepsHandler({
    isFrequencyDictionaryEnabled: () =>
      getResolvedConfig().subtitleStyle.frequencyDictionary.enabled,
    getDictionaryRoots: () => buildFrequencyDictionaryRootsHandler(),
    getFrequencyDictionarySearchPaths,
    getSourcePath: () => getResolvedConfig().subtitleStyle.frequencyDictionary.sourcePath,
    setFrequencyRankLookup: (lookup) => {
      appState.frequencyRankLookup = lookup;
    },
    logInfo: (message) => logger.info(message),
  })(),
);

const buildGetFieldGroupingResolverMainDepsHandler =
  createBuildGetFieldGroupingResolverMainDepsHandler({
    getResolver: () => appState.fieldGroupingResolver,
  });
const getFieldGroupingResolverMainDeps = buildGetFieldGroupingResolverMainDepsHandler();
const getFieldGroupingResolverHandler = createGetFieldGroupingResolverHandler(
  getFieldGroupingResolverMainDeps,
);

function getFieldGroupingResolver(): ((choice: KikuFieldGroupingChoice) => void) | null {
  return getFieldGroupingResolverHandler();
}

const buildSetFieldGroupingResolverMainDepsHandler =
  createBuildSetFieldGroupingResolverMainDepsHandler({
    setResolver: (resolver) => {
      appState.fieldGroupingResolver = resolver;
    },
    nextSequence: () => {
      appState.fieldGroupingResolverSequence += 1;
      return appState.fieldGroupingResolverSequence;
    },
    getSequence: () => appState.fieldGroupingResolverSequence,
  });
const setFieldGroupingResolverMainDeps = buildSetFieldGroupingResolverMainDepsHandler();
const setFieldGroupingResolverHandler = createSetFieldGroupingResolverHandler(
  setFieldGroupingResolverMainDeps,
);

function setFieldGroupingResolver(
  resolver: ((choice: KikuFieldGroupingChoice) => void) | null,
): void {
  setFieldGroupingResolverHandler(resolver);
}

const fieldGroupingOverlayRuntime = createFieldGroupingOverlayRuntime<OverlayHostedModal>(
  createBuildFieldGroupingOverlayMainDepsHandler<OverlayHostedModal>({
    getMainWindow: () => overlayManager.getMainWindow(),
    getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    setVisibleOverlayVisible: (visible) => setVisibleOverlayVisible(visible),
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
const JELLYFIN_SUBTITLE_DELAYS_PATH = path.join(CONFIG_DIR, 'jellyfin-subtitle-delays.json');

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

const characterDictionaryRuntime = createCharacterDictionaryRuntimeService({
  userDataPath: USER_DATA_PATH,
  getCurrentMediaPath: () => appState.currentMediaPath,
  getCurrentMediaTitle: () => appState.currentMediaTitle,
  resolveMediaPathForJimaku: (mediaPath) => mediaRuntime.resolveMediaPathForJimaku(mediaPath),
  guessAnilistMediaInfo: (mediaPath, mediaTitle) => guessAnilistMediaInfo(mediaPath, mediaTitle),
  getCollapsibleSectionOpenState: (section) =>
    getResolvedConfig().anilist.characterDictionary.collapsibleSections[section],
  now: () => Date.now(),
  logInfo: (message) => logger.info(message),
  logWarn: (message) => logger.warn(message),
});

const characterDictionaryAutoSyncRuntime = createCharacterDictionaryAutoSyncRuntimeService({
  userDataPath: USER_DATA_PATH,
  getConfig: () => {
    const config = getResolvedConfig().anilist.characterDictionary;
    return {
      enabled:
        config.enabled &&
        yomitanProfilePolicy.isCharacterDictionaryEnabled() &&
        !isYoutubePlaybackActiveNow(),
      maxLoaded: config.maxLoaded,
      profileScope: config.profileScope,
    };
  },
  getOrCreateCurrentSnapshot: (targetPath, progress) =>
    characterDictionaryRuntime.getOrCreateCurrentSnapshot(targetPath, progress),
  buildMergedDictionary: (mediaIds) => characterDictionaryRuntime.buildMergedDictionary(mediaIds),
  waitForYomitanMutationReady: () =>
    currentMediaTokenizationGate.waitUntilReady(
      appState.currentMediaPath?.trim() || appState.mpvClient?.currentVideoPath?.trim() || null,
    ),
  getYomitanDictionaryInfo: async () => {
    await ensureYomitanExtensionLoaded();
    return await getYomitanDictionaryInfo(getYomitanParserRuntimeDeps(), {
      error: (message, ...args) => logger.error(message, ...args),
      info: (message, ...args) => logger.info(message, ...args),
    });
  },
  importYomitanDictionary: async (zipPath) => {
    if (yomitanProfilePolicy.isExternalReadOnlyMode()) {
      yomitanProfilePolicy.logSkippedWrite(
        formatSkippedYomitanWriteAction('importYomitanDictionary', zipPath),
      );
      return false;
    }
    await ensureYomitanExtensionLoaded();
    return await importYomitanDictionaryFromZip(zipPath, getYomitanParserRuntimeDeps(), {
      error: (message, ...args) => logger.error(message, ...args),
      info: (message, ...args) => logger.info(message, ...args),
    });
  },
  deleteYomitanDictionary: async (dictionaryTitle) => {
    if (yomitanProfilePolicy.isExternalReadOnlyMode()) {
      yomitanProfilePolicy.logSkippedWrite(
        formatSkippedYomitanWriteAction('deleteYomitanDictionary', dictionaryTitle),
      );
      return false;
    }
    await ensureYomitanExtensionLoaded();
    return await deleteYomitanDictionaryByTitle(dictionaryTitle, getYomitanParserRuntimeDeps(), {
      error: (message, ...args) => logger.error(message, ...args),
      info: (message, ...args) => logger.info(message, ...args),
    });
  },
  upsertYomitanDictionarySettings: async (dictionaryTitle, profileScope) => {
    if (yomitanProfilePolicy.isExternalReadOnlyMode()) {
      yomitanProfilePolicy.logSkippedWrite(
        formatSkippedYomitanWriteAction('upsertYomitanDictionarySettings', dictionaryTitle),
      );
      return false;
    }
    await ensureYomitanExtensionLoaded();
    return await upsertYomitanDictionarySettings(
      dictionaryTitle,
      profileScope,
      getYomitanParserRuntimeDeps(),
      {
        error: (message, ...args) => logger.error(message, ...args),
        info: (message, ...args) => logger.info(message, ...args),
      },
    );
  },
  now: () => Date.now(),
  schedule: (fn, delayMs) => setTimeout(fn, delayMs),
  clearSchedule: (timer) => clearTimeout(timer),
  logInfo: (message) => logger.info(message),
  logWarn: (message) => logger.warn(message),
  onSyncStatus: (event) => {
    notifyCharacterDictionaryAutoSyncStatus(event, {
      getNotificationType: () => getResolvedConfig().ankiConnect.behavior.notificationType,
      showOsd: (message) => showMpvOsd(message),
      showDesktopNotification: (title, options) => showDesktopNotification(title, options),
      startupOsdSequencer,
    });
  },
  onSyncComplete: ({ mediaId, mediaTitle, changed }) => {
    handleCharacterDictionaryAutoSyncComplete(
      {
        mediaId,
        mediaTitle,
        changed,
      },
      {
        hasParserWindow: () => Boolean(appState.yomitanParserWindow),
        clearParserCaches: () => {
          if (appState.yomitanParserWindow) {
            clearYomitanParserCachesForWindow(appState.yomitanParserWindow);
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
        logInfo: (message) => logger.info(message),
      },
    );
  },
});

const overlayVisibilityRuntime = createOverlayVisibilityRuntimeService(
  createBuildOverlayVisibilityRuntimeMainDepsHandler({
    getMainWindow: () => overlayManager.getMainWindow(),
    getModalActive: () => overlayModalInputState.getModalInputExclusive(),
    getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    getForceMousePassthrough: () => appState.statsOverlayVisible,
    getSuspendVisibleOverlay: () => appState.statsOverlayVisible,
    getOverlayInteractionActive: () => visibleOverlayInteractionActive,
    getWindowTracker: () => appState.windowTracker,
    getLastKnownWindowsForegroundProcessName: () => lastWindowsVisibleOverlayForegroundProcessName,
    getWindowsOverlayProcessName: () => path.parse(process.execPath).name.toLowerCase(),
    getWindowsFocusHandoffGraceActive: () => hasWindowsVisibleOverlayFocusHandoffGrace(),
    getMacOSForegroundProbeActive: () => macOSVisibleOverlayForegroundProbeActive,
    getTrackerNotReadyWarningShown: () => appState.trackerNotReadyWarningShown,
    setTrackerNotReadyWarningShown: (shown: boolean) => {
      appState.trackerNotReadyWarningShown = shown;
    },
    updateVisibleOverlayBounds: (geometry: WindowGeometry) => updateVisibleOverlayBounds(geometry),
    ensureOverlayWindowLevel: (window) => {
      ensureOverlayWindowLevel(window);
    },
    syncWindowsOverlayToMpvZOrder: (_window) => {
      requestWindowsVisibleOverlayZOrderSync();
    },
    syncPrimaryOverlayWindowLayer: (layer) => {
      syncPrimaryOverlayWindowLayer(layer);
    },
    enforceOverlayLayerOrder: () => {
      enforceOverlayLayerOrder();
    },
    syncOverlayShortcuts: () => {
      overlayShortcutsRuntime.syncOverlayShortcuts();
    },
    isMacOSPlatform: () => process.platform === 'darwin',
    isWindowsPlatform: () => process.platform === 'win32',
    showOverlayLoadingOsd: (message: string) => {
      showMpvOsd(message);
    },
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
  })(),
);

const VISIBLE_OVERLAY_BLUR_REFRESH_DELAYS_MS = [0, 25, 100, 250] as const;
const WINDOWS_VISIBLE_OVERLAY_Z_ORDER_RETRY_DELAYS_MS = [0, 48, 120, 240, 480] as const;
const WINDOWS_VISIBLE_OVERLAY_FOREGROUND_POLL_INTERVAL_MS = 75;
const WINDOWS_VISIBLE_OVERLAY_FOCUS_HANDOFF_GRACE_MS = 200;
const MACOS_VISIBLE_OVERLAY_FOREGROUND_PROBE_TIMEOUT_MS = 1_200;
let visibleOverlayBlurRefreshTimeouts: Array<ReturnType<typeof setTimeout>> = [];
let windowsVisibleOverlayZOrderRetryTimeouts: Array<ReturnType<typeof setTimeout>> = [];
let windowsVisibleOverlayZOrderSyncInFlight = false;
let windowsVisibleOverlayZOrderSyncQueued = false;
let windowsVisibleOverlayForegroundPollInterval: ReturnType<typeof setInterval> | null = null;
let lastWindowsVisibleOverlayForegroundProcessName: string | null = null;
let lastWindowsVisibleOverlayBlurredAtMs = 0;
let visibleOverlayInteractionActive = false;
let macOSVisibleOverlayForegroundProbeActive = false;
let macOSVisibleOverlayForegroundProbeToken = 0;
let macOSVisibleOverlayForegroundProbeTimeout: ReturnType<typeof setTimeout> | null = null;

const handleStatsOverlayVisibilityChanged = createStatsOverlayVisibilityChangeHandler({
  setStatsOverlayVisibleState: (visible) => {
    appState.statsOverlayVisible = visible;
  },
  resetVisibleOverlayInteraction: () => {
    visibleOverlayInteractionActive = false;
  },
  getMainWindow: () => overlayManager.getMainWindow(),
  updateVisibleOverlayVisibility: () => overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
});

function clearVisibleOverlayBlurRefreshTimeouts(): void {
  for (const timeout of visibleOverlayBlurRefreshTimeouts) {
    clearTimeout(timeout);
  }
  visibleOverlayBlurRefreshTimeouts = [];
}

function clearWindowsVisibleOverlayZOrderRetryTimeouts(): void {
  for (const timeout of windowsVisibleOverlayZOrderRetryTimeouts) {
    clearTimeout(timeout);
  }
  windowsVisibleOverlayZOrderRetryTimeouts = [];
}

function finishMacOSVisibleOverlayForegroundProbe(token: number): void {
  if (token !== macOSVisibleOverlayForegroundProbeToken) {
    return;
  }
  if (macOSVisibleOverlayForegroundProbeTimeout !== null) {
    clearTimeout(macOSVisibleOverlayForegroundProbeTimeout);
    macOSVisibleOverlayForegroundProbeTimeout = null;
  }
  if (!macOSVisibleOverlayForegroundProbeActive) {
    return;
  }
  macOSVisibleOverlayForegroundProbeActive = false;
  overlayVisibilityRuntime.updateVisibleOverlayVisibility();
}

function startMacOSVisibleOverlayForegroundProbe(): void {
  if (process.platform !== 'darwin') {
    return;
  }
  const tracker = appState.windowTracker;
  if (!tracker) {
    return;
  }

  macOSVisibleOverlayForegroundProbeActive = true;
  const token = ++macOSVisibleOverlayForegroundProbeToken;
  if (macOSVisibleOverlayForegroundProbeTimeout !== null) {
    clearTimeout(macOSVisibleOverlayForegroundProbeTimeout);
  }
  macOSVisibleOverlayForegroundProbeTimeout = setTimeout(() => {
    finishMacOSVisibleOverlayForegroundProbe(token);
  }, MACOS_VISIBLE_OVERLAY_FOREGROUND_PROBE_TIMEOUT_MS);

  void tracker
    .refreshNow()
    .catch((error) => {
      logger.warn('Failed to refresh macOS frontmost app after overlay blur', error);
    })
    .finally(() => {
      finishMacOSVisibleOverlayForegroundProbe(token);
    });
}

function getWindowsNativeWindowHandle(window: BrowserWindow): string {
  const handle = window.getNativeWindowHandle();
  return handle.length >= 8
    ? handle.readBigUInt64LE(0).toString()
    : BigInt(handle.readUInt32LE(0)).toString();
}

function getWindowsNativeWindowHandleNumber(window: BrowserWindow): number {
  const handle = window.getNativeWindowHandle();
  return handle.length >= 8 ? Number(handle.readBigUInt64LE(0)) : handle.readUInt32LE(0);
}

function resolveWindowsOverlayBindTargetHandle(targetMpvSocketPath?: string | null): number | null {
  if (process.platform !== 'win32') {
    return null;
  }

  try {
    if (targetMpvSocketPath) {
      const windowTracker = appState.windowTracker as {
        getTargetWindowHandle?: () => number | null;
      } | null;
      const trackedHandle = windowTracker?.getTargetWindowHandle?.();
      if (typeof trackedHandle === 'number' && Number.isFinite(trackedHandle)) {
        return trackedHandle;
      }
      return null;
    }
    return findWindowsMpvTargetWindowHandle();
  } catch {
    return null;
  }
}

function createOverlayWindowTracker(override?: string | null, targetMpvSocketPath?: string | null) {
  if (appState.initialArgs && isHeadlessInitialCommand(appState.initialArgs)) {
    return null;
  }
  return createWindowTrackerCore(override, targetMpvSocketPath);
}

function bindVisibleOverlayOwner(): void {
  const mainWindow = overlayManager.getMainWindow();
  if (process.platform !== 'win32' || !mainWindow || mainWindow.isDestroyed()) return;
  const overlayHwnd = getWindowsNativeWindowHandleNumber(mainWindow);
  const targetSocketPath = appState.mpvSocketPath;
  const targetWindowHwnd = resolveWindowsOverlayBindTargetHandle(targetSocketPath);
  if (targetWindowHwnd !== null && bindWindowsOverlayAboveMpv(overlayHwnd, targetWindowHwnd)) {
    return;
  }
  if (targetSocketPath) {
    return;
  }
  const tracker = appState.windowTracker;
  const mpvResult = tracker
    ? (() => {
        try {
          const win32 =
            require('./window-trackers/win32') as typeof import('./window-trackers/win32');
          const poll = win32.findMpvWindows();
          const focused = poll.matches.find((m) => m.isForeground);
          return focused ?? [...poll.matches].sort((a, b) => b.area - a.area)[0] ?? null;
        } catch {
          return null;
        }
      })()
    : null;
  if (!mpvResult) return;
  if (!setWindowsOverlayOwner(overlayHwnd, mpvResult.hwnd)) {
    logger.warn('Failed to set overlay owner via koffi');
  }
}

function releaseVisibleOverlayOwner(): void {
  const mainWindow = overlayManager.getMainWindow();
  if (process.platform !== 'win32' || !mainWindow || mainWindow.isDestroyed()) return;
  const overlayHwnd = getWindowsNativeWindowHandleNumber(mainWindow);
  if (!clearWindowsOverlayOwner(overlayHwnd)) {
    logger.warn('Failed to clear overlay owner via koffi');
  }
}

function startOverlayWindowTrackerForCurrentSocket(): void {
  startOverlayWindowTrackerCore({
    backendOverride: appState.backendOverride,
    getMpvSocketPath: () => appState.mpvSocketPath,
    createWindowTracker: createOverlayWindowTracker,
    setWindowTracker: (tracker) => {
      appState.windowTracker = tracker;
    },
    updateVisibleOverlayBounds: (geometry: WindowGeometry) => updateVisibleOverlayBounds(geometry),
    isVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    updateVisibleOverlayVisibility: () => overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
    refreshCurrentSubtitle: () => {
      subtitleProcessingController.refreshCurrentSubtitle(appState.currentSubText);
    },
    getOverlayWindows: () => getOverlayWindows(),
    syncOverlayShortcuts: () => overlayShortcutsRuntime.syncOverlayShortcuts(),
    bindOverlayOwner: () => bindVisibleOverlayOwner(),
    releaseOverlayOwner: () => releaseVisibleOverlayOwner(),
  });
}

function retargetOverlayWindowTrackerForMpvSocket(
  nextSocketPath: string,
  previousSocketPath: string,
): void {
  if (nextSocketPath === previousSocketPath || !appState.overlayRuntimeInitialized) {
    return;
  }

  const previousTracker = appState.windowTracker;
  if (previousTracker) {
    try {
      previousTracker.stop();
    } catch (error) {
      logger.warn('Failed to stop previous overlay window tracker before retargeting', error);
    }
  }

  releaseVisibleOverlayOwner();
  appState.windowTracker = null;
  appState.trackerNotReadyWarningShown = false;
  lastOverlayWindowGeometry = null;
  startOverlayWindowTrackerForCurrentSocket();
  overlayVisibilityRuntime.updateVisibleOverlayVisibility();
  overlayShortcutsRuntime.syncOverlayShortcuts();
  logger.info(
    `Retargeted overlay window tracker for MPV socket: ${previousSocketPath} -> ${nextSocketPath}`,
  );
}

async function syncWindowsVisibleOverlayToMpvZOrder(): Promise<boolean> {
  if (process.platform !== 'win32') {
    return false;
  }

  const mainWindow = overlayManager.getMainWindow();
  if (
    !mainWindow ||
    mainWindow.isDestroyed() ||
    !mainWindow.isVisible() ||
    !overlayManager.getVisibleOverlayVisible()
  ) {
    return false;
  }

  const windowTracker = appState.windowTracker;
  if (!windowTracker) {
    return false;
  }

  if (
    typeof windowTracker.isTargetWindowMinimized === 'function' &&
    windowTracker.isTargetWindowMinimized()
  ) {
    return false;
  }

  if (!windowTracker.isTracking() && windowTracker.getGeometry() === null) {
    return false;
  }

  const overlayHwnd = getWindowsNativeWindowHandleNumber(mainWindow);
  const targetWindowHwnd = resolveWindowsOverlayBindTargetHandle(appState.mpvSocketPath);
  if (targetWindowHwnd !== null && bindWindowsOverlayAboveMpv(overlayHwnd, targetWindowHwnd)) {
    (mainWindow as BrowserWindow & { setOpacity?: (opacity: number) => void }).setOpacity?.(1);
    return true;
  }
  return false;
}

function requestWindowsVisibleOverlayZOrderSync(): void {
  if (process.platform !== 'win32') {
    return;
  }

  if (windowsVisibleOverlayZOrderSyncInFlight) {
    windowsVisibleOverlayZOrderSyncQueued = true;
    return;
  }

  windowsVisibleOverlayZOrderSyncInFlight = true;
  void syncWindowsVisibleOverlayToMpvZOrder()
    .catch((error) => {
      logger.warn('Failed to bind Windows overlay z-order to mpv', error);
    })
    .finally(() => {
      windowsVisibleOverlayZOrderSyncInFlight = false;
      if (!windowsVisibleOverlayZOrderSyncQueued) {
        return;
      }

      windowsVisibleOverlayZOrderSyncQueued = false;
      requestWindowsVisibleOverlayZOrderSync();
    });
}

function scheduleWindowsVisibleOverlayZOrderSyncBurst(): void {
  if (process.platform !== 'win32') {
    return;
  }

  clearWindowsVisibleOverlayZOrderRetryTimeouts();
  for (const delayMs of WINDOWS_VISIBLE_OVERLAY_Z_ORDER_RETRY_DELAYS_MS) {
    const retryTimeout = setTimeout(() => {
      windowsVisibleOverlayZOrderRetryTimeouts = windowsVisibleOverlayZOrderRetryTimeouts.filter(
        (timeout) => timeout !== retryTimeout,
      );
      requestWindowsVisibleOverlayZOrderSync();
    }, delayMs);
    windowsVisibleOverlayZOrderRetryTimeouts.push(retryTimeout);
  }
}

function hasWindowsVisibleOverlayFocusHandoffGrace(): boolean {
  return (
    process.platform === 'win32' &&
    lastWindowsVisibleOverlayBlurredAtMs > 0 &&
    Date.now() - lastWindowsVisibleOverlayBlurredAtMs <=
      WINDOWS_VISIBLE_OVERLAY_FOCUS_HANDOFF_GRACE_MS
  );
}

function shouldPollWindowsVisibleOverlayForegroundProcess(): boolean {
  if (process.platform !== 'win32' || !overlayManager.getVisibleOverlayVisible()) {
    return false;
  }

  const mainWindow = overlayManager.getMainWindow();
  if (!mainWindow || mainWindow.isDestroyed()) {
    return false;
  }

  const windowTracker = appState.windowTracker;
  if (!windowTracker) {
    return false;
  }

  if (
    typeof windowTracker.isTargetWindowMinimized === 'function' &&
    windowTracker.isTargetWindowMinimized()
  ) {
    return false;
  }

  const overlayFocused = mainWindow.isFocused();
  const trackerFocused = windowTracker.isTargetWindowFocused?.() ?? false;
  return !overlayFocused && !trackerFocused;
}

function maybePollWindowsVisibleOverlayForegroundProcess(): void {
  if (!shouldPollWindowsVisibleOverlayForegroundProcess()) {
    lastWindowsVisibleOverlayForegroundProcessName = null;
    return;
  }

  const processName = getWindowsForegroundProcessName();
  const normalizedProcessName = processName?.trim().toLowerCase() ?? null;
  const previousProcessName = lastWindowsVisibleOverlayForegroundProcessName;
  lastWindowsVisibleOverlayForegroundProcessName = normalizedProcessName;

  if (normalizedProcessName !== previousProcessName) {
    overlayVisibilityRuntime.updateVisibleOverlayVisibility();
  }
  if (normalizedProcessName === 'mpv' && previousProcessName !== 'mpv') {
    requestWindowsVisibleOverlayZOrderSync();
  }
}

function ensureWindowsVisibleOverlayForegroundPollLoop(): void {
  if (process.platform !== 'win32' || windowsVisibleOverlayForegroundPollInterval !== null) {
    return;
  }

  windowsVisibleOverlayForegroundPollInterval = setInterval(() => {
    maybePollWindowsVisibleOverlayForegroundProcess();
  }, WINDOWS_VISIBLE_OVERLAY_FOREGROUND_POLL_INTERVAL_MS);
}

function clearWindowsVisibleOverlayForegroundPollLoop(): void {
  if (windowsVisibleOverlayForegroundPollInterval === null) {
    return;
  }

  clearInterval(windowsVisibleOverlayForegroundPollInterval);
  windowsVisibleOverlayForegroundPollInterval = null;
}

function scheduleVisibleOverlayBlurRefresh(): void {
  if (process.platform !== 'win32' && process.platform !== 'darwin') {
    return;
  }

  if (process.platform === 'win32') {
    lastWindowsVisibleOverlayBlurredAtMs = Date.now();
  }
  startMacOSVisibleOverlayForegroundProbe();
  clearVisibleOverlayBlurRefreshTimeouts();
  for (const delayMs of VISIBLE_OVERLAY_BLUR_REFRESH_DELAYS_MS) {
    const refreshTimeout = setTimeout(() => {
      visibleOverlayBlurRefreshTimeouts = visibleOverlayBlurRefreshTimeouts.filter(
        (timeout) => timeout !== refreshTimeout,
      );
      overlayVisibilityRuntime.updateVisibleOverlayVisibility();
    }, delayMs);
    visibleOverlayBlurRefreshTimeouts.push(refreshTimeout);
  }
}

ensureWindowsVisibleOverlayForegroundPollLoop();

const buildGetRuntimeOptionsStateMainDepsHandler = createBuildGetRuntimeOptionsStateMainDepsHandler(
  {
    getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
  },
);
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
syncOverlayVisibilityForModal = () => {
  overlayVisibilityRuntime.updateVisibleOverlayVisibility();
};

function broadcastToOverlayWindows(channel: string, ...args: unknown[]): void {
  overlayManager.broadcastToOverlayWindows(channel, ...args);
}

const buildBroadcastRuntimeOptionsChangedMainDepsHandler =
  createBuildBroadcastRuntimeOptionsChangedMainDepsHandler({
    broadcastRuntimeOptionsChangedRuntime,
    getRuntimeOptionsState: () => getRuntimeOptionsState(),
    broadcastToOverlayWindows: (channel, ...args) => broadcastToOverlayWindows(channel, ...args),
  });

const buildSendToActiveOverlayWindowMainDepsHandler =
  createBuildSendToActiveOverlayWindowMainDepsHandler({
    sendToActiveOverlayWindowRuntime: (channel, payload, runtimeOptions) =>
      overlayModalRuntime.sendToActiveOverlayWindow(channel, payload, runtimeOptions),
  });

const buildSetOverlayDebugVisualizationEnabledMainDepsHandler =
  createBuildSetOverlayDebugVisualizationEnabledMainDepsHandler({
    setOverlayDebugVisualizationEnabledRuntime,
    getCurrentEnabled: () => appState.overlayDebugVisualizationEnabled,
    setCurrentEnabled: (next) => {
      appState.overlayDebugVisualizationEnabled = next;
    },
  });

const buildOpenRuntimeOptionsPaletteMainDepsHandler =
  createBuildOpenRuntimeOptionsPaletteMainDepsHandler({
    openRuntimeOptionsPaletteRuntime: () => overlayModalRuntime.openRuntimeOptionsPalette(),
  });
const overlayVisibilityComposer = composeOverlayVisibilityRuntime({
  overlayVisibilityRuntime,
  restorePreviousSecondarySubVisibilityMainDeps:
    buildRestorePreviousSecondarySubVisibilityMainDepsHandler(),
  broadcastRuntimeOptionsChangedMainDeps: buildBroadcastRuntimeOptionsChangedMainDepsHandler(),
  sendToActiveOverlayWindowMainDeps: buildSendToActiveOverlayWindowMainDepsHandler(),
  setOverlayDebugVisualizationEnabledMainDeps:
    buildSetOverlayDebugVisualizationEnabledMainDepsHandler(),
  openRuntimeOptionsPaletteMainDeps: buildOpenRuntimeOptionsPaletteMainDepsHandler(),
});

function restorePreviousSecondarySubVisibility(): void {
  overlayVisibilityComposer.restorePreviousSecondarySubVisibility();
}

function broadcastRuntimeOptionsChanged(): void {
  overlayVisibilityComposer.broadcastRuntimeOptionsChanged();
}

function sendToActiveOverlayWindow(
  channel: string,
  payload?: unknown,
  runtimeOptions?: { restoreOnModalClose?: OverlayHostedModal },
): boolean {
  return overlayVisibilityComposer.sendToActiveOverlayWindow(channel, payload, runtimeOptions);
}

function setOverlayDebugVisualizationEnabled(enabled: boolean): void {
  overlayVisibilityComposer.setOverlayDebugVisualizationEnabled(enabled);
}

function createOverlayHostedModalOpenDeps(): {
  ensureOverlayStartupPrereqs: () => void;
  ensureOverlayWindowsReadyForVisibilityActions: () => void;
  sendToActiveOverlayWindow: (
    channel: string,
    payload?: unknown,
    runtimeOptions?: {
      restoreOnModalClose?: OverlayHostedModal;
      preferModalWindow?: boolean;
    },
  ) => boolean;
  waitForModalOpen: (modal: OverlayHostedModal, timeoutMs: number) => Promise<boolean>;
  logWarn: (message: string) => void;
} {
  return {
    ensureOverlayStartupPrereqs: () => ensureOverlayStartupPrereqs(),
    ensureOverlayWindowsReadyForVisibilityActions: () =>
      ensureOverlayWindowsReadyForVisibilityActions(),
    sendToActiveOverlayWindow: (channel, payload, runtimeOptions) =>
      sendToActiveOverlayWindow(channel, payload, runtimeOptions),
    waitForModalOpen: (modal, timeoutMs) => overlayModalRuntime.waitForModalOpen(modal, timeoutMs),
    logWarn: (message) => logger.warn(message),
  };
}

function openOverlayHostedModalWithOsd(
  openModal: (deps: ReturnType<typeof createOverlayHostedModalOpenDeps>) => Promise<boolean>,
  unavailableMessage: string,
  failureLogMessage: string,
): void {
  void openModal(createOverlayHostedModalOpenDeps())
    .then((opened) => {
      if (!opened) {
        showMpvOsd(unavailableMessage);
      }
    })
    .catch((error) => {
      logger.error(failureLogMessage, error);
      showMpvOsd(unavailableMessage);
    });
}

function openRuntimeOptionsPalette(): void {
  openOverlayHostedModalWithOsd(
    openRuntimeOptionsModalRuntime,
    'Runtime options overlay unavailable.',
    'Failed to open runtime options overlay.',
  );
}

function openJimakuOverlay(): void {
  openOverlayHostedModalWithOsd(
    openJimakuModalRuntime,
    'Jimaku overlay unavailable.',
    'Failed to open Jimaku overlay.',
  );
}

function openSessionHelpOverlay(): void {
  openOverlayHostedModalWithOsd(
    openSessionHelpModalRuntime,
    'Session help overlay unavailable.',
    'Failed to open session help overlay.',
  );
}

function openCharacterDictionaryOverlay(): void {
  openOverlayHostedModalWithOsd(
    openCharacterDictionaryModalRuntime,
    'Character dictionary overlay unavailable.',
    'Failed to open character dictionary overlay.',
  );
}

function openControllerSelectOverlay(): void {
  openOverlayHostedModalWithOsd(
    openControllerSelectModalRuntime,
    'Controller select overlay unavailable.',
    'Failed to open controller select overlay.',
  );
}

function openControllerDebugOverlay(): void {
  openOverlayHostedModalWithOsd(
    openControllerDebugModalRuntime,
    'Controller debug overlay unavailable.',
    'Failed to open controller debug overlay.',
  );
}

function openPlaylistBrowser(): void {
  if (!appState.mpvClient?.connected) {
    showMpvOsd('Playlist browser requires active playback.');
    return;
  }
  openOverlayHostedModalWithOsd(
    openPlaylistBrowserRuntime,
    'Playlist browser overlay unavailable.',
    'Failed to open playlist browser overlay.',
  );
}

function getResolvedConfig() {
  return configService.getConfig();
}

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
  const config = getResolvedConfig();
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

const {
  getResolvedJellyfinConfig,
  reportJellyfinRemoteProgress,
  reportJellyfinRemoteStopped,
  startJellyfinRemoteSession,
  stopJellyfinRemoteSession,
  cleanupJellyfinSubtitleCache,
  runJellyfinCommand,
  openJellyfinSetupWindow,
  getJellyfinClientInfo,
} = composeJellyfinRuntimeHandlers({
  getResolvedJellyfinConfigMainDeps: {
    getResolvedConfig: () => getResolvedConfig(),
    loadStoredSession: () => jellyfinTokenStore.loadSession(),
    getEnv: (name) => process.env[name],
  },
  getJellyfinClientInfoMainDeps: {
    getResolvedJellyfinConfig: () => getResolvedJellyfinConfig(),
    getHostName: () => os.hostname(),
    defaultClientName: DEFAULT_JELLYFIN_CLIENT_NAME,
    defaultClientVersion: DEFAULT_JELLYFIN_CLIENT_VERSION,
  },
  waitForMpvConnectedMainDeps: {
    getMpvClient: () => appState.mpvClient,
    now: () => Date.now(),
    sleep: (delayMs) => new Promise((resolve) => setTimeout(resolve, delayMs)),
  },
  launchMpvIdleForJellyfinPlaybackMainDeps: {
    getSocketPath: () => appState.mpvSocketPath,
    getLaunchMode: () => getResolvedConfig().mpv.launchMode,
    platform: process.platform,
    execPath: process.execPath,
    getRuntimePluginEntrypoint: () => resolveBundledMpvRuntimePluginEntrypoint(),
    getInstalledPluginDetection: () =>
      detectInstalledMpvPlugin({
        platform: process.platform,
        homeDir: os.homedir(),
        xdgConfigHome: process.env.XDG_CONFIG_HOME,
        appDataDir: app.getPath('appData'),
        mpvExecutablePath: getResolvedConfig().mpv.executablePath,
      }),
    getPluginRuntimeConfig: () => getMpvPluginRuntimeConfig(),
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
  },
  ensureMpvConnectedForJellyfinPlaybackMainDeps: {
    getMpvClient: () => appState.mpvClient,
    setMpvClient: (client) => {
      appState.mpvClient = client as MpvIpcClient | null;
    },
    createMpvClient: () => createMpvClientRuntimeService(),
    getAutoLaunchInFlight: () => jellyfinMpvAutoLaunchInFlight,
    setAutoLaunchInFlight: (promise) => {
      jellyfinMpvAutoLaunchInFlight = promise;
    },
    connectTimeoutMs: JELLYFIN_MPV_CONNECT_TIMEOUT_MS,
    autoLaunchTimeoutMs: JELLYFIN_MPV_AUTO_LAUNCH_TIMEOUT_MS,
  },
  preloadJellyfinExternalSubtitlesMainDeps: {
    listJellyfinSubtitleTracks: (session, clientInfo, itemId) =>
      listJellyfinSubtitleTracksRuntime(session, clientInfo, itemId),
    getMpvClient: () => appState.mpvClient,
    sendMpvCommand: (command) => {
      sendMpvCommandRuntime(appState.mpvClient, command);
    },
    wait: (ms) => new Promise<void>((resolve) => setTimeout(resolve, ms)),
    cacheSubtitleTrack: (track) => jellyfinSubtitleCacheIo.cacheSubtitleTrack(track),
    cleanupCachedSubtitles: (dirs) => jellyfinSubtitleCacheIo.cleanupCachedSubtitles(dirs),
    getSavedSubtitleDelay: (itemId, streamIndex) =>
      loadJellyfinSubtitleDelay({
        filePath: JELLYFIN_SUBTITLE_DELAYS_PATH,
        itemId,
        streamIndex,
      }),
    setActiveSubtitleDelayKey: (key) => {
      activeJellyfinSubtitleDelayKey = key;
    },
    loadSubtitleSourceText,
    saveSubtitleDelay: (itemId, streamIndex, delaySeconds) =>
      saveJellyfinSubtitleDelay({
        filePath: JELLYFIN_SUBTITLE_DELAYS_PATH,
        itemId,
        streamIndex,
        delaySeconds,
      }),
    logDebug: (message, error) => {
      logger.debug(message, error);
    },
  },
  playJellyfinItemInMpvMainDeps: {
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
    applyJellyfinMpvDefaults: (mpvClient) => applyJellyfinMpvDefaults(mpvClient),
    showVisibleOverlay: () => setVisibleOverlayVisible(true),
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
    setActivePlayback: (state) => {
      activeJellyfinRemotePlayback = {
        ...(state as ActiveJellyfinRemotePlaybackState),
        stopReportsAfterMs:
          state.stopReportsAfterMs ?? Date.now() + JELLYFIN_REMOTE_STARTUP_STOP_GRACE_MS,
      };
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
    updateCurrentMediaTitle: (title) => {
      mediaRuntime.updateCurrentMediaTitle(title);
    },
    recordJellyfinPlaybackMetadata: (metadata) => {
      ensureImmersionTrackerStarted();
      appState.immersionTracker?.recordJellyfinPlaybackMetadata(metadata);
    },
  },
  remoteComposerOptions: {
    getConfiguredSession: () => getConfiguredJellyfinSession(getResolvedJellyfinConfig()),
    logWarn: (message) => logger.warn(message),
    getMpvClient: () => appState.mpvClient,
    sendMpvCommand: (client, command) => sendMpvCommandRuntime(client as MpvIpcClient, command),
    jellyfinTicksToSeconds: (ticks) => jellyfinTicksToSecondsRuntime(ticks),
    getActivePlayback: () => activeJellyfinRemotePlayback,
    clearActivePlayback: () => {
      activeJellyfinRemotePlayback = null;
      activeJellyfinSubtitleDelayKey = null;
    },
    getSession: () => appState.jellyfinRemoteSession,
    getNow: () => Date.now(),
    getLastProgressAtMs: () => jellyfinRemoteLastProgressAtMs,
    setLastProgressAtMs: (value) => {
      jellyfinRemoteLastProgressAtMs = value;
    },
    progressIntervalMs: JELLYFIN_REMOTE_PROGRESS_INTERVAL_MS,
    ticksPerSecond: JELLYFIN_TICKS_PER_SECOND,
    logDebug: (message, error) => logger.debug(message, error),
  },
  handleJellyfinAuthCommandsMainDeps: {
    patchRawConfig: (patch) => {
      configService.patchRawConfig(patch);
      refreshTrayMenuIfPresent();
    },
    authenticateWithPassword: (serverUrl, username, password, clientInfo) =>
      authenticateWithPasswordRuntime(serverUrl, username, password, clientInfo),
    saveStoredSession: (session) => jellyfinTokenStore.saveSession(session),
    clearStoredSession: () =>
      clearJellyfinAuthSessionAndRefreshTrayRuntime(getJellyfinTrayDiscoveryDeps()),
    logInfo: (message) => logger.info(message),
  },
  handleJellyfinListCommandsMainDeps: {
    listJellyfinLibraries: (session, clientInfo) =>
      listJellyfinLibrariesRuntime(session, clientInfo),
    listJellyfinItems: (session, clientInfo, params) =>
      listJellyfinItemsRuntime(session, clientInfo, params),
    listJellyfinSubtitleTracks: (session, clientInfo, itemId) =>
      listJellyfinSubtitleTracksRuntime(session, clientInfo, itemId),
    writeJellyfinPreviewAuth: (responsePath, payload) => {
      fs.mkdirSync(path.dirname(responsePath), { recursive: true });
      fs.writeFileSync(responsePath, JSON.stringify(payload, null, 2), 'utf-8');
    },
    logInfo: (message) => logger.info(message),
  },
  handleJellyfinPlayCommandMainDeps: {
    logWarn: (message) => logger.warn(message),
  },
  handleJellyfinRemoteAnnounceCommandMainDeps: {
    getRemoteSession: () => appState.jellyfinRemoteSession,
    logInfo: (message) => logger.info(message),
    logWarn: (message) => logger.warn(message),
  },
  startJellyfinRemoteSessionMainDeps: {
    getCurrentSession: () => appState.jellyfinRemoteSession,
    setCurrentSession: (session) => {
      appState.jellyfinRemoteSession = session as typeof appState.jellyfinRemoteSession;
    },
    createRemoteSessionService: (options) => new JellyfinRemoteSessionService(options),
    getHostName: () => os.hostname(),
    defaultDeviceId: createHostDerivedJellyfinDeviceId(os.hostname()),
    defaultClientName: DEFAULT_JELLYFIN_CLIENT_NAME,
    defaultClientVersion: DEFAULT_JELLYFIN_CLIENT_VERSION,
    logInfo: (message) => logger.info(message),
    logWarn: (message, details) => logger.warn(message, details),
    onSessionStateChanged: () => refreshTrayMenuIfPresent(),
  },
  stopJellyfinRemoteSessionMainDeps: {
    getCurrentSession: () => appState.jellyfinRemoteSession,
    setCurrentSession: (session) => {
      appState.jellyfinRemoteSession = session as typeof appState.jellyfinRemoteSession;
    },
    clearActivePlayback: () => {
      activeJellyfinRemotePlayback = null;
    },
    onSessionStateChanged: () => refreshTrayMenuIfPresent(),
  },
  runJellyfinCommandMainDeps: {
    defaultServerUrl: DEFAULT_CONFIG.jellyfin.serverUrl,
  },
  maybeFocusExistingJellyfinSetupWindowMainDeps: {
    getSetupWindow: () => appState.jellyfinSetupWindow,
  },
  openJellyfinSetupWindowMainDeps: {
    createSetupWindow: createCreateJellyfinSetupWindowHandler({
      createBrowserWindow: (options) => new BrowserWindow(options),
      preloadPath: JELLYFIN_SETUP_PRELOAD_PATH,
    }),
    buildSetupFormHtml: (state) => buildJellyfinSetupFormHtml(state),
    parseSubmissionUrl: (rawUrl) => parseJellyfinSetupSubmissionUrl(rawUrl),
    authenticateWithPassword: (server, username, password, clientInfo) =>
      authenticateWithPasswordRuntime(server, username, password, clientInfo),
    saveStoredSession: (session) => jellyfinTokenStore.saveSession(session),
    clearStoredSession: () =>
      clearJellyfinAuthSessionAndRefreshTrayRuntime(getJellyfinTrayDiscoveryDeps()),
    patchJellyfinConfig: (session) => {
      const recentServers = mergeJellyfinRecentServers(
        session.serverUrl,
        getResolvedConfig().jellyfin.recentServers || [],
      );
      configService.patchRawConfig({
        jellyfin: {
          enabled: true,
          serverUrl: session.serverUrl,
          username: session.username,
          recentServers,
        },
      });
      refreshTrayMenuIfPresent();
    },
    persistAuthenticatedSession: (session, clientInfo) =>
      persistJellyfinAuthSession({
        session,
        clientInfo,
        existingRecentServers: getResolvedConfig().jellyfin.recentServers || [],
        saveStoredSession: (storedSession) => jellyfinTokenStore.saveSession(storedSession),
        patchRawConfig: (patch) => {
          configService.patchRawConfig(patch);
          refreshTrayMenuIfPresent();
        },
      }),
    logInfo: (message) => logger.info(message),
    logError: (message, error) => logger.error(message, error),
    showMpvOsd: (message) => showMpvOsd(message),
    clearSetupWindow: () => {
      appState.jellyfinSetupWindow = null;
    },
    setSetupWindow: (window) => {
      appState.jellyfinSetupWindow = window as BrowserWindow;
    },
    registerSetupIpcHandler: (handler) => {
      const channel = IPC_CHANNELS.request.jellyfinSetupSubmit;
      ipcMain.removeHandler(channel);
      ipcMain.handle(channel, async (event, payload) => {
        const setupWindow = appState.jellyfinSetupWindow;
        if (!setupWindow || event.sender !== setupWindow.webContents) {
          return {
            handled: false,
            statusMessage: 'This Jellyfin setup window is no longer active.',
            statusKind: 'error',
          };
        }
        return handler(payload);
      });
      return () => {
        ipcMain.removeHandler(channel);
      };
    },
    encodeURIComponent: (value) => encodeURIComponent(value),
    defaultServerUrl: DEFAULT_CONFIG.jellyfin.serverUrl || 'http://127.0.0.1:8096',
    hasStoredSession: () => Boolean(jellyfinTokenStore.loadSession()),
  },
});

const maybeFocusExistingFirstRunSetupWindow = createMaybeFocusExistingFirstRunSetupWindowHandler({
  getSetupWindow: () => appState.firstRunSetupWindow,
});
const openFirstRunSetupWindowHandler = createOpenFirstRunSetupWindowHandler({
  maybeFocusExistingSetupWindow: maybeFocusExistingFirstRunSetupWindow,
  createSetupWindow: createCreateFirstRunSetupWindowHandler({
    createBrowserWindow: (options) => new BrowserWindow(options),
  }),
  getSetupSnapshot: async () => {
    const snapshot = await firstRunSetupService.getSetupStatus();
    const mpvExecutablePath = getResolvedConfig().mpv.executablePath;
    return {
      configReady: snapshot.configReady,
      dictionaryCount: snapshot.dictionaryCount,
      canFinish: snapshot.canFinish,
      externalYomitanConfigured: snapshot.externalYomitanConfigured,
      pluginStatus: snapshot.pluginStatus,
      pluginInstallPathSummary: snapshot.pluginInstallPathSummary,
      legacyMpvPluginPaths: snapshot.legacyMpvPluginPaths,
      mpvExecutablePath,
      mpvExecutablePathStatus: getConfiguredWindowsMpvPathStatus(mpvExecutablePath),
      windowsMpvShortcuts: snapshot.windowsMpvShortcuts,
      commandLineLauncher: snapshot.commandLineLauncher,
      message: firstRunSetupMessage,
    };
  },
  buildSetupHtml: (model) => buildFirstRunSetupHtml(model),
  parseSubmissionUrl: (rawUrl) => parseFirstRunSetupSubmissionUrl(rawUrl),
  handleAction: async (submission: FirstRunSetupSubmission) => {
    if (submission.action === 'remove-legacy-plugin') {
      const snapshot = await firstRunSetupService.removeLegacyMpvPlugin();
      firstRunSetupMessage = snapshot.message;
      return;
    }
    if (submission.action === 'configure-mpv-executable-path') {
      const mpvExecutablePath = submission.mpvExecutablePath?.trim() ?? '';
      const pathStatus = getConfiguredWindowsMpvPathStatus(mpvExecutablePath);
      configService.patchRawConfig({
        mpv: {
          executablePath: mpvExecutablePath,
        },
      });
      firstRunSetupMessage =
        pathStatus === 'invalid'
          ? `Saved mpv executable path, but the file was not found: ${mpvExecutablePath}`
          : mpvExecutablePath
            ? `Saved mpv executable path: ${mpvExecutablePath}`
            : 'Cleared mpv executable path. SubMiner will auto-discover mpv.exe from SUBMINER_MPV_PATH or PATH.';
      return;
    }
    if (submission.action === 'configure-windows-mpv-shortcuts') {
      const snapshot = await firstRunSetupService.configureWindowsMpvShortcuts({
        startMenuEnabled: submission.startMenuEnabled === true,
        desktopEnabled: submission.desktopEnabled === true,
      });
      firstRunSetupMessage = snapshot.message;
      return;
    }
    if (submission.action === 'install-bun') {
      const snapshot = await firstRunSetupService.installBun();
      firstRunSetupMessage = snapshot.message;
      return;
    }
    if (submission.action === 'install-command-line-launcher') {
      const snapshot = await firstRunSetupService.installCommandLineLauncher();
      firstRunSetupMessage = snapshot.message;
      return;
    }
    if (submission.action === 'open-yomitan-settings') {
      firstRunSetupMessage = openYomitanSettings()
        ? 'Opened Yomitan settings. Install dictionaries, then refresh status.'
        : 'Yomitan settings are unavailable while external read-only profile mode is enabled.';
      return;
    }
    if (submission.action === 'open-config-settings') {
      const opened = openConfigSettingsWindow();
      firstRunSetupMessage = opened
        ? 'Opened SubMiner settings.'
        : 'SubMiner settings are unavailable.';
      if (opened) {
        return { skipRender: true };
      }
      return;
    }
    if (submission.action === 'refresh') {
      const snapshot = await firstRunSetupService.refreshStatus('Status refreshed.');
      firstRunSetupMessage = snapshot.message;
      return;
    }

    const snapshot = await firstRunSetupService.markSetupCompleted();
    if (snapshot.state.status === 'completed') {
      firstRunSetupMessage = null;
      return { closeWindow: true };
    }
    firstRunSetupMessage =
      getFirstRunSetupCompletionMessage(snapshot) ??
      'Finish setup requires the mpv plugin and Yomitan dictionaries.';
    return;
  },
  markSetupInProgress: async () => {
    firstRunSetupMessage = null;
    await firstRunSetupService.markSetupInProgress();
  },
  markSetupCancelled: async () => {
    firstRunSetupMessage = null;
    await firstRunSetupService.markSetupCancelled();
  },
  isSetupCompleted: () => firstRunSetupService.isSetupCompleted(),
  shouldQuitWhenClosedIncomplete: () => !appState.backgroundMode,
  shouldQuitWhenClosedCompleted: () =>
    Boolean(appState.initialArgs && isStandaloneFirstRunSetupCommand(appState.initialArgs)),
  quitApp: () => requestAppQuit(),
  clearSetupWindow: () => {
    appState.firstRunSetupWindow = null;
  },
  setSetupWindow: (window) => {
    appState.firstRunSetupWindow = window as BrowserWindow;
  },
  encodeURIComponent: (value) => encodeURIComponent(value),
  logError: (message, error) => logger.error(message, error),
});

function openFirstRunSetupWindow(force = false): void {
  if (!force && firstRunSetupService.isSetupCompleted()) {
    return;
  }
  openFirstRunSetupWindowHandler();
}

const {
  notifyAnilistSetup,
  consumeAnilistSetupTokenFromUrl,
  handleAnilistSetupProtocolUrl,
  registerSubminerProtocolClient,
} = composeAnilistSetupHandlers({
  notifyDeps: {
    hasMpvClient: () => Boolean(appState.mpvClient),
    showMpvOsd: (message) => showMpvOsd(message),
    showDesktopNotification: (title, options) => showDesktopNotification(title, options),
    logInfo: (message) => logger.info(message),
  },
  consumeTokenDeps: {
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
  },
  handleProtocolDeps: {
    consumeAnilistSetupTokenFromUrl: (rawUrl) => consumeAnilistSetupTokenFromUrl(rawUrl),
    logWarn: (message, details) => logger.warn(message, details),
  },
  registerProtocolClientDeps: {
    isDefaultApp: () => Boolean(process.defaultApp),
    getArgv: () => process.argv,
    execPath: process.execPath,
    resolvePath: (value) => path.resolve(value),
    setAsDefaultProtocolClient: (scheme, appPath, args) =>
      appPath
        ? app.setAsDefaultProtocolClient(scheme, appPath, args)
        : app.setAsDefaultProtocolClient(scheme),
    logDebug: (message, details) => logger.debug(message, details),
  },
});

const maybeFocusExistingAnilistSetupWindow = createMaybeFocusExistingAnilistSetupWindowHandler({
  getSetupWindow: () => appState.anilistSetupWindow,
});
const buildOpenAnilistSetupWindowMainDepsHandler = createBuildOpenAnilistSetupWindowMainDepsHandler(
  {
    maybeFocusExistingSetupWindow: maybeFocusExistingAnilistSetupWindow,
    createSetupWindow: createCreateAnilistSetupWindowHandler({
      createBrowserWindow: (options) => new BrowserWindow(options),
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
  },
);

function openAnilistSetupWindow(): void {
  createOpenAnilistSetupWindowHandler(buildOpenAnilistSetupWindowMainDepsHandler())();
}

const {
  refreshAnilistClientSecretState,
  getCurrentAnilistMediaKey,
  resetAnilistMediaTracking,
  getAnilistMediaGuessRuntimeState,
  setAnilistMediaGuessRuntimeState,
  recordAnilistMediaDuration,
  resetAnilistMediaGuessState,
  maybeProbeAnilistDuration,
  ensureAnilistMediaGuess,
  processNextAnilistRetryUpdate,
  maybeRunAnilistPostWatchUpdate,
} = composeAnilistTrackingHandlers({
  refreshClientSecretMainDeps: {
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
  },
  getCurrentMediaKeyMainDeps: {
    getCurrentMediaPath: () => appState.currentMediaPath,
  },
  resetMediaTrackingMainDeps: {
    setMediaKey: (value) => {
      anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
        anilistMediaGuessRuntimeState,
        { mediaKey: value },
      );
    },
    setMediaDurationSec: (value) => {
      anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
        anilistMediaGuessRuntimeState,
        { mediaDurationSec: value },
      );
    },
    setMediaGuess: (value) => {
      anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
        anilistMediaGuessRuntimeState,
        { mediaGuess: value },
      );
    },
    setMediaGuessPromise: (value) => {
      anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
        anilistMediaGuessRuntimeState,
        { mediaGuessPromise: value },
      );
    },
    setLastDurationProbeAtMs: (value) => {
      anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
        anilistMediaGuessRuntimeState,
        { lastDurationProbeAtMs: value },
      );
    },
  },
  getMediaGuessRuntimeStateMainDeps: {
    getMediaKey: () => anilistMediaGuessRuntimeState.mediaKey,
    getMediaDurationSec: () => anilistMediaGuessRuntimeState.mediaDurationSec,
    getMediaGuess: () => anilistMediaGuessRuntimeState.mediaGuess,
    getMediaGuessPromise: () => anilistMediaGuessRuntimeState.mediaGuessPromise,
    getLastDurationProbeAtMs: () => anilistMediaGuessRuntimeState.lastDurationProbeAtMs,
  },
  setMediaGuessRuntimeStateMainDeps: {
    setMediaKey: (value) => {
      anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
        anilistMediaGuessRuntimeState,
        { mediaKey: value },
      );
    },
    setMediaDurationSec: (value) => {
      anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
        anilistMediaGuessRuntimeState,
        { mediaDurationSec: value },
      );
    },
    setMediaGuess: (value) => {
      anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
        anilistMediaGuessRuntimeState,
        { mediaGuess: value },
      );
    },
    setMediaGuessPromise: (value) => {
      anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
        anilistMediaGuessRuntimeState,
        { mediaGuessPromise: value },
      );
    },
    setLastDurationProbeAtMs: (value) => {
      anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
        anilistMediaGuessRuntimeState,
        { lastDurationProbeAtMs: value },
      );
    },
  },
  recordMediaDurationMainDeps: {
    getCurrentMediaKey: () => getCurrentAnilistMediaKey(),
    getState: () => getAnilistMediaGuessRuntimeState(),
    setState: (state) => {
      setAnilistMediaGuessRuntimeState(state);
    },
  },
  resetMediaGuessStateMainDeps: {
    setMediaGuess: (value) => {
      anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
        anilistMediaGuessRuntimeState,
        { mediaGuess: value },
      );
    },
    setMediaGuessPromise: (value) => {
      anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
        anilistMediaGuessRuntimeState,
        { mediaGuessPromise: value },
      );
    },
  },
  maybeProbeDurationMainDeps: {
    getState: () => getAnilistMediaGuessRuntimeState(),
    setState: (state) => {
      setAnilistMediaGuessRuntimeState(state);
    },
    durationRetryIntervalMs: ANILIST_DURATION_RETRY_INTERVAL_MS,
    now: () => Date.now(),
    requestMpvDuration: async () => appState.mpvClient?.requestProperty('duration'),
    logWarn: (message, error) => logger.warn(message, error),
  },
  ensureMediaGuessMainDeps: {
    getState: () => getAnilistMediaGuessRuntimeState(),
    setState: (state) => {
      setAnilistMediaGuessRuntimeState(state);
    },
    resolveMediaPathForJimaku: (currentMediaPath) =>
      mediaRuntime.resolveMediaPathForJimaku(currentMediaPath),
    getCurrentMediaPath: () => appState.currentMediaPath,
    getCurrentMediaTitle: () => appState.currentMediaTitle,
    guessAnilistMediaInfo: (mediaPath, mediaTitle) => guessAnilistMediaInfo(mediaPath, mediaTitle),
  },
  processNextRetryUpdateMainDeps: {
    nextReady: () => anilistUpdateQueue.nextReady(),
    refreshRetryQueueState: () => anilistStateRuntime.refreshRetryQueueState(),
    setLastAttemptAt: (value) => {
      appState.anilistRetryQueueState = transitionAnilistRetryQueueLastAttemptAt(
        appState.anilistRetryQueueState,
        value,
      );
    },
    setLastError: (value) => {
      appState.anilistRetryQueueState = transitionAnilistRetryQueueLastError(
        appState.anilistRetryQueueState,
        value,
      );
    },
    refreshAnilistClientSecretState: () => refreshAnilistClientSecretState(),
    updateAnilistPostWatchProgress: (accessToken, title, episode, season) =>
      updateAnilistPostWatchProgress(accessToken, title, episode, {
        rateLimiter: anilistRateLimiter,
        season,
      }),
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
  },
  maybeRunPostWatchUpdateMainDeps: {
    getInFlight: () => anilistUpdateInFlightState.inFlight,
    setInFlight: (value) => {
      anilistUpdateInFlightState = transitionAnilistUpdateInFlightState(
        anilistUpdateInFlightState,
        value,
      );
    },
    getResolvedConfig: () => getResolvedConfig(),
    isAnilistTrackingEnabled: (config) => isAnilistTrackingEnabled(config as ResolvedConfig),
    getCurrentMediaKey: () => getCurrentAnilistMediaKey(),
    hasMpvClient: () => Boolean(appState.mpvClient),
    getTrackedMediaKey: () => anilistMediaGuessRuntimeState.mediaKey,
    resetTrackedMedia: (mediaKey) => {
      resetAnilistMediaTracking(mediaKey);
    },
    getWatchedSeconds: () => appState.mpvClient?.currentTimePos ?? Number.NaN,
    maybeProbeAnilistDuration: (mediaKey, options) => maybeProbeAnilistDuration(mediaKey, options),
    ensureAnilistMediaGuess: (mediaKey) => ensureAnilistMediaGuess(mediaKey),
    hasAttemptedUpdateKey: (key) => anilistAttemptedUpdateKeys.has(key),
    processNextAnilistRetryUpdate: () => processNextAnilistRetryUpdate(),
    refreshAnilistClientSecretState: () => refreshAnilistClientSecretState(),
    enqueueRetry: (key, title, episode, season) => {
      anilistUpdateQueue.enqueue(key, title, episode, season);
    },
    markRetryFailure: (key, message) => {
      anilistUpdateQueue.markFailure(key, message);
    },
    markRetrySuccess: (key) => {
      anilistUpdateQueue.markSuccess(key);
    },
    refreshRetryQueueState: () => anilistStateRuntime.refreshRetryQueueState(),
    updateAnilistPostWatchProgress: (accessToken, title, episode, season) =>
      updateAnilistPostWatchProgress(accessToken, title, episode, {
        rateLimiter: anilistRateLimiter,
        season,
      }),
    rememberAttemptedUpdateKey: (key) => {
      rememberAnilistAttemptedUpdate(key);
    },
    showMpvOsd: (message) => showMpvOsd(message),
    logInfo: (message) => logger.info(message),
    logWarn: (message) => logger.warn(message),
    minWatchSeconds: ANILIST_UPDATE_MIN_WATCH_SECONDS,
    minWatchRatio: DEFAULT_MIN_WATCH_RATIO,
  },
});

function refreshAnilistClientSecretStateIfEnabled(options?: {
  force?: boolean;
  allowSetupPrompt?: boolean;
}): Promise<string | null> {
  if (!isAnilistTrackingEnabled(getResolvedConfig())) {
    return Promise.resolve(null);
  }
  return refreshAnilistClientSecretState(options);
}

const rememberAnilistAttemptedUpdate = (key: string): void => {
  rememberAnilistAttemptedUpdateKey(
    anilistAttemptedUpdateKeys,
    key,
    ANILIST_MAX_ATTEMPTED_UPDATE_KEYS,
  );
};

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
const loadSubtitlePosition = createLoadSubtitlePositionHandler(loadSubtitlePositionMainDeps);

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
const saveSubtitlePosition = createSaveSubtitlePositionHandler(saveSubtitlePositionMainDeps);

registerSubminerProtocolClient();
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
    handleAnilistSetupProtocolUrl: (rawUrl) => handleAnilistSetupProtocolUrl(rawUrl),
    findAnilistSetupDeepLinkArgvUrl: (argv) => findAnilistSetupDeepLinkArgvUrl(argv),
    logUnhandledOpenUrl: (rawUrl) => {
      logger.warn('Unhandled app protocol URL', { rawUrl });
    },
    logUnhandledSecondInstanceUrl: (rawUrl) => {
      logger.warn('Unhandled second-instance protocol URL', { rawUrl });
    },
  },
  onWillQuitCleanupMainDeps: {
    destroyTray: () => destroyTray(),
    stopConfigHotReload: () => configHotReloadRuntime.stop(),
    restorePreviousSecondarySubVisibility: () => restorePreviousSecondarySubVisibility(),
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
    clearWindowsVisibleOverlayForegroundPollLoop: () =>
      clearWindowsVisibleOverlayForegroundPollLoop(),
    clearLinuxMpvFullscreenOverlayRefreshTimeouts: () => {
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
    clearImmersionTracker: () => {
      stopStatsServer();
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
    stopJellyfinRemoteSession: () => stopJellyfinRemoteSession(),
    cleanupYoutubeSubtitleTempDirs: () => youtubeFlowRuntime.cleanupSubtitleTempDirs(),
    cleanupJellyfinSubtitleCache: () => cleanupJellyfinSubtitleCache(),
    stopDiscordPresenceService: () => {
      void appState.discordPresenceService?.stop();
      appState.discordPresenceService = null;
    },
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

const statsDistPath = path.join(__dirname, '..', 'stats', 'dist');
const statsPreloadPath = path.join(__dirname, 'preload-stats.js');

const startLocalStatsServer = (): void => {
  const tracker = appState.immersionTracker;
  if (!tracker) {
    throw new Error('Immersion tracker failed to initialize.');
  }
  if (!statsServer) {
    const yomitanDeps = {
      getYomitanExt: () => appState.yomitanExt,
      getYomitanSession: () => appState.yomitanSession,
      getYomitanParserWindow: () => appState.yomitanParserWindow,
      setYomitanParserWindow: (w: BrowserWindow | null) => {
        appState.yomitanParserWindow = w;
      },
      getYomitanParserReadyPromise: () => appState.yomitanParserReadyPromise,
      setYomitanParserReadyPromise: (p: Promise<void> | null) => {
        appState.yomitanParserReadyPromise = p;
      },
      getYomitanParserInitPromise: () => appState.yomitanParserInitPromise,
      setYomitanParserInitPromise: (p: Promise<boolean> | null) => {
        appState.yomitanParserInitPromise = p;
      },
    };
    const yomitanLogger = createLogger('main:yomitan-stats');
    statsServer = startStatsServer({
      port: getResolvedConfig().stats.serverPort,
      staticDir: statsDistPath,
      tracker,
      knownWordCachePath: path.join(USER_DATA_PATH, 'known-words-cache.json'),
      mpvSocketPath: appState.mpvSocketPath,
      ankiConnectConfig: getResolvedConfig().ankiConnect,
      anilistRateLimiter,
      resolveAnkiNoteId: (noteId: number) =>
        appState.ankiIntegration?.resolveCurrentNoteId(noteId) ?? noteId,
      addYomitanNote: async (word: string) => {
        const ankiUrl = getResolvedConfig().ankiConnect.url || 'http://127.0.0.1:8765';
        await syncYomitanDefaultAnkiServerCore(ankiUrl, yomitanDeps, yomitanLogger, {
          forceOverride: true,
        });
        const result = await addYomitanNoteViaSearch(word, yomitanDeps, yomitanLogger);
        if (result.noteId && result.duplicateNoteIds.length > 0) {
          appState.ankiIntegration?.trackDuplicateNoteIdsForNote(
            result.noteId,
            result.duplicateNoteIds,
          );
        }
        return result.noteId;
      },
    });
    appState.statsServer = statsServer;
  }
  appState.statsServer = statsServer;
};

const ensureStatsServerStarted = createEnsureStatsServerUrlHandler({
  currentPid: process.pid,
  readBackgroundState: () => readBackgroundStatsServerState(statsDaemonStatePath),
  removeBackgroundState: () => {
    removeBackgroundStatsServerState(statsDaemonStatePath);
  },
  isProcessAlive: (pid) => isBackgroundStatsServerProcessAlive(pid),
  hasLocalStatsServer: () => statsServer !== null,
  startLocalStatsServer,
  getConfiguredPort: () => getResolvedConfig().stats.serverPort,
});

const ensureBackgroundStatsServerStarted = (): {
  url: string;
  runningInCurrentProcess: boolean;
} => {
  const liveDaemon = readLiveBackgroundStatsDaemonState();
  if (liveDaemon && liveDaemon.pid !== process.pid) {
    return {
      url: resolveBackgroundStatsServerUrl(liveDaemon),
      runningInCurrentProcess: false,
    };
  }

  appState.statsStartupInProgress = true;
  try {
    ensureImmersionTrackerStarted();
  } finally {
    appState.statsStartupInProgress = false;
  }

  const port = getResolvedConfig().stats.serverPort;
  const result = ensureStatsServerStarted();
  if (result.source === 'local') {
    writeBackgroundStatsServerState(statsDaemonStatePath, {
      pid: process.pid,
      port,
      startedAtMs: Date.now(),
    });
  }
  return { url: result.url, runningInCurrentProcess: result.source === 'local' };
};

const stopBackgroundStatsServer = async (): Promise<{ ok: boolean; stale: boolean }> => {
  const state = readBackgroundStatsServerState(statsDaemonStatePath);
  if (!state) {
    removeBackgroundStatsServerState(statsDaemonStatePath);
    return { ok: true, stale: true };
  }
  if (!isBackgroundStatsServerProcessAlive(state.pid)) {
    removeBackgroundStatsServerState(statsDaemonStatePath);
    return { ok: true, stale: true };
  }

  try {
    process.kill(state.pid, 'SIGTERM');
  } catch (error) {
    if ((error as NodeJS.ErrnoException)?.code === 'ESRCH') {
      removeBackgroundStatsServerState(statsDaemonStatePath);
      return { ok: true, stale: true };
    }
    if ((error as NodeJS.ErrnoException)?.code === 'EPERM') {
      throw new Error(
        `Insufficient permissions to stop background stats server (pid ${state.pid}).`,
      );
    }
    throw error;
  }

  const deadline = Date.now() + 2_000;
  while (Date.now() < deadline) {
    if (!isBackgroundStatsServerProcessAlive(state.pid)) {
      removeBackgroundStatsServerState(statsDaemonStatePath);
      return { ok: true, stale: false };
    }
    await new Promise((resolve) => setTimeout(resolve, 50));
  }

  throw new Error('Timed out stopping background stats server.');
};

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

const immersionTrackerStartupMainDeps: Parameters<
  typeof createBuildImmersionTrackerStartupMainDepsHandler
>[0] = {
  getResolvedConfig: () => getResolvedConfig(),
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
      stopStatsServer();
      appState.statsServer = null;
    }

    appState.immersionTracker = tracker as ImmersionTrackerService | null;
    appState.immersionTracker?.setCoverArtFetcher(statsCoverArtFetcher);
    if (tracker) {
      // Start HTTP stats server
      if (!appState.statsServer) {
        const config = getResolvedConfig();
        if (config.stats.autoStartServer) {
          ensureStatsServerStarted();
        }
      }

      // Register stats overlay toggle IPC handler (idempotent)
      registerStatsOverlayToggle({
        staticDir: statsDistPath,
        preloadPath: statsPreloadPath,
        getApiBaseUrl: () => ensureStatsServerStarted().url,
        getToggleKey: () => getResolvedConfig().stats.toggleKey,
        resolveBounds: () => getCurrentOverlayGeometry(),
        onVisibilityChanged: (visible) => {
          handleStatsOverlayVisibilityChanged(visible);
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
  createBuildImmersionTrackerStartupMainDepsHandler(immersionTrackerStartupMainDeps)(),
);
const recordTrackedCardsMined = (count: number, noteIds?: number[]): void => {
  ensureImmersionTrackerStarted();
  appState.immersionTracker?.recordCardsMined(count, noteIds);
};
const refreshCurrentSubtitleAfterKnownWordUpdate = (): void => {
  subtitleProcessingController.refreshCurrentSubtitle(appState.currentSubText);
};
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

const runStatsCliCommand = createRunStatsCliCommandHandler({
  getResolvedConfig: () => getResolvedConfig(),
  ensureImmersionTrackerStarted: () => statsStartupRuntime.ensureImmersionTrackerStarted(),
  ensureVocabularyCleanupTokenizerReady: async () => {
    await createMecabTokenizerAndCheck();
  },
  getImmersionTracker: () => appState.immersionTracker,
  ensureStatsServerStarted: () => statsStartupRuntime.ensureStatsServerStarted().url,
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

  const resolvedConfig = getResolvedConfig();
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
    startConfigHotReload: () => configHotReloadRuntime.start(),
    shouldRefreshAnilistClientSecretState: () =>
      shouldRefreshAnilistOnConfigReload(appState.initialArgs),
    refreshAnilistClientSecretState: (options) => refreshAnilistClientSecretStateIfEnabled(options),
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
      appState.keybindings = resolveKeybindings(getResolvedConfig(), DEFAULT_KEYBINDINGS);
      refreshCurrentSessionBindings();
    },
    createMpvClient: () => {
      appState.mpvClient = createMpvClientRuntimeService();
    },
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
          getSubtitleStyleConfig: () => configService.getConfig().subtitleStyle,
          onOptionsChanged: () => {
            subtitleProcessingController.invalidateTokenizationCache();
            subtitlePrefetchService?.onSeek(lastObservedTimePos);
            broadcastRuntimeOptionsChanged();
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
        () => ({
          enabled: getResolvedConfig().subtitleStyle.frequencyDictionary.enabled,
          topX: getResolvedConfig().subtitleStyle.frequencyDictionary.topX,
          mode: getResolvedConfig().subtitleStyle.frequencyDictionary.mode,
        }),
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
        () => ({
          enabled: getResolvedConfig().subtitleStyle.frequencyDictionary.enabled,
          topX: getResolvedConfig().subtitleStyle.frequencyDictionary.topX,
          mode: getResolvedConfig().subtitleStyle.frequencyDictionary.mode,
        }),
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
      const snapshot = await firstRunSetupService.ensureSetupStateInitialized();
      appState.firstRunSetupCompleted = snapshot.state.status === 'completed';
      const args = appState.initialArgs;
      if (args && shouldAutoOpenFirstRunSetup(args)) {
        const force = Boolean(args.setup);
        if (force || snapshot.state.status !== 'completed') {
          openFirstRunSetupWindow(force);
        }
      }
    },
    startJellyfinRemoteSession: async () => {
      await startJellyfinRemoteSession();
    },
    prewarmSubtitleDictionaries: async () => {
      await prewarmSubtitleDictionaries();
    },
    startBackgroundWarmups: () => {
      startBackgroundWarmupsIfAllowed();
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
    appState.keybindings = resolveKeybindings(getResolvedConfig(), DEFAULT_KEYBINDINGS);
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
          broadcastRuntimeOptionsChanged();
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
  tokenizeSubtitle,
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
      jellyfinPlayQuitOnDisconnectArmed || youtubePlaybackRuntime.getQuitOnDisconnectArmed(),
    scheduleQuitCheck: (callback) => {
      setTimeout(callback, 500);
    },
    quitApp: () => requestAppQuit(),
    reportJellyfinRemoteStopped: () => {
      void reportJellyfinRemoteStopped();
    },
    onMpvConnected: () => {
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
    maybeRunAnilistPostWatchUpdate: (options) => maybeRunAnilistPostWatchUpdate(options),
    recordAnilistMediaDuration: (durationSec) => {
      recordAnilistMediaDuration(durationSec);
    },
    logSubtitleTimingError: (message, error) => logger.error(message, error),
    broadcastToOverlayWindows: (channel, payload) => {
      broadcastToOverlayWindows(channel, payload);
    },
    getImmediateSubtitlePayload: (text) => subtitleProcessingController.consumeCachedSubtitle(text),
    emitImmediateSubtitle: (payload) => {
      emitSubtitlePayload(payload);
    },
    onSubtitleChange: (text) => {
      subtitlePrefetchService?.pause();
      subtitleProcessingController.onSubtitleChange(text);
    },
    refreshDiscordPresence: () => {
      discordPresenceRuntime.publishDiscordPresence();
    },
    ensureImmersionTrackerInitialized: () => {
      ensureImmersionTrackerStarted();
    },
    tokenizeSubtitleForImmersion: async (text): Promise<SubtitleData | null> =>
      tokenizeSubtitleDeferred ? await tokenizeSubtitleDeferred(text) : null,
    updateCurrentMediaPath: (path) => {
      const normalizedPath = path.trim();
      const previousPath = appState.currentMediaPath?.trim() || null;
      const preserveParsedSubtitleCues = isSameYoutubeMediaPath(
        normalizedPath,
        appState.activeParsedSubtitleMediaPath,
      );
      if ((normalizedPath || null) !== previousPath) {
        const resetSubtitlePayload = { text: '', tokens: null };
        const frequencyDictionary = getResolvedConfig().subtitleStyle.frequencyDictionary;
        const frequencyOptions = {
          enabled: frequencyDictionary.enabled,
          topX: frequencyDictionary.topX,
          mode: frequencyDictionary.mode,
        };
        autoplaySubtitlePrimedMediaPath = null;
        lastObservedTimePos = 0;
        appState.currentSubText = '';
        appState.currentSubAssText = '';
        appState.currentSubtitleData = null;
        if (!preserveParsedSubtitleCues) {
          appState.activeParsedSubtitleCues = [];
          appState.activeParsedSubtitleSource = null;
          appState.activeParsedSubtitleMediaPath = null;
        }
        activeJellyfinSubtitleDelayKey = null;
        broadcastToOverlayWindows('subtitle:set', resetSubtitlePayload);
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
      youtubePrimarySubtitleNotificationRuntime.handleMediaPathChange(path);
      if (path) {
        ensureImmersionTrackerStarted();
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
    signalAutoplayReadyIfWarm: (path) => signalAutoplayReadyFromWarmTokenization?.(path),
    markJellyfinRemotePlaybackLoaded: (path) => {
      markJellyfinRemotePlaybackLoadedState(activeJellyfinRemotePlayback, path);
    },
    scheduleCharacterDictionarySync: () => {
      if (!yomitanProfilePolicy.isCharacterDictionaryEnabled() || isYoutubePlaybackActiveNow()) {
        return;
      }
      characterDictionaryAutoSyncRuntime.scheduleSync();
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
    onTimePosUpdate: (time) => {
      const delta = time - lastObservedTimePos;
      if (subtitlePrefetchService && (delta > SEEK_THRESHOLD_SECONDS || delta < 0)) {
        subtitlePrefetchService.onSeek(time);
      }
      lastObservedTimePos = time;
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
          ensureOverlayWindowLevel: (window) => ensureOverlayWindowLevel(window),
        },
        cancelLinuxMpvFullscreenOverlayRefreshBurst,
      );
    },
    onSubtitleTrackChange: (sid) => {
      scheduleSubtitlePrefetchRefresh();
      youtubePrimarySubtitleNotificationRuntime.handleSubtitleTrackChange(sid);
    },
    onSubtitleTrackListChange: (trackList) => {
      managedLocalSubtitleSelectionRuntime.handleSubtitleTrackListChange(trackList);
      scheduleSubtitlePrefetchRefresh();
      youtubePrimarySubtitleNotificationRuntime.handleSubtitleTrackListChange(trackList);
    },
    updateSubtitleRenderMetrics: (patch) => {
      updateMpvSubtitleRenderMetrics(patch as Partial<MpvSubtitleRenderMetrics>);
    },
    syncOverlayMpvSubtitleSuppression: () => {
      syncOverlayMpvSubtitleSuppression();
    },
  },
  mpvClientRuntimeServiceFactoryMainDeps: {
    createClient: MpvIpcClient,
    getSocketPath: () => appState.mpvSocketPath,
    getResolvedConfig: () => getResolvedConfig(),
    isAutoStartOverlayEnabled: () => appState.autoStartOverlay,
    setOverlayVisible: (visible: boolean) => setOverlayVisible(visible),
    isVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    getReconnectTimer: () => appState.reconnectTimer,
    setReconnectTimer: (timer: ReturnType<typeof setTimeout> | null) => {
      appState.reconnectTimer = timer;
    },
    shouldAutoLoadSecondarySubTrack: (path: string) =>
      shouldAutoLoadSecondarySubTrackForJellyfinPlayback(activeJellyfinRemotePlayback, path),
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
      isKnownWord: (text) => Boolean(appState.ankiIntegration?.isKnownWord(text)),
      recordLookup: (hit) => {
        ensureImmersionTrackerStarted();
        appState.immersionTracker?.recordLookup(hit);
      },
      getKnownWordMatchMode: () =>
        appState.ankiIntegration?.getKnownWordMatchMode() ??
        getResolvedConfig().ankiConnect.knownWords.matchMode,
      getKnownWordsEnabled: () =>
        getRuntimeBooleanOption(
          'subtitle.annotation.knownWords.highlightEnabled',
          getResolvedConfig().ankiConnect.knownWords.highlightEnabled,
        ),
      getNPlusOneEnabled: () =>
        getRuntimeBooleanOption(
          'subtitle.annotation.nPlusOne',
          getResolvedConfig().ankiConnect.nPlusOne.enabled,
        ),
      getMinSentenceWordsForNPlusOne: () =>
        getResolvedConfig().ankiConnect.nPlusOne.minSentenceWords,
      getJlptLevel: (text) => appState.jlptLevelLookup(text),
      getJlptEnabled: () =>
        getRuntimeBooleanOption(
          'subtitle.annotation.jlpt',
          getResolvedConfig().subtitleStyle.enableJlpt,
        ),
      getCharacterDictionaryEnabled: () =>
        getResolvedConfig().anilist.characterDictionary.enabled &&
        yomitanProfilePolicy.isCharacterDictionaryEnabled() &&
        !isYoutubePlaybackActiveNow(),
      getNameMatchEnabled: () => getResolvedConfig().subtitleStyle.nameMatchEnabled,
      getFrequencyDictionaryEnabled: () =>
        getRuntimeBooleanOption(
          'subtitle.annotation.frequency',
          getResolvedConfig().subtitleStyle.frequencyDictionary.enabled,
        ),
      getFrequencyDictionaryMatchMode: () =>
        getResolvedConfig().subtitleStyle.frequencyDictionary.matchMode,
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
    createTokenizerRuntimeDeps: (deps) =>
      createTokenizerDepsRuntime(deps as Parameters<typeof createTokenizerDepsRuntime>[0]),
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
      ensureJlptDictionaryLookup: () => jlptDictionaryRuntime.ensureJlptDictionaryLookup(),
      ensureFrequencyDictionaryLookup: () =>
        frequencyDictionaryRuntime.ensureFrequencyDictionaryLookup(),
      showMpvOsd: (message: string) => showMpvOsd(message),
      showLoadingOsd: (message: string) => startupOsdSequencer.showAnnotationLoading(message),
      showLoadedOsd: (message: string) =>
        startupOsdSequencer.markAnnotationLoadingComplete(message),
      shouldShowOsdNotification: () => {
        const type = getResolvedConfig().ankiConnect.behavior.notificationType;
        return type === 'osd' || type === 'both';
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
        const startupWarmups = getResolvedConfig().startupWarmups;
        if (startupWarmups.lowPowerMode) {
          return false;
        }
        if (!startupWarmups.mecab) {
          return false;
        }
        return shouldInitializeMecabForAnnotations();
      },
      shouldWarmupYomitanExtension: () => getResolvedConfig().startupWarmups.yomitanExtension,
      shouldWarmupSubtitleDictionaries: () => {
        const startupWarmups = getResolvedConfig().startupWarmups;
        if (startupWarmups.lowPowerMode) {
          return false;
        }
        return startupWarmups.subtitleDictionaries;
      },
      shouldWarmupJellyfinRemoteSession: () => {
        const startupWarmups = getResolvedConfig().startupWarmups;
        if (startupWarmups.lowPowerMode) {
          return false;
        }
        return startupWarmups.jellyfinRemoteSession;
      },
      shouldAutoConnectJellyfinRemote: () => {
        const jellyfin = getResolvedConfig().jellyfin;
        return (
          jellyfin.enabled && jellyfin.remoteControlEnabled && jellyfin.remoteControlAutoConnect
        );
      },
      startJellyfinRemoteSession: () => startJellyfinRemoteSession(),
      logDebug: (message) => logger.debug(message),
    },
  },
});
signalAutoplayReadyFromWarmTokenization = createAutoplayTokenizationWarmRelease({
  isTokenizationWarmupReady: () => isTokenizationWarmupReady(),
  startTokenizationWarmups: async () => {
    await startTokenizationWarmups();
  },
  getCurrentMediaPath: () =>
    appState.currentMediaPath?.trim() || appState.mpvClient?.currentVideoPath?.trim() || null,
  primeCurrentSubtitle: (mediaPath) => primeCurrentSubtitleForAutoplay(mediaPath),
  signalAutoplayReady: () => {
    autoplayReadyGate.maybeSignalPluginAutoplayReady(
      { text: '__warm__', tokens: null },
      { forceWhilePaused: true },
    );
  },
  warn: (message, error) => logger.warn(message, error),
});
tokenizeSubtitleDeferred = tokenizeSubtitle;

function createMpvClientRuntimeService(): MpvIpcClient {
  const client = createMpvClientRuntimeServiceHandler() as MpvIpcClient;
  client.on('connection-change', ({ connected }) => {
    if (connected) {
      return;
    }
    if (!youtubeFlowRuntime.hasActiveSession()) {
      return;
    }
    youtubeFlowRuntime.cancelActivePicker();
    broadcastToOverlayWindows(IPC_CHANNELS.event.youtubePickerCancel, null);
    overlayModalRuntime.handleOverlayModalClosed('youtube-track-picker');
  });
  return client;
}

function resetSubtitleSidebarEmbeddedLayoutRuntime(): void {
  sendMpvCommandRuntime(appState.mpvClient, ['set_property', 'video-margin-ratio-right', 0]);
  sendMpvCommandRuntime(appState.mpvClient, ['set_property', 'video-pan-x', 0]);
}

function updateMpvSubtitleRenderMetrics(patch: Partial<MpvSubtitleRenderMetrics>): void {
  updateMpvSubtitleRenderMetricsHandler(patch);
}

let lastOverlayWindowGeometry: WindowGeometry | null = null;

function getOverlayGeometryFallback(): WindowGeometry {
  const cursorPoint = screen.getCursorScreenPoint();
  const display = screen.getDisplayNearestPoint(cursorPoint);
  const bounds = display.workArea;
  return {
    x: bounds.x,
    y: bounds.y,
    width: bounds.width,
    height: bounds.height,
  };
}

function getCurrentOverlayGeometry(): WindowGeometry {
  if (lastOverlayWindowGeometry) return lastOverlayWindowGeometry;
  const trackerGeometry = appState.windowTracker?.getGeometry();
  if (trackerGeometry) return trackerGeometry;
  return getOverlayGeometryFallback();
}

function geometryMatches(a: WindowGeometry | null, b: WindowGeometry | null): boolean {
  if (!a || !b) return false;
  return a.x === b.x && a.y === b.y && a.width === b.width && a.height === b.height;
}

function applyOverlayRegions(geometry: WindowGeometry): void {
  lastOverlayWindowGeometry = geometry;
  overlayManager.setOverlayWindowBounds(geometry);
  overlayManager.setModalWindowBounds(geometry);
}

const buildUpdateVisibleOverlayBoundsMainDepsHandler =
  createBuildUpdateVisibleOverlayBoundsMainDepsHandler({
    setOverlayWindowBounds: (geometry) => applyOverlayRegions(geometry),
    afterSetOverlayWindowBounds: () => {
      if (!overlayManager.getVisibleOverlayVisible()) {
        return;
      }
      if (process.platform === 'win32') {
        scheduleWindowsVisibleOverlayZOrderSyncBurst();
        return;
      }
      const mainWindow = overlayManager.getMainWindow();
      if (!mainWindow || mainWindow.isDestroyed()) {
        return;
      }
      ensureOverlayWindowLevel(mainWindow);
    },
  });
const updateVisibleOverlayBoundsMainDeps = buildUpdateVisibleOverlayBoundsMainDepsHandler();
const updateVisibleOverlayBounds = createUpdateVisibleOverlayBoundsHandler(
  updateVisibleOverlayBoundsMainDeps,
);

const buildEnsureOverlayWindowLevelMainDepsHandler =
  createBuildEnsureOverlayWindowLevelMainDepsHandler({
    shouldSuppressOverlayWindowLevel: (window) =>
      appState.statsOverlayVisible && window === overlayManager.getMainWindow(),
    ensureOverlayWindowLevelCore: (window) => ensureOverlayWindowLevelCore(window as BrowserWindow),
    afterEnsureOverlayWindowLevel: () => {
      promoteStatsOverlayAbovePlayback();
    },
  });
const ensureOverlayWindowLevelMainDeps = buildEnsureOverlayWindowLevelMainDepsHandler();
const ensureOverlayWindowLevel = createEnsureOverlayWindowLevelHandler(
  ensureOverlayWindowLevelMainDeps,
);

function syncPrimaryOverlayWindowLayer(layer: 'visible'): void {
  const mainWindow = overlayManager.getMainWindow();
  if (!mainWindow || mainWindow.isDestroyed()) return;
  syncOverlayWindowLayer(mainWindow, layer);
}

const buildEnforceOverlayLayerOrderMainDepsHandler =
  createBuildEnforceOverlayLayerOrderMainDepsHandler({
    enforceOverlayLayerOrderCore: (params) =>
      enforceOverlayLayerOrderCore({
        visibleOverlayVisible: params.visibleOverlayVisible,
        mainWindow: params.mainWindow as BrowserWindow | null,
        ensureOverlayWindowLevel: (window) =>
          params.ensureOverlayWindowLevel(window as BrowserWindow),
      }),
    getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    getMainWindow: () => overlayManager.getMainWindow(),
    ensureOverlayWindowLevel: (window) => ensureOverlayWindowLevel(window as BrowserWindow),
  });
const enforceOverlayLayerOrderMainDeps = buildEnforceOverlayLayerOrderMainDepsHandler();
const enforceOverlayLayerOrder = createEnforceOverlayLayerOrderHandler(
  enforceOverlayLayerOrderMainDeps,
);

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

let lastSyncedYomitanAnkiServer: string | null = null;

function getPreferredYomitanAnkiServerUrl(): string {
  return getPreferredYomitanAnkiServerUrlRuntime(getResolvedConfig().ankiConnect);
}

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

async function syncYomitanDefaultProfileAnkiServer(): Promise<void> {
  if (yomitanProfilePolicy.isExternalReadOnlyMode()) {
    return;
  }

  const targetUrl = getPreferredYomitanAnkiServerUrl().trim();
  if (!targetUrl || targetUrl === lastSyncedYomitanAnkiServer) {
    return;
  }

  const synced = await syncYomitanDefaultAnkiServerCore(
    targetUrl,
    getYomitanParserRuntimeDeps(),
    {
      error: (message, ...args) => {
        logger.error(message, ...args);
      },
      info: (message, ...args) => {
        logger.info(message, ...args);
      },
    },
    {
      forceOverride: shouldForceOverrideYomitanAnkiServer(getResolvedConfig().ankiConnect),
    },
  );

  if (synced) {
    lastSyncedYomitanAnkiServer = targetUrl;
  }
}

function createModalWindow(): BrowserWindow {
  const existingWindow = overlayManager.getModalWindow();
  if (existingWindow && !existingWindow.isDestroyed()) {
    return existingWindow;
  }
  const window = createModalWindowHandler();
  overlayManager.setModalWindowBounds(getCurrentOverlayGeometry());
  return window;
}

function createMainWindow(): BrowserWindow {
  const window = createMainWindowHandler();
  if (process.platform === 'win32') {
    const overlayHwnd = getWindowsNativeWindowHandleNumber(window);
    if (!ensureWindowsOverlayTransparency(overlayHwnd)) {
      logger.warn('Failed to eagerly extend Windows overlay transparency via koffi');
    }
  }
  return window;
}

function ensureTray(): void {
  ensureTrayHandler();
}

function destroyTray(): void {
  destroyTrayHandler();
}

function initializeOverlayRuntime(): void {
  initializeOverlayRuntimeHandler();
  appState.ankiIntegration?.setRecordCardsMinedCallback(recordTrackedCardsMined);
  appState.ankiIntegration?.setKnownWordCacheUpdatedCallback(
    refreshCurrentSubtitleAfterKnownWordUpdate,
  );
  syncOverlayMpvSubtitleSuppression();
}

function openYomitanSettings(): boolean {
  if (yomitanProfilePolicy.isExternalReadOnlyMode()) {
    const message = 'Yomitan settings unavailable while using read-only external-profile mode.';
    logger.warn(
      'Yomitan settings window disabled while yomitan.externalProfilePath is configured because external profile mode is read-only.',
    );
    showDesktopNotification('SubMiner', { body: message });
    showMpvOsd(message);
    return false;
  }
  openYomitanSettingsHandler();
  return true;
}

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
      getResolvedConfig: () => getResolvedConfig(),
      defaultConfig: DEFAULT_CONFIG,
      resolveConfiguredShortcuts,
    },
    buildRegisterGlobalShortcutsMainDeps: (getConfiguredShortcutsHandler) => ({
      getConfiguredShortcuts: () => getConfiguredShortcutsHandler(),
      registerGlobalShortcutsCore,
      toggleVisibleOverlay: () => toggleVisibleOverlay(),
      openYomitanSettings: () => openYomitanSettings(),
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
    showMpvOsd: (text) => showMpvOsd(text),
    setTimer: (handler, timeoutMs) => setTimeout(handler, timeoutMs),
    clearTimer: (timer) => clearTimeout(timer),
  },
  numericSessions: {
    onMultiCopyDigit: (count) => handleMultiCopyDigit(count),
    onMineSentenceDigit: (count) => handleMineSentenceDigit(count),
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

function resolveSessionBindingPlatform(): 'darwin' | 'win32' | 'linux' {
  if (process.platform === 'darwin') return 'darwin';
  if (process.platform === 'win32') return 'win32';
  return 'linux';
}

function compileCurrentSessionBindings(): {
  bindings: CompiledSessionBinding[];
  warnings: ReturnType<typeof compileSessionBindings>['warnings'];
} {
  return compileSessionBindings({
    keybindings: appState.keybindings,
    shortcuts: getConfiguredShortcuts(),
    statsToggleKey: getResolvedConfig().stats.toggleKey,
    statsMarkWatchedKey: getResolvedConfig().stats.markWatchedKey,
    platform: resolveSessionBindingPlatform(),
    rawConfig: getResolvedConfig(),
  });
}

function persistSessionBindings(
  bindings: CompiledSessionBinding[],
  warnings: ReturnType<typeof compileSessionBindings>['warnings'] = [],
): void {
  const artifact = buildPluginSessionBindingsArtifact({
    bindings,
    warnings,
    numericSelectionTimeoutMs: getConfiguredShortcuts().multiCopyTimeoutMs,
  });
  writeSessionBindingsArtifact(CONFIG_DIR, artifact);
  appState.sessionBindings = bindings;
  appState.sessionBindingsInitialized = true;
  if (appState.mpvClient?.connected) {
    sendMpvCommandRuntime(appState.mpvClient, [
      'script-message',
      'subminer-reload-session-bindings',
    ]);
  }
}

function refreshCurrentSessionBindings(): void {
  const compiled = compileCurrentSessionBindings();
  for (const warning of compiled.warnings) {
    logger.warn(`[session-bindings] ${warning.message}`);
  }
  persistSessionBindings(compiled.bindings, compiled.warnings);
}

const { flushMpvLog, showMpvOsd } = createMpvOsdRuntimeHandlers({
  appendToMpvLogMainDeps: {
    logPath: DEFAULT_MPV_LOG_PATH,
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

const updateStateStore = createFileUpdateStateStore(path.join(USER_DATA_PATH, 'update-state.json'));
let updateService: ReturnType<typeof createUpdateService> | null = null;
const globalFetchForUpdater = createGlobalFetch();
const curlFetch = createCurlFetch();

function createNativeUpdaterHttpExecutor() {
  if (process.platform === 'win32') {
    return createFetchHttpExecutor();
  }
  return createCurlHttpExecutor();
}

function getFetchForUpdater() {
  if (process.platform === 'win32') return globalFetchForUpdater;
  return curlFetch;
}

async function updateLauncherFromSelectedRelease(
  launcherPath?: string,
  channel: UpdateChannel = getResolvedConfig().updates.channel,
  release: GitHubRelease | null = null,
) {
  const fetchForUpdater = getFetchForUpdater();
  if (!release) {
    return { status: 'missing-asset', message: `No ${channel} GitHub release found.` };
  }
  const sumsAsset = findReleaseAsset(release, 'SHA256SUMS.txt');
  if (!sumsAsset) {
    return { status: 'missing-asset', message: 'Release has no SHA256SUMS.txt asset.' };
  }
  const sums = parseSha256Sums(
    await fetchReleaseAssetText(fetchForUpdater, sumsAsset.browser_download_url),
  );
  const launcherResult = await updateLauncherFromRelease({
    release,
    sha256Sums: sums,
    launcherPath,
    downloadAsset: (url) => fetchReleaseAssetBuffer(fetchForUpdater, url),
  });
  const supportResults = await updateSupportAssetsFromRelease({
    release,
    sha256Sums: sums,
    downloadAsset: (url) => fetchReleaseAssetBuffer(fetchForUpdater, url),
  });
  for (const result of supportResults) {
    if (result.status === 'protected' && result.command) {
      logger.warn(`Rofi theme update requires manual command: ${result.command}`);
    } else if (result.status === 'hash-mismatch' || result.status === 'missing-asset') {
      logger.warn(`Rofi theme update skipped: ${result.message ?? result.status}`);
    }
  }
  return launcherResult;
}

function getUpdateService() {
  if (updateService) return updateService;
  const appUpdater = createElectronAppUpdater({
    currentVersion: app.getVersion(),
    isPackaged: app.isPackaged,
    log: (message) => logger.info(message),
    getChannel: () => getResolvedConfig().updates.channel,
    configureHttpExecutor: createNativeUpdaterHttpExecutor,
    disableDifferentialDownload: true,
    isNativeUpdaterSupported: () =>
      isNativeUpdaterSupported({
        platform: process.platform,
        isPackaged: app.isPackaged,
        execPath: process.execPath,
        env: process.env,
        log: (message) => logger.warn(message),
      }),
  });
  const updateDialogPresenter = createUpdateDialogPresenter({
    platform: process.platform,
    focusApp: async () => {
      if (process.platform !== 'darwin') {
        app.focus({ steal: true });
        return;
      }
      try {
        await app.dock?.show();
      } catch (error) {
        logger.warn('Failed to show macOS dock before update dialog', error);
      }
      // app.focus({ steal: true }) alone does not reliably activate the process
      // when SubMiner was reached via `subminer -u` (single-instance forwarding
      // from a CLI-spawned child). osascript's `activate` uses LaunchServices,
      // which is the only path that reliably brings the running app forward.
      await new Promise<void>((resolve) => {
        execFile(
          '/usr/bin/osascript',
          ['-e', `tell application id "${SUBMINER_BUNDLE_ID}" to activate`],
          { timeout: 2000 },
          (error) => {
            if (error) {
              logger.warn(
                `Failed to activate SubMiner via osascript: ${error instanceof Error ? error.message : String(error)}`,
              );
            }
            resolve();
          },
        );
      });
      app.focus({ steal: true });
    },
    withStatsWindowLayerSuspended: (showDialog) =>
      withStatsWindowLayerSuspendedForNativeDialog(showDialog),
    showMessageBox: (options) => dialog.showMessageBox(options),
  });
  updateService = createUpdateService({
    getConfig: () => getResolvedConfig().updates,
    getCurrentVersion: () => app.getVersion(),
    now: () => Date.now(),
    readState: () => updateStateStore.readState(),
    writeState: (state) => updateStateStore.writeState(state),
    checkAppUpdate: (channel) => appUpdater.checkForUpdates(channel),
    shouldFetchReleaseMetadata: ({ appUpdate }) =>
      shouldFetchReleaseMetadataForPlatform(process.platform, appUpdate),
    fetchLatestStableRelease: (channel) =>
      fetchLatestStableRelease({ fetch: getFetchForUpdater(), channel }),
    updateLauncher: (launcherPath, channel, release) =>
      updateLauncherFromSelectedRelease(launcherPath, channel, release),
    showNoUpdateDialog: (version) => updateDialogPresenter.showNoUpdateDialog(version),
    showUpdateAvailableDialog: (version) =>
      updateDialogPresenter.showUpdateAvailableDialog(version),
    showUpdateFailedDialog: (message) => updateDialogPresenter.showUpdateFailedDialog(message),
    showManualUpdateRequiredDialog: (version) =>
      updateDialogPresenter.showManualUpdateRequiredDialog(version),
    downloadAppUpdate: () => appUpdater.downloadUpdate(),
    showRestartDialog: () => updateDialogPresenter.showRestartDialog(),
    quitAndInstall: () => appUpdater.quitAndInstall(),
    notifyUpdateAvailable: (version) =>
      notifyUpdateAvailable(
        { notificationType: getResolvedConfig().updates.notificationType, version },
        {
          showSystemNotification: (title, body) => showDesktopNotification(title, { body }),
          showOsdNotification: (message) => showMpvOsd(message),
          log: (message) => logger.warn(message),
        },
      ),
    log: (message) => logger.warn(message),
  });
  return updateService;
}

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
    broadcastToOverlayWindows: (channel, mode) => {
      broadcastToOverlayWindows(channel, mode);
    },
    showMpvOsd: (text: string) => showMpvOsd(text),
  },
  cycleSecondarySubMode: (deps) => cycleSecondarySubModeCore(deps),
});

function setSecondarySubMode(mode: SecondarySubMode): void {
  appState.secondarySubMode = mode;
}

function handleCycleSecondarySubMode(): void {
  cycleSecondarySubMode();
}

function toggleSubtitleSidebar(): void {
  broadcastToOverlayWindows(IPC_CHANNELS.event.subtitleSidebarToggle);
}

function togglePrimarySubtitleBar(): void {
  broadcastToOverlayWindows(IPC_CHANNELS.event.primarySubtitleBarToggle);
}

async function triggerSubsyncFromConfig(): Promise<void> {
  await subsyncRuntime.triggerFromConfig();
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
const updateLastCardFromClipboardMainDeps = buildUpdateLastCardFromClipboardMainDepsHandler();
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
const triggerFieldGroupingHandler = createTriggerFieldGroupingHandler(triggerFieldGroupingMainDeps);

const buildMarkLastCardAsAudioCardMainDepsHandler =
  createBuildMarkLastCardAsAudioCardMainDepsHandler({
    getAnkiIntegration: () => appState.ankiIntegration,
    showMpvOsd: (text) => showMpvOsd(text),
    markLastCardAsAudioCardCore,
  });
const markLastCardAsAudioCardMainDeps = buildMarkLastCardAsAudioCardMainDepsHandler();
const markLastCardAsAudioCardHandler = createMarkLastCardAsAudioCardHandler(
  markLastCardAsAudioCardMainDeps,
);

const buildMineSentenceCardMainDepsHandler = createBuildMineSentenceCardMainDepsHandler({
  getAnkiIntegration: () => appState.ankiIntegration,
  getMpvClient: () => appState.mpvClient,
  showMpvOsd: (text) => showMpvOsd(text),
  mineSentenceCardCore,
  recordCardsMined: (count, noteIds) => {
    ensureImmersionTrackerStarted();
    appState.immersionTracker?.recordCardsMined(count, noteIds);
  },
});
const mineSentenceCardHandler = createMineSentenceCardHandler(
  buildMineSentenceCardMainDepsHandler(),
);

const buildHandleMultiCopyDigitMainDepsHandler = createBuildHandleMultiCopyDigitMainDepsHandler({
  getSubtitleTimingTracker: () => appState.subtitleTimingTracker,
  writeClipboardText: (text) => clipboard.writeText(text),
  showMpvOsd: (text) => showMpvOsd(text),
  handleMultiCopyDigitCore,
});
const handleMultiCopyDigitMainDeps = buildHandleMultiCopyDigitMainDepsHandler();
const handleMultiCopyDigitHandler = createHandleMultiCopyDigitHandler(handleMultiCopyDigitMainDeps);

const buildCopyCurrentSubtitleMainDepsHandler = createBuildCopyCurrentSubtitleMainDepsHandler({
  getSubtitleTimingTracker: () => appState.subtitleTimingTracker,
  writeClipboardText: (text) => clipboard.writeText(text),
  showMpvOsd: (text) => showMpvOsd(text),
  copyCurrentSubtitleCore,
});
const copyCurrentSubtitleMainDeps = buildCopyCurrentSubtitleMainDepsHandler();
const copyCurrentSubtitleHandler = createCopyCurrentSubtitleHandler(copyCurrentSubtitleMainDeps);

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
      ensureImmersionTrackerStarted();
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

const buildHandleOverlayModalClosedMainDepsHandler =
  createBuildHandleOverlayModalClosedMainDepsHandler({
    handleOverlayModalClosedRuntime: (modal) => overlayModalRuntime.handleOverlayModalClosed(modal),
  });
const handleOverlayModalClosedMainDeps = buildHandleOverlayModalClosedMainDepsHandler();
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
const appendClipboardVideoToQueueMainDeps = buildAppendClipboardVideoToQueueMainDepsHandler();
const appendClipboardVideoToQueueHandler = createAppendClipboardVideoToQueueHandler(
  appendClipboardVideoToQueueMainDeps,
);

async function loadSubtitleSourceText(source: string): Promise<string> {
  if (/^https?:\/\//i.test(source)) {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 4000);
    try {
      const response = await fetch(source, { signal: controller.signal });
      if (!response.ok) {
        throw new Error(`Failed to download subtitle source (${response.status})`);
      }
      return await response.text();
    } finally {
      clearTimeout(timeoutId);
    }
  }

  const filePath = resolveSubtitleSourcePath(source);
  return fs.promises.readFile(filePath, 'utf8');
}

type MpvSubtitleTrackLike = {
  type?: unknown;
  id?: unknown;
  selected?: unknown;
  external?: unknown;
  codec?: unknown;
  'ff-index'?: unknown;
  'external-filename'?: unknown;
};

function parseTrackId(value: unknown): number | null {
  if (typeof value === 'number' && Number.isInteger(value)) {
    return value;
  }
  if (typeof value === 'string') {
    const parsed = Number(value.trim());
    return Number.isInteger(parsed) ? parsed : null;
  }
  return null;
}

function buildFfmpegSubtitleExtractionArgs(
  videoPath: string,
  ffIndex: number,
  outputPath: string,
): string[] {
  return [
    '-hide_banner',
    '-nostdin',
    '-y',
    '-loglevel',
    'error',
    '-an',
    '-vn',
    '-i',
    videoPath,
    '-map',
    `0:${ffIndex}`,
    '-f',
    path.extname(outputPath).slice(1),
    outputPath,
  ];
}

async function extractInternalSubtitleTrackToTempFile(
  ffmpegPath: string,
  videoPath: string,
  track: MpvSubtitleTrackLike,
): Promise<{ path: string; cleanup: () => Promise<void> } | null> {
  const ffIndex = parseTrackId(track['ff-index']);
  const codec = typeof track.codec === 'string' ? track.codec : null;
  const extension = codecToExtension(codec ?? undefined);
  if (ffIndex === null || extension === null) {
    return null;
  }

  const tempDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'subminer-sidebar-'));
  const outputPath = path.join(tempDir, `track_${ffIndex}.${extension}`);

  try {
    await new Promise<void>((resolve, reject) => {
      const child = spawn(
        ffmpegPath,
        buildFfmpegSubtitleExtractionArgs(videoPath, ffIndex, outputPath),
      );
      let stderr = '';
      child.stderr.on('data', (chunk: Buffer) => {
        stderr += chunk.toString();
      });
      child.on('error', (error) => {
        reject(error);
      });
      child.on('close', (code) => {
        if (code === 0) {
          resolve();
          return;
        }
        reject(new Error(stderr.trim() || `ffmpeg exited with code ${code ?? 'unknown'}`));
      });
    });
  } catch (error) {
    await fs.promises.rm(tempDir, { recursive: true, force: true }).catch(() => undefined);
    throw error;
  }

  return {
    path: outputPath,
    cleanup: async () => {
      await fs.promises.rm(tempDir, { recursive: true, force: true });
    },
  };
}

const shiftSubtitleDelayToAdjacentCueHandler = createShiftSubtitleDelayToAdjacentCueHandler({
  getMpvClient: () => appState.mpvClient,
  loadSubtitleSourceText,
  sendMpvCommand: (command) => sendMpvCommandRuntime(appState.mpvClient, command),
  onSubtitleDelayShifted: (delaySeconds) => {
    const key = activeJellyfinSubtitleDelayKey;
    if (!key) return;
    const saved = saveJellyfinSubtitleDelay({
      filePath: JELLYFIN_SUBTITLE_DELAYS_PATH,
      itemId: key.itemId,
      streamIndex: key.streamIndex,
      delaySeconds,
    });
    if (!saved) {
      logger.warn('Failed to save Jellyfin subtitle delay.');
    }
  },
  showMpvOsd: (text) => showMpvOsd(text),
});

async function dispatchSessionAction(request: SessionActionDispatchRequest): Promise<void> {
  await dispatchSessionActionCore(request, {
    toggleStatsOverlay: () =>
      toggleStatsOverlayWindow({
        staticDir: statsDistPath,
        preloadPath: statsPreloadPath,
        getApiBaseUrl: () => ensureStatsServerStarted().url,
        getToggleKey: () => getResolvedConfig().stats.toggleKey,
        resolveBounds: () => getCurrentOverlayGeometry(),
        onVisibilityChanged: (visible) => {
          handleStatsOverlayVisibilityChanged(visible);
        },
      }),
    toggleVisibleOverlay: () => toggleVisibleOverlay(),
    copyCurrentSubtitle: () => copyCurrentSubtitle(),
    copySubtitleCount: (count) => handleMultiCopyDigit(count),
    updateLastCardFromClipboard: () => updateLastCardFromClipboard(),
    triggerFieldGrouping: () => triggerFieldGrouping(),
    triggerSubsyncFromConfig: () => triggerSubsyncFromConfig(),
    mineSentenceCard: () => mineSentenceCard(),
    mineSentenceCount: (count) => handleMineSentenceDigit(count),
    toggleSecondarySub: () => handleCycleSecondarySubMode(),
    toggleSubtitleSidebar: () => toggleSubtitleSidebar(),
    markLastCardAsAudioCard: () => markLastCardAsAudioCard(),
    markActiveVideoWatched: async () => {
      ensureImmersionTrackerStarted();
      const marked = (await appState.immersionTracker?.markActiveVideoWatched()) ?? false;
      if (marked) {
        try {
          await maybeRunAnilistPostWatchUpdate({ force: true });
        } catch (error) {
          logger.warn('Failed to run AniList post-watch update after manual watched mark:', error);
        }
      }
      return marked;
    },
    openRuntimeOptionsPalette: () => openRuntimeOptionsPalette(),
    openJimaku: () => openJimakuOverlay(),
    openSessionHelp: () => openSessionHelpOverlay(),
    openCharacterDictionary: () => openCharacterDictionaryOverlay(),
    openControllerSelect: () => openControllerSelectOverlay(),
    openControllerDebug: () => openControllerDebugOverlay(),
    openYoutubeTrackPicker: () => openYoutubeTrackPickerFromPlayback(),
    openPlaylistBrowser: () => openPlaylistBrowser(),
    replayCurrentSubtitle: () => replayCurrentSubtitleRuntime(appState.mpvClient),
    playNextSubtitle: () => playNextSubtitleRuntime(appState.mpvClient),
    shiftSubDelayToAdjacentSubtitle: (direction) =>
      shiftSubtitleDelayToAdjacentCueHandler(direction),
    cycleRuntimeOption: (id, direction) => {
      if (!appState.runtimeOptionsManager) {
        return { ok: false, error: 'Runtime options manager unavailable' };
      }
      return applyRuntimeOptionResultRuntime(
        appState.runtimeOptionsManager.cycleOption(id, direction),
        (text) => showMpvOsd(text),
      );
    },
    playNextPlaylistItem: () =>
      sendMpvCommandRuntime(appState.mpvClient, ['playlist-next', 'force']),
    showMpvOsd: (text) => showMpvOsd(text),
  });
}

const { playlistBrowserMainDeps } = createPlaylistBrowserIpcRuntime(() => appState.mpvClient, {
  getPrimarySubtitleLanguages: () => getResolvedConfig().youtube.primarySubLanguages,
  getSecondarySubtitleLanguages: () => getResolvedConfig().secondarySub.secondarySubLanguages,
});

const { registerIpcRuntimeHandlers } = composeIpcRuntimeHandlers({
  mpvCommandMainDeps: {
    triggerSubsyncFromConfig: () => triggerSubsyncFromConfig(),
    openRuntimeOptionsPalette: () => openRuntimeOptionsPalette(),
    openJimaku: () => openJimakuOverlay(),
    openYoutubeTrackPicker: () => openYoutubeTrackPickerFromPlayback(),
    openPlaylistBrowser: () => openPlaylistBrowser(),
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
    shiftSubDelayToAdjacentSubtitle: (direction) =>
      shiftSubtitleDelayToAdjacentCueHandler(direction),
    sendMpvCommand: (rawCommand: (string | number)[]) =>
      sendMpvCommandRuntime(appState.mpvClient, rawCommand),
    getMpvClient: () => appState.mpvClient,
    isMpvConnected: () => Boolean(appState.mpvClient && appState.mpvClient.connected),
    hasRuntimeOptionsManager: () => appState.runtimeOptionsManager !== null,
  },
  handleMpvCommandFromIpcRuntime,
  runSubsyncManualFromIpc: (request) => subsyncRuntime.runManualFromIpc(request),
  registration: {
    runtimeOptions: {
      getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
      showMpvOsd: (text: string) => showMpvOsd(text),
    },
    mainDeps: {
      getMainWindow: () => overlayManager.getMainWindow(),
      getVisibleOverlayVisibility: () => overlayManager.getVisibleOverlayVisible(),
      focusMainWindow: () => {
        const mainWindow = overlayManager.getMainWindow();
        if (!mainWindow || mainWindow.isDestroyed()) return;
        if (!mainWindow.isFocused()) {
          mainWindow.focus();
        }
      },
      onOverlayModalClosed: (modal, senderWindow) => {
        const modalWindow = overlayManager.getModalWindow();
        if (
          senderWindow &&
          modalWindow &&
          senderWindow === modalWindow &&
          !senderWindow.isDestroyed()
        ) {
          senderWindow.setIgnoreMouseEvents(true, { forward: true });
          senderWindow.hide();
        }
        handleOverlayModalClosed(modal);
      },
      onOverlayModalOpened: (modal) => {
        overlayModalRuntime.notifyOverlayModalOpened(modal);
      },
      onOverlayMouseInteractionChanged: (active, senderWindow) => {
        const mainWindow = overlayManager.getMainWindow();
        if (!mainWindow || senderWindow !== mainWindow) {
          return;
        }
        if (visibleOverlayInteractionActive === active) {
          if (active && process.platform === 'darwin' && !mainWindow.isFocused()) {
            overlayVisibilityRuntime.updateVisibleOverlayVisibility();
          }
          return;
        }
        visibleOverlayInteractionActive = active;
        overlayVisibilityRuntime.updateVisibleOverlayVisibility();
      },
      onYoutubePickerResolve: (request) => youtubeFlowRuntime.resolveActivePicker(request),
      openYomitanSettings: () => openYomitanSettings(),
      quitApp: () => requestAppQuit(),
      toggleVisibleOverlay: () => toggleVisibleOverlay(),
      tokenizeCurrentSubtitle: async () => {
        const tokenizeSubtitleForCurrent = tokenizeSubtitleDeferred;
        return resolveCurrentSubtitleForRenderer({
          currentSubText: appState.currentSubText,
          currentSubtitleData: appState.currentSubtitleData,
          withCurrentSubtitleTiming: (payload) => withCurrentSubtitleTiming(payload),
          tokenizeSubtitle: tokenizeSubtitleForCurrent
            ? (text) => tokenizeSubtitleForCurrent(text)
            : undefined,
        });
      },
      getCurrentSubtitleRaw: () => appState.currentSubText,
      getCurrentSubtitleAss: () => appState.currentSubAssText,
      getSubtitleSidebarSnapshot: async () => {
        const currentSubtitle = {
          text: appState.currentSubText,
          startTime: appState.mpvClient?.currentSubStart ?? null,
          endTime: appState.mpvClient?.currentSubEnd ?? null,
        };
        const currentTimeSec = appState.mpvClient?.currentTimePos ?? null;
        const config = getResolvedConfig().subtitleSidebar;
        const client = appState.mpvClient;
        if (!client?.connected) {
          return {
            cues: appState.activeParsedSubtitleCues,
            currentTimeSec,
            currentSubtitle,
            config,
          };
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
            return {
              cues: appState.activeParsedSubtitleCues,
              currentTimeSec,
              currentSubtitle,
              config,
            };
          }
          if (
            shouldUseCachedYoutubeParsedCues({
              videoPath,
              cachedMediaPath: appState.activeParsedSubtitleMediaPath,
              cachedCueCount: appState.activeParsedSubtitleCues.length,
            })
          ) {
            return {
              cues: appState.activeParsedSubtitleCues,
              currentTimeSec,
              currentSubtitle,
              config,
            };
          }

          const resolvedSource = await resolveActiveSubtitleSidebarSourceHandler({
            currentExternalFilenameRaw,
            currentTrackRaw,
            trackListRaw,
            sidRaw,
            videoPath,
          });
          if (!resolvedSource) {
            return {
              cues: appState.activeParsedSubtitleCues,
              currentTimeSec,
              currentSubtitle,
              config,
            };
          }

          try {
            if (appState.activeParsedSubtitleSource === resolvedSource.sourceKey) {
              return {
                cues: appState.activeParsedSubtitleCues,
                currentTimeSec,
                currentSubtitle,
                config,
              };
            }

            const content = await loadSubtitleSourceText(resolvedSource.path);
            const cues = parseSubtitleCues(content, resolvedSource.path);
            appState.activeParsedSubtitleCues = cues;
            appState.activeParsedSubtitleSource = resolvedSource.sourceKey;
            appState.activeParsedSubtitleMediaPath = videoPath || null;
            return {
              cues,
              currentTimeSec,
              currentSubtitle,
              config,
            };
          } finally {
            await resolvedSource.cleanup?.();
          }
        } catch {
          return {
            cues: appState.activeParsedSubtitleCues,
            currentTimeSec,
            currentSubtitle,
            config,
          };
        }
      },
      getPlaybackPaused: () => appState.playbackPaused,
      getSubtitlePosition: () => loadSubtitlePosition(),
      getSubtitleStyle: () => {
        const resolvedConfig = getResolvedConfig();
        return resolveSubtitleStyleForRenderer(resolvedConfig);
      },
      saveSubtitlePosition: (position) => saveSubtitlePosition(position),
      getMecabTokenizer: () => appState.mecabTokenizer,
      getKeybindings: () => appState.keybindings,
      getSessionBindings: () => appState.sessionBindings,
      getConfiguredShortcuts: () => getConfiguredShortcuts(),
      dispatchSessionAction: (request) => dispatchSessionAction(request),
      getStatsToggleKey: () => getResolvedConfig().stats.toggleKey,
      getMarkWatchedKey: () => getResolvedConfig().stats.markWatchedKey,
      getControllerConfig: () => getResolvedConfig().controller,
      saveControllerConfig: (update) => {
        const currentRawConfig = configService.getRawConfig();
        configService.patchRawConfig({
          controller: applyControllerConfigUpdate(currentRawConfig.controller, update),
        });
      },
      saveControllerPreference: ({ preferredGamepadId, preferredGamepadLabel }) => {
        configService.patchRawConfig({
          controller: {
            preferredGamepadId,
            preferredGamepadLabel,
          },
        });
      },
      getSecondarySubMode: () => appState.secondarySubMode,
      getMpvClient: () => appState.mpvClient,
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
      runAnilistPostWatchUpdateOnManualMark: () => maybeRunAnilistPostWatchUpdate({ force: true }),
      getCharacterDictionarySelection: () =>
        characterDictionaryRuntime.getManualSelectionSnapshot(),
      setCharacterDictionarySelection: async (mediaId: number) =>
        applyCharacterDictionarySelection(
          { mediaId },
          {
            setManualSelection: (request) => characterDictionaryRuntime.setManualSelection(request),
            resetAnilistMediaGuessState,
            runSyncNow: () => characterDictionaryAutoSyncRuntime.runSyncNow(),
            warn: (message, error) => logger.warn(message, error),
          },
        ),
      appendClipboardVideoToQueue: () => appendClipboardVideoToQueue(),
      ...playlistBrowserMainDeps,
      getImmersionTracker: () => appState.immersionTracker,
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
        appState.ankiIntegration?.setRecordCardsMinedCallback(recordTrackedCardsMined);
        appState.ankiIntegration?.setKnownWordCacheUpdatedCallback(
          refreshCurrentSubtitleAfterKnownWordUpdate,
        );
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
    registerIpcRuntimeServices,
  },
});
const { handleCliCommand, handleInitialArgs } = composeCliStartupHandlers({
  cliCommandContextMainDeps: {
    appState,
    setLogLevel: (level) => setLogLevel(level, 'cli'),
    onMpvSocketPathChanged: (nextSocketPath, previousSocketPath) =>
      retargetOverlayWindowTrackerForMpvSocket(nextSocketPath, previousSocketPath),
    texthookerService,
    getResolvedConfig: () => getResolvedConfig(),
    defaultWebsocketPort: DEFAULT_CONFIG.websocket.port,
    defaultAnnotationWebsocketPort: DEFAULT_CONFIG.annotationWebsocket.port,
    hasMpvWebsocketPlugin: () => hasMpvWebsocketPlugin(),
    openExternal: (url: string) => shell.openExternal(url),
    logBrowserOpenError: (url: string, error: unknown) =>
      logger.error(`Failed to open browser for texthooker URL: ${url}`, error),
    showMpvOsd: (text: string) => showMpvOsd(text),
    initializeOverlayRuntime: () => initializeOverlayRuntime(),
    toggleVisibleOverlay: () => toggleVisibleOverlay(),
    togglePrimarySubtitleBar: () => togglePrimarySubtitleBar(),
    openFirstRunSetupWindow: (force?: boolean) => openFirstRunSetupWindow(force),
    setVisibleOverlayVisible: (visible: boolean) => setVisibleOverlayVisible(visible),
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
    generateCharacterDictionary: async (targetPath?: string) => {
      const disabledReason = yomitanProfilePolicy.getCharacterDictionaryDisabledReason();
      if (disabledReason) {
        throw new Error(disabledReason);
      }
      return await characterDictionaryRuntime.generateForCurrentMedia(targetPath);
    },
    getCharacterDictionarySelection: async (targetPath?: string) =>
      characterDictionaryRuntime.getManualSelectionSnapshot(targetPath),
    setCharacterDictionarySelection: async (request) =>
      applyCharacterDictionarySelection(request, {
        setManualSelection: (selectionRequest) =>
          characterDictionaryRuntime.setManualSelection(selectionRequest),
        resetAnilistMediaGuessState,
        runSyncNow: () => characterDictionaryAutoSyncRuntime.runSyncNow(),
        warn: (message, error) => logger.warn(message, error),
      }),
    runJellyfinCommand: (argsFromCommand: CliArgs) => runJellyfinCommand(argsFromCommand),
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
    runYoutubePlaybackFlow: (request) => youtubePlaybackRuntime.runYoutubePlaybackFlow(request),
    openYomitanSettings: () => openYomitanSettings(),
    openConfigSettingsWindow: () => openConfigSettingsWindow(),
    cycleSecondarySubMode: () => handleCycleSecondarySubMode(),
    openRuntimeOptionsPalette: () => openRuntimeOptionsPalette(),
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
        ensureTray();
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
    ensureTray: () => ensureTray(),
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
      getDefaultSocketPath: () => getResolvedConfig().mpv.socketPath || getDefaultSocketPath(),
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
  if (!shouldStartAutomaticUpdateChecks(appState.initialArgs)) {
    return;
  }
  getUpdateService().startAutomaticChecks();
});
const startupModeFlags = getStartupModeFlags(appState.initialArgs);
const shouldUseMinimalStartup = startupModeFlags.shouldUseMinimalStartup;
const shouldSkipHeavyStartup = startupModeFlags.shouldSkipHeavyStartup;
if (!appState.initialArgs || (!shouldUseMinimalStartup && !shouldSkipHeavyStartup)) {
  if (isAnilistTrackingEnabled(getResolvedConfig())) {
    void refreshAnilistClientSecretStateIfEnabled({
      force: true,
      allowSetupPrompt: false,
    }).catch((error) => {
      logger.error('Failed to refresh AniList client secret state during startup', error);
    });
    anilistStateRuntime.refreshRetryQueueState();
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
      ensureOverlayWindowLevel: (window) => ensureOverlayWindowLevel(window),
      onRuntimeOptionsChanged: () => broadcastRuntimeOptionsChanged(),
      setOverlayDebugVisualizationEnabled: (enabled) =>
        setOverlayDebugVisualizationEnabled(enabled),
      isOverlayVisible: (windowKind) =>
        windowKind === 'visible' ? overlayManager.getVisibleOverlayVisible() : false,
      getYomitanSession: () => appState.yomitanSession,
      tryHandleOverlayShortcutLocalFallback: (input) =>
        overlayShortcutsRuntime.tryHandleOverlayShortcutLocalFallback(input),
      forwardTabToMpv: () => sendMpvCommandRuntime(appState.mpvClient, ['keypress', 'TAB']),
      onVisibleWindowBlurred: () => scheduleVisibleOverlayBlurRefresh(),
      onWindowContentReady: () => {
        overlayVisibilityRuntime.updateVisibleOverlayVisibility();
        if (appState.currentSubText.trim()) {
          subtitleProcessingController.refreshCurrentSubtitle(appState.currentSubText);
        }
        autoplayReadyGate.flushPendingAutoplayReadySignal();
      },
      onWindowClosed: (windowKind) => {
        if (windowKind === 'visible') {
          cancelPendingLinuxMpvFullscreenOverlayRefreshBurst();
          overlayManager.setMainWindow(null);
        } else {
          overlayManager.setModalWindow(null);
        }
      },
    },
    setMainWindow: (window) => overlayManager.setMainWindow(window),
    setModalWindow: (window) => overlayManager.setModalWindow(window),
  });

function refreshTrayMenuIfPresent(): void {
  if (appTray) {
    ensureTrayHandler();
  }
}

function getJellyfinTrayDiscoveryDeps() {
  return {
    getResolvedJellyfinConfig: () => getResolvedJellyfinConfig(),
    getRemoteSession: () => appState.jellyfinRemoteSession,
    clearStoredSession: () => jellyfinTokenStore.clearSession(),
    stopRemoteSession: () => stopJellyfinRemoteSession(),
    startRemoteSession: (options: { explicit: true }) => startJellyfinRemoteSession(options),
    refreshTrayMenu: () => refreshTrayMenuIfPresent(),
    logger,
    showMpvOsd: (message: string) => showMpvOsd(message),
  };
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
      openSessionHelpModal: () => openSessionHelpOverlay(),
      openTexthookerInBrowser: () =>
        handleCliCommand(parseArgs(['--texthooker', '--open-browser'])),
      showTexthookerPage: () => shouldShowTexthookerTrayEntry(getResolvedConfig()),
      showFirstRunSetup: () => !firstRunSetupService.isSetupCompleted(),
      openFirstRunSetupWindow: (force?: boolean) => openFirstRunSetupWindow(force),
      showWindowsMpvLauncherSetup: () => process.platform === 'win32',
      openYomitanSettings: () => openYomitanSettings(),
      openConfigSettingsWindow: () => openConfigSettingsWindow(),
      openJellyfinSetupWindow: () => openJellyfinSetupWindow(),
      isJellyfinConfigured: () =>
        isJellyfinConfiguredForTrayRuntime(getJellyfinTrayDiscoveryDeps()),
      isJellyfinDiscoveryActive: () => Boolean(appState.jellyfinRemoteSession),
      toggleJellyfinDiscovery: (checked: boolean) =>
        toggleJellyfinDiscoveryFromTrayRuntime(getJellyfinTrayDiscoveryDeps(), {
          desiredActive: checked,
        }),
      openAnilistSetupWindow: () => openAnilistSetupWindow(),
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
  externalProfilePath: getResolvedConfig().yomitan.externalProfilePath,
  logInfo: (message) => logger.info(message),
});
const configuredExternalYomitanProfilePath = yomitanProfilePolicy.externalProfilePath;
const yomitanExtensionRuntime = createYomitanExtensionRuntime({
  loadYomitanExtensionCore,
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
      getOverlayWindows(),
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
        createOverlayWindowTracker(override, targetMpvSocketPath),
      updateVisibleOverlayBounds: (geometry: WindowGeometry) =>
        updateVisibleOverlayBounds(geometry),
      bindOverlayOwner: () => bindVisibleOverlayOwner(),
      releaseOverlayOwner: () => releaseVisibleOverlayOwner(),
      getOverlayWindows: () => getOverlayWindows(),
      getResolvedConfig: () => getResolvedConfig(),
      showDesktopNotification,
      createFieldGroupingCallback: () => createFieldGroupingCallback(),
      getKnownWordCacheStatePath: () => path.join(USER_DATA_PATH, 'known-words-cache.json'),
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

function handleMineSentenceDigit(count: number): void {
  handleMineSentenceDigitHandler(count);
}

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
    autoplayReadyGate.markCurrentMediaAutoplayReady();
    cancelPendingLinuxMpvFullscreenOverlayRefreshBurst();
  }
  if (visible) {
    void ensureOverlayMpvSubtitlesHidden();
  }
  setVisibleOverlayVisibleHandler(visible);
  notifyMpvPluginVisibleOverlayVisibility(visible);
  syncOverlayMpvSubtitleSuppression();
}

function toggleVisibleOverlay(): void {
  ensureOverlayWindowsReadyForVisibilityActions();
  const nextVisible = !overlayManager.getVisibleOverlayVisible();
  if (!nextVisible) {
    autoplayReadyGate.markCurrentMediaAutoplayReady();
    cancelPendingLinuxMpvFullscreenOverlayRefreshBurst();
  } else {
    void ensureOverlayMpvSubtitlesHidden();
  }
  toggleVisibleOverlayHandler();
  notifyMpvPluginVisibleOverlayVisibility(nextVisible);
  syncOverlayMpvSubtitleSuppression();
}
function setOverlayVisible(visible: boolean): void {
  if (!visible) {
    autoplayReadyGate.markCurrentMediaAutoplayReady();
    cancelPendingLinuxMpvFullscreenOverlayRefreshBurst();
  }
  if (visible) {
    void ensureOverlayMpvSubtitlesHidden();
  }
  setOverlayVisibleHandler(visible);
  notifyMpvPluginVisibleOverlayVisibility(visible);
  syncOverlayMpvSubtitleSuppression();
}
function handleOverlayModalClosed(modal: OverlayHostedModal): void {
  handleOverlayModalClosedHandler(modal);
}

function appendClipboardVideoToQueue(): { ok: boolean; message: string } {
  return appendClipboardVideoToQueueHandler();
}

registerIpcRuntimeHandlers();
