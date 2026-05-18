"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
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
const electron_1 = require("electron");
const controller_config_update_js_1 = require("./main/controller-config-update.js");
const playlist_browser_open_1 = require("./main/runtime/playlist-browser-open");
const discord_rpc_client_js_1 = require("./main/runtime/discord-rpc-client.js");
const linux_mpv_fullscreen_overlay_refresh_1 = require("./main/runtime/linux-mpv-fullscreen-overlay-refresh");
const config_1 = require("./ai/config");
function getPasswordStoreArg(argv) {
    for (let i = 0; i < argv.length; i += 1) {
        const arg = argv[i];
        if (!arg?.startsWith('--password-store')) {
            continue;
        }
        if (arg === '--password-store') {
            const value = argv[i + 1];
            if (value && !value.startsWith('--')) {
                return value;
            }
            return null;
        }
        const [prefix, value] = arg.split('=', 2);
        if (prefix === '--password-store' && value && value.trim().length > 0) {
            return value.trim();
        }
    }
    return null;
}
function normalizePasswordStoreArg(value) {
    const normalized = value.trim();
    if (normalized.toLowerCase() === 'gnome') {
        return 'gnome-libsecret';
    }
    return normalized;
}
function getDefaultPasswordStore() {
    return 'gnome-libsecret';
}
electron_1.protocol.registerSchemesAsPrivileged([
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
const fs = __importStar(require("fs"));
const node_child_process_1 = require("node:child_process");
const os = __importStar(require("os"));
const path = __importStar(require("path"));
const mecab_tokenizer_1 = require("./mecab-tokenizer");
const anki_integration_1 = require("./anki-integration");
const subtitle_timing_tracker_1 = require("./subtitle-timing-tracker");
const runtime_options_1 = require("./runtime-options");
const utils_1 = require("./jimaku/utils");
const logger_1 = require("./logger");
const window_trackers_1 = require("./window-trackers");
const windows_helper_1 = require("./window-trackers/windows-helper");
const args_1 = require("./cli/args");
const help_1 = require("./cli/help");
const contracts_1 = require("./shared/ipc/contracts");
const anki_connect_1 = require("./anki-connect");
const startup_mode_flags_1 = require("./main/runtime/startup-mode-flags");
const config_validation_1 = require("./main/config-validation");
const anilist_1 = require("./main/runtime/domains/anilist");
const watch_threshold_1 = require("./shared/watch-threshold");
const jellyfin_1 = require("./main/runtime/domains/jellyfin");
const overlay_1 = require("./main/runtime/domains/overlay");
const startup_1 = require("./main/runtime/domains/startup");
const mpv_1 = require("./main/runtime/domains/mpv");
const mining_1 = require("./main/runtime/domains/mining");
const utils_2 = require("./core/utils");
const setup_state_1 = require("./shared/setup-state");
const services_1 = require("./core/services");
const generate_1 = require("./core/services/youtube/generate");
const playback_resolve_1 = require("./core/services/youtube/playback-resolve");
const track_probe_1 = require("./core/services/youtube/track-probe");
const stats_server_1 = require("./core/services/stats-server");
const stats_window_js_1 = require("./core/services/stats-window.js");
const stats_window_js_2 = require("./core/services/stats-window.js");
const first_run_setup_service_1 = require("./main/runtime/first-run-setup-service");
const youtube_flow_1 = require("./main/runtime/youtube-flow");
const youtube_playback_runtime_1 = require("./main/runtime/youtube-playback-runtime");
const youtube_primary_subtitle_notification_1 = require("./main/runtime/youtube-primary-subtitle-notification");
const autoplay_ready_gate_1 = require("./main/runtime/autoplay-ready-gate");
const local_subtitle_selection_1 = require("./main/runtime/local-subtitle-selection");
const first_run_setup_window_1 = require("./main/runtime/first-run-setup-window");
const first_run_setup_plugin_1 = require("./main/runtime/first-run-setup-plugin");
const windows_mpv_shortcuts_1 = require("./main/runtime/windows-mpv-shortcuts");
const command_line_launcher_1 = require("./main/runtime/command-line-launcher");
const windows_mpv_launch_1 = require("./main/runtime/windows-mpv-launch");
const jellyfin_remote_connection_1 = require("./main/runtime/jellyfin-remote-connection");
const jellyfin_tray_discovery_1 = require("./main/runtime/jellyfin-tray-discovery");
const youtube_playback_launch_1 = require("./main/runtime/youtube-playback-launch");
const startup_tray_policy_1 = require("./main/runtime/startup-tray-policy");
const immersion_startup_1 = require("./main/runtime/immersion-startup");
const immersion_startup_main_deps_1 = require("./main/runtime/immersion-startup-main-deps");
const stats_cli_command_1 = require("./main/runtime/stats-cli-command");
const stats_daemon_1 = require("./main/runtime/stats-daemon");
const stats_server_routing_1 = require("./main/runtime/stats-server-routing");
const legacy_vocabulary_pos_1 = require("./core/services/immersion-tracker/legacy-vocabulary-pos");
const anilist_update_queue_1 = require("./core/services/anilist/anilist-update-queue");
const anilist_updater_1 = require("./core/services/anilist/anilist-updater");
const cover_art_fetcher_1 = require("./core/services/anilist/cover-art-fetcher");
const rate_limiter_1 = require("./core/services/anilist/rate-limiter");
const jellyfin_token_store_1 = require("./core/services/jellyfin-token-store");
const runtime_options_ipc_1 = require("./core/services/runtime-options-ipc");
const anilist_token_store_1 = require("./core/services/anilist/anilist-token-store");
const session_bindings_1 = require("./core/services/session-bindings");
const session_actions_1 = require("./core/services/session-actions");
const shortcuts_1 = require("./main/runtime/domains/shortcuts");
const registry_1 = require("./main/runtime/registry");
const overlay_mpv_sub_visibility_1 = require("./main/runtime/overlay-mpv-sub-visibility");
const composers_1 = require("./main/runtime/composers");
const overlay_window_runtime_handlers_1 = require("./main/runtime/overlay-window-runtime-handlers");
const startup_2 = require("./main/startup");
const startup_lifecycle_1 = require("./main/startup-lifecycle");
const early_single_instance_1 = require("./main/early-single-instance");
const ipc_mpv_command_1 = require("./main/ipc-mpv-command");
const ipc_runtime_1 = require("./main/ipc-runtime");
const dependencies_1 = require("./main/dependencies");
const services_2 = require("./main/boot/services");
const cli_runtime_1 = require("./main/cli-runtime");
const overlay_runtime_1 = require("./main/overlay-runtime");
const overlay_modal_input_state_1 = require("./main/runtime/overlay-modal-input-state");
const youtube_picker_open_1 = require("./main/runtime/youtube-picker-open");
const runtime_options_open_1 = require("./main/runtime/runtime-options-open");
const jimaku_open_1 = require("./main/runtime/jimaku-open");
const subsync_open_1 = require("./main/runtime/subsync-open");
const session_help_open_1 = require("./main/runtime/session-help-open");
const character_dictionary_open_1 = require("./main/runtime/character-dictionary-open");
const controller_select_open_1 = require("./main/runtime/controller-select-open");
const controller_debug_open_1 = require("./main/runtime/controller-debug-open");
const playlist_browser_ipc_1 = require("./main/runtime/playlist-browser-ipc");
const session_bindings_artifact_1 = require("./main/runtime/session-bindings-artifact");
const overlay_shortcuts_runtime_1 = require("./main/overlay-shortcuts-runtime");
const frequency_dictionary_runtime_1 = require("./main/frequency-dictionary-runtime");
const jlpt_runtime_1 = require("./main/jlpt-runtime");
const media_runtime_1 = require("./main/media-runtime");
const overlay_visibility_runtime_1 = require("./main/overlay-visibility-runtime");
const discord_presence_runtime_1 = require("./main/runtime/discord-presence-runtime");
const character_dictionary_runtime_1 = require("./main/character-dictionary-runtime");
const character_dictionary_auto_sync_1 = require("./main/runtime/character-dictionary-auto-sync");
const character_dictionary_auto_sync_completion_1 = require("./main/runtime/character-dictionary-auto-sync-completion");
const character_dictionary_auto_sync_notifications_1 = require("./main/runtime/character-dictionary-auto-sync-notifications");
const current_media_tokenization_gate_1 = require("./main/runtime/current-media-tokenization-gate");
const current_subtitle_snapshot_1 = require("./main/runtime/current-subtitle-snapshot");
const startup_osd_sequencer_1 = require("./main/runtime/startup-osd-sequencer");
const app_updater_1 = require("./main/runtime/update/app-updater");
const fetch_adapter_1 = require("./main/runtime/update/fetch-adapter");
const curl_http_executor_1 = require("./main/runtime/update/curl-http-executor");
const release_assets_1 = require("./main/runtime/update/release-assets");
const release_metadata_policy_1 = require("./main/runtime/update/release-metadata-policy");
const launcher_updater_1 = require("./main/runtime/update/launcher-updater");
const update_notifications_1 = require("./main/runtime/update/update-notifications");
const update_dialogs_1 = require("./main/runtime/update/update-dialogs");
const update_cli_command_1 = require("./main/runtime/update/update-cli-command");
const update_service_1 = require("./main/runtime/update/update-service");
const support_assets_1 = require("./main/runtime/update/support-assets");
const subtitle_prefetch_runtime_1 = require("./main/runtime/subtitle-prefetch-runtime");
const setup_window_factory_1 = require("./main/runtime/setup-window-factory");
const config_settings_runtime_1 = require("./main/runtime/config-settings-runtime");
const youtube_playback_1 = require("./main/runtime/youtube-playback");
const yomitan_profile_policy_1 = require("./main/runtime/yomitan-profile-policy");
const yomitan_read_only_log_1 = require("./main/runtime/yomitan-read-only-log");
const yomitan_anki_server_1 = require("./main/runtime/yomitan-anki-server");
const state_1 = require("./main/state");
const anilist_url_guard_1 = require("./main/anilist-url-guard");
const config_2 = require("./config");
const path_resolution_1 = require("./config/path-resolution");
const registry_2 = require("./config/settings/registry");
const subtitle_cue_parser_1 = require("./core/services/subtitle-cue-parser");
const subtitle_prefetch_1 = require("./core/services/subtitle-prefetch");
const subtitle_prefetch_source_1 = require("./main/runtime/subtitle-prefetch-source");
const subtitle_prefetch_init_1 = require("./main/runtime/subtitle-prefetch-init");
const character_dictionary_selection_1 = require("./main/character-dictionary-selection");
const utils_3 = require("./subsync/utils");
if (process.platform === 'linux') {
    electron_1.app.commandLine.appendSwitch('enable-features', 'GlobalShortcutsPortal');
    const passwordStore = normalizePasswordStoreArg(getPasswordStoreArg(process.argv) ?? getDefaultPasswordStore());
    electron_1.app.commandLine.appendSwitch('password-store', passwordStore);
    (0, logger_1.createLogger)('main').debug(`Applied --password-store ${passwordStore}`);
}
electron_1.app.setName('SubMiner');
const DEFAULT_TEXTHOOKER_PORT = 5174;
const DEFAULT_MPV_LOG_FILE = (0, logger_1.resolveDefaultLogFilePath)({
    platform: process.platform,
    homeDir: os.homedir(),
    appDataDir: process.env.APPDATA,
});
const ANILIST_SETUP_CLIENT_ID_URL = 'https://anilist.co/api/v2/oauth/authorize';
const ANILIST_SETUP_RESPONSE_TYPE = 'token';
const ANILIST_DEFAULT_CLIENT_ID = '36084';
const ANILIST_REDIRECT_URI = 'https://anilist.subminer.moe/';
const ANILIST_DEVELOPER_SETTINGS_URL = 'https://anilist.co/settings/developer';
const ANILIST_UPDATE_MIN_WATCH_SECONDS = 10 * 60;
const ANILIST_DURATION_RETRY_INTERVAL_MS = 15_000;
const ANILIST_MAX_ATTEMPTED_UPDATE_KEYS = 1000;
const TRAY_TOOLTIP = 'SubMiner';
let anilistMediaGuessRuntimeState = (0, state_1.createInitialAnilistMediaGuessRuntimeState)();
let anilistUpdateInFlightState = (0, state_1.createInitialAnilistUpdateInFlightState)();
const anilistAttemptedUpdateKeys = new Set();
let anilistCachedAccessToken = null;
let jellyfinPlayQuitOnDisconnectArmed = false;
const JELLYFIN_LANG_PREF = 'ja,jp,jpn,japanese,en,eng,english,enUS,en-US';
const JELLYFIN_TICKS_PER_SECOND = 10_000_000;
const JELLYFIN_REMOTE_PROGRESS_INTERVAL_MS = 3000;
const DISCORD_PRESENCE_APP_ID = '1475264834730856619';
const JELLYFIN_MPV_CONNECT_TIMEOUT_MS = 3000;
const JELLYFIN_MPV_AUTO_LAUNCH_TIMEOUT_MS = 20000;
const YOUTUBE_MPV_CONNECT_TIMEOUT_MS = 3000;
const YOUTUBE_MPV_AUTO_LAUNCH_TIMEOUT_MS = 10000;
const YOUTUBE_MPV_YTDL_FORMAT = 'bestvideo*+bestaudio/best';
const YOUTUBE_DIRECT_PLAYBACK_FORMAT = 'b';
const MPV_JELLYFIN_DEFAULT_ARGS = [
    '--sub-auto=fuzzy',
    '--sub-file-paths=.;subs;subtitles',
    '--sid=auto',
    '--secondary-sid=auto',
    '--secondary-sub-visibility=no',
    '--alang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
    '--slang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
];
let activeJellyfinRemotePlayback = null;
let jellyfinRemoteLastProgressAtMs = 0;
let jellyfinMpvAutoLaunchInFlight = null;
let backgroundWarmupsStarted = false;
let yomitanLoadInFlight = null;
let notifyAnilistTokenStoreWarning = () => { };
const buildApplyJellyfinMpvDefaultsMainDepsHandler = (0, jellyfin_1.createBuildApplyJellyfinMpvDefaultsMainDepsHandler)({
    sendMpvCommandRuntime: (client, command) => (0, services_1.sendMpvCommandRuntime)(client, command),
    jellyfinLangPref: JELLYFIN_LANG_PREF,
});
const applyJellyfinMpvDefaultsMainDeps = buildApplyJellyfinMpvDefaultsMainDepsHandler();
const applyJellyfinMpvDefaultsHandler = (0, jellyfin_1.createApplyJellyfinMpvDefaultsHandler)(applyJellyfinMpvDefaultsMainDeps);
function applyJellyfinMpvDefaults(client) {
    applyJellyfinMpvDefaultsHandler(client);
}
const isDev = process.argv.includes('--dev') || process.argv.includes('--debug');
const texthookerService = new services_1.Texthooker(() => {
    const config = getResolvedConfig();
    const characterDictionaryEnabled = config.anilist.characterDictionary.enabled &&
        yomitanProfilePolicy.isCharacterDictionaryEnabled();
    const knownWordColoringEnabled = getRuntimeBooleanOption('subtitle.annotation.knownWords.highlightEnabled', config.ankiConnect.knownWords.highlightEnabled);
    const nPlusOneColoringEnabled = getRuntimeBooleanOption('subtitle.annotation.nPlusOne', config.ankiConnect.nPlusOne.enabled);
    return {
        enableKnownWordColoring: knownWordColoringEnabled,
        enableNPlusOneColoring: nPlusOneColoringEnabled,
        enableNameMatchColoring: config.subtitleStyle.nameMatchEnabled && characterDictionaryEnabled,
        enableFrequencyColoring: getRuntimeBooleanOption('subtitle.annotation.frequency', config.subtitleStyle.frequencyDictionary.enabled),
        enableJlptColoring: getRuntimeBooleanOption('subtitle.annotation.jlpt', config.subtitleStyle.enableJlpt),
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
let syncOverlayShortcutsForModal = () => { };
let syncOverlayVisibilityForModal = () => { };
const buildGetDefaultSocketPathMainDepsHandler = (0, jellyfin_1.createBuildGetDefaultSocketPathMainDepsHandler)({
    platform: process.platform,
});
const getDefaultSocketPathMainDeps = buildGetDefaultSocketPathMainDepsHandler();
const getDefaultSocketPathHandler = (0, jellyfin_1.createGetDefaultSocketPathHandler)(getDefaultSocketPathMainDeps);
function getDefaultSocketPath() {
    return getDefaultSocketPathHandler();
}
const bootServices = (0, services_2.createMainBootServices)({
    platform: process.platform,
    argv: process.argv,
    appDataDir: process.env.APPDATA,
    xdgConfigHome: process.env.XDG_CONFIG_HOME,
    homeDir: os.homedir(),
    defaultMpvLogFile: DEFAULT_MPV_LOG_FILE,
    envMpvLog: process.env.SUBMINER_MPV_LOG,
    defaultTexthookerPort: DEFAULT_TEXTHOOKER_PORT,
    getDefaultSocketPath: () => getDefaultSocketPath(),
    resolveConfigDir: path_resolution_1.resolveConfigDir,
    existsSync: fs.existsSync,
    mkdirSync: fs.mkdirSync,
    joinPath: (...parts) => path.join(...parts),
    app: electron_1.app,
    shouldBypassSingleInstanceLock: () => (0, early_single_instance_1.shouldBypassSingleInstanceLockForArgv)(process.argv),
    requestSingleInstanceLockEarly: () => (0, early_single_instance_1.requestSingleInstanceLockEarly)(electron_1.app),
    registerSecondInstanceHandlerEarly: (listener) => {
        (0, early_single_instance_1.registerSecondInstanceHandlerEarly)(electron_1.app, listener);
    },
    onConfigStartupParseError: (error) => {
        (0, config_validation_1.failStartupFromConfig)('SubMiner config parse error', (0, config_validation_1.buildConfigParseErrorDetails)(error.path, error.parseError), {
            logError: (details) => console.error(details),
            showErrorBox: (title, details) => electron_1.dialog.showErrorBox(title, details),
            quit: () => requestAppQuit(),
        });
    },
    createConfigService: (configDir) => new config_2.ConfigService(configDir),
    createAnilistTokenStore: (targetPath) => (0, anilist_token_store_1.createAnilistTokenStore)(targetPath, {
        info: (message) => console.info(message),
        warn: (message, details) => console.warn(message, details),
        error: (message, details) => console.error(message, details),
        warnUser: (message) => notifyAnilistTokenStoreWarning(message),
    }),
    createJellyfinTokenStore: (targetPath) => (0, jellyfin_token_store_1.createJellyfinTokenStore)(targetPath, {
        info: (message) => console.info(message),
        warn: (message, details) => console.warn(message, details),
        error: (message, details) => console.error(message, details),
    }),
    createAnilistUpdateQueue: (targetPath) => (0, anilist_update_queue_1.createAnilistUpdateQueue)(targetPath, {
        info: (message) => console.info(message),
        warn: (message, details) => console.warn(message, details),
        error: (message, details) => console.error(message, details),
    }),
    createSubtitleWebSocket: () => new services_1.SubtitleWebSocket(),
    createLogger: logger_1.createLogger,
    createMainRuntimeRegistry: registry_1.createMainRuntimeRegistry,
    createOverlayManager: services_1.createOverlayManager,
    createOverlayModalInputState: overlay_modal_input_state_1.createOverlayModalInputState,
    createOverlayContentMeasurementStore: ({ logger }) => {
        const buildHandler = (0, overlay_1.createBuildOverlayContentMeasurementStoreMainDepsHandler)({
            now: () => Date.now(),
            warn: (message) => logger.warn(message),
        });
        return (0, services_1.createOverlayContentMeasurementStore)(buildHandler());
    },
    getSyncOverlayShortcutsForModal: () => syncOverlayShortcutsForModal,
    getSyncOverlayVisibilityForModal: () => syncOverlayVisibilityForModal,
    createOverlayModalRuntime: ({ overlayManager, overlayModalInputState }) => {
        const buildHandler = (0, overlay_1.createBuildOverlayModalRuntimeMainDepsHandler)({
            getMainWindow: () => overlayManager.getMainWindow(),
            getModalWindow: () => overlayManager.getModalWindow(),
            createModalWindow: () => createModalWindow(),
            getModalGeometry: () => getCurrentOverlayGeometry(),
            setModalWindowBounds: (geometry) => overlayManager.setModalWindowBounds(geometry),
        });
        return (0, overlay_runtime_1.createOverlayModalRuntimeService)(buildHandler(), {
            onModalStateChange: (isActive) => overlayModalInputState.handleModalInputStateChange(isActive),
        });
    },
    createAppState: state_1.createAppState,
});
const { configDir: CONFIG_DIR, userDataPath: USER_DATA_PATH, defaultMpvLogPath: DEFAULT_MPV_LOG_PATH, defaultImmersionDbPath: DEFAULT_IMMERSION_DB_PATH, configService, anilistTokenStore, jellyfinTokenStore, anilistUpdateQueue, subtitleWsService, annotationSubtitleWsService, logger, runtimeRegistry, overlayManager, overlayModalInputState, overlayContentMeasurementStore, overlayModalRuntime, appState, appLifecycleApp, } = bootServices;
const configSettingsFields = (0, registry_2.buildConfigSettingsRegistry)(config_2.DEFAULT_CONFIG);
notifyAnilistTokenStoreWarning = (message) => {
    logger.warn(`[AniList] ${message}`);
    try {
        (0, utils_2.showDesktopNotification)('SubMiner AniList', {
            body: message,
        });
    }
    catch {
        // Notification may fail if desktop notifications are unavailable early in startup.
    }
};
const appLogger = {
    logInfo: (message) => {
        logger.info(message);
    },
    logDebug: (message) => {
        logger.debug(message);
    },
    logWarning: (message) => {
        logger.warn(message);
    },
    logError: (message, details) => {
        logger.error(message, details);
    },
    logNoRunningInstance: () => {
        logger.error('No running instance. Use --start to launch the app.');
    },
    logConfigWarning: (warning) => {
        logger.warn(`[config] ${warning.path}: ${warning.message} value=${JSON.stringify(warning.value)} fallback=${JSON.stringify(warning.fallback)}`);
    },
};
let forceQuitTimer = null;
let statsServer = null;
const statsDaemonStatePath = path.join(USER_DATA_PATH, 'stats-daemon.json');
function readLiveBackgroundStatsDaemonState() {
    const state = (0, stats_daemon_1.readBackgroundStatsServerState)(statsDaemonStatePath);
    if (!state) {
        (0, stats_daemon_1.removeBackgroundStatsServerState)(statsDaemonStatePath);
        return null;
    }
    if (state.pid === process.pid && !statsServer) {
        (0, stats_daemon_1.removeBackgroundStatsServerState)(statsDaemonStatePath);
        return null;
    }
    if (!(0, stats_daemon_1.isBackgroundStatsServerProcessAlive)(state.pid)) {
        (0, stats_daemon_1.removeBackgroundStatsServerState)(statsDaemonStatePath);
        return null;
    }
    return state;
}
function clearOwnedBackgroundStatsDaemonState() {
    const state = (0, stats_daemon_1.readBackgroundStatsServerState)(statsDaemonStatePath);
    if (state?.pid === process.pid) {
        (0, stats_daemon_1.removeBackgroundStatsServerState)(statsDaemonStatePath);
    }
}
function stopStatsServer() {
    if (!statsServer) {
        return;
    }
    statsServer.close();
    statsServer = null;
    clearOwnedBackgroundStatsDaemonState();
}
function requestAppQuit() {
    (0, services_1.destroyYomitanSettingsWindow)(appState.yomitanSettingsWindow);
    appState.yomitanSettingsWindow = null;
    (0, stats_window_js_1.destroyStatsWindow)();
    stopStatsServer();
    if (!forceQuitTimer) {
        forceQuitTimer = setTimeout(() => {
            logger.warn('App quit timed out; forcing process exit.');
            electron_1.app.exit(0);
        }, 2000);
    }
    electron_1.app.quit();
}
process.on('SIGINT', () => {
    requestAppQuit();
});
process.on('SIGTERM', () => {
    requestAppQuit();
});
const startBackgroundWarmupsIfAllowed = () => {
    startBackgroundWarmups();
};
const youtubeFlowRuntime = (0, youtube_flow_1.createYoutubeFlowRuntime)({
    probeYoutubeTracks: (url) => (0, track_probe_1.probeYoutubeTracks)(url),
    acquireYoutubeSubtitleTrack: (input) => (0, generate_1.acquireYoutubeSubtitleTrack)(input),
    acquireYoutubeSubtitleTracks: (input) => (0, generate_1.acquireYoutubeSubtitleTracks)(input),
    openPicker: async (payload) => {
        return await (0, youtube_picker_open_1.openYoutubeTrackPicker)({
            sendToActiveOverlayWindow: (channel, nextPayload, runtimeOptions) => overlayModalRuntime.sendToActiveOverlayWindow(channel, nextPayload, runtimeOptions),
            waitForModalOpen: (modal, timeoutMs) => overlayModalRuntime.waitForModalOpen(modal, timeoutMs),
            logWarn: (message) => logger.warn(message),
        }, payload);
    },
    pauseMpv: () => {
        (0, services_1.sendMpvCommandRuntime)(appState.mpvClient, ['set_property', 'pause', 'yes']);
    },
    resumeMpv: () => {
        (0, services_1.sendMpvCommandRuntime)(appState.mpvClient, ['set_property', 'pause', 'no']);
    },
    sendMpvCommand: (command) => {
        (0, services_1.sendMpvCommandRuntime)(appState.mpvClient, command);
    },
    requestMpvProperty: async (name) => {
        const client = appState.mpvClient;
        if (!client)
            return null;
        return await client.requestProperty(name);
    },
    refreshCurrentSubtitle: (text) => {
        subtitleProcessingController.refreshCurrentSubtitle(text);
    },
    refreshSubtitleSidebarSource: async (sourcePath) => {
        await subtitlePrefetchRuntime.refreshSubtitleSidebarFromSource(sourcePath);
    },
    startTokenizationWarmups: async () => {
        await startTokenizationWarmups();
    },
    waitForTokenizationReady: async () => {
        await currentMediaTokenizationGate.waitUntilReady(appState.currentMediaPath?.trim() || appState.mpvClient?.currentVideoPath?.trim() || null);
    },
    waitForAnkiReady: async () => {
        const integration = appState.ankiIntegration;
        if (!integration) {
            return;
        }
        try {
            await Promise.race([
                integration.waitUntilReady(),
                new Promise((_, reject) => {
                    setTimeout(() => reject(new Error('Timed out waiting for AnkiConnect integration')), 2500);
                }),
            ]);
        }
        catch (error) {
            logger.warn('Continuing YouTube playback before AnkiConnect integration reported ready:', error instanceof Error ? error.message : String(error));
        }
    },
    wait: (ms) => new Promise((resolve) => setTimeout(resolve, ms)),
    waitForPlaybackWindowReady: async () => {
        const deadline = Date.now() + 4000;
        let stableGeometry = null;
        let stableSinceMs = 0;
        while (Date.now() < deadline) {
            const tracker = appState.windowTracker;
            const trackerGeometry = tracker?.getGeometry() ?? null;
            const mediaPath = appState.currentMediaPath?.trim() || appState.mpvClient?.currentVideoPath?.trim() || '';
            const trackerFocused = tracker?.isTargetWindowFocused() ?? false;
            if (tracker && tracker.isTracking() && trackerGeometry && trackerFocused && mediaPath) {
                if (!geometryMatches(stableGeometry, trackerGeometry)) {
                    stableGeometry = trackerGeometry;
                    stableSinceMs = Date.now();
                }
                else if (Date.now() - stableSinceMs >= 200) {
                    return;
                }
            }
            else {
                stableGeometry = null;
                stableSinceMs = 0;
            }
            await new Promise((resolve) => setTimeout(resolve, 100));
        }
        logger.warn('Timed out waiting for tracked playback window focus/media readiness before opening YouTube subtitle picker.');
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
    showMpvOsd: (text) => showMpvOsd(text),
    reportSubtitleFailure: (message) => reportYoutubeSubtitleFailure(message),
    warn: (message) => logger.warn(message),
    log: (message) => logger.info(message),
    getYoutubeOutputDir: () => path.join(os.homedir(), '.cache', 'subminer', 'youtube-subs'),
});
const prepareYoutubePlaybackInMpv = (0, youtube_playback_launch_1.createPrepareYoutubePlaybackInMpvHandler)({
    requestPath: async () => {
        const client = appState.mpvClient;
        if (!client)
            return null;
        const value = await client.requestProperty('path').catch(() => null);
        return typeof value === 'string' ? value : null;
    },
    requestProperty: async (name) => {
        const client = appState.mpvClient;
        if (!client)
            return null;
        return await client.requestProperty(name);
    },
    sendMpvCommand: (command) => {
        (0, services_1.sendMpvCommandRuntime)(appState.mpvClient, command);
    },
    wait: (ms) => new Promise((resolve) => setTimeout(resolve, ms)),
});
const waitForYoutubeMpvConnected = (0, jellyfin_remote_connection_1.createWaitForMpvConnectedHandler)({
    getMpvClient: () => appState.mpvClient,
    now: () => Date.now(),
    sleep: (delayMs) => new Promise((resolve) => setTimeout(resolve, delayMs)),
});
const autoplayReadyGate = (0, autoplay_ready_gate_1.createAutoplayReadyGate)({
    isAppOwnedFlowInFlight: () => youtubePrimarySubtitleNotificationRuntime.isAppOwnedFlowInFlight(),
    getCurrentMediaPath: () => appState.currentMediaPath,
    getCurrentVideoPath: () => appState.mpvClient?.currentVideoPath ?? null,
    getPlaybackPaused: () => appState.playbackPaused,
    getMpvClient: () => appState.mpvClient,
    signalPluginAutoplayReady: () => {
        (0, services_1.sendMpvCommandRuntime)(appState.mpvClient, ['script-message', 'subminer-autoplay-ready']);
    },
    schedule: (callback, delayMs) => setTimeout(callback, delayMs),
    logDebug: (message) => logger.debug(message),
});
const managedLocalSubtitleSelectionRuntime = (0, local_subtitle_selection_1.createManagedLocalSubtitleSelectionRuntime)({
    getCurrentMediaPath: () => appState.currentMediaPath,
    getMpvClient: () => appState.mpvClient,
    getPrimarySubtitleLanguages: () => getResolvedConfig().youtube.primarySubLanguages,
    getSecondarySubtitleLanguages: () => getResolvedConfig().secondarySub.secondarySubLanguages,
    sendMpvCommand: (command) => {
        (0, services_1.sendMpvCommandRuntime)(appState.mpvClient, command);
    },
    schedule: (callback, delayMs) => setTimeout(callback, delayMs),
    clearScheduled: (timer) => clearTimeout(timer),
});
function resolveBundledMpvRuntimePluginEntrypoint() {
    return ((0, first_run_setup_plugin_1.resolvePackagedRuntimePluginPath)({
        dirname: __dirname,
        appPath: electron_1.app.getAppPath(),
        resourcesPath: process.resourcesPath,
    }) ?? undefined);
}
function detectWindowsInstalledMpvPlugin(mpvExecutablePath) {
    return (0, first_run_setup_plugin_1.detectInstalledMpvPlugin)({
        platform: 'win32',
        homeDir: os.homedir(),
        appDataDir: electron_1.app.getPath('appData'),
        mpvExecutablePath,
    });
}
function logInstalledMpvPluginDetected(detection) {
    if (!detection.path)
        return;
    logger.warn(`SubMiner detected an installed mpv plugin at ${detection.path}. This mpv session will use the installed plugin. Remove it to use the bundled runtime plugin automatically. Detected plugin version: ${detection.version ?? 'unknown or legacy'}.`);
}
async function promptForLegacyMpvPluginRemovalBeforeWindowsLaunch(mpvPath, detection) {
    const response = await electron_1.dialog.showMessageBox({
        type: 'warning',
        title: 'SubMiner mpv plugin detected',
        message: [
            'SubMiner detected an installed mpv plugin at:',
            detection.path ?? 'unknown path',
            '',
            "This mpv session will use the installed plugin unless it is removed. Remove it now to use SubMiner's bundled runtime plugin automatically.",
            `Detected plugin version: ${detection.version ?? 'unknown or legacy'}`,
        ].join('\n'),
        detail: 'Remove the legacy SubMiner mpv plugin files from mpv before launching this video? This moves the files to the OS trash.',
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
    const result = await (0, first_run_setup_plugin_1.removeLegacyMpvPluginCandidates)({
        candidates: (0, first_run_setup_plugin_1.detectInstalledFirstRunPluginCandidates)({
            platform: 'win32',
            homeDir: os.homedir(),
            appDataDir: electron_1.app.getPath('appData'),
            mpvExecutablePath: mpvPath,
        }),
        trashItem: (candidatePath) => electron_1.shell.trashItem(candidatePath),
    });
    if (result.ok) {
        await electron_1.dialog.showMessageBox({
            type: 'info',
            title: 'Legacy mpv plugin removed',
            message: 'Legacy mpv plugin removed. SubMiner-managed playback will use the bundled runtime plugin.',
        });
        return 'removed';
    }
    await electron_1.dialog.showMessageBox({
        type: 'error',
        title: 'Could not remove legacy mpv plugin',
        message: 'Some legacy SubMiner mpv plugin files could not be moved to the trash.',
        detail: result.failedPaths.map((failure) => `${failure.path}: ${failure.message}`).join('\n'),
    });
    return 'cancel';
}
const youtubePlaybackRuntime = (0, youtube_playback_runtime_1.createYoutubePlaybackRuntime)({
    platform: process.platform,
    directPlaybackFormat: YOUTUBE_DIRECT_PLAYBACK_FORMAT,
    mpvYtdlFormat: YOUTUBE_MPV_YTDL_FORMAT,
    autoLaunchTimeoutMs: YOUTUBE_MPV_AUTO_LAUNCH_TIMEOUT_MS,
    connectTimeoutMs: YOUTUBE_MPV_CONNECT_TIMEOUT_MS,
    getSocketPath: () => appState.mpvSocketPath,
    getMpvConnected: () => Boolean(appState.mpvClient?.connected),
    invalidatePendingAutoplayReadyFallbacks: () => autoplayReadyGate.invalidatePendingAutoplayReadyFallbacks(),
    setAppOwnedFlowInFlight: (next) => {
        youtubePrimarySubtitleNotificationRuntime.setAppOwnedFlowInFlight(next);
    },
    ensureYoutubePlaybackRuntimeReady: async () => {
        await ensureYoutubePlaybackRuntimeReady();
    },
    resolveYoutubePlaybackUrl: (url, format) => (0, playback_resolve_1.resolveYoutubePlaybackUrl)(url, format),
    launchWindowsMpv: (playbackUrl, args) => (0, windows_mpv_launch_1.launchWindowsMpv)([playbackUrl], (0, windows_mpv_launch_1.createWindowsMpvLaunchDeps)({
        showError: (title, content) => electron_1.dialog.showErrorBox(title, content),
    }), [...args, `--log-file=${DEFAULT_MPV_LOG_PATH}`], process.execPath, resolveBundledMpvRuntimePluginEntrypoint(), getResolvedConfig().mpv.executablePath, getResolvedConfig().mpv.launchMode, {
        detectInstalledMpvPlugin: detectWindowsInstalledMpvPlugin,
        notifyInstalledPluginDetected: logInstalledMpvPluginDetected,
        resolveInstalledPluginBeforeLaunch: (detection, mpvPath) => promptForLegacyMpvPluginRemovalBeforeWindowsLaunch(mpvPath, detection),
    }, {
        socketPath: appState.mpvSocketPath,
        binaryPath: getResolvedConfig().mpv.subminerBinaryPath,
        backend: getResolvedConfig().mpv.backend,
        autoStart: getResolvedConfig().mpv.autoStartSubMiner,
        autoStartVisibleOverlay: getResolvedConfig().auto_start_overlay,
        autoStartPauseUntilReady: getResolvedConfig().mpv.pauseUntilOverlayReady,
        texthookerEnabled: getResolvedConfig().texthooker.launchAtStartup,
        aniskipEnabled: getResolvedConfig().mpv.aniskipEnabled,
        aniskipButtonKey: getResolvedConfig().mpv.aniskipButtonKey,
    }),
    waitForYoutubeMpvConnected: (timeoutMs) => waitForYoutubeMpvConnected(timeoutMs),
    prepareYoutubePlaybackInMpv: (request) => prepareYoutubePlaybackInMpv(request),
    runYoutubePlaybackFlow: (request) => youtubeFlowRuntime.runYoutubePlaybackFlow(request),
    logInfo: (message) => logger.info(message),
    logWarn: (message) => logger.warn(message),
    schedule: (callback, delayMs) => setTimeout(callback, delayMs),
    clearScheduled: (timer) => clearTimeout(timer),
});
let firstRunSetupMessage = null;
const resolveWindowsMpvShortcutRuntimePaths = () => (0, windows_mpv_shortcuts_1.resolveWindowsMpvShortcutPaths)({
    appDataDir: electron_1.app.getPath('appData'),
    desktopDir: electron_1.app.getPath('desktop'),
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
const firstRunSetupService = (0, first_run_setup_service_1.createFirstRunSetupService)({
    platform: process.platform,
    configDir: CONFIG_DIR,
    getYomitanDictionaryCount: async () => {
        await ensureYomitanExtensionLoaded();
        const dictionaries = await (0, services_1.getYomitanDictionaryInfo)(getYomitanParserRuntimeDeps(), {
            error: (message, ...args) => logger.error(message, ...args),
            info: (message, ...args) => logger.info(message, ...args),
        });
        return dictionaries.length;
    },
    isExternalYomitanConfigured: () => getResolvedConfig().yomitan.externalProfilePath.trim().length > 0,
    detectPluginInstalled: () => {
        const candidates = (0, first_run_setup_plugin_1.detectInstalledFirstRunPluginCandidates)({
            platform: process.platform,
            homeDir: os.homedir(),
            xdgConfigHome: process.env.XDG_CONFIG_HOME,
            appDataDir: electron_1.app.getPath('appData'),
            mpvExecutablePath: getResolvedConfig().mpv.executablePath,
        });
        if (candidates.length > 0) {
            return true;
        }
        const installPaths = (0, setup_state_1.resolveDefaultMpvInstallPaths)(process.platform, os.homedir(), process.env.XDG_CONFIG_HOME);
        return (0, first_run_setup_plugin_1.detectInstalledFirstRunPlugin)(installPaths);
    },
    detectLegacyMpvPluginCandidates: () => (0, first_run_setup_plugin_1.detectInstalledFirstRunPluginCandidates)({
        platform: process.platform,
        homeDir: os.homedir(),
        xdgConfigHome: process.env.XDG_CONFIG_HOME,
        appDataDir: electron_1.app.getPath('appData'),
        mpvExecutablePath: getResolvedConfig().mpv.executablePath,
    }),
    removeLegacyMpvPlugins: (candidates) => (0, first_run_setup_plugin_1.removeLegacyMpvPluginCandidates)({
        candidates,
        trashItem: (candidatePath) => electron_1.shell.trashItem(candidatePath),
    }),
    detectWindowsMpvShortcuts: () => {
        if (process.platform !== 'win32') {
            return {
                startMenuInstalled: false,
                desktopInstalled: false,
            };
        }
        return (0, windows_mpv_shortcuts_1.detectWindowsMpvShortcuts)(resolveWindowsMpvShortcutRuntimePaths());
    },
    applyWindowsMpvShortcuts: async (preferences) => {
        if (process.platform !== 'win32') {
            return {
                ok: true,
                status: 'unknown',
                message: '',
            };
        }
        return (0, windows_mpv_shortcuts_1.applyWindowsMpvShortcuts)({
            preferences,
            paths: resolveWindowsMpvShortcutRuntimePaths(),
            exePath: process.execPath,
            writeShortcutLink: (shortcutPath, operation, details) => electron_1.shell.writeShortcutLink(shortcutPath, operation, details),
        });
    },
    detectCommandLineLauncher: () => (0, command_line_launcher_1.detectCommandLineLauncher)(createCommandLineLauncherRuntimeOptions()),
    installBun: async () => {
        const snapshot = await (0, command_line_launcher_1.installBun)(createCommandLineLauncherRuntimeOptions());
        return {
            ok: snapshot.status === 'ready',
            message: snapshot.message ??
                (snapshot.status === 'ready'
                    ? 'Bun is ready. Open a new terminal.'
                    : 'Bun installation failed.'),
        };
    },
    installCommandLineLauncher: async () => {
        const snapshot = await (0, command_line_launcher_1.installLauncher)(createCommandLineLauncherRuntimeOptions());
        const ok = snapshot.status === 'ready' || snapshot.status === 'installed_bun_missing';
        return {
            ok,
            installPath: snapshot.installPath,
            message: snapshot.message ??
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
let discordPresenceMediaDurationSec = null;
const discordPresenceRuntime = (0, discord_presence_runtime_1.createDiscordPresenceRuntime)({
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
async function initializeDiscordPresenceService() {
    if (getResolvedConfig().discordPresence.enabled !== true) {
        appState.discordPresenceService = null;
        return;
    }
    appState.discordPresenceService = (0, services_1.createDiscordPresenceService)({
        config: getResolvedConfig().discordPresence,
        createClient: () => (0, discord_rpc_client_js_1.createDiscordRpcClient)(DISCORD_PRESENCE_APP_ID),
        logDebug: (message, meta) => logger.debug(message, meta),
    });
    await appState.discordPresenceService.start();
    discordPresenceRuntime.publishDiscordPresence();
}
const ensureOverlayMpvSubtitlesHidden = (0, overlay_mpv_sub_visibility_1.createEnsureOverlayMpvSubtitlesHiddenHandler)({
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
        (0, services_1.setMpvSubVisibilityRuntime)(appState.mpvClient, visible);
    },
    logWarn: (message, error) => {
        logger.warn(message, error);
    },
});
const restoreOverlayMpvSubtitles = (0, overlay_mpv_sub_visibility_1.createRestoreOverlayMpvSubtitlesHandler)({
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
        (0, services_1.setMpvSubVisibilityRuntime)(appState.mpvClient, visible);
    },
});
function shouldSuppressMpvSubtitlesForOverlay() {
    return overlayManager.getVisibleOverlayVisible();
}
function syncOverlayMpvSubtitleSuppression() {
    if (shouldSuppressMpvSubtitlesForOverlay()) {
        void ensureOverlayMpvSubtitlesHidden();
        return;
    }
    restoreOverlayMpvSubtitles();
}
const buildImmersionMediaRuntimeMainDepsHandler = (0, startup_1.createBuildImmersionMediaRuntimeMainDepsHandler)({
    getResolvedConfig: () => getResolvedConfig(),
    defaultImmersionDbPath: DEFAULT_IMMERSION_DB_PATH,
    getTracker: () => appState.immersionTracker,
    getMpvClient: () => appState.mpvClient,
    getCurrentMediaPath: () => appState.currentMediaPath,
    getCurrentMediaTitle: () => appState.currentMediaTitle,
    logDebug: (message) => logger.debug(message),
    logInfo: (message) => logger.info(message),
});
const buildAnilistStateRuntimeMainDepsHandler = (0, startup_1.createBuildAnilistStateRuntimeMainDepsHandler)({
    getClientSecretState: () => appState.anilistClientSecretState,
    setClientSecretState: (next) => {
        appState.anilistClientSecretState = (0, state_1.transitionAnilistClientSecretState)(appState.anilistClientSecretState, next);
    },
    getRetryQueueState: () => appState.anilistRetryQueueState,
    setRetryQueueState: (next) => {
        appState.anilistRetryQueueState = (0, state_1.transitionAnilistRetryQueueState)(appState.anilistRetryQueueState, next);
    },
    getUpdateQueueSnapshot: () => anilistUpdateQueue.getSnapshot(),
    clearStoredToken: () => anilistTokenStore.clearToken(),
    clearCachedAccessToken: () => {
        anilistCachedAccessToken = null;
    },
});
const buildConfigDerivedRuntimeMainDepsHandler = (0, startup_1.createBuildConfigDerivedRuntimeMainDepsHandler)({
    getResolvedConfig: () => getResolvedConfig(),
    getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
    defaultJimakuLanguagePreference: config_2.DEFAULT_CONFIG.jimaku.languagePreference,
    defaultJimakuMaxEntryResults: config_2.DEFAULT_CONFIG.jimaku.maxEntryResults,
    defaultJimakuApiBaseUrl: config_2.DEFAULT_CONFIG.jimaku.apiBaseUrl,
});
const buildMainSubsyncRuntimeMainDepsHandler = (0, startup_1.createBuildMainSubsyncRuntimeMainDepsHandler)({
    getMpvClient: () => appState.mpvClient,
    getResolvedConfig: () => getResolvedConfig(),
    getSubsyncInProgress: () => appState.subsyncInProgress,
    setSubsyncInProgress: (inProgress) => {
        appState.subsyncInProgress = inProgress;
    },
    showMpvOsd: (text) => showMpvOsd(text),
    openManualPicker: (payload) => {
        openOverlayHostedModalWithOsd((deps) => (0, subsync_open_1.openSubsyncManualModal)(deps, payload), 'Subsync overlay unavailable.', 'Failed to open subsync overlay.');
    },
});
const immersionMediaRuntime = (0, startup_1.createImmersionMediaRuntime)(buildImmersionMediaRuntimeMainDepsHandler());
const anilistRateLimiter = (0, rate_limiter_1.createAnilistRateLimiter)();
const statsCoverArtFetcher = (0, cover_art_fetcher_1.createCoverArtFetcher)(anilistRateLimiter, (0, logger_1.createLogger)('main:stats-cover-art'));
const anilistStateRuntime = (0, anilist_1.createAnilistStateRuntime)(buildAnilistStateRuntimeMainDepsHandler());
const configDerivedRuntime = (0, startup_1.createConfigDerivedRuntime)(buildConfigDerivedRuntimeMainDepsHandler());
const subsyncRuntime = (0, startup_1.createMainSubsyncRuntime)(buildMainSubsyncRuntimeMainDepsHandler());
const currentMediaTokenizationGate = (0, current_media_tokenization_gate_1.createCurrentMediaTokenizationGate)();
const startupOsdSequencer = (0, startup_osd_sequencer_1.createStartupOsdSequencer)({
    showOsd: (message) => showMpvOsd(message),
});
const youtubePrimarySubtitleNotificationRuntime = (0, youtube_primary_subtitle_notification_1.createYoutubePrimarySubtitleNotificationRuntime)({
    getPrimarySubtitleLanguages: () => getResolvedConfig().youtube.primarySubLanguages,
    notifyFailure: (message) => reportYoutubeSubtitleFailure(message),
    schedule: (fn, delayMs) => setTimeout(fn, delayMs),
    clearSchedule: youtube_primary_subtitle_notification_1.clearYoutubePrimarySubtitleNotificationTimer,
});
function isYoutubePlaybackActiveNow() {
    return (0, youtube_playback_1.isYoutubePlaybackActive)(appState.currentMediaPath, appState.mpvClient?.currentVideoPath ?? null);
}
function reportYoutubeSubtitleFailure(message) {
    const type = getResolvedConfig().ankiConnect.behavior.notificationType;
    if (type === 'osd' || type === 'both') {
        showMpvOsd(message);
    }
    if (type === 'system' || type === 'both') {
        try {
            (0, utils_2.showDesktopNotification)('SubMiner', { body: message });
        }
        catch {
            logger.warn(`Unable to show desktop notification: ${message}`);
        }
    }
}
async function openYoutubeTrackPickerFromPlayback() {
    if (youtubeFlowRuntime.hasActiveSession()) {
        showMpvOsd('YouTube subtitle flow already in progress.');
        return;
    }
    const currentMediaPath = appState.currentMediaPath?.trim() || appState.mpvClient?.currentVideoPath?.trim() || '';
    if (!isYoutubePlaybackActiveNow() || !currentMediaPath) {
        showMpvOsd('YouTube subtitle picker is only available during YouTube playback.');
        return;
    }
    await youtubeFlowRuntime.openManualPicker({
        url: currentMediaPath,
    });
}
let appTray = null;
let tokenizeSubtitleDeferred = null;
function withCurrentSubtitleTiming(payload) {
    return {
        ...payload,
        startTime: appState.mpvClient?.currentSubStart ?? null,
        endTime: appState.mpvClient?.currentSubEnd ?? null,
    };
}
function emitSubtitlePayload(payload) {
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
    subtitlePrefetchService?.resume();
}
const buildSubtitleProcessingControllerMainDepsHandler = (0, startup_1.createBuildSubtitleProcessingControllerMainDepsHandler)({
    tokenizeSubtitle: async (text) => tokenizeSubtitleDeferred ? await tokenizeSubtitleDeferred(text) : { text, tokens: null },
    emitSubtitle: (payload) => emitSubtitlePayload(payload),
    logDebug: (message) => {
        logger.debug(`[subtitle-processing] ${message}`);
    },
    now: () => Date.now(),
});
const subtitleProcessingControllerMainDeps = buildSubtitleProcessingControllerMainDepsHandler();
const subtitleProcessingController = (0, services_1.createSubtitleProcessingController)(subtitleProcessingControllerMainDeps);
let subtitlePrefetchService = null;
let subtitlePrefetchRefreshTimer = null;
let lastObservedTimePos = 0;
let cancelLinuxMpvFullscreenOverlayRefreshBurst = null;
const SEEK_THRESHOLD_SECONDS = 3;
function clearScheduledSubtitlePrefetchRefresh() {
    if (subtitlePrefetchRefreshTimer) {
        clearTimeout(subtitlePrefetchRefreshTimer);
        subtitlePrefetchRefreshTimer = null;
    }
}
function cancelPendingLinuxMpvFullscreenOverlayRefreshBurst() {
    cancelLinuxMpvFullscreenOverlayRefreshBurst?.();
    cancelLinuxMpvFullscreenOverlayRefreshBurst = null;
}
const subtitlePrefetchInitController = (0, subtitle_prefetch_init_1.createSubtitlePrefetchInitController)({
    getCurrentService: () => subtitlePrefetchService,
    setCurrentService: (service) => {
        subtitlePrefetchService = service;
    },
    loadSubtitleSourceText,
    parseSubtitleCues: (content, filename) => (0, subtitle_cue_parser_1.parseSubtitleCues)(content, filename),
    createSubtitlePrefetchService: (deps) => (0, subtitle_prefetch_1.createSubtitlePrefetchService)(deps),
    tokenizeSubtitle: async (text) => tokenizeSubtitleDeferred ? await tokenizeSubtitleDeferred(text) : null,
    preCacheTokenization: (text, data) => {
        subtitleProcessingController.preCacheTokenization(text, data);
    },
    isCacheFull: () => subtitleProcessingController.isCacheFull(),
    logInfo: (message) => logger.info(message),
    logWarn: (message) => logger.warn(message),
    onParsedSubtitleCuesChanged: (cues, sourceKey) => {
        appState.activeParsedSubtitleCues = cues ?? [];
        appState.activeParsedSubtitleSource = sourceKey;
    },
});
const resolveActiveSubtitleSidebarSourceHandler = (0, subtitle_prefetch_runtime_1.createResolveActiveSubtitleSidebarSourceHandler)({
    getFfmpegPath: () => getResolvedConfig().subsync.ffmpeg_path.trim() || 'ffmpeg',
    extractInternalSubtitleTrack: (ffmpegPath, videoPath, track) => extractInternalSubtitleTrackToTempFile(ffmpegPath, videoPath, track),
});
async function refreshSubtitleSidebarFromSource(sourcePath) {
    const normalizedSourcePath = (0, subtitle_prefetch_source_1.resolveSubtitleSourcePath)(sourcePath.trim());
    if (!normalizedSourcePath) {
        return;
    }
    await subtitlePrefetchInitController.initSubtitlePrefetch(normalizedSourcePath, lastObservedTimePos, normalizedSourcePath);
}
const refreshSubtitlePrefetchFromActiveTrackHandler = (0, subtitle_prefetch_runtime_1.createRefreshSubtitlePrefetchFromActiveTrackHandler)({
    getMpvClient: () => appState.mpvClient,
    getLastObservedTimePos: () => lastObservedTimePos,
    subtitlePrefetchInitController,
    resolveActiveSubtitleSidebarSource: (input) => resolveActiveSubtitleSidebarSourceHandler(input),
});
function scheduleSubtitlePrefetchRefresh(delayMs = 0) {
    clearScheduledSubtitlePrefetchRefresh();
    subtitlePrefetchRefreshTimer = setTimeout(() => {
        subtitlePrefetchRefreshTimer = null;
        void refreshSubtitlePrefetchFromActiveTrackHandler();
    }, delayMs);
}
const subtitlePrefetchRuntime = {
    cancelPendingInit: () => subtitlePrefetchInitController.cancelPendingInit(),
    initSubtitlePrefetch: subtitlePrefetchInitController.initSubtitlePrefetch,
    refreshSubtitleSidebarFromSource: (sourcePath) => refreshSubtitleSidebarFromSource(sourcePath),
    refreshSubtitlePrefetchFromActiveTrack: () => refreshSubtitlePrefetchFromActiveTrackHandler(),
    scheduleSubtitlePrefetchRefresh: (delayMs) => scheduleSubtitlePrefetchRefresh(delayMs),
    clearScheduledSubtitlePrefetchRefresh: () => clearScheduledSubtitlePrefetchRefresh(),
};
const overlayShortcutsRuntime = (0, overlay_shortcuts_runtime_1.createOverlayShortcutsRuntimeService)((0, shortcuts_1.createBuildOverlayShortcutsRuntimeMainDepsHandler)({
    getConfiguredShortcuts: () => getConfiguredShortcuts(),
    getShortcutsRegistered: () => appState.shortcutsRegistered,
    setShortcutsRegistered: (registered) => {
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
    showMpvOsd: (text) => showMpvOsd(text),
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
    copySubtitleMultiple: (timeoutMs) => {
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
    mineSentenceMultiple: (timeoutMs) => {
        startPendingMineSentenceMultiple(timeoutMs);
    },
    cancelPendingMultiCopy: () => {
        cancelPendingMultiCopy();
    },
    cancelPendingMineSentenceMultiple: () => {
        cancelPendingMineSentenceMultiple();
    },
})());
syncOverlayShortcutsForModal = (isActive) => {
    if (isActive) {
        overlayShortcutsRuntime.unregisterOverlayShortcuts();
    }
    else {
        overlayShortcutsRuntime.syncOverlayShortcuts();
    }
};
const buildConfigHotReloadMessageMainDepsHandler = (0, overlay_1.createBuildConfigHotReloadMessageMainDepsHandler)({
    showMpvOsd: (message) => showMpvOsd(message),
    showDesktopNotification: (title, options) => (0, utils_2.showDesktopNotification)(title, options),
});
const configHotReloadMessageMainDeps = buildConfigHotReloadMessageMainDepsHandler();
const notifyConfigHotReloadMessage = (0, overlay_1.createConfigHotReloadMessageHandler)(configHotReloadMessageMainDeps);
const buildWatchConfigPathMainDepsHandler = (0, overlay_1.createBuildWatchConfigPathMainDepsHandler)({
    fileExists: (targetPath) => fs.existsSync(targetPath),
    dirname: (targetPath) => path.dirname(targetPath),
    watchPath: (targetPath, listener) => fs.watch(targetPath, listener),
});
const watchConfigPathHandler = (0, overlay_1.createWatchConfigPathHandler)(buildWatchConfigPathMainDepsHandler());
const buildConfigHotReloadAppliedMainDepsHandler = (0, overlay_1.createBuildConfigHotReloadAppliedMainDepsHandler)({
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
});
const applyConfigHotReloadDiff = (0, overlay_1.createConfigHotReloadAppliedHandler)(buildConfigHotReloadAppliedMainDepsHandler());
const buildConfigHotReloadRuntimeMainDepsHandler = (0, overlay_1.createBuildConfigHotReloadRuntimeMainDepsHandler)({
    getCurrentConfig: () => getResolvedConfig(),
    reloadConfigStrict: () => configService.reloadConfigStrict(),
    watchConfigPath: (configPath, onChange) => watchConfigPathHandler(configPath, onChange),
    setTimeout: (callback, delayMs) => setTimeout(callback, delayMs),
    clearTimeout: (timeout) => clearTimeout(timeout),
    debounceMs: 250,
    onHotReloadApplied: applyConfigHotReloadDiff,
    onRestartRequired: (fields) => notifyConfigHotReloadMessage((0, overlay_1.buildRestartRequiredConfigMessage)(fields)),
    onInvalidConfig: notifyConfigHotReloadMessage,
    onValidationWarnings: (configPath, warnings) => {
        (0, utils_2.showDesktopNotification)('SubMiner', {
            body: (0, config_validation_1.buildConfigWarningNotificationBody)(configPath, warnings),
        });
        if (process.platform === 'darwin') {
            electron_1.dialog.showErrorBox('SubMiner config validation warning', (0, config_validation_1.buildConfigWarningDialogDetails)(configPath, warnings));
        }
    },
});
const configHotReloadRuntime = (0, services_1.createConfigHotReloadRuntime)(buildConfigHotReloadRuntimeMainDepsHandler());
const configSettingsRuntime = (0, config_settings_runtime_1.createConfigSettingsRuntime)({
    fields: configSettingsFields,
    getConfigPath: () => configService.getConfigPath(),
    getRawConfig: () => configService.getRawConfig(),
    getConfig: () => configService.getConfig(),
    getWarnings: () => configService.getWarnings(),
    reloadConfigStrict: () => configService.reloadConfigStrict(),
    defaultAnkiConnectUrl: config_2.DEFAULT_CONFIG.ankiConnect.url,
    createAnkiClient: (url) => new anki_connect_1.AnkiConnectClient(url),
    getSettingsWindow: () => appState.configSettingsWindow,
    setSettingsWindow: (window) => {
        appState.configSettingsWindow = window;
    },
    createSettingsWindow: (0, setup_window_factory_1.createCreateConfigSettingsWindowHandler)({
        createBrowserWindow: (options) => new electron_1.BrowserWindow(options),
        preloadPath: path.join(__dirname, 'preload-settings.js'),
    }),
    settingsHtmlPath: path.join(__dirname, 'settings', 'index.html'),
    openPath: (targetPath) => electron_1.shell.openPath(targetPath),
    ipcMain: electron_1.ipcMain,
    ipcChannels: contracts_1.IPC_CHANNELS.request,
    log: (message) => logger.error(message),
});
configSettingsRuntime.registerHandlers();
const openConfigSettingsWindow = () => configSettingsRuntime.openWindow();
const buildDictionaryRootsHandler = (0, startup_1.createBuildDictionaryRootsMainHandler)({
    platform: process.platform,
    dirname: __dirname,
    appPath: electron_1.app.getAppPath(),
    resourcesPath: process.resourcesPath,
    userDataPath: USER_DATA_PATH,
    appUserDataPath: electron_1.app.getPath('userData'),
    homeDir: os.homedir(),
    appDataDir: process.env.APPDATA,
    cwd: process.cwd(),
    joinPath: (...parts) => path.join(...parts),
});
const buildFrequencyDictionaryRootsHandler = (0, startup_1.createBuildFrequencyDictionaryRootsMainHandler)({
    platform: process.platform,
    dirname: __dirname,
    appPath: electron_1.app.getAppPath(),
    resourcesPath: process.resourcesPath,
    userDataPath: USER_DATA_PATH,
    appUserDataPath: electron_1.app.getPath('userData'),
    homeDir: os.homedir(),
    appDataDir: process.env.APPDATA,
    cwd: process.cwd(),
    joinPath: (...parts) => path.join(...parts),
});
const jlptDictionaryRuntime = (0, jlpt_runtime_1.createJlptDictionaryRuntimeService)((0, startup_1.createBuildJlptDictionaryRuntimeMainDepsHandler)({
    isJlptEnabled: () => getResolvedConfig().subtitleStyle.enableJlpt,
    getDictionaryRoots: () => buildDictionaryRootsHandler(),
    getJlptDictionarySearchPaths: jlpt_runtime_1.getJlptDictionarySearchPaths,
    setJlptLevelLookup: (lookup) => {
        appState.jlptLevelLookup = lookup;
    },
    logInfo: (message) => logger.info(message),
})());
const frequencyDictionaryRuntime = (0, frequency_dictionary_runtime_1.createFrequencyDictionaryRuntimeService)((0, startup_1.createBuildFrequencyDictionaryRuntimeMainDepsHandler)({
    isFrequencyDictionaryEnabled: () => getResolvedConfig().subtitleStyle.frequencyDictionary.enabled,
    getDictionaryRoots: () => buildFrequencyDictionaryRootsHandler(),
    getFrequencyDictionarySearchPaths: frequency_dictionary_runtime_1.getFrequencyDictionarySearchPaths,
    getSourcePath: () => getResolvedConfig().subtitleStyle.frequencyDictionary.sourcePath,
    setFrequencyRankLookup: (lookup) => {
        appState.frequencyRankLookup = lookup;
    },
    logInfo: (message) => logger.info(message),
})());
const buildGetFieldGroupingResolverMainDepsHandler = (0, overlay_1.createBuildGetFieldGroupingResolverMainDepsHandler)({
    getResolver: () => appState.fieldGroupingResolver,
});
const getFieldGroupingResolverMainDeps = buildGetFieldGroupingResolverMainDepsHandler();
const getFieldGroupingResolverHandler = (0, overlay_1.createGetFieldGroupingResolverHandler)(getFieldGroupingResolverMainDeps);
function getFieldGroupingResolver() {
    return getFieldGroupingResolverHandler();
}
const buildSetFieldGroupingResolverMainDepsHandler = (0, overlay_1.createBuildSetFieldGroupingResolverMainDepsHandler)({
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
const setFieldGroupingResolverHandler = (0, overlay_1.createSetFieldGroupingResolverHandler)(setFieldGroupingResolverMainDeps);
function setFieldGroupingResolver(resolver) {
    setFieldGroupingResolverHandler(resolver);
}
const fieldGroupingOverlayRuntime = (0, services_1.createFieldGroupingOverlayRuntime)((0, overlay_1.createBuildFieldGroupingOverlayMainDepsHandler)({
    getMainWindow: () => overlayManager.getMainWindow(),
    getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    setVisibleOverlayVisible: (visible) => setVisibleOverlayVisible(visible),
    getResolver: () => getFieldGroupingResolver(),
    setResolver: (resolver) => setFieldGroupingResolver(resolver),
    getRestoreVisibleOverlayOnModalClose: () => overlayModalRuntime.getRestoreVisibleOverlayOnModalClose(),
    sendToActiveOverlayWindow: (channel, payload, runtimeOptions) => overlayModalRuntime.sendToActiveOverlayWindow(channel, payload, runtimeOptions),
})());
const createFieldGroupingCallback = fieldGroupingOverlayRuntime.createFieldGroupingCallback;
const SUBTITLE_POSITIONS_DIR = path.join(CONFIG_DIR, 'subtitle-positions');
const mediaRuntime = (0, media_runtime_1.createMediaRuntimeService)((0, startup_1.createBuildMediaRuntimeMainDepsHandler)({
    isRemoteMediaPath: (mediaPath) => (0, utils_1.isRemoteMediaPath)(mediaPath),
    loadSubtitlePosition: () => loadSubtitlePosition(),
    getCurrentMediaPath: () => appState.currentMediaPath,
    getPendingSubtitlePosition: () => appState.pendingSubtitlePosition,
    getSubtitlePositionsDir: () => SUBTITLE_POSITIONS_DIR,
    setCurrentMediaPath: (nextPath) => {
        appState.currentMediaPath = nextPath;
    },
    clearPendingSubtitlePosition: () => {
        appState.pendingSubtitlePosition = null;
    },
    setSubtitlePosition: (position) => {
        appState.subtitlePosition = position;
    },
    broadcastToOverlayWindows: (channel, payload) => {
        broadcastToOverlayWindows(channel, payload);
    },
    getCurrentMediaTitle: () => appState.currentMediaTitle,
    setCurrentMediaTitle: (title) => {
        appState.currentMediaTitle = title;
    },
})());
const characterDictionaryRuntime = (0, character_dictionary_runtime_1.createCharacterDictionaryRuntimeService)({
    userDataPath: USER_DATA_PATH,
    getCurrentMediaPath: () => appState.currentMediaPath,
    getCurrentMediaTitle: () => appState.currentMediaTitle,
    resolveMediaPathForJimaku: (mediaPath) => mediaRuntime.resolveMediaPathForJimaku(mediaPath),
    guessAnilistMediaInfo: (mediaPath, mediaTitle) => (0, anilist_updater_1.guessAnilistMediaInfo)(mediaPath, mediaTitle),
    getCollapsibleSectionOpenState: (section) => getResolvedConfig().anilist.characterDictionary.collapsibleSections[section],
    now: () => Date.now(),
    logInfo: (message) => logger.info(message),
    logWarn: (message) => logger.warn(message),
});
const characterDictionaryAutoSyncRuntime = (0, character_dictionary_auto_sync_1.createCharacterDictionaryAutoSyncRuntimeService)({
    userDataPath: USER_DATA_PATH,
    getConfig: () => {
        const config = getResolvedConfig().anilist.characterDictionary;
        return {
            enabled: config.enabled &&
                yomitanProfilePolicy.isCharacterDictionaryEnabled() &&
                !isYoutubePlaybackActiveNow(),
            maxLoaded: config.maxLoaded,
            profileScope: config.profileScope,
        };
    },
    getOrCreateCurrentSnapshot: (targetPath, progress) => characterDictionaryRuntime.getOrCreateCurrentSnapshot(targetPath, progress),
    buildMergedDictionary: (mediaIds) => characterDictionaryRuntime.buildMergedDictionary(mediaIds),
    waitForYomitanMutationReady: () => currentMediaTokenizationGate.waitUntilReady(appState.currentMediaPath?.trim() || appState.mpvClient?.currentVideoPath?.trim() || null),
    getYomitanDictionaryInfo: async () => {
        await ensureYomitanExtensionLoaded();
        return await (0, services_1.getYomitanDictionaryInfo)(getYomitanParserRuntimeDeps(), {
            error: (message, ...args) => logger.error(message, ...args),
            info: (message, ...args) => logger.info(message, ...args),
        });
    },
    importYomitanDictionary: async (zipPath) => {
        if (yomitanProfilePolicy.isExternalReadOnlyMode()) {
            yomitanProfilePolicy.logSkippedWrite((0, yomitan_read_only_log_1.formatSkippedYomitanWriteAction)('importYomitanDictionary', zipPath));
            return false;
        }
        await ensureYomitanExtensionLoaded();
        return await (0, services_1.importYomitanDictionaryFromZip)(zipPath, getYomitanParserRuntimeDeps(), {
            error: (message, ...args) => logger.error(message, ...args),
            info: (message, ...args) => logger.info(message, ...args),
        });
    },
    deleteYomitanDictionary: async (dictionaryTitle) => {
        if (yomitanProfilePolicy.isExternalReadOnlyMode()) {
            yomitanProfilePolicy.logSkippedWrite((0, yomitan_read_only_log_1.formatSkippedYomitanWriteAction)('deleteYomitanDictionary', dictionaryTitle));
            return false;
        }
        await ensureYomitanExtensionLoaded();
        return await (0, services_1.deleteYomitanDictionaryByTitle)(dictionaryTitle, getYomitanParserRuntimeDeps(), {
            error: (message, ...args) => logger.error(message, ...args),
            info: (message, ...args) => logger.info(message, ...args),
        });
    },
    upsertYomitanDictionarySettings: async (dictionaryTitle, profileScope) => {
        if (yomitanProfilePolicy.isExternalReadOnlyMode()) {
            yomitanProfilePolicy.logSkippedWrite((0, yomitan_read_only_log_1.formatSkippedYomitanWriteAction)('upsertYomitanDictionarySettings', dictionaryTitle));
            return false;
        }
        await ensureYomitanExtensionLoaded();
        return await (0, services_1.upsertYomitanDictionarySettings)(dictionaryTitle, profileScope, getYomitanParserRuntimeDeps(), {
            error: (message, ...args) => logger.error(message, ...args),
            info: (message, ...args) => logger.info(message, ...args),
        });
    },
    now: () => Date.now(),
    schedule: (fn, delayMs) => setTimeout(fn, delayMs),
    clearSchedule: (timer) => clearTimeout(timer),
    logInfo: (message) => logger.info(message),
    logWarn: (message) => logger.warn(message),
    onSyncStatus: (event) => {
        (0, character_dictionary_auto_sync_notifications_1.notifyCharacterDictionaryAutoSyncStatus)(event, {
            getNotificationType: () => getResolvedConfig().ankiConnect.behavior.notificationType,
            showOsd: (message) => showMpvOsd(message),
            showDesktopNotification: (title, options) => (0, utils_2.showDesktopNotification)(title, options),
            startupOsdSequencer,
        });
    },
    onSyncComplete: ({ mediaId, mediaTitle, changed }) => {
        (0, character_dictionary_auto_sync_completion_1.handleCharacterDictionaryAutoSyncComplete)({
            mediaId,
            mediaTitle,
            changed,
        }, {
            hasParserWindow: () => Boolean(appState.yomitanParserWindow),
            clearParserCaches: () => {
                if (appState.yomitanParserWindow) {
                    (0, services_1.clearYomitanParserCachesForWindow)(appState.yomitanParserWindow);
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
        });
    },
});
const overlayVisibilityRuntime = (0, overlay_visibility_runtime_1.createOverlayVisibilityRuntimeService)((0, overlay_1.createBuildOverlayVisibilityRuntimeMainDepsHandler)({
    getMainWindow: () => overlayManager.getMainWindow(),
    getModalActive: () => overlayModalInputState.getModalInputExclusive(),
    getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    getForceMousePassthrough: () => appState.statsOverlayVisible,
    getOverlayInteractionActive: () => visibleOverlayInteractionActive,
    getWindowTracker: () => appState.windowTracker,
    getLastKnownWindowsForegroundProcessName: () => lastWindowsVisibleOverlayForegroundProcessName,
    getWindowsOverlayProcessName: () => path.parse(process.execPath).name.toLowerCase(),
    getWindowsFocusHandoffGraceActive: () => hasWindowsVisibleOverlayFocusHandoffGrace(),
    getTrackerNotReadyWarningShown: () => appState.trackerNotReadyWarningShown,
    setTrackerNotReadyWarningShown: (shown) => {
        appState.trackerNotReadyWarningShown = shown;
    },
    updateVisibleOverlayBounds: (geometry) => updateVisibleOverlayBounds(geometry),
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
    showOverlayLoadingOsd: (message) => {
        showMpvOsd(message);
    },
    resolveFallbackBounds: () => {
        const cursorPoint = electron_1.screen.getCursorScreenPoint();
        const display = electron_1.screen.getDisplayNearestPoint(cursorPoint);
        const fallbackBounds = display.workArea;
        return {
            x: fallbackBounds.x,
            y: fallbackBounds.y,
            width: fallbackBounds.width,
            height: fallbackBounds.height,
        };
    },
})());
const VISIBLE_OVERLAY_BLUR_REFRESH_DELAYS_MS = [0, 25, 100, 250];
const WINDOWS_VISIBLE_OVERLAY_Z_ORDER_RETRY_DELAYS_MS = [0, 48, 120, 240, 480];
const WINDOWS_VISIBLE_OVERLAY_FOREGROUND_POLL_INTERVAL_MS = 75;
const WINDOWS_VISIBLE_OVERLAY_FOCUS_HANDOFF_GRACE_MS = 200;
let visibleOverlayBlurRefreshTimeouts = [];
let windowsVisibleOverlayZOrderRetryTimeouts = [];
let windowsVisibleOverlayZOrderSyncInFlight = false;
let windowsVisibleOverlayZOrderSyncQueued = false;
let windowsVisibleOverlayForegroundPollInterval = null;
let lastWindowsVisibleOverlayForegroundProcessName = null;
let lastWindowsVisibleOverlayBlurredAtMs = 0;
let visibleOverlayInteractionActive = false;
function clearVisibleOverlayBlurRefreshTimeouts() {
    for (const timeout of visibleOverlayBlurRefreshTimeouts) {
        clearTimeout(timeout);
    }
    visibleOverlayBlurRefreshTimeouts = [];
}
function clearWindowsVisibleOverlayZOrderRetryTimeouts() {
    for (const timeout of windowsVisibleOverlayZOrderRetryTimeouts) {
        clearTimeout(timeout);
    }
    windowsVisibleOverlayZOrderRetryTimeouts = [];
}
function getWindowsNativeWindowHandle(window) {
    const handle = window.getNativeWindowHandle();
    return handle.length >= 8
        ? handle.readBigUInt64LE(0).toString()
        : BigInt(handle.readUInt32LE(0)).toString();
}
function getWindowsNativeWindowHandleNumber(window) {
    const handle = window.getNativeWindowHandle();
    return handle.length >= 8 ? Number(handle.readBigUInt64LE(0)) : handle.readUInt32LE(0);
}
function resolveWindowsOverlayBindTargetHandle(targetMpvSocketPath) {
    if (process.platform !== 'win32') {
        return null;
    }
    try {
        if (targetMpvSocketPath) {
            const windowTracker = appState.windowTracker;
            const trackedHandle = windowTracker?.getTargetWindowHandle?.();
            if (typeof trackedHandle === 'number' && Number.isFinite(trackedHandle)) {
                return trackedHandle;
            }
        }
        return (0, windows_helper_1.findWindowsMpvTargetWindowHandle)();
    }
    catch {
        return null;
    }
}
async function syncWindowsVisibleOverlayToMpvZOrder() {
    if (process.platform !== 'win32') {
        return false;
    }
    const mainWindow = overlayManager.getMainWindow();
    if (!mainWindow ||
        mainWindow.isDestroyed() ||
        !mainWindow.isVisible() ||
        !overlayManager.getVisibleOverlayVisible()) {
        return false;
    }
    const windowTracker = appState.windowTracker;
    if (!windowTracker) {
        return false;
    }
    if (typeof windowTracker.isTargetWindowMinimized === 'function' &&
        windowTracker.isTargetWindowMinimized()) {
        return false;
    }
    if (!windowTracker.isTracking() && windowTracker.getGeometry() === null) {
        return false;
    }
    const overlayHwnd = getWindowsNativeWindowHandleNumber(mainWindow);
    const targetWindowHwnd = resolveWindowsOverlayBindTargetHandle(appState.mpvSocketPath);
    if (targetWindowHwnd !== null && (0, windows_helper_1.bindWindowsOverlayAboveMpv)(overlayHwnd, targetWindowHwnd)) {
        mainWindow.setOpacity?.(1);
        return true;
    }
    return false;
}
function requestWindowsVisibleOverlayZOrderSync() {
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
function scheduleWindowsVisibleOverlayZOrderSyncBurst() {
    if (process.platform !== 'win32') {
        return;
    }
    clearWindowsVisibleOverlayZOrderRetryTimeouts();
    for (const delayMs of WINDOWS_VISIBLE_OVERLAY_Z_ORDER_RETRY_DELAYS_MS) {
        const retryTimeout = setTimeout(() => {
            windowsVisibleOverlayZOrderRetryTimeouts = windowsVisibleOverlayZOrderRetryTimeouts.filter((timeout) => timeout !== retryTimeout);
            requestWindowsVisibleOverlayZOrderSync();
        }, delayMs);
        windowsVisibleOverlayZOrderRetryTimeouts.push(retryTimeout);
    }
}
function hasWindowsVisibleOverlayFocusHandoffGrace() {
    return (process.platform === 'win32' &&
        lastWindowsVisibleOverlayBlurredAtMs > 0 &&
        Date.now() - lastWindowsVisibleOverlayBlurredAtMs <=
            WINDOWS_VISIBLE_OVERLAY_FOCUS_HANDOFF_GRACE_MS);
}
function shouldPollWindowsVisibleOverlayForegroundProcess() {
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
    if (typeof windowTracker.isTargetWindowMinimized === 'function' &&
        windowTracker.isTargetWindowMinimized()) {
        return false;
    }
    const overlayFocused = mainWindow.isFocused();
    const trackerFocused = windowTracker.isTargetWindowFocused?.() ?? false;
    return !overlayFocused && !trackerFocused;
}
function maybePollWindowsVisibleOverlayForegroundProcess() {
    if (!shouldPollWindowsVisibleOverlayForegroundProcess()) {
        lastWindowsVisibleOverlayForegroundProcessName = null;
        return;
    }
    const processName = (0, windows_helper_1.getWindowsForegroundProcessName)();
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
function ensureWindowsVisibleOverlayForegroundPollLoop() {
    if (process.platform !== 'win32' || windowsVisibleOverlayForegroundPollInterval !== null) {
        return;
    }
    windowsVisibleOverlayForegroundPollInterval = setInterval(() => {
        maybePollWindowsVisibleOverlayForegroundProcess();
    }, WINDOWS_VISIBLE_OVERLAY_FOREGROUND_POLL_INTERVAL_MS);
}
function clearWindowsVisibleOverlayForegroundPollLoop() {
    if (windowsVisibleOverlayForegroundPollInterval === null) {
        return;
    }
    clearInterval(windowsVisibleOverlayForegroundPollInterval);
    windowsVisibleOverlayForegroundPollInterval = null;
}
function scheduleVisibleOverlayBlurRefresh() {
    if (process.platform !== 'win32' && process.platform !== 'darwin') {
        return;
    }
    if (process.platform === 'win32') {
        lastWindowsVisibleOverlayBlurredAtMs = Date.now();
    }
    clearVisibleOverlayBlurRefreshTimeouts();
    for (const delayMs of VISIBLE_OVERLAY_BLUR_REFRESH_DELAYS_MS) {
        const refreshTimeout = setTimeout(() => {
            visibleOverlayBlurRefreshTimeouts = visibleOverlayBlurRefreshTimeouts.filter((timeout) => timeout !== refreshTimeout);
            overlayVisibilityRuntime.updateVisibleOverlayVisibility();
        }, delayMs);
        visibleOverlayBlurRefreshTimeouts.push(refreshTimeout);
    }
}
ensureWindowsVisibleOverlayForegroundPollLoop();
const buildGetRuntimeOptionsStateMainDepsHandler = (0, overlay_1.createBuildGetRuntimeOptionsStateMainDepsHandler)({
    getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
});
const getRuntimeOptionsStateMainDeps = buildGetRuntimeOptionsStateMainDepsHandler();
const getRuntimeOptionsStateHandler = (0, overlay_1.createGetRuntimeOptionsStateHandler)(getRuntimeOptionsStateMainDeps);
function getRuntimeOptionsState() {
    return getRuntimeOptionsStateHandler();
}
function getOverlayWindows() {
    return overlayManager.getOverlayWindows();
}
const buildRestorePreviousSecondarySubVisibilityMainDepsHandler = (0, overlay_1.createBuildRestorePreviousSecondarySubVisibilityMainDepsHandler)({
    getMpvClient: () => appState.mpvClient,
});
syncOverlayVisibilityForModal = () => {
    overlayVisibilityRuntime.updateVisibleOverlayVisibility();
};
function broadcastToOverlayWindows(channel, ...args) {
    overlayManager.broadcastToOverlayWindows(channel, ...args);
}
const buildBroadcastRuntimeOptionsChangedMainDepsHandler = (0, overlay_1.createBuildBroadcastRuntimeOptionsChangedMainDepsHandler)({
    broadcastRuntimeOptionsChangedRuntime: services_1.broadcastRuntimeOptionsChangedRuntime,
    getRuntimeOptionsState: () => getRuntimeOptionsState(),
    broadcastToOverlayWindows: (channel, ...args) => broadcastToOverlayWindows(channel, ...args),
});
const buildSendToActiveOverlayWindowMainDepsHandler = (0, overlay_1.createBuildSendToActiveOverlayWindowMainDepsHandler)({
    sendToActiveOverlayWindowRuntime: (channel, payload, runtimeOptions) => overlayModalRuntime.sendToActiveOverlayWindow(channel, payload, runtimeOptions),
});
const buildSetOverlayDebugVisualizationEnabledMainDepsHandler = (0, overlay_1.createBuildSetOverlayDebugVisualizationEnabledMainDepsHandler)({
    setOverlayDebugVisualizationEnabledRuntime: services_1.setOverlayDebugVisualizationEnabledRuntime,
    getCurrentEnabled: () => appState.overlayDebugVisualizationEnabled,
    setCurrentEnabled: (next) => {
        appState.overlayDebugVisualizationEnabled = next;
    },
});
const buildOpenRuntimeOptionsPaletteMainDepsHandler = (0, overlay_1.createBuildOpenRuntimeOptionsPaletteMainDepsHandler)({
    openRuntimeOptionsPaletteRuntime: () => overlayModalRuntime.openRuntimeOptionsPalette(),
});
const overlayVisibilityComposer = (0, composers_1.composeOverlayVisibilityRuntime)({
    overlayVisibilityRuntime,
    restorePreviousSecondarySubVisibilityMainDeps: buildRestorePreviousSecondarySubVisibilityMainDepsHandler(),
    broadcastRuntimeOptionsChangedMainDeps: buildBroadcastRuntimeOptionsChangedMainDepsHandler(),
    sendToActiveOverlayWindowMainDeps: buildSendToActiveOverlayWindowMainDepsHandler(),
    setOverlayDebugVisualizationEnabledMainDeps: buildSetOverlayDebugVisualizationEnabledMainDepsHandler(),
    openRuntimeOptionsPaletteMainDeps: buildOpenRuntimeOptionsPaletteMainDepsHandler(),
});
function restorePreviousSecondarySubVisibility() {
    overlayVisibilityComposer.restorePreviousSecondarySubVisibility();
}
function broadcastRuntimeOptionsChanged() {
    overlayVisibilityComposer.broadcastRuntimeOptionsChanged();
}
function sendToActiveOverlayWindow(channel, payload, runtimeOptions) {
    return overlayVisibilityComposer.sendToActiveOverlayWindow(channel, payload, runtimeOptions);
}
function setOverlayDebugVisualizationEnabled(enabled) {
    overlayVisibilityComposer.setOverlayDebugVisualizationEnabled(enabled);
}
function createOverlayHostedModalOpenDeps() {
    return {
        ensureOverlayStartupPrereqs: () => ensureOverlayStartupPrereqs(),
        ensureOverlayWindowsReadyForVisibilityActions: () => ensureOverlayWindowsReadyForVisibilityActions(),
        sendToActiveOverlayWindow: (channel, payload, runtimeOptions) => sendToActiveOverlayWindow(channel, payload, runtimeOptions),
        waitForModalOpen: (modal, timeoutMs) => overlayModalRuntime.waitForModalOpen(modal, timeoutMs),
        logWarn: (message) => logger.warn(message),
    };
}
function openOverlayHostedModalWithOsd(openModal, unavailableMessage, failureLogMessage) {
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
function openRuntimeOptionsPalette() {
    openOverlayHostedModalWithOsd(runtime_options_open_1.openRuntimeOptionsModal, 'Runtime options overlay unavailable.', 'Failed to open runtime options overlay.');
}
function openJimakuOverlay() {
    openOverlayHostedModalWithOsd(jimaku_open_1.openJimakuModal, 'Jimaku overlay unavailable.', 'Failed to open Jimaku overlay.');
}
function openSessionHelpOverlay() {
    openOverlayHostedModalWithOsd(session_help_open_1.openSessionHelpModal, 'Session help overlay unavailable.', 'Failed to open session help overlay.');
}
function openCharacterDictionaryOverlay() {
    openOverlayHostedModalWithOsd(character_dictionary_open_1.openCharacterDictionaryModal, 'Character dictionary overlay unavailable.', 'Failed to open character dictionary overlay.');
}
function openControllerSelectOverlay() {
    openOverlayHostedModalWithOsd(controller_select_open_1.openControllerSelectModal, 'Controller select overlay unavailable.', 'Failed to open controller select overlay.');
}
function openControllerDebugOverlay() {
    openOverlayHostedModalWithOsd(controller_debug_open_1.openControllerDebugModal, 'Controller debug overlay unavailable.', 'Failed to open controller debug overlay.');
}
function openPlaylistBrowser() {
    if (!appState.mpvClient?.connected) {
        showMpvOsd('Playlist browser requires active playback.');
        return;
    }
    openOverlayHostedModalWithOsd(playlist_browser_open_1.openPlaylistBrowser, 'Playlist browser overlay unavailable.', 'Failed to open playlist browser overlay.');
}
function getResolvedConfig() {
    return configService.getConfig();
}
function getRuntimeBooleanOption(id, fallback) {
    const value = appState.runtimeOptionsManager?.getOptionValue(id);
    return typeof value === 'boolean' ? value : fallback;
}
function shouldInitializeMecabForAnnotations() {
    const config = getResolvedConfig();
    const knownWordsEnabled = getRuntimeBooleanOption('subtitle.annotation.knownWords.highlightEnabled', config.ankiConnect.knownWords.highlightEnabled);
    const nPlusOneEnabled = getRuntimeBooleanOption('subtitle.annotation.nPlusOne', config.ankiConnect.nPlusOne.enabled);
    const jlptEnabled = getRuntimeBooleanOption('subtitle.annotation.jlpt', config.subtitleStyle.enableJlpt);
    const frequencyEnabled = getRuntimeBooleanOption('subtitle.annotation.frequency', config.subtitleStyle.frequencyDictionary.enabled);
    return knownWordsEnabled || nPlusOneEnabled || jlptEnabled || frequencyEnabled;
}
const { getResolvedJellyfinConfig, reportJellyfinRemoteProgress, reportJellyfinRemoteStopped, startJellyfinRemoteSession, stopJellyfinRemoteSession, runJellyfinCommand, openJellyfinSetupWindow, getJellyfinClientInfo, } = (0, composers_1.composeJellyfinRuntimeHandlers)({
    getResolvedJellyfinConfigMainDeps: {
        getResolvedConfig: () => getResolvedConfig(),
        loadStoredSession: () => jellyfinTokenStore.loadSession(),
        getEnv: (name) => process.env[name],
    },
    getJellyfinClientInfoMainDeps: {
        getResolvedJellyfinConfig: () => getResolvedJellyfinConfig(),
        getDefaultJellyfinConfig: () => config_2.DEFAULT_CONFIG.jellyfin,
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
        getPluginRuntimeConfig: () => ({
            socketPath: appState.mpvSocketPath,
            binaryPath: getResolvedConfig().mpv.subminerBinaryPath,
            backend: getResolvedConfig().mpv.backend,
            autoStart: getResolvedConfig().mpv.autoStartSubMiner,
            autoStartVisibleOverlay: getResolvedConfig().auto_start_overlay,
            autoStartPauseUntilReady: getResolvedConfig().mpv.pauseUntilOverlayReady,
            texthookerEnabled: getResolvedConfig().texthooker.launchAtStartup,
            aniskipEnabled: getResolvedConfig().mpv.aniskipEnabled,
            aniskipButtonKey: getResolvedConfig().mpv.aniskipButtonKey,
        }),
        defaultMpvLogPath: DEFAULT_MPV_LOG_PATH,
        defaultMpvArgs: MPV_JELLYFIN_DEFAULT_ARGS,
        removeSocketPath: (socketPath) => {
            fs.rmSync(socketPath, { force: true });
        },
        spawnMpv: (args) => (0, node_child_process_1.spawn)('mpv', args, {
            detached: true,
            stdio: 'ignore',
        }),
        logWarn: (message, error) => logger.warn(message, error),
        logInfo: (message) => logger.info(message),
    },
    ensureMpvConnectedForJellyfinPlaybackMainDeps: {
        getMpvClient: () => appState.mpvClient,
        setMpvClient: (client) => {
            appState.mpvClient = client;
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
        listJellyfinSubtitleTracks: (session, clientInfo, itemId) => (0, services_1.listJellyfinSubtitleTracksRuntime)(session, clientInfo, itemId),
        getMpvClient: () => appState.mpvClient,
        sendMpvCommand: (command) => {
            (0, services_1.sendMpvCommandRuntime)(appState.mpvClient, command);
        },
        wait: (ms) => new Promise((resolve) => setTimeout(resolve, ms)),
        logDebug: (message, error) => {
            logger.debug(message, error);
        },
    },
    playJellyfinItemInMpvMainDeps: {
        getMpvClient: () => appState.mpvClient,
        resolvePlaybackPlan: (params) => (0, services_1.resolveJellyfinPlaybackPlanRuntime)(params.session, params.clientInfo, params.jellyfinConfig, {
            itemId: params.itemId,
            audioStreamIndex: params.audioStreamIndex ?? undefined,
            subtitleStreamIndex: params.subtitleStreamIndex ?? undefined,
        }),
        applyJellyfinMpvDefaults: (mpvClient) => applyJellyfinMpvDefaults(mpvClient),
        sendMpvCommand: (command) => (0, services_1.sendMpvCommandRuntime)(appState.mpvClient, command),
        armQuitOnDisconnect: () => {
            jellyfinPlayQuitOnDisconnectArmed = false;
            setTimeout(() => {
                jellyfinPlayQuitOnDisconnectArmed = true;
            }, 3000);
        },
        schedule: (callback, delayMs) => {
            setTimeout(callback, delayMs);
        },
        convertTicksToSeconds: (ticks) => (0, services_1.jellyfinTicksToSecondsRuntime)(ticks),
        setActivePlayback: (state) => {
            activeJellyfinRemotePlayback = state;
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
    },
    remoteComposerOptions: {
        getConfiguredSession: () => (0, jellyfin_1.getConfiguredJellyfinSession)(getResolvedJellyfinConfig()),
        logWarn: (message) => logger.warn(message),
        getMpvClient: () => appState.mpvClient,
        sendMpvCommand: (client, command) => (0, services_1.sendMpvCommandRuntime)(client, command),
        jellyfinTicksToSeconds: (ticks) => (0, services_1.jellyfinTicksToSecondsRuntime)(ticks),
        getActivePlayback: () => activeJellyfinRemotePlayback,
        clearActivePlayback: () => {
            activeJellyfinRemotePlayback = null;
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
        authenticateWithPassword: (serverUrl, username, password, clientInfo) => (0, services_1.authenticateWithPasswordRuntime)(serverUrl, username, password, clientInfo),
        saveStoredSession: (session) => jellyfinTokenStore.saveSession(session),
        clearStoredSession: () => (0, jellyfin_tray_discovery_1.clearJellyfinAuthSessionAndRefreshTray)(getJellyfinTrayDiscoveryDeps()),
        logInfo: (message) => logger.info(message),
    },
    handleJellyfinListCommandsMainDeps: {
        listJellyfinLibraries: (session, clientInfo) => (0, services_1.listJellyfinLibrariesRuntime)(session, clientInfo),
        listJellyfinItems: (session, clientInfo, params) => (0, services_1.listJellyfinItemsRuntime)(session, clientInfo, params),
        listJellyfinSubtitleTracks: (session, clientInfo, itemId) => (0, services_1.listJellyfinSubtitleTracksRuntime)(session, clientInfo, itemId),
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
            appState.jellyfinRemoteSession = session;
        },
        createRemoteSessionService: (options) => new services_1.JellyfinRemoteSessionService(options),
        defaultDeviceId: config_2.DEFAULT_CONFIG.jellyfin.deviceId,
        defaultClientName: config_2.DEFAULT_CONFIG.jellyfin.clientName,
        defaultClientVersion: config_2.DEFAULT_CONFIG.jellyfin.clientVersion,
        logInfo: (message) => logger.info(message),
        logWarn: (message, details) => logger.warn(message, details),
    },
    stopJellyfinRemoteSessionMainDeps: {
        getCurrentSession: () => appState.jellyfinRemoteSession,
        setCurrentSession: (session) => {
            appState.jellyfinRemoteSession = session;
        },
        clearActivePlayback: () => {
            activeJellyfinRemotePlayback = null;
        },
    },
    runJellyfinCommandMainDeps: {
        defaultServerUrl: config_2.DEFAULT_CONFIG.jellyfin.serverUrl,
    },
    maybeFocusExistingJellyfinSetupWindowMainDeps: {
        getSetupWindow: () => appState.jellyfinSetupWindow,
    },
    openJellyfinSetupWindowMainDeps: {
        createSetupWindow: (0, setup_window_factory_1.createCreateJellyfinSetupWindowHandler)({
            createBrowserWindow: (options) => new electron_1.BrowserWindow(options),
        }),
        buildSetupFormHtml: (state) => (0, jellyfin_1.buildJellyfinSetupFormHtml)(state),
        parseSubmissionUrl: (rawUrl) => (0, jellyfin_1.parseJellyfinSetupSubmissionUrl)(rawUrl),
        authenticateWithPassword: (server, username, password, clientInfo) => (0, services_1.authenticateWithPasswordRuntime)(server, username, password, clientInfo),
        saveStoredSession: (session) => jellyfinTokenStore.saveSession(session),
        clearStoredSession: () => (0, jellyfin_tray_discovery_1.clearJellyfinAuthSessionAndRefreshTray)(getJellyfinTrayDiscoveryDeps()),
        patchJellyfinConfig: (session) => {
            const clientInfo = getJellyfinClientInfo();
            const recentServers = (0, jellyfin_1.mergeJellyfinRecentServers)(session.serverUrl, getResolvedConfig().jellyfin.recentServers || []);
            configService.patchRawConfig({
                jellyfin: {
                    enabled: true,
                    serverUrl: session.serverUrl,
                    username: session.username,
                    deviceId: clientInfo.deviceId,
                    clientName: clientInfo.clientName,
                    clientVersion: clientInfo.clientVersion,
                    recentServers,
                },
            });
            refreshTrayMenuIfPresent();
        },
        persistAuthenticatedSession: (session, clientInfo) => (0, jellyfin_1.persistJellyfinAuthSession)({
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
            appState.jellyfinSetupWindow = window;
        },
        encodeURIComponent: (value) => encodeURIComponent(value),
        defaultServerUrl: config_2.DEFAULT_CONFIG.jellyfin.serverUrl || 'http://127.0.0.1:8096',
        hasStoredSession: () => Boolean(jellyfinTokenStore.loadSession()),
    },
});
const maybeFocusExistingFirstRunSetupWindow = (0, first_run_setup_window_1.createMaybeFocusExistingFirstRunSetupWindowHandler)({
    getSetupWindow: () => appState.firstRunSetupWindow,
});
const openFirstRunSetupWindowHandler = (0, first_run_setup_window_1.createOpenFirstRunSetupWindowHandler)({
    maybeFocusExistingSetupWindow: maybeFocusExistingFirstRunSetupWindow,
    createSetupWindow: (0, setup_window_factory_1.createCreateFirstRunSetupWindowHandler)({
        createBrowserWindow: (options) => new electron_1.BrowserWindow(options),
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
            mpvExecutablePathStatus: (0, windows_mpv_launch_1.getConfiguredWindowsMpvPathStatus)(mpvExecutablePath),
            windowsMpvShortcuts: snapshot.windowsMpvShortcuts,
            commandLineLauncher: snapshot.commandLineLauncher,
            message: firstRunSetupMessage,
        };
    },
    buildSetupHtml: (model) => (0, first_run_setup_window_1.buildFirstRunSetupHtml)(model),
    parseSubmissionUrl: (rawUrl) => (0, first_run_setup_window_1.parseFirstRunSetupSubmissionUrl)(rawUrl),
    handleAction: async (submission) => {
        if (submission.action === 'remove-legacy-plugin') {
            const snapshot = await firstRunSetupService.removeLegacyMpvPlugin();
            firstRunSetupMessage = snapshot.message;
            return;
        }
        if (submission.action === 'configure-mpv-executable-path') {
            const mpvExecutablePath = submission.mpvExecutablePath?.trim() ?? '';
            const pathStatus = (0, windows_mpv_launch_1.getConfiguredWindowsMpvPathStatus)(mpvExecutablePath);
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
            (0, first_run_setup_service_1.getFirstRunSetupCompletionMessage)(snapshot) ??
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
    shouldQuitWhenClosedCompleted: () => Boolean(appState.initialArgs && (0, first_run_setup_service_1.isStandaloneFirstRunSetupCommand)(appState.initialArgs)),
    quitApp: () => requestAppQuit(),
    clearSetupWindow: () => {
        appState.firstRunSetupWindow = null;
    },
    setSetupWindow: (window) => {
        appState.firstRunSetupWindow = window;
    },
    encodeURIComponent: (value) => encodeURIComponent(value),
    logError: (message, error) => logger.error(message, error),
});
function openFirstRunSetupWindow(force = false) {
    if (!force && firstRunSetupService.isSetupCompleted()) {
        return;
    }
    openFirstRunSetupWindowHandler();
}
const { notifyAnilistSetup, consumeAnilistSetupTokenFromUrl, handleAnilistSetupProtocolUrl, registerSubminerProtocolClient, } = (0, composers_1.composeAnilistSetupHandlers)({
    notifyDeps: {
        hasMpvClient: () => Boolean(appState.mpvClient),
        showMpvOsd: (message) => showMpvOsd(message),
        showDesktopNotification: (title, options) => (0, utils_2.showDesktopNotification)(title, options),
        logInfo: (message) => logger.info(message),
    },
    consumeTokenDeps: {
        consumeAnilistSetupCallbackUrl: anilist_1.consumeAnilistSetupCallbackUrl,
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
        setAsDefaultProtocolClient: (scheme, appPath, args) => appPath
            ? electron_1.app.setAsDefaultProtocolClient(scheme, appPath, args)
            : electron_1.app.setAsDefaultProtocolClient(scheme),
        logDebug: (message, details) => logger.debug(message, details),
    },
});
const maybeFocusExistingAnilistSetupWindow = (0, anilist_1.createMaybeFocusExistingAnilistSetupWindowHandler)({
    getSetupWindow: () => appState.anilistSetupWindow,
});
const buildOpenAnilistSetupWindowMainDepsHandler = (0, anilist_1.createBuildOpenAnilistSetupWindowMainDepsHandler)({
    maybeFocusExistingSetupWindow: maybeFocusExistingAnilistSetupWindow,
    createSetupWindow: (0, setup_window_factory_1.createCreateAnilistSetupWindowHandler)({
        createBrowserWindow: (options) => new electron_1.BrowserWindow(options),
    }),
    buildAuthorizeUrl: () => (0, anilist_1.buildAnilistSetupUrl)({
        authorizeUrl: ANILIST_SETUP_CLIENT_ID_URL,
        clientId: ANILIST_DEFAULT_CLIENT_ID,
        responseType: ANILIST_SETUP_RESPONSE_TYPE,
    }),
    consumeCallbackUrl: (rawUrl) => consumeAnilistSetupTokenFromUrl(rawUrl),
    openSetupInBrowser: (authorizeUrl) => (0, anilist_1.openAnilistSetupInBrowser)({
        authorizeUrl,
        openExternal: (url) => electron_1.shell.openExternal(url),
        logError: (message, error) => logger.error(message, error),
    }),
    loadManualTokenEntry: (setupWindow, authorizeUrl) => (0, anilist_1.loadAnilistManualTokenEntry)({
        setupWindow: setupWindow,
        authorizeUrl,
        developerSettingsUrl: ANILIST_DEVELOPER_SETTINGS_URL,
        logWarn: (message, data) => logger.warn(message, data),
    }),
    redirectUri: ANILIST_REDIRECT_URI,
    developerSettingsUrl: ANILIST_DEVELOPER_SETTINGS_URL,
    isAllowedExternalUrl: (url) => (0, anilist_url_guard_1.isAllowedAnilistExternalUrl)(url),
    isAllowedNavigationUrl: (url) => (0, anilist_url_guard_1.isAllowedAnilistSetupNavigationUrl)(url),
    logWarn: (message, details) => logger.warn(message, details),
    logError: (message, details) => logger.error(message, details),
    clearSetupWindow: () => {
        appState.anilistSetupWindow = null;
    },
    setSetupPageOpened: (opened) => {
        appState.anilistSetupPageOpened = opened;
    },
    setSetupWindow: (setupWindow) => {
        appState.anilistSetupWindow = setupWindow;
    },
    openExternal: (url) => {
        void electron_1.shell.openExternal(url);
    },
});
function openAnilistSetupWindow() {
    (0, anilist_1.createOpenAnilistSetupWindowHandler)(buildOpenAnilistSetupWindowMainDepsHandler())();
}
const { refreshAnilistClientSecretState, getCurrentAnilistMediaKey, resetAnilistMediaTracking, getAnilistMediaGuessRuntimeState, setAnilistMediaGuessRuntimeState, recordAnilistMediaDuration, resetAnilistMediaGuessState, maybeProbeAnilistDuration, ensureAnilistMediaGuess, processNextAnilistRetryUpdate, maybeRunAnilistPostWatchUpdate, } = (0, composers_1.composeAnilistTrackingHandlers)({
    refreshClientSecretMainDeps: {
        getResolvedConfig: () => getResolvedConfig(),
        isAnilistTrackingEnabled: (config) => (0, anilist_1.isAnilistTrackingEnabled)(config),
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
            anilistMediaGuessRuntimeState = (0, state_1.transitionAnilistMediaGuessRuntimeState)(anilistMediaGuessRuntimeState, { mediaKey: value });
        },
        setMediaDurationSec: (value) => {
            anilistMediaGuessRuntimeState = (0, state_1.transitionAnilistMediaGuessRuntimeState)(anilistMediaGuessRuntimeState, { mediaDurationSec: value });
        },
        setMediaGuess: (value) => {
            anilistMediaGuessRuntimeState = (0, state_1.transitionAnilistMediaGuessRuntimeState)(anilistMediaGuessRuntimeState, { mediaGuess: value });
        },
        setMediaGuessPromise: (value) => {
            anilistMediaGuessRuntimeState = (0, state_1.transitionAnilistMediaGuessRuntimeState)(anilistMediaGuessRuntimeState, { mediaGuessPromise: value });
        },
        setLastDurationProbeAtMs: (value) => {
            anilistMediaGuessRuntimeState = (0, state_1.transitionAnilistMediaGuessRuntimeState)(anilistMediaGuessRuntimeState, { lastDurationProbeAtMs: value });
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
            anilistMediaGuessRuntimeState = (0, state_1.transitionAnilistMediaGuessRuntimeState)(anilistMediaGuessRuntimeState, { mediaKey: value });
        },
        setMediaDurationSec: (value) => {
            anilistMediaGuessRuntimeState = (0, state_1.transitionAnilistMediaGuessRuntimeState)(anilistMediaGuessRuntimeState, { mediaDurationSec: value });
        },
        setMediaGuess: (value) => {
            anilistMediaGuessRuntimeState = (0, state_1.transitionAnilistMediaGuessRuntimeState)(anilistMediaGuessRuntimeState, { mediaGuess: value });
        },
        setMediaGuessPromise: (value) => {
            anilistMediaGuessRuntimeState = (0, state_1.transitionAnilistMediaGuessRuntimeState)(anilistMediaGuessRuntimeState, { mediaGuessPromise: value });
        },
        setLastDurationProbeAtMs: (value) => {
            anilistMediaGuessRuntimeState = (0, state_1.transitionAnilistMediaGuessRuntimeState)(anilistMediaGuessRuntimeState, { lastDurationProbeAtMs: value });
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
            anilistMediaGuessRuntimeState = (0, state_1.transitionAnilistMediaGuessRuntimeState)(anilistMediaGuessRuntimeState, { mediaGuess: value });
        },
        setMediaGuessPromise: (value) => {
            anilistMediaGuessRuntimeState = (0, state_1.transitionAnilistMediaGuessRuntimeState)(anilistMediaGuessRuntimeState, { mediaGuessPromise: value });
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
        resolveMediaPathForJimaku: (currentMediaPath) => mediaRuntime.resolveMediaPathForJimaku(currentMediaPath),
        getCurrentMediaPath: () => appState.currentMediaPath,
        getCurrentMediaTitle: () => appState.currentMediaTitle,
        guessAnilistMediaInfo: (mediaPath, mediaTitle) => (0, anilist_updater_1.guessAnilistMediaInfo)(mediaPath, mediaTitle),
    },
    processNextRetryUpdateMainDeps: {
        nextReady: () => anilistUpdateQueue.nextReady(),
        refreshRetryQueueState: () => anilistStateRuntime.refreshRetryQueueState(),
        setLastAttemptAt: (value) => {
            appState.anilistRetryQueueState = (0, state_1.transitionAnilistRetryQueueLastAttemptAt)(appState.anilistRetryQueueState, value);
        },
        setLastError: (value) => {
            appState.anilistRetryQueueState = (0, state_1.transitionAnilistRetryQueueLastError)(appState.anilistRetryQueueState, value);
        },
        refreshAnilistClientSecretState: () => refreshAnilistClientSecretState(),
        updateAnilistPostWatchProgress: (accessToken, title, episode, season) => (0, anilist_updater_1.updateAnilistPostWatchProgress)(accessToken, title, episode, {
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
            anilistUpdateInFlightState = (0, state_1.transitionAnilistUpdateInFlightState)(anilistUpdateInFlightState, value);
        },
        getResolvedConfig: () => getResolvedConfig(),
        isAnilistTrackingEnabled: (config) => (0, anilist_1.isAnilistTrackingEnabled)(config),
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
        updateAnilistPostWatchProgress: (accessToken, title, episode, season) => (0, anilist_updater_1.updateAnilistPostWatchProgress)(accessToken, title, episode, {
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
        minWatchRatio: watch_threshold_1.DEFAULT_MIN_WATCH_RATIO,
    },
});
function refreshAnilistClientSecretStateIfEnabled(options) {
    if (!(0, anilist_1.isAnilistTrackingEnabled)(getResolvedConfig())) {
        return Promise.resolve(null);
    }
    return refreshAnilistClientSecretState(options);
}
const rememberAnilistAttemptedUpdate = (key) => {
    (0, anilist_1.rememberAnilistAttemptedUpdateKey)(anilistAttemptedUpdateKeys, key, ANILIST_MAX_ATTEMPTED_UPDATE_KEYS);
};
const buildLoadSubtitlePositionMainDepsHandler = (0, overlay_1.createBuildLoadSubtitlePositionMainDepsHandler)({
    loadSubtitlePositionCore: () => (0, services_1.loadSubtitlePosition)({
        currentMediaPath: appState.currentMediaPath,
        fallbackPosition: getResolvedConfig().subtitlePosition,
        subtitlePositionsDir: SUBTITLE_POSITIONS_DIR,
    }),
    setSubtitlePosition: (position) => {
        appState.subtitlePosition = position;
    },
});
const loadSubtitlePositionMainDeps = buildLoadSubtitlePositionMainDepsHandler();
const loadSubtitlePosition = (0, overlay_1.createLoadSubtitlePositionHandler)(loadSubtitlePositionMainDeps);
const buildSaveSubtitlePositionMainDepsHandler = (0, overlay_1.createBuildSaveSubtitlePositionMainDepsHandler)({
    saveSubtitlePositionCore: (position) => {
        (0, services_1.saveSubtitlePosition)({
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
const saveSubtitlePosition = (0, overlay_1.createSaveSubtitlePositionHandler)(saveSubtitlePositionMainDeps);
registerSubminerProtocolClient();
let flushPendingMpvLogWrites = () => { };
const { registerProtocolUrlHandlers: registerProtocolUrlHandlersHandler, onWillQuitCleanup: onWillQuitCleanupHandler, shouldRestoreWindowsOnActivate: shouldRestoreWindowsOnActivateHandler, restoreWindowsOnActivate: restoreWindowsOnActivateHandler, } = (0, composers_1.composeStartupLifecycleHandlers)({
    registerProtocolUrlHandlersMainDeps: {
        registerOpenUrl: (listener) => {
            electron_1.app.on('open-url', listener);
        },
        registerSecondInstance: (listener) => {
            (0, early_single_instance_1.registerSecondInstanceHandlerEarly)(electron_1.app, listener);
        },
        handleAnilistSetupProtocolUrl: (rawUrl) => handleAnilistSetupProtocolUrl(rawUrl),
        findAnilistSetupDeepLinkArgvUrl: (argv) => (0, anilist_1.findAnilistSetupDeepLinkArgvUrl)(argv),
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
            restoreOverlayMpvSubtitles();
        },
        unregisterAllGlobalShortcuts: () => electron_1.globalShortcut.unregisterAll(),
        stopSubtitleWebsocket: () => {
            subtitleWsService.stop();
            annotationSubtitleWsService.stop();
        },
        stopTexthookerService: () => texthookerService.stop(),
        clearWindowsVisibleOverlayForegroundPollLoop: () => clearWindowsVisibleOverlayForegroundPollLoop(),
        clearLinuxMpvFullscreenOverlayRefreshTimeouts: () => {
            cancelLinuxMpvFullscreenOverlayRefreshBurst = null;
            (0, linux_mpv_fullscreen_overlay_refresh_1.clearLinuxMpvFullscreenOverlayRefreshTimeouts)();
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
        stopDiscordPresenceService: () => {
            void appState.discordPresenceService?.stop();
            appState.discordPresenceService = null;
        },
    },
    shouldRestoreWindowsOnActivateMainDeps: {
        isOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
        getAllWindowCount: () => electron_1.BrowserWindow.getAllWindows().length,
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
const startLocalStatsServer = () => {
    const tracker = appState.immersionTracker;
    if (!tracker) {
        throw new Error('Immersion tracker failed to initialize.');
    }
    if (!statsServer) {
        const yomitanDeps = {
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
        };
        const yomitanLogger = (0, logger_1.createLogger)('main:yomitan-stats');
        statsServer = (0, stats_server_1.startStatsServer)({
            port: getResolvedConfig().stats.serverPort,
            staticDir: statsDistPath,
            tracker,
            knownWordCachePath: path.join(USER_DATA_PATH, 'known-words-cache.json'),
            mpvSocketPath: appState.mpvSocketPath,
            ankiConnectConfig: getResolvedConfig().ankiConnect,
            anilistRateLimiter,
            resolveAnkiNoteId: (noteId) => appState.ankiIntegration?.resolveCurrentNoteId(noteId) ?? noteId,
            addYomitanNote: async (word) => {
                const ankiUrl = getResolvedConfig().ankiConnect.url || 'http://127.0.0.1:8765';
                await (0, services_1.syncYomitanDefaultAnkiServer)(ankiUrl, yomitanDeps, yomitanLogger, {
                    forceOverride: true,
                });
                const result = await (0, services_1.addYomitanNoteViaSearch)(word, yomitanDeps, yomitanLogger);
                if (result.noteId && result.duplicateNoteIds.length > 0) {
                    appState.ankiIntegration?.trackDuplicateNoteIdsForNote(result.noteId, result.duplicateNoteIds);
                }
                return result.noteId;
            },
        });
        appState.statsServer = statsServer;
    }
    appState.statsServer = statsServer;
};
const ensureStatsServerStarted = (0, stats_server_routing_1.createEnsureStatsServerUrlHandler)({
    currentPid: process.pid,
    readBackgroundState: () => (0, stats_daemon_1.readBackgroundStatsServerState)(statsDaemonStatePath),
    removeBackgroundState: () => {
        (0, stats_daemon_1.removeBackgroundStatsServerState)(statsDaemonStatePath);
    },
    isProcessAlive: (pid) => (0, stats_daemon_1.isBackgroundStatsServerProcessAlive)(pid),
    hasLocalStatsServer: () => statsServer !== null,
    startLocalStatsServer,
    getConfiguredPort: () => getResolvedConfig().stats.serverPort,
});
const ensureBackgroundStatsServerStarted = () => {
    const liveDaemon = readLiveBackgroundStatsDaemonState();
    if (liveDaemon && liveDaemon.pid !== process.pid) {
        return {
            url: (0, stats_daemon_1.resolveBackgroundStatsServerUrl)(liveDaemon),
            runningInCurrentProcess: false,
        };
    }
    appState.statsStartupInProgress = true;
    try {
        ensureImmersionTrackerStarted();
    }
    finally {
        appState.statsStartupInProgress = false;
    }
    const port = getResolvedConfig().stats.serverPort;
    const result = ensureStatsServerStarted();
    if (result.source === 'local') {
        (0, stats_daemon_1.writeBackgroundStatsServerState)(statsDaemonStatePath, {
            pid: process.pid,
            port,
            startedAtMs: Date.now(),
        });
    }
    return { url: result.url, runningInCurrentProcess: result.source === 'local' };
};
const stopBackgroundStatsServer = async () => {
    const state = (0, stats_daemon_1.readBackgroundStatsServerState)(statsDaemonStatePath);
    if (!state) {
        (0, stats_daemon_1.removeBackgroundStatsServerState)(statsDaemonStatePath);
        return { ok: true, stale: true };
    }
    if (!(0, stats_daemon_1.isBackgroundStatsServerProcessAlive)(state.pid)) {
        (0, stats_daemon_1.removeBackgroundStatsServerState)(statsDaemonStatePath);
        return { ok: true, stale: true };
    }
    try {
        process.kill(state.pid, 'SIGTERM');
    }
    catch (error) {
        if (error?.code === 'ESRCH') {
            (0, stats_daemon_1.removeBackgroundStatsServerState)(statsDaemonStatePath);
            return { ok: true, stale: true };
        }
        if (error?.code === 'EPERM') {
            throw new Error(`Insufficient permissions to stop background stats server (pid ${state.pid}).`);
        }
        throw error;
    }
    const deadline = Date.now() + 2_000;
    while (Date.now() < deadline) {
        if (!(0, stats_daemon_1.isBackgroundStatsServerProcessAlive)(state.pid)) {
            (0, stats_daemon_1.removeBackgroundStatsServerState)(statsDaemonStatePath);
            return { ok: true, stale: false };
        }
        await new Promise((resolve) => setTimeout(resolve, 50));
    }
    throw new Error('Timed out stopping background stats server.');
};
const resolveLegacyVocabularyPos = async (row) => {
    const tokenizer = appState.mecabTokenizer;
    if (!tokenizer) {
        return null;
    }
    const lookupTexts = [...new Set([row.headword, row.word, row.reading ?? ''])]
        .map((value) => value.trim())
        .filter((value) => value.length > 0);
    for (const lookupText of lookupTexts) {
        const tokens = await tokenizer.tokenize(lookupText);
        const resolved = (0, legacy_vocabulary_pos_1.resolveLegacyVocabularyPosFromTokens)(lookupText, tokens);
        if (resolved) {
            return resolved;
        }
    }
    return null;
};
const immersionTrackerStartupMainDeps = {
    getResolvedConfig: () => getResolvedConfig(),
    getConfiguredDbPath: () => immersionMediaRuntime.getConfiguredDbPath(),
    createTrackerService: (params) => new services_1.ImmersionTrackerService({
        ...params,
        resolveLegacyVocabularyPos,
    }),
    setTracker: (tracker) => {
        const trackerHasChanged = appState.immersionTracker !== null && appState.immersionTracker !== tracker;
        if (trackerHasChanged && appState.statsServer) {
            stopStatsServer();
            appState.statsServer = null;
        }
        appState.immersionTracker = tracker;
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
            (0, stats_window_js_1.registerStatsOverlayToggle)({
                staticDir: statsDistPath,
                preloadPath: statsPreloadPath,
                getApiBaseUrl: () => ensureStatsServerStarted().url,
                getToggleKey: () => getResolvedConfig().stats.toggleKey,
                resolveBounds: () => getCurrentOverlayGeometry(),
                onVisibilityChanged: (visible) => {
                    appState.statsOverlayVisible = visible;
                    overlayVisibilityRuntime.updateVisibleOverlayVisibility();
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
const createImmersionTrackerStartup = (0, immersion_startup_1.createImmersionTrackerStartupHandler)((0, immersion_startup_main_deps_1.createBuildImmersionTrackerStartupMainDepsHandler)(immersionTrackerStartupMainDeps)());
const recordTrackedCardsMined = (count, noteIds) => {
    ensureImmersionTrackerStarted();
    appState.immersionTracker?.recordCardsMined(count, noteIds);
};
const refreshCurrentSubtitleAfterKnownWordUpdate = () => {
    subtitleProcessingController.refreshCurrentSubtitle(appState.currentSubText);
};
let hasAttemptedImmersionTrackerStartup = false;
const ensureImmersionTrackerStarted = () => {
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
        }
        finally {
            appState.statsStartupInProgress = false;
        }
    },
};
const runStatsCliCommand = (0, stats_cli_command_1.createRunStatsCliCommandHandler)({
    getResolvedConfig: () => getResolvedConfig(),
    ensureImmersionTrackerStarted: () => statsStartupRuntime.ensureImmersionTrackerStarted(),
    ensureVocabularyCleanupTokenizerReady: async () => {
        await createMecabTokenizerAndCheck();
    },
    getImmersionTracker: () => appState.immersionTracker,
    ensureStatsServerStarted: () => statsStartupRuntime.ensureStatsServerStarted().url,
    ensureBackgroundStatsServerStarted: () => statsStartupRuntime.ensureBackgroundStatsServerStarted(),
    stopBackgroundStatsServer: () => statsStartupRuntime.stopBackgroundStatsServer(),
    openExternal: (url) => electron_1.shell.openExternal(url),
    writeResponse: (responsePath, payload) => {
        (0, stats_cli_command_1.writeStatsCliCommandResponse)(responsePath, payload);
    },
    exitAppWithCode: (code) => {
        process.exitCode = code;
        requestAppQuit();
    },
    logInfo: (message) => logger.info(message),
    logWarn: (message, error) => logger.warn(message, error),
    logError: (message, error) => logger.error(message, error),
});
async function runHeadlessInitialCommand() {
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
    const effectiveAnkiConfig = appState.runtimeOptionsManager?.getEffectiveAnkiConnectConfig(resolvedConfig.ankiConnect) ??
        resolvedConfig.ankiConnect;
    const integration = new anki_integration_1.AnkiIntegration(effectiveAnkiConfig, new subtitle_timing_tracker_1.SubtitleTimingTracker(), { send: () => undefined }, undefined, undefined, async () => ({
        keepNoteId: 0,
        deleteNoteId: 0,
        deleteDuplicate: false,
        cancelled: true,
    }), path.join(USER_DATA_PATH, 'known-words-cache.json'), (0, config_1.mergeAiConfig)(resolvedConfig.ai, resolvedConfig.ankiConnect?.ai));
    try {
        await integration.refreshKnownWordCache();
    }
    catch (error) {
        logger.error('Headless known-word refresh failed:', error);
        process.exitCode = 1;
    }
    finally {
        integration.stop();
        requestAppQuit();
    }
}
const { appReadyRuntimeRunner } = (0, composers_1.composeAppReadyRuntime)({
    reloadConfigMainDeps: {
        reloadConfigStrict: () => configService.reloadConfigStrict(),
        logInfo: (message) => appLogger.logInfo(message),
        logDebug: (message) => appLogger.logDebug(message),
        logWarning: (message) => appLogger.logWarning(message),
        showDesktopNotification: (title, options) => (0, utils_2.showDesktopNotification)(title, options),
        startConfigHotReload: () => configHotReloadRuntime.start(),
        shouldRefreshAnilistClientSecretState: () => (0, startup_mode_flags_1.shouldRefreshAnilistOnConfigReload)(appState.initialArgs),
        refreshAnilistClientSecretState: (options) => refreshAnilistClientSecretStateIfEnabled(options),
        failHandlers: {
            logError: (details) => logger.error(details),
            showErrorBox: (title, details) => electron_1.dialog.showErrorBox(title, details),
            quit: () => requestAppQuit(),
        },
    },
    criticalConfigErrorMainDeps: {
        getConfigPath: () => configService.getConfigPath(),
        failHandlers: {
            logError: (message) => logger.error(message),
            showErrorBox: (title, message) => electron_1.dialog.showErrorBox(title, message),
            quit: () => requestAppQuit(),
        },
    },
    appReadyRuntimeMainDeps: {
        ensureDefaultConfigBootstrap: () => {
            (0, setup_state_1.ensureDefaultConfigBootstrap)({
                configDir: CONFIG_DIR,
                configFilePaths: (0, setup_state_1.getDefaultConfigFilePaths)(CONFIG_DIR),
                generateTemplate: () => (0, config_2.generateConfigTemplate)(config_2.DEFAULT_CONFIG),
            });
        },
        loadSubtitlePosition: () => loadSubtitlePosition(),
        resolveKeybindings: () => {
            appState.keybindings = (0, utils_2.resolveKeybindings)(getResolvedConfig(), config_2.DEFAULT_KEYBINDINGS);
            refreshCurrentSessionBindings();
        },
        createMpvClient: () => {
            appState.mpvClient = createMpvClientRuntimeService();
        },
        getResolvedConfig: () => getResolvedConfig(),
        getConfigWarnings: () => configService.getWarnings(),
        logConfigWarning: (warning) => appLogger.logConfigWarning(warning),
        setLogLevel: (level, source) => (0, logger_1.setLogLevel)(level, source),
        initRuntimeOptionsManager: () => {
            appState.runtimeOptionsManager = new runtime_options_1.RuntimeOptionsManager(() => configService.getConfig().ankiConnect, {
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
            });
        },
        setSecondarySubMode: (mode) => {
            setSecondarySubMode(mode);
        },
        defaultSecondarySubMode: 'hover',
        defaultWebsocketPort: config_2.DEFAULT_CONFIG.websocket.port,
        defaultAnnotationWebsocketPort: config_2.DEFAULT_CONFIG.annotationWebsocket.port,
        defaultTexthookerPort: DEFAULT_TEXTHOOKER_PORT,
        hasMpvWebsocketPlugin: () => (0, services_1.hasMpvWebsocketPlugin)(),
        startSubtitleWebsocket: (port) => {
            subtitleWsService.start(port, () => appState.currentSubtitleData ??
                (appState.currentSubText
                    ? {
                        text: appState.currentSubText,
                        tokens: null,
                        startTime: appState.mpvClient?.currentSubStart ?? null,
                        endTime: appState.mpvClient?.currentSubEnd ?? null,
                    }
                    : null), () => ({
                enabled: getResolvedConfig().subtitleStyle.frequencyDictionary.enabled,
                topX: getResolvedConfig().subtitleStyle.frequencyDictionary.topX,
                mode: getResolvedConfig().subtitleStyle.frequencyDictionary.mode,
            }));
        },
        startAnnotationWebsocket: (port) => {
            annotationSubtitleWsService.start(port, () => appState.currentSubtitleData ??
                (appState.currentSubText
                    ? {
                        text: appState.currentSubText,
                        tokens: null,
                        startTime: appState.mpvClient?.currentSubStart ?? null,
                        endTime: appState.mpvClient?.currentSubEnd ?? null,
                    }
                    : null), () => ({
                enabled: getResolvedConfig().subtitleStyle.frequencyDictionary.enabled,
                topX: getResolvedConfig().subtitleStyle.frequencyDictionary.topX,
                mode: getResolvedConfig().subtitleStyle.frequencyDictionary.mode,
            }));
        },
        startTexthooker: (port, websocketUrl) => {
            if (!texthookerService.isRunning()) {
                texthookerService.start(port, websocketUrl);
            }
        },
        log: (message) => appLogger.logInfo(message),
        createMecabTokenizerAndCheck: async () => {
            await createMecabTokenizerAndCheck();
        },
        createSubtitleTimingTracker: () => {
            const tracker = new subtitle_timing_tracker_1.SubtitleTimingTracker();
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
            if (args && (0, first_run_setup_service_1.shouldAutoOpenFirstRunSetup)(args)) {
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
        shouldAutoInitializeOverlayRuntimeFromConfig: () => appState.backgroundMode
            ? false
            : configDerivedRuntime.shouldAutoInitializeOverlayRuntimeFromConfig(),
        setVisibleOverlayVisible: (visible) => setVisibleOverlayVisible(visible),
        initializeOverlayRuntime: () => initializeOverlayRuntime(),
        runHeadlessInitialCommand: () => runHeadlessInitialCommand(),
        handleInitialArgs: () => handleInitialArgs(),
        shouldRunHeadlessInitialCommand: () => Boolean(appState.initialArgs && (0, args_1.isHeadlessInitialCommand)(appState.initialArgs)),
        shouldUseMinimalStartup: () => (0, startup_mode_flags_1.getStartupModeFlags)(appState.initialArgs).shouldUseMinimalStartup,
        shouldSkipHeavyStartup: () => (0, startup_mode_flags_1.getStartupModeFlags)(appState.initialArgs).shouldSkipHeavyStartup,
        createImmersionTracker: () => {
            ensureImmersionTrackerStarted();
        },
        logDebug: (message) => {
            logger.debug(message);
        },
        now: () => Date.now(),
    },
    immersionTrackerStartupMainDeps,
});
function ensureOverlayStartupPrereqs() {
    if (appState.subtitlePosition === null) {
        loadSubtitlePosition();
    }
    if (appState.keybindings.length === 0) {
        appState.keybindings = (0, utils_2.resolveKeybindings)(getResolvedConfig(), config_2.DEFAULT_KEYBINDINGS);
        refreshCurrentSessionBindings();
    }
    else if (!appState.sessionBindingsInitialized) {
        refreshCurrentSessionBindings();
    }
    if (!appState.mpvClient) {
        appState.mpvClient = createMpvClientRuntimeService();
    }
    if (!appState.runtimeOptionsManager) {
        appState.runtimeOptionsManager = new runtime_options_1.RuntimeOptionsManager(() => configService.getConfig().ankiConnect, {
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
        });
    }
    if (!appState.subtitleTimingTracker) {
        appState.subtitleTimingTracker = new subtitle_timing_tracker_1.SubtitleTimingTracker();
    }
}
async function ensureYoutubePlaybackRuntimeReady() {
    ensureOverlayStartupPrereqs();
    await ensureYomitanExtensionLoaded();
    if (!appState.overlayRuntimeInitialized) {
        initializeOverlayRuntime();
        return;
    }
    ensureOverlayWindowsReadyForVisibilityActions();
}
const { createMpvClientRuntimeService: createMpvClientRuntimeServiceHandler, updateMpvSubtitleRenderMetrics: updateMpvSubtitleRenderMetricsHandler, tokenizeSubtitle, createMecabTokenizerAndCheck, prewarmSubtitleDictionaries, startBackgroundWarmups, startTokenizationWarmups, isTokenizationWarmupReady, } = (0, composers_1.composeMpvRuntimeHandlers)({
    bindMpvMainEventHandlersMainDeps: {
        appState,
        getQuitOnDisconnectArmed: () => jellyfinPlayQuitOnDisconnectArmed || youtubePlaybackRuntime.getQuitOnDisconnectArmed(),
        scheduleQuitCheck: (callback) => {
            setTimeout(callback, 500);
        },
        quitApp: () => requestAppQuit(),
        reportJellyfinRemoteStopped: () => {
            void reportJellyfinRemoteStopped();
        },
        onMpvConnected: () => {
            if (appState.sessionBindingsInitialized) {
                (0, services_1.sendMpvCommandRuntime)(appState.mpvClient, [
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
        tokenizeSubtitleForImmersion: async (text) => tokenizeSubtitleDeferred ? await tokenizeSubtitleDeferred(text) : null,
        updateCurrentMediaPath: (path) => {
            autoplayReadyGate.invalidatePendingAutoplayReadyFallbacks();
            currentMediaTokenizationGate.updateCurrentMediaPath(path);
            managedLocalSubtitleSelectionRuntime.handleMediaPathChange(path);
            startupOsdSequencer.reset();
            subtitlePrefetchRuntime.clearScheduledSubtitlePrefetchRefresh();
            subtitlePrefetchRuntime.cancelPendingInit();
            youtubePrimarySubtitleNotificationRuntime.handleMediaPathChange(path);
            if (path) {
                ensureImmersionTrackerStarted();
                // Delay slightly to allow MPV's track-list to be populated.
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
        signalAutoplayReadyIfWarm: () => {
            if (!isTokenizationWarmupReady()) {
                return;
            }
            autoplayReadyGate.maybeSignalPluginAutoplayReady({ text: '__warm__', tokens: null }, { forceWhilePaused: true });
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
            cancelLinuxMpvFullscreenOverlayRefreshBurst = (0, linux_mpv_fullscreen_overlay_refresh_1.updateLinuxMpvFullscreenOverlayRefreshBurst)(fullscreen, {
                overlayManager: {
                    getMainWindow: () => overlayManager.getMainWindow(),
                    getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
                },
                overlayVisibilityRuntime,
                ensureOverlayWindowLevel: (window) => ensureOverlayWindowLevel(window),
            }, cancelLinuxMpvFullscreenOverlayRefreshBurst);
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
            updateMpvSubtitleRenderMetrics(patch);
        },
        syncOverlayMpvSubtitleSuppression: () => {
            syncOverlayMpvSubtitleSuppression();
        },
    },
    mpvClientRuntimeServiceFactoryMainDeps: {
        createClient: services_1.MpvIpcClient,
        getSocketPath: () => appState.mpvSocketPath,
        getResolvedConfig: () => getResolvedConfig(),
        isAutoStartOverlayEnabled: () => appState.autoStartOverlay,
        setOverlayVisible: (visible) => setOverlayVisible(visible),
        isVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
        getReconnectTimer: () => appState.reconnectTimer,
        setReconnectTimer: (timer) => {
            appState.reconnectTimer = timer;
        },
        shouldQuitOnMpvShutdown: () => appState.initialArgs?.managedPlayback === true,
        requestAppQuit: () => requestAppQuit(),
    },
    updateMpvSubtitleRenderMetricsMainDeps: {
        getCurrentMetrics: () => appState.mpvSubtitleRenderMetrics,
        setCurrentMetrics: (metrics) => {
            appState.mpvSubtitleRenderMetrics = metrics;
        },
        applyPatch: (current, patch) => (0, services_1.applyMpvSubtitleRenderMetricsPatch)(current, patch),
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
                appState.yomitanParserWindow = window;
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
            getKnownWordMatchMode: () => appState.ankiIntegration?.getKnownWordMatchMode() ??
                getResolvedConfig().ankiConnect.knownWords.matchMode,
            getNPlusOneEnabled: () => getRuntimeBooleanOption('subtitle.annotation.nPlusOne', getResolvedConfig().ankiConnect.nPlusOne.enabled),
            getMinSentenceWordsForNPlusOne: () => getResolvedConfig().ankiConnect.nPlusOne.minSentenceWords,
            getJlptLevel: (text) => appState.jlptLevelLookup(text),
            getJlptEnabled: () => getRuntimeBooleanOption('subtitle.annotation.jlpt', getResolvedConfig().subtitleStyle.enableJlpt),
            getCharacterDictionaryEnabled: () => getResolvedConfig().anilist.characterDictionary.enabled &&
                yomitanProfilePolicy.isCharacterDictionaryEnabled() &&
                !isYoutubePlaybackActiveNow(),
            getNameMatchEnabled: () => getResolvedConfig().subtitleStyle.nameMatchEnabled,
            getFrequencyDictionaryEnabled: () => getRuntimeBooleanOption('subtitle.annotation.frequency', getResolvedConfig().subtitleStyle.frequencyDictionary.enabled),
            getFrequencyDictionaryMatchMode: () => getResolvedConfig().subtitleStyle.frequencyDictionary.matchMode,
            getFrequencyRank: (text) => appState.frequencyRankLookup(text),
            getYomitanGroupDebugEnabled: () => appState.overlayDebugVisualizationEnabled,
            getMecabTokenizer: () => appState.mecabTokenizer,
            onTokenizationReady: (text) => {
                currentMediaTokenizationGate.markReady(appState.currentMediaPath?.trim() || appState.mpvClient?.currentVideoPath?.trim() || null);
                startupOsdSequencer.markTokenizationReady();
                autoplayReadyGate.maybeSignalPluginAutoplayReady({ text, tokens: null }, { forceWhilePaused: true });
            },
        },
        createTokenizerRuntimeDeps: (deps) => (0, services_1.createTokenizerDepsRuntime)(deps),
        tokenizeSubtitle: (text, deps) => (0, services_1.tokenizeSubtitle)(text, deps),
        createMecabTokenizerAndCheckMainDeps: {
            getMecabTokenizer: () => appState.mecabTokenizer,
            setMecabTokenizer: (tokenizer) => {
                appState.mecabTokenizer = tokenizer;
            },
            createMecabTokenizer: () => new mecab_tokenizer_1.MecabTokenizer(),
            checkAvailability: async (tokenizer) => tokenizer.checkAvailability(),
        },
        prewarmSubtitleDictionariesMainDeps: {
            ensureJlptDictionaryLookup: () => jlptDictionaryRuntime.ensureJlptDictionaryLookup(),
            ensureFrequencyDictionaryLookup: () => frequencyDictionaryRuntime.ensureFrequencyDictionaryLookup(),
            showMpvOsd: (message) => showMpvOsd(message),
            showLoadingOsd: (message) => startupOsdSequencer.showAnnotationLoading(message),
            showLoadedOsd: (message) => startupOsdSequencer.markAnnotationLoadingComplete(message),
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
            ensureYomitanExtensionLoaded: () => ensureYomitanExtensionLoaded().then(() => { }),
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
                return (jellyfin.enabled && jellyfin.remoteControlEnabled && jellyfin.remoteControlAutoConnect);
            },
            startJellyfinRemoteSession: () => startJellyfinRemoteSession(),
            logDebug: (message) => logger.debug(message),
        },
    },
});
tokenizeSubtitleDeferred = tokenizeSubtitle;
function createMpvClientRuntimeService() {
    const client = createMpvClientRuntimeServiceHandler();
    client.on('connection-change', ({ connected }) => {
        if (connected) {
            return;
        }
        if (!youtubeFlowRuntime.hasActiveSession()) {
            return;
        }
        youtubeFlowRuntime.cancelActivePicker();
        broadcastToOverlayWindows(contracts_1.IPC_CHANNELS.event.youtubePickerCancel, null);
        overlayModalRuntime.handleOverlayModalClosed('youtube-track-picker');
    });
    return client;
}
function resetSubtitleSidebarEmbeddedLayoutRuntime() {
    (0, services_1.sendMpvCommandRuntime)(appState.mpvClient, ['set_property', 'video-margin-ratio-right', 0]);
    (0, services_1.sendMpvCommandRuntime)(appState.mpvClient, ['set_property', 'video-pan-x', 0]);
}
function updateMpvSubtitleRenderMetrics(patch) {
    updateMpvSubtitleRenderMetricsHandler(patch);
}
let lastOverlayWindowGeometry = null;
function getOverlayGeometryFallback() {
    const cursorPoint = electron_1.screen.getCursorScreenPoint();
    const display = electron_1.screen.getDisplayNearestPoint(cursorPoint);
    const bounds = display.workArea;
    return {
        x: bounds.x,
        y: bounds.y,
        width: bounds.width,
        height: bounds.height,
    };
}
function getCurrentOverlayGeometry() {
    if (lastOverlayWindowGeometry)
        return lastOverlayWindowGeometry;
    const trackerGeometry = appState.windowTracker?.getGeometry();
    if (trackerGeometry)
        return trackerGeometry;
    return getOverlayGeometryFallback();
}
function geometryMatches(a, b) {
    if (!a || !b)
        return false;
    return a.x === b.x && a.y === b.y && a.width === b.width && a.height === b.height;
}
function applyOverlayRegions(geometry) {
    lastOverlayWindowGeometry = geometry;
    overlayManager.setOverlayWindowBounds(geometry);
    overlayManager.setModalWindowBounds(geometry);
}
const buildUpdateVisibleOverlayBoundsMainDepsHandler = (0, overlay_1.createBuildUpdateVisibleOverlayBoundsMainDepsHandler)({
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
const updateVisibleOverlayBounds = (0, overlay_1.createUpdateVisibleOverlayBoundsHandler)(updateVisibleOverlayBoundsMainDeps);
const buildEnsureOverlayWindowLevelMainDepsHandler = (0, overlay_1.createBuildEnsureOverlayWindowLevelMainDepsHandler)({
    ensureOverlayWindowLevelCore: (window) => (0, services_1.ensureOverlayWindowLevel)(window),
});
const ensureOverlayWindowLevelMainDeps = buildEnsureOverlayWindowLevelMainDepsHandler();
const ensureOverlayWindowLevel = (0, overlay_1.createEnsureOverlayWindowLevelHandler)(ensureOverlayWindowLevelMainDeps);
function syncPrimaryOverlayWindowLayer(layer) {
    const mainWindow = overlayManager.getMainWindow();
    if (!mainWindow || mainWindow.isDestroyed())
        return;
    (0, services_1.syncOverlayWindowLayer)(mainWindow, layer);
}
const buildEnforceOverlayLayerOrderMainDepsHandler = (0, overlay_1.createBuildEnforceOverlayLayerOrderMainDepsHandler)({
    enforceOverlayLayerOrderCore: (params) => (0, services_1.enforceOverlayLayerOrder)({
        visibleOverlayVisible: params.visibleOverlayVisible,
        mainWindow: params.mainWindow,
        ensureOverlayWindowLevel: (window) => params.ensureOverlayWindowLevel(window),
    }),
    getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    getMainWindow: () => overlayManager.getMainWindow(),
    ensureOverlayWindowLevel: (window) => ensureOverlayWindowLevel(window),
});
const enforceOverlayLayerOrderMainDeps = buildEnforceOverlayLayerOrderMainDepsHandler();
const enforceOverlayLayerOrder = (0, overlay_1.createEnforceOverlayLayerOrderHandler)(enforceOverlayLayerOrderMainDeps);
async function loadYomitanExtension() {
    const extension = await yomitanExtensionRuntime.loadYomitanExtension();
    if (extension && !yomitanProfilePolicy.isExternalReadOnlyMode()) {
        await syncYomitanDefaultProfileAnkiServer();
    }
    return extension;
}
async function ensureYomitanExtensionLoaded() {
    const extension = await yomitanExtensionRuntime.ensureYomitanExtensionLoaded();
    if (extension && !yomitanProfilePolicy.isExternalReadOnlyMode()) {
        await syncYomitanDefaultProfileAnkiServer();
    }
    return extension;
}
let lastSyncedYomitanAnkiServer = null;
function getPreferredYomitanAnkiServerUrl() {
    return (0, yomitan_anki_server_1.getPreferredYomitanAnkiServerUrl)(getResolvedConfig().ankiConnect);
}
function getYomitanParserRuntimeDeps() {
    return {
        getYomitanExt: () => appState.yomitanExt,
        getYomitanSession: () => appState.yomitanSession,
        getYomitanParserWindow: () => appState.yomitanParserWindow,
        setYomitanParserWindow: (window) => {
            appState.yomitanParserWindow = window;
        },
        getYomitanParserReadyPromise: () => appState.yomitanParserReadyPromise,
        setYomitanParserReadyPromise: (promise) => {
            appState.yomitanParserReadyPromise = promise;
        },
        getYomitanParserInitPromise: () => appState.yomitanParserInitPromise,
        setYomitanParserInitPromise: (promise) => {
            appState.yomitanParserInitPromise = promise;
        },
    };
}
async function syncYomitanDefaultProfileAnkiServer() {
    if (yomitanProfilePolicy.isExternalReadOnlyMode()) {
        return;
    }
    const targetUrl = getPreferredYomitanAnkiServerUrl().trim();
    if (!targetUrl || targetUrl === lastSyncedYomitanAnkiServer) {
        return;
    }
    const synced = await (0, services_1.syncYomitanDefaultAnkiServer)(targetUrl, getYomitanParserRuntimeDeps(), {
        error: (message, ...args) => {
            logger.error(message, ...args);
        },
        info: (message, ...args) => {
            logger.info(message, ...args);
        },
    }, {
        forceOverride: (0, yomitan_anki_server_1.shouldForceOverrideYomitanAnkiServer)(getResolvedConfig().ankiConnect),
    });
    if (synced) {
        lastSyncedYomitanAnkiServer = targetUrl;
    }
}
function createModalWindow() {
    const existingWindow = overlayManager.getModalWindow();
    if (existingWindow && !existingWindow.isDestroyed()) {
        return existingWindow;
    }
    const window = createModalWindowHandler();
    overlayManager.setModalWindowBounds(getCurrentOverlayGeometry());
    return window;
}
function createMainWindow() {
    const window = createMainWindowHandler();
    if (process.platform === 'win32') {
        const overlayHwnd = getWindowsNativeWindowHandleNumber(window);
        if (!(0, windows_helper_1.ensureWindowsOverlayTransparency)(overlayHwnd)) {
            logger.warn('Failed to eagerly extend Windows overlay transparency via koffi');
        }
    }
    return window;
}
function ensureTray() {
    ensureTrayHandler();
}
function destroyTray() {
    destroyTrayHandler();
}
function initializeOverlayRuntime() {
    initializeOverlayRuntimeHandler();
    appState.ankiIntegration?.setRecordCardsMinedCallback(recordTrackedCardsMined);
    appState.ankiIntegration?.setKnownWordCacheUpdatedCallback(refreshCurrentSubtitleAfterKnownWordUpdate);
    syncOverlayMpvSubtitleSuppression();
}
function openYomitanSettings() {
    if (yomitanProfilePolicy.isExternalReadOnlyMode()) {
        const message = 'Yomitan settings unavailable while using read-only external-profile mode.';
        logger.warn('Yomitan settings window disabled while yomitan.externalProfilePath is configured because external profile mode is read-only.');
        (0, utils_2.showDesktopNotification)('SubMiner', { body: message });
        showMpvOsd(message);
        return false;
    }
    openYomitanSettingsHandler();
    return true;
}
const { getConfiguredShortcuts, registerGlobalShortcuts, refreshGlobalAndOverlayShortcuts, cancelPendingMultiCopy, startPendingMultiCopy, cancelPendingMineSentenceMultiple, startPendingMineSentenceMultiple, syncOverlayShortcuts, refreshOverlayShortcuts, } = (0, composers_1.composeShortcutRuntimes)({
    globalShortcuts: {
        getConfiguredShortcutsMainDeps: {
            getResolvedConfig: () => getResolvedConfig(),
            defaultConfig: config_2.DEFAULT_CONFIG,
            resolveConfiguredShortcuts: utils_2.resolveConfiguredShortcuts,
        },
        buildRegisterGlobalShortcutsMainDeps: (getConfiguredShortcutsHandler) => ({
            getConfiguredShortcuts: () => getConfiguredShortcutsHandler(),
            registerGlobalShortcutsCore: services_1.registerGlobalShortcuts,
            toggleVisibleOverlay: () => toggleVisibleOverlay(),
            openYomitanSettings: () => openYomitanSettings(),
            isDev,
            getMainWindow: () => overlayManager.getMainWindow(),
        }),
        buildRefreshGlobalAndOverlayShortcutsMainDeps: (registerGlobalShortcutsHandler) => ({
            unregisterAllGlobalShortcuts: () => electron_1.globalShortcut.unregisterAll(),
            registerGlobalShortcuts: () => registerGlobalShortcutsHandler(),
            syncOverlayShortcuts: () => syncOverlayShortcuts(),
        }),
    },
    numericShortcutRuntimeMainDeps: {
        globalShortcut: electron_1.globalShortcut,
        showMpvOsd: (text) => showMpvOsd(text),
        setTimer: (handler, timeoutMs) => setTimeout(handler, timeoutMs),
        clearTimer: (timer) => clearTimeout(timer),
    },
    numericSessions: {
        onMultiCopyDigit: (count) => handleMultiCopyDigit(count),
        onMineSentenceDigit: (count) => handleMineSentenceDigit(count),
    },
    overlayShortcutsRuntimeMainDeps: {
        overlayShortcutsRuntime,
    },
});
function resolveSessionBindingPlatform() {
    if (process.platform === 'darwin')
        return 'darwin';
    if (process.platform === 'win32')
        return 'win32';
    return 'linux';
}
function compileCurrentSessionBindings() {
    return (0, session_bindings_1.compileSessionBindings)({
        keybindings: appState.keybindings,
        shortcuts: getConfiguredShortcuts(),
        statsToggleKey: getResolvedConfig().stats.toggleKey,
        platform: resolveSessionBindingPlatform(),
        rawConfig: getResolvedConfig(),
    });
}
function persistSessionBindings(bindings, warnings = []) {
    const artifact = (0, session_bindings_1.buildPluginSessionBindingsArtifact)({
        bindings,
        warnings,
        numericSelectionTimeoutMs: getConfiguredShortcuts().multiCopyTimeoutMs,
    });
    (0, session_bindings_artifact_1.writeSessionBindingsArtifact)(CONFIG_DIR, artifact);
    appState.sessionBindings = bindings;
    appState.sessionBindingsInitialized = true;
    if (appState.mpvClient?.connected) {
        (0, services_1.sendMpvCommandRuntime)(appState.mpvClient, [
            'script-message',
            'subminer-reload-session-bindings',
        ]);
    }
}
function refreshCurrentSessionBindings() {
    const compiled = compileCurrentSessionBindings();
    for (const warning of compiled.warnings) {
        logger.warn(`[session-bindings] ${warning.message}`);
    }
    persistSessionBindings(compiled.bindings, compiled.warnings);
}
const { flushMpvLog, showMpvOsd } = (0, mpv_1.createMpvOsdRuntimeHandlers)({
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
        showMpvOsdRuntime: (mpvClient, text, fallbackLog) => (0, services_1.showMpvOsdRuntime)(mpvClient, text, fallbackLog),
        getMpvClient: () => appState.mpvClient,
        logInfo: (line) => logger.info(line),
    }),
});
flushPendingMpvLogWrites = () => {
    void flushMpvLog();
};
const updateStateStore = (0, update_service_1.createFileUpdateStateStore)(path.join(USER_DATA_PATH, 'update-state.json'));
let updateService = null;
const electronNetFetch = (0, fetch_adapter_1.createElectronNetFetch)({
    fetch: (url, init) => electron_1.net.fetch(url, init),
});
function getFetchForUpdater() {
    return electronNetFetch;
}
async function updateLauncherFromSelectedRelease(launcherPath, channel = getResolvedConfig().updates.channel, release = null) {
    const fetchForUpdater = getFetchForUpdater();
    if (!release) {
        return { status: 'missing-asset', message: `No ${channel} GitHub release found.` };
    }
    const sumsAsset = (0, release_assets_1.findReleaseAsset)(release, 'SHA256SUMS.txt');
    if (!sumsAsset) {
        return { status: 'missing-asset', message: 'Release has no SHA256SUMS.txt asset.' };
    }
    const sums = (0, release_assets_1.parseSha256Sums)(await (0, release_assets_1.fetchReleaseAssetText)(fetchForUpdater, sumsAsset.browser_download_url));
    const launcherResult = await (0, launcher_updater_1.updateLauncherFromRelease)({
        release,
        sha256Sums: sums,
        launcherPath,
        downloadAsset: (url) => (0, release_assets_1.fetchReleaseAssetBuffer)(fetchForUpdater, url),
    });
    const supportResults = await (0, support_assets_1.updateSupportAssetsFromRelease)({
        release,
        sha256Sums: sums,
        downloadAsset: (url) => (0, release_assets_1.fetchReleaseAssetBuffer)(fetchForUpdater, url),
    });
    for (const result of supportResults) {
        if (result.status === 'protected' && result.command) {
            logger.warn(`Rofi theme update requires manual command: ${result.command}`);
        }
        else if (result.status === 'hash-mismatch' || result.status === 'missing-asset') {
            logger.warn(`Rofi theme update skipped: ${result.message ?? result.status}`);
        }
    }
    return launcherResult;
}
function getUpdateService() {
    if (updateService)
        return updateService;
    const appUpdater = (0, app_updater_1.createElectronAppUpdater)({
        currentVersion: electron_1.app.getVersion(),
        isPackaged: electron_1.app.isPackaged,
        log: (message) => logger.info(message),
        getChannel: () => getResolvedConfig().updates.channel,
        configureHttpExecutor: process.platform === 'darwin' ? () => (0, curl_http_executor_1.createCurlHttpExecutor)() : undefined,
        disableDifferentialDownload: process.platform === 'darwin',
        isNativeUpdaterSupported: () => (0, app_updater_1.isNativeUpdaterSupported)({
            platform: process.platform,
            isPackaged: electron_1.app.isPackaged,
            execPath: process.execPath,
            env: process.env,
            log: (message) => logger.warn(message),
        }),
    });
    const updateDialogPresenter = (0, update_dialogs_1.createUpdateDialogPresenter)({
        platform: process.platform,
        focusApp: () => electron_1.app.focus({ steal: true }),
        showMessageBox: (options) => electron_1.dialog.showMessageBox(options),
    });
    updateService = (0, update_service_1.createUpdateService)({
        getConfig: () => getResolvedConfig().updates,
        getCurrentVersion: () => electron_1.app.getVersion(),
        now: () => Date.now(),
        readState: () => updateStateStore.readState(),
        writeState: (state) => updateStateStore.writeState(state),
        checkAppUpdate: (channel) => appUpdater.checkForUpdates(channel),
        shouldFetchReleaseMetadata: ({ appUpdate }) => (0, release_metadata_policy_1.shouldFetchReleaseMetadataForPlatform)(process.platform, appUpdate),
        fetchLatestStableRelease: (channel) => (0, release_assets_1.fetchLatestStableRelease)({ fetch: getFetchForUpdater(), channel }),
        updateLauncher: (launcherPath, channel, release) => updateLauncherFromSelectedRelease(launcherPath, channel, release),
        showNoUpdateDialog: (version) => updateDialogPresenter.showNoUpdateDialog(version),
        showUpdateAvailableDialog: (version) => updateDialogPresenter.showUpdateAvailableDialog(version),
        showUpdateFailedDialog: (message) => updateDialogPresenter.showUpdateFailedDialog(message),
        showManualUpdateRequiredDialog: (version) => updateDialogPresenter.showManualUpdateRequiredDialog(version),
        downloadAppUpdate: () => appUpdater.downloadUpdate(),
        showRestartDialog: () => updateDialogPresenter.showRestartDialog(),
        quitAndInstall: () => appUpdater.quitAndInstall(),
        notifyUpdateAvailable: (version) => (0, update_notifications_1.notifyUpdateAvailable)({ notificationType: getResolvedConfig().updates.notificationType, version }, {
            showSystemNotification: (title, body) => (0, utils_2.showDesktopNotification)(title, { body }),
            showOsdNotification: (message) => showMpvOsd(message),
            log: (message) => logger.warn(message),
        }),
        log: (message) => logger.warn(message),
    });
    return updateService;
}
const cycleSecondarySubMode = (0, mpv_1.createCycleSecondarySubModeRuntimeHandler)({
    cycleSecondarySubModeMainDeps: {
        getSecondarySubMode: () => appState.secondarySubMode,
        setSecondarySubMode: (mode) => {
            setSecondarySubMode(mode);
        },
        getLastSecondarySubToggleAtMs: () => appState.lastSecondarySubToggleAtMs,
        setLastSecondarySubToggleAtMs: (timestampMs) => {
            appState.lastSecondarySubToggleAtMs = timestampMs;
        },
        broadcastToOverlayWindows: (channel, mode) => {
            broadcastToOverlayWindows(channel, mode);
        },
        showMpvOsd: (text) => showMpvOsd(text),
    },
    cycleSecondarySubMode: (deps) => (0, services_1.cycleSecondarySubMode)(deps),
});
function setSecondarySubMode(mode) {
    appState.secondarySubMode = mode;
}
function handleCycleSecondarySubMode() {
    cycleSecondarySubMode();
}
function toggleSubtitleSidebar() {
    broadcastToOverlayWindows(contracts_1.IPC_CHANNELS.event.subtitleSidebarToggle);
}
function togglePrimarySubtitleBar() {
    broadcastToOverlayWindows(contracts_1.IPC_CHANNELS.event.primarySubtitleBarToggle);
}
async function triggerSubsyncFromConfig() {
    await subsyncRuntime.triggerFromConfig();
}
function handleMultiCopyDigit(count) {
    handleMultiCopyDigitHandler(count);
}
function copyCurrentSubtitle() {
    copyCurrentSubtitleHandler();
}
const buildUpdateLastCardFromClipboardMainDepsHandler = (0, mining_1.createBuildUpdateLastCardFromClipboardMainDepsHandler)({
    getAnkiIntegration: () => appState.ankiIntegration,
    readClipboardText: () => electron_1.clipboard.readText(),
    showMpvOsd: (text) => showMpvOsd(text),
    updateLastCardFromClipboardCore: services_1.updateLastCardFromClipboard,
});
const updateLastCardFromClipboardMainDeps = buildUpdateLastCardFromClipboardMainDepsHandler();
const updateLastCardFromClipboardHandler = (0, mining_1.createUpdateLastCardFromClipboardHandler)(updateLastCardFromClipboardMainDeps);
const buildRefreshKnownWordCacheMainDepsHandler = (0, mining_1.createBuildRefreshKnownWordCacheMainDepsHandler)({
    getAnkiIntegration: () => appState.ankiIntegration,
    missingIntegrationMessage: 'AnkiConnect integration not enabled',
});
const refreshKnownWordCacheMainDeps = buildRefreshKnownWordCacheMainDepsHandler();
const refreshKnownWordCacheHandler = (0, mining_1.createRefreshKnownWordCacheHandler)(refreshKnownWordCacheMainDeps);
const buildTriggerFieldGroupingMainDepsHandler = (0, mining_1.createBuildTriggerFieldGroupingMainDepsHandler)({
    getAnkiIntegration: () => appState.ankiIntegration,
    showMpvOsd: (text) => showMpvOsd(text),
    triggerFieldGroupingCore: services_1.triggerFieldGrouping,
});
const triggerFieldGroupingMainDeps = buildTriggerFieldGroupingMainDepsHandler();
const triggerFieldGroupingHandler = (0, mining_1.createTriggerFieldGroupingHandler)(triggerFieldGroupingMainDeps);
const buildMarkLastCardAsAudioCardMainDepsHandler = (0, mining_1.createBuildMarkLastCardAsAudioCardMainDepsHandler)({
    getAnkiIntegration: () => appState.ankiIntegration,
    showMpvOsd: (text) => showMpvOsd(text),
    markLastCardAsAudioCardCore: services_1.markLastCardAsAudioCard,
});
const markLastCardAsAudioCardMainDeps = buildMarkLastCardAsAudioCardMainDepsHandler();
const markLastCardAsAudioCardHandler = (0, mining_1.createMarkLastCardAsAudioCardHandler)(markLastCardAsAudioCardMainDeps);
const buildMineSentenceCardMainDepsHandler = (0, mining_1.createBuildMineSentenceCardMainDepsHandler)({
    getAnkiIntegration: () => appState.ankiIntegration,
    getMpvClient: () => appState.mpvClient,
    showMpvOsd: (text) => showMpvOsd(text),
    mineSentenceCardCore: services_1.mineSentenceCard,
    recordCardsMined: (count, noteIds) => {
        ensureImmersionTrackerStarted();
        appState.immersionTracker?.recordCardsMined(count, noteIds);
    },
});
const mineSentenceCardHandler = (0, mining_1.createMineSentenceCardHandler)(buildMineSentenceCardMainDepsHandler());
const buildHandleMultiCopyDigitMainDepsHandler = (0, mining_1.createBuildHandleMultiCopyDigitMainDepsHandler)({
    getSubtitleTimingTracker: () => appState.subtitleTimingTracker,
    writeClipboardText: (text) => electron_1.clipboard.writeText(text),
    showMpvOsd: (text) => showMpvOsd(text),
    handleMultiCopyDigitCore: services_1.handleMultiCopyDigit,
});
const handleMultiCopyDigitMainDeps = buildHandleMultiCopyDigitMainDepsHandler();
const handleMultiCopyDigitHandler = (0, mining_1.createHandleMultiCopyDigitHandler)(handleMultiCopyDigitMainDeps);
const buildCopyCurrentSubtitleMainDepsHandler = (0, mining_1.createBuildCopyCurrentSubtitleMainDepsHandler)({
    getSubtitleTimingTracker: () => appState.subtitleTimingTracker,
    writeClipboardText: (text) => electron_1.clipboard.writeText(text),
    showMpvOsd: (text) => showMpvOsd(text),
    copyCurrentSubtitleCore: services_1.copyCurrentSubtitle,
});
const copyCurrentSubtitleMainDeps = buildCopyCurrentSubtitleMainDepsHandler();
const copyCurrentSubtitleHandler = (0, mining_1.createCopyCurrentSubtitleHandler)(copyCurrentSubtitleMainDeps);
const buildHandleMineSentenceDigitMainDepsHandler = (0, mining_1.createBuildHandleMineSentenceDigitMainDepsHandler)({
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
    handleMineSentenceDigitCore: services_1.handleMineSentenceDigit,
});
const handleMineSentenceDigitMainDeps = buildHandleMineSentenceDigitMainDepsHandler();
const handleMineSentenceDigitHandler = (0, mining_1.createHandleMineSentenceDigitHandler)(handleMineSentenceDigitMainDeps);
const { setVisibleOverlayVisible: setVisibleOverlayVisibleHandler, toggleVisibleOverlay: toggleVisibleOverlayHandler, setOverlayVisible: setOverlayVisibleHandler, } = (0, overlay_1.createOverlayVisibilityRuntime)({
    setVisibleOverlayVisibleDeps: {
        setVisibleOverlayVisibleCore: services_1.setVisibleOverlayVisible,
        setVisibleOverlayVisibleState: (nextVisible) => {
            overlayManager.setVisibleOverlayVisible(nextVisible);
        },
        updateVisibleOverlayVisibility: () => overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
    },
    getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
});
const buildHandleOverlayModalClosedMainDepsHandler = (0, overlay_1.createBuildHandleOverlayModalClosedMainDepsHandler)({
    handleOverlayModalClosedRuntime: (modal) => overlayModalRuntime.handleOverlayModalClosed(modal),
});
const handleOverlayModalClosedMainDeps = buildHandleOverlayModalClosedMainDepsHandler();
const handleOverlayModalClosedHandler = (0, overlay_1.createHandleOverlayModalClosedHandler)(handleOverlayModalClosedMainDeps);
const buildAppendClipboardVideoToQueueMainDepsHandler = (0, overlay_1.createBuildAppendClipboardVideoToQueueMainDepsHandler)({
    appendClipboardVideoToQueueRuntime: startup_1.appendClipboardVideoToQueueRuntime,
    getMpvClient: () => appState.mpvClient,
    readClipboardText: () => electron_1.clipboard.readText(),
    showMpvOsd: (text) => showMpvOsd(text),
    sendMpvCommand: (command) => {
        (0, services_1.sendMpvCommandRuntime)(appState.mpvClient, command);
    },
});
const appendClipboardVideoToQueueMainDeps = buildAppendClipboardVideoToQueueMainDepsHandler();
const appendClipboardVideoToQueueHandler = (0, overlay_1.createAppendClipboardVideoToQueueHandler)(appendClipboardVideoToQueueMainDeps);
async function loadSubtitleSourceText(source) {
    if (/^https?:\/\//i.test(source)) {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 4000);
        try {
            const response = await fetch(source, { signal: controller.signal });
            if (!response.ok) {
                throw new Error(`Failed to download subtitle source (${response.status})`);
            }
            return await response.text();
        }
        finally {
            clearTimeout(timeoutId);
        }
    }
    const filePath = (0, subtitle_prefetch_source_1.resolveSubtitleSourcePath)(source);
    return fs.promises.readFile(filePath, 'utf8');
}
function parseTrackId(value) {
    if (typeof value === 'number' && Number.isInteger(value)) {
        return value;
    }
    if (typeof value === 'string') {
        const parsed = Number(value.trim());
        return Number.isInteger(parsed) ? parsed : null;
    }
    return null;
}
function buildFfmpegSubtitleExtractionArgs(videoPath, ffIndex, outputPath) {
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
async function extractInternalSubtitleTrackToTempFile(ffmpegPath, videoPath, track) {
    const ffIndex = parseTrackId(track['ff-index']);
    const codec = typeof track.codec === 'string' ? track.codec : null;
    const extension = (0, utils_3.codecToExtension)(codec ?? undefined);
    if (ffIndex === null || extension === null) {
        return null;
    }
    const tempDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'subminer-sidebar-'));
    const outputPath = path.join(tempDir, `track_${ffIndex}.${extension}`);
    try {
        await new Promise((resolve, reject) => {
            const child = (0, node_child_process_1.spawn)(ffmpegPath, buildFfmpegSubtitleExtractionArgs(videoPath, ffIndex, outputPath));
            let stderr = '';
            child.stderr.on('data', (chunk) => {
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
    }
    catch (error) {
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
const shiftSubtitleDelayToAdjacentCueHandler = (0, services_1.createShiftSubtitleDelayToAdjacentCueHandler)({
    getMpvClient: () => appState.mpvClient,
    loadSubtitleSourceText,
    sendMpvCommand: (command) => (0, services_1.sendMpvCommandRuntime)(appState.mpvClient, command),
    showMpvOsd: (text) => showMpvOsd(text),
});
async function dispatchSessionAction(request) {
    await (0, session_actions_1.dispatchSessionAction)(request, {
        toggleStatsOverlay: () => (0, stats_window_js_2.toggleStatsOverlay)({
            staticDir: statsDistPath,
            preloadPath: statsPreloadPath,
            getApiBaseUrl: () => ensureStatsServerStarted().url,
            getToggleKey: () => getResolvedConfig().stats.toggleKey,
            resolveBounds: () => getCurrentOverlayGeometry(),
            onVisibilityChanged: (visible) => {
                appState.statsOverlayVisible = visible;
                overlayVisibilityRuntime.updateVisibleOverlayVisibility();
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
        openRuntimeOptionsPalette: () => openRuntimeOptionsPalette(),
        openJimaku: () => openJimakuOverlay(),
        openSessionHelp: () => openSessionHelpOverlay(),
        openCharacterDictionary: () => openCharacterDictionaryOverlay(),
        openControllerSelect: () => openControllerSelectOverlay(),
        openControllerDebug: () => openControllerDebugOverlay(),
        openYoutubeTrackPicker: () => openYoutubeTrackPickerFromPlayback(),
        openPlaylistBrowser: () => openPlaylistBrowser(),
        replayCurrentSubtitle: () => (0, services_1.replayCurrentSubtitleRuntime)(appState.mpvClient),
        playNextSubtitle: () => (0, services_1.playNextSubtitleRuntime)(appState.mpvClient),
        shiftSubDelayToAdjacentSubtitle: (direction) => shiftSubtitleDelayToAdjacentCueHandler(direction),
        cycleRuntimeOption: (id, direction) => {
            if (!appState.runtimeOptionsManager) {
                return { ok: false, error: 'Runtime options manager unavailable' };
            }
            return (0, runtime_options_ipc_1.applyRuntimeOptionResultRuntime)(appState.runtimeOptionsManager.cycleOption(id, direction), (text) => showMpvOsd(text));
        },
        showMpvOsd: (text) => showMpvOsd(text),
    });
}
const { playlistBrowserMainDeps } = (0, playlist_browser_ipc_1.createPlaylistBrowserIpcRuntime)(() => appState.mpvClient, {
    getPrimarySubtitleLanguages: () => getResolvedConfig().youtube.primarySubLanguages,
    getSecondarySubtitleLanguages: () => getResolvedConfig().secondarySub.secondarySubLanguages,
});
const { registerIpcRuntimeHandlers } = (0, composers_1.composeIpcRuntimeHandlers)({
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
            return (0, runtime_options_ipc_1.applyRuntimeOptionResultRuntime)(appState.runtimeOptionsManager.cycleOption(id, direction), (text) => showMpvOsd(text));
        },
        showMpvOsd: (text) => showMpvOsd(text),
        replayCurrentSubtitle: () => (0, services_1.replayCurrentSubtitleRuntime)(appState.mpvClient),
        playNextSubtitle: () => (0, services_1.playNextSubtitleRuntime)(appState.mpvClient),
        shiftSubDelayToAdjacentSubtitle: (direction) => shiftSubtitleDelayToAdjacentCueHandler(direction),
        sendMpvCommand: (rawCommand) => (0, services_1.sendMpvCommandRuntime)(appState.mpvClient, rawCommand),
        getMpvClient: () => appState.mpvClient,
        isMpvConnected: () => Boolean(appState.mpvClient && appState.mpvClient.connected),
        hasRuntimeOptionsManager: () => appState.runtimeOptionsManager !== null,
    },
    handleMpvCommandFromIpcRuntime: ipc_mpv_command_1.handleMpvCommandFromIpcRuntime,
    runSubsyncManualFromIpc: (request) => subsyncRuntime.runManualFromIpc(request),
    registration: {
        runtimeOptions: {
            getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
            showMpvOsd: (text) => showMpvOsd(text),
        },
        mainDeps: {
            getMainWindow: () => overlayManager.getMainWindow(),
            getVisibleOverlayVisibility: () => overlayManager.getVisibleOverlayVisible(),
            focusMainWindow: () => {
                const mainWindow = overlayManager.getMainWindow();
                if (!mainWindow || mainWindow.isDestroyed())
                    return;
                if (!mainWindow.isFocused()) {
                    mainWindow.focus();
                }
            },
            onOverlayModalClosed: (modal, senderWindow) => {
                const modalWindow = overlayManager.getModalWindow();
                if (senderWindow &&
                    modalWindow &&
                    senderWindow === modalWindow &&
                    !senderWindow.isDestroyed()) {
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
            tokenizeCurrentSubtitle: async () => (0, current_subtitle_snapshot_1.resolveCurrentSubtitleForRenderer)({
                currentSubText: appState.currentSubText,
                currentSubtitleData: appState.currentSubtitleData,
                withCurrentSubtitleTiming: (payload) => withCurrentSubtitleTiming(payload),
            }),
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
                    const [currentExternalFilenameRaw, currentTrackRaw, trackListRaw, sidRaw, videoPathRaw] = await Promise.all([
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
                        const cues = (0, subtitle_cue_parser_1.parseSubtitleCues)(content, resolvedSource.path);
                        appState.activeParsedSubtitleCues = cues;
                        appState.activeParsedSubtitleSource = resolvedSource.sourceKey;
                        return {
                            cues,
                            currentTimeSec,
                            currentSubtitle,
                            config,
                        };
                    }
                    finally {
                        await resolvedSource.cleanup?.();
                    }
                }
                catch {
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
                return (0, overlay_1.resolveSubtitleStyleForRenderer)(resolvedConfig);
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
                    controller: (0, controller_config_update_js_1.applyControllerConfigUpdate)(currentRawConfig.controller, update),
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
            reportOverlayContentBounds: (payload) => {
                overlayContentMeasurementStore.report(payload);
            },
            getAnilistStatus: () => anilistStateRuntime.getStatusSnapshot(),
            clearAnilistToken: () => anilistStateRuntime.clearTokenState(),
            openAnilistSetup: () => openAnilistSetupWindow(),
            getAnilistQueueStatus: () => anilistStateRuntime.getQueueStatusSnapshot(),
            retryAnilistQueueNow: () => processNextAnilistRetryUpdate(),
            runAnilistPostWatchUpdateOnManualMark: () => maybeRunAnilistPostWatchUpdate({ force: true }),
            getCharacterDictionarySelection: () => characterDictionaryRuntime.getManualSelectionSnapshot(),
            setCharacterDictionarySelection: async (mediaId) => (0, character_dictionary_selection_1.applyCharacterDictionarySelection)({ mediaId }, {
                setManualSelection: (request) => characterDictionaryRuntime.setManualSelection(request),
                resetAnilistMediaGuessState,
                runSyncNow: () => characterDictionaryAutoSyncRuntime.runSyncNow(),
                warn: (message, error) => logger.warn(message, error),
            }),
            appendClipboardVideoToQueue: () => appendClipboardVideoToQueue(),
            ...playlistBrowserMainDeps,
            getImmersionTracker: () => appState.immersionTracker,
        },
        ankiJimakuDeps: (0, dependencies_1.createAnkiJimakuIpcRuntimeServiceDeps)({
            patchAnkiConnectEnabled: (enabled) => {
                configService.patchRawConfig({ ankiConnect: { enabled } });
            },
            getResolvedConfig: () => getResolvedConfig(),
            getRuntimeOptionsManager: () => appState.runtimeOptionsManager,
            getSubtitleTimingTracker: () => appState.subtitleTimingTracker,
            getMpvClient: () => appState.mpvClient,
            getAnkiIntegration: () => appState.ankiIntegration,
            setAnkiIntegration: (integration) => {
                appState.ankiIntegration = integration;
                appState.ankiIntegration?.setRecordCardsMinedCallback(recordTrackedCardsMined);
                appState.ankiIntegration?.setKnownWordCacheUpdatedCallback(refreshCurrentSubtitleAfterKnownWordUpdate);
            },
            getKnownWordCacheStatePath: () => path.join(USER_DATA_PATH, 'known-words-cache.json'),
            showDesktopNotification: utils_2.showDesktopNotification,
            createFieldGroupingCallback: () => createFieldGroupingCallback(),
            broadcastRuntimeOptionsChanged: () => broadcastRuntimeOptionsChanged(),
            getFieldGroupingResolver: () => getFieldGroupingResolver(),
            setFieldGroupingResolver: (resolver) => setFieldGroupingResolver(resolver),
            parseMediaInfo: (mediaPath) => (0, utils_1.parseMediaInfo)(mediaRuntime.resolveMediaPathForJimaku(mediaPath)),
            getCurrentMediaPath: () => appState.currentMediaPath,
            jimakuFetchJson: (endpoint, query) => configDerivedRuntime.jimakuFetchJson(endpoint, query),
            getJimakuMaxEntryResults: () => configDerivedRuntime.getJimakuMaxEntryResults(),
            getJimakuLanguagePreference: () => configDerivedRuntime.getJimakuLanguagePreference(),
            resolveJimakuApiKey: () => configDerivedRuntime.resolveJimakuApiKey(),
            isRemoteMediaPath: (mediaPath) => (0, utils_1.isRemoteMediaPath)(mediaPath),
            downloadToFile: (url, destPath, headers) => (0, utils_1.downloadToFile)(url, destPath, headers),
        }),
        registerIpcRuntimeServices: ipc_runtime_1.registerIpcRuntimeServices,
    },
});
const { handleCliCommand, handleInitialArgs } = (0, composers_1.composeCliStartupHandlers)({
    cliCommandContextMainDeps: {
        appState,
        setLogLevel: (level) => (0, logger_1.setLogLevel)(level, 'cli'),
        texthookerService,
        getResolvedConfig: () => getResolvedConfig(),
        defaultWebsocketPort: config_2.DEFAULT_CONFIG.websocket.port,
        defaultAnnotationWebsocketPort: config_2.DEFAULT_CONFIG.annotationWebsocket.port,
        hasMpvWebsocketPlugin: () => (0, services_1.hasMpvWebsocketPlugin)(),
        openExternal: (url) => electron_1.shell.openExternal(url),
        logBrowserOpenError: (url, error) => logger.error(`Failed to open browser for texthooker URL: ${url}`, error),
        showMpvOsd: (text) => showMpvOsd(text),
        initializeOverlayRuntime: () => initializeOverlayRuntime(),
        toggleVisibleOverlay: () => toggleVisibleOverlay(),
        togglePrimarySubtitleBar: () => togglePrimarySubtitleBar(),
        openFirstRunSetupWindow: (force) => openFirstRunSetupWindow(force),
        setVisibleOverlayVisible: (visible) => setVisibleOverlayVisible(visible),
        copyCurrentSubtitle: () => copyCurrentSubtitle(),
        startPendingMultiCopy: (timeoutMs) => startPendingMultiCopy(timeoutMs),
        mineSentenceCard: () => mineSentenceCard(),
        startPendingMineSentenceMultiple: (timeoutMs) => startPendingMineSentenceMultiple(timeoutMs),
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
        generateCharacterDictionary: async (targetPath) => {
            const disabledReason = yomitanProfilePolicy.getCharacterDictionaryDisabledReason();
            if (disabledReason) {
                throw new Error(disabledReason);
            }
            return await characterDictionaryRuntime.generateForCurrentMedia(targetPath);
        },
        getCharacterDictionarySelection: async (targetPath) => characterDictionaryRuntime.getManualSelectionSnapshot(targetPath),
        setCharacterDictionarySelection: async (request) => (0, character_dictionary_selection_1.applyCharacterDictionarySelection)(request, {
            setManualSelection: (selectionRequest) => characterDictionaryRuntime.setManualSelection(selectionRequest),
            resetAnilistMediaGuessState,
            runSyncNow: () => characterDictionaryAutoSyncRuntime.runSyncNow(),
            warn: (message, error) => logger.warn(message, error),
        }),
        runJellyfinCommand: (argsFromCommand) => runJellyfinCommand(argsFromCommand),
        runStatsCommand: (argsFromCommand, source) => runStatsCliCommand(argsFromCommand, source),
        runUpdateCommand: async (argsFromCommand, source) => {
            await (0, update_cli_command_1.runUpdateCliCommand)(argsFromCommand, source, {
                checkForUpdates: (request) => getUpdateService().checkForUpdates(request),
                writeResponse: (responsePath, payload) => (0, update_cli_command_1.writeUpdateCliCommandResponse)(responsePath, payload),
                logWarn: (message, error) => logger.warn(message, error),
            });
        },
        runYoutubePlaybackFlow: (request) => youtubePlaybackRuntime.runYoutubePlaybackFlow(request),
        openYomitanSettings: () => openYomitanSettings(),
        openConfigSettingsWindow: () => openConfigSettingsWindow(),
        cycleSecondarySubMode: () => handleCycleSecondarySubMode(),
        openRuntimeOptionsPalette: () => openRuntimeOptionsPalette(),
        printHelp: () => (0, help_1.printHelp)(DEFAULT_TEXTHOOKER_PORT),
        stopApp: () => requestAppQuit(),
        hasMainWindow: () => Boolean(overlayManager.getMainWindow()),
        dispatchSessionAction: (request) => dispatchSessionAction(request),
        getMultiCopyTimeoutMs: () => getConfiguredShortcuts().multiCopyTimeoutMs,
        schedule: (fn, delayMs) => setTimeout(fn, delayMs),
        logInfo: (message) => logger.info(message),
        logDebug: (message) => logger.debug(message),
        logWarn: (message) => logger.warn(message),
        logError: (message, err) => logger.error(message, err),
    },
    cliCommandRuntimeHandlerMainDeps: {
        handleTexthookerOnlyModeTransitionMainDeps: {
            isTexthookerOnlyMode: () => appState.texthookerOnlyMode,
            ensureOverlayStartupPrereqs: () => ensureOverlayStartupPrereqs(),
            setTexthookerOnlyMode: (enabled) => {
                appState.texthookerOnlyMode = enabled;
            },
            commandNeedsOverlayStartupPrereqs: (inputArgs) => (0, args_1.commandNeedsOverlayStartupPrereqs)(inputArgs),
            startBackgroundWarmups: () => startBackgroundWarmups(),
            logInfo: (message) => logger.info(message),
        },
        handleCliCommandRuntimeServiceWithContext: (args, source, cliContext) => (0, cli_runtime_1.handleCliCommandRuntimeServiceWithContext)(args, source, cliContext),
    },
    initialArgsRuntimeHandlerMainDeps: {
        getInitialArgs: () => appState.initialArgs,
        isBackgroundMode: () => appState.backgroundMode,
        shouldEnsureTrayOnStartup: () => (0, startup_tray_policy_1.shouldEnsureTrayOnStartupForInitialArgs)(process.platform, appState.initialArgs),
        shouldRunHeadlessInitialCommand: (args) => (0, args_1.isHeadlessInitialCommand)(args),
        ensureTray: () => ensureTray(),
        isTexthookerOnlyMode: () => appState.texthookerOnlyMode,
        hasImmersionTracker: () => Boolean(appState.immersionTracker),
        getMpvClient: () => appState.mpvClient,
        commandNeedsOverlayStartupPrereqs: (args) => (0, args_1.commandNeedsOverlayStartupPrereqs)(args),
        commandNeedsOverlayRuntime: (args) => (0, args_1.commandNeedsOverlayRuntime)(args),
        ensureOverlayStartupPrereqs: () => ensureOverlayStartupPrereqs(),
        isOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
        initializeOverlayRuntime: () => initializeOverlayRuntime(),
        logInfo: (message) => logger.info(message),
    },
});
const { runAndApplyStartupState } = (0, composers_1.composeHeadlessStartupHandlers)({
    startupRuntimeHandlersDeps: {
        appLifecycleRuntimeRunnerMainDeps: {
            app: appLifecycleApp,
            platform: process.platform,
            shouldStartApp: (nextArgs) => (0, args_1.shouldStartApp)(nextArgs),
            parseArgs: (argv) => (0, args_1.parseArgs)(argv),
            handleCliCommand: (nextArgs, source) => handleCliCommand(nextArgs, source),
            printHelp: () => (0, help_1.printHelp)(DEFAULT_TEXTHOOKER_PORT),
            logNoRunningInstance: () => appLogger.logNoRunningInstance(),
            onReady: appReadyRuntimeRunner,
            onWillQuitCleanup: () => onWillQuitCleanupHandler(),
            shouldRestoreWindowsOnActivate: () => shouldRestoreWindowsOnActivateHandler(),
            restoreWindowsOnActivate: () => restoreWindowsOnActivateHandler(),
            shouldQuitOnWindowAllClosed: () => (0, startup_tray_policy_1.shouldQuitOnWindowAllClosedForTrayState)({
                backgroundMode: appState.backgroundMode,
                hasTray: Boolean(appTray),
            }),
        },
        createAppLifecycleRuntimeRunner: (params) => (0, startup_lifecycle_1.createAppLifecycleRuntimeRunner)(params),
        buildStartupBootstrapMainDeps: (startAppLifecycle) => ({
            argv: process.argv,
            parseArgs: (argv) => (0, args_1.parseArgs)(argv),
            setLogLevel: (level, source) => {
                (0, logger_1.setLogLevel)(level, source);
            },
            forceX11Backend: (args) => {
                (0, utils_2.forceX11Backend)(args);
            },
            enforceUnsupportedWaylandMode: (args) => {
                (0, utils_2.enforceUnsupportedWaylandMode)(args);
            },
            shouldStartApp: (args) => (0, args_1.shouldStartApp)(args),
            getDefaultSocketPath: () => getResolvedConfig().mpv.socketPath || getDefaultSocketPath(),
            defaultTexthookerPort: DEFAULT_TEXTHOOKER_PORT,
            configDir: CONFIG_DIR,
            defaultConfig: config_2.DEFAULT_CONFIG,
            generateConfigTemplate: (config) => (0, config_2.generateConfigTemplate)(config),
            generateDefaultConfigFile: (args, options) => (0, utils_2.generateDefaultConfigFile)(args, options),
            setExitCode: (code) => {
                process.exitCode = code;
            },
            quitApp: () => requestAppQuit(),
            logGenerateConfigError: (message) => logger.error(message),
            startAppLifecycle,
        }),
        createStartupBootstrapRuntimeDeps: (deps) => (0, startup_2.createStartupBootstrapRuntimeDeps)(deps),
        runStartupBootstrapRuntime: services_1.runStartupBootstrapRuntime,
        applyStartupState: (startupState) => (0, state_1.applyStartupState)(appState, startupState),
    },
});
runAndApplyStartupState();
void electron_1.app.whenReady().then(() => {
    if (!(0, startup_mode_flags_1.shouldStartAutomaticUpdateChecks)(appState.initialArgs)) {
        return;
    }
    getUpdateService().startAutomaticChecks();
});
const startupModeFlags = (0, startup_mode_flags_1.getStartupModeFlags)(appState.initialArgs);
const shouldUseMinimalStartup = startupModeFlags.shouldUseMinimalStartup;
const shouldSkipHeavyStartup = startupModeFlags.shouldSkipHeavyStartup;
if (!appState.initialArgs || (!shouldUseMinimalStartup && !shouldSkipHeavyStartup)) {
    if ((0, anilist_1.isAnilistTrackingEnabled)(getResolvedConfig())) {
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
const { createMainWindow: createMainWindowHandler, createModalWindow: createModalWindowHandler } = (0, overlay_window_runtime_handlers_1.createOverlayWindowRuntimeHandlers)({
    createOverlayWindowDeps: {
        createOverlayWindowCore: (kind, options) => (0, services_1.createOverlayWindow)(kind, options),
        isDev,
        ensureOverlayWindowLevel: (window) => ensureOverlayWindowLevel(window),
        onRuntimeOptionsChanged: () => broadcastRuntimeOptionsChanged(),
        setOverlayDebugVisualizationEnabled: (enabled) => setOverlayDebugVisualizationEnabled(enabled),
        isOverlayVisible: (windowKind) => windowKind === 'visible' ? overlayManager.getVisibleOverlayVisible() : false,
        getYomitanSession: () => appState.yomitanSession,
        tryHandleOverlayShortcutLocalFallback: (input) => overlayShortcutsRuntime.tryHandleOverlayShortcutLocalFallback(input),
        forwardTabToMpv: () => (0, services_1.sendMpvCommandRuntime)(appState.mpvClient, ['keypress', 'TAB']),
        onVisibleWindowBlurred: () => scheduleVisibleOverlayBlurRefresh(),
        onWindowContentReady: () => {
            overlayVisibilityRuntime.updateVisibleOverlayVisibility();
            if (appState.currentSubText.trim()) {
                subtitleProcessingController.refreshCurrentSubtitle(appState.currentSubText);
            }
        },
        onWindowClosed: (windowKind) => {
            if (windowKind === 'visible') {
                cancelPendingLinuxMpvFullscreenOverlayRefreshBurst();
                overlayManager.setMainWindow(null);
            }
            else {
                overlayManager.setModalWindow(null);
            }
        },
    },
    setMainWindow: (window) => overlayManager.setMainWindow(window),
    setModalWindow: (window) => overlayManager.setModalWindow(window),
});
function refreshTrayMenuIfPresent() {
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
        startRemoteSession: (options) => startJellyfinRemoteSession(options),
        refreshTrayMenu: () => refreshTrayMenuIfPresent(),
        logger,
        showMpvOsd: (message) => showMpvOsd(message),
    };
}
const { ensureTray: ensureTrayHandler, destroyTray: destroyTrayHandler } = (0, overlay_1.createTrayRuntimeHandlers)({
    resolveTrayIconPathDeps: {
        resolveTrayIconPathRuntime: overlay_1.resolveTrayIconPathRuntime,
        platform: process.platform,
        resourcesPath: process.resourcesPath,
        appPath: electron_1.app.getAppPath(),
        dirname: __dirname,
        joinPath: (...parts) => path.join(...parts),
        fileExists: (candidate) => fs.existsSync(candidate),
    },
    buildTrayMenuTemplateDeps: {
        buildTrayMenuTemplateRuntime: overlay_1.buildTrayMenuTemplateRuntime,
        initializeOverlayRuntime: () => initializeOverlayRuntime(),
        isOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
        openSessionHelpModal: () => openSessionHelpOverlay(),
        openTexthookerInBrowser: () => handleCliCommand((0, args_1.parseArgs)(['--texthooker', '--open-browser'])),
        showTexthookerPage: () => getResolvedConfig().texthooker.launchAtStartup !== false,
        showFirstRunSetup: () => !firstRunSetupService.isSetupCompleted(),
        openFirstRunSetupWindow: () => openFirstRunSetupWindow(),
        showWindowsMpvLauncherSetup: () => process.platform === 'win32',
        openYomitanSettings: () => openYomitanSettings(),
        openRuntimeOptionsPalette: () => openRuntimeOptionsPalette(),
        openConfigSettingsWindow: () => openConfigSettingsWindow(),
        openJellyfinSetupWindow: () => openJellyfinSetupWindow(),
        isJellyfinConfigured: () => (0, jellyfin_tray_discovery_1.isJellyfinConfiguredForTray)(getJellyfinTrayDiscoveryDeps()),
        isJellyfinDiscoveryActive: () => Boolean(appState.jellyfinRemoteSession),
        toggleJellyfinDiscovery: () => (0, jellyfin_tray_discovery_1.toggleJellyfinDiscoveryFromTray)(getJellyfinTrayDiscoveryDeps()),
        openAnilistSetupWindow: () => openAnilistSetupWindow(),
        checkForUpdates: () => {
            void getUpdateService().checkForUpdates({ source: 'manual' });
        },
        quitApp: () => requestAppQuit(),
    },
    ensureTrayDeps: {
        getTray: () => appTray,
        setTray: (tray) => {
            appTray = tray;
        },
        createImageFromPath: (iconPath) => electron_1.nativeImage.createFromPath(iconPath),
        createEmptyImage: () => electron_1.nativeImage.createEmpty(),
        createTray: (icon) => new electron_1.Tray(icon),
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
            appTray = tray;
        },
    },
    buildMenuFromTemplate: (template) => electron_1.Menu.buildFromTemplate(template),
});
const yomitanProfilePolicy = (0, yomitan_profile_policy_1.createYomitanProfilePolicy)({
    externalProfilePath: getResolvedConfig().yomitan.externalProfilePath,
    logInfo: (message) => logger.info(message),
});
const configuredExternalYomitanProfilePath = yomitanProfilePolicy.externalProfilePath;
const yomitanExtensionRuntime = (0, overlay_1.createYomitanExtensionRuntime)({
    loadYomitanExtensionCore: services_1.loadYomitanExtension,
    userDataPath: USER_DATA_PATH,
    externalProfilePath: configuredExternalYomitanProfilePath,
    getYomitanParserWindow: () => appState.yomitanParserWindow,
    setYomitanParserWindow: (window) => {
        appState.yomitanParserWindow = window;
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
});
const { initializeOverlayRuntime: initializeOverlayRuntimeHandler } = runtimeRegistry.overlay.createOverlayRuntimeBootstrapHandlers({
    initializeOverlayRuntimeMainDeps: {
        appState,
        overlayManager: {
            getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
        },
        overlayVisibilityRuntime: {
            updateVisibleOverlayVisibility: () => overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
        },
        refreshCurrentSubtitle: () => {
            subtitleProcessingController.refreshCurrentSubtitle(appState.currentSubText);
        },
        overlayShortcutsRuntime: {
            syncOverlayShortcuts: () => overlayShortcutsRuntime.syncOverlayShortcuts(),
        },
        createMainWindow: () => {
            if (appState.initialArgs && (0, args_1.isHeadlessInitialCommand)(appState.initialArgs)) {
                return;
            }
            createMainWindow();
        },
        registerGlobalShortcuts: () => {
            if (appState.initialArgs && (0, args_1.isHeadlessInitialCommand)(appState.initialArgs)) {
                return;
            }
            registerGlobalShortcuts();
        },
        createWindowTracker: (override, targetMpvSocketPath) => {
            if (appState.initialArgs && (0, args_1.isHeadlessInitialCommand)(appState.initialArgs)) {
                return null;
            }
            return (0, window_trackers_1.createWindowTracker)(override, targetMpvSocketPath);
        },
        updateVisibleOverlayBounds: (geometry) => updateVisibleOverlayBounds(geometry),
        bindOverlayOwner: () => {
            const mainWindow = overlayManager.getMainWindow();
            if (process.platform !== 'win32' || !mainWindow || mainWindow.isDestroyed())
                return;
            const overlayHwnd = getWindowsNativeWindowHandleNumber(mainWindow);
            const targetWindowHwnd = resolveWindowsOverlayBindTargetHandle(appState.mpvSocketPath);
            if (targetWindowHwnd !== null &&
                (0, windows_helper_1.bindWindowsOverlayAboveMpv)(overlayHwnd, targetWindowHwnd)) {
                return;
            }
            const tracker = appState.windowTracker;
            const mpvResult = tracker
                ? (() => {
                    try {
                        const win32 = require('./window-trackers/win32');
                        const poll = win32.findMpvWindows();
                        const focused = poll.matches.find((m) => m.isForeground);
                        return focused ?? [...poll.matches].sort((a, b) => b.area - a.area)[0] ?? null;
                    }
                    catch {
                        return null;
                    }
                })()
                : null;
            if (!mpvResult)
                return;
            if (!(0, windows_helper_1.setWindowsOverlayOwner)(overlayHwnd, mpvResult.hwnd)) {
                logger.warn('Failed to set overlay owner via koffi');
            }
        },
        releaseOverlayOwner: () => {
            const mainWindow = overlayManager.getMainWindow();
            if (process.platform !== 'win32' || !mainWindow || mainWindow.isDestroyed())
                return;
            const overlayHwnd = getWindowsNativeWindowHandleNumber(mainWindow);
            if (!(0, windows_helper_1.clearWindowsOverlayOwner)(overlayHwnd)) {
                logger.warn('Failed to clear overlay owner via koffi');
            }
        },
        getOverlayWindows: () => getOverlayWindows(),
        getResolvedConfig: () => getResolvedConfig(),
        showDesktopNotification: utils_2.showDesktopNotification,
        createFieldGroupingCallback: () => createFieldGroupingCallback(),
        getKnownWordCacheStatePath: () => path.join(USER_DATA_PATH, 'known-words-cache.json'),
        shouldStartAnkiIntegration: () => !(appState.initialArgs && (0, args_1.isHeadlessInitialCommand)(appState.initialArgs)),
    },
    initializeOverlayRuntimeBootstrapDeps: {
        isOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
        initializeOverlayRuntimeCore: services_1.initializeOverlayRuntime,
        setOverlayRuntimeInitialized: (initialized) => {
            appState.overlayRuntimeInitialized = initialized;
        },
        startBackgroundWarmups: () => {
            if (appState.initialArgs && (0, args_1.isHeadlessInitialCommand)(appState.initialArgs)) {
                return;
            }
            startBackgroundWarmups();
        },
    },
});
const { openYomitanSettings: openYomitanSettingsHandler } = (0, overlay_1.createYomitanSettingsRuntime)({
    ensureYomitanExtensionLoaded: () => ensureYomitanExtensionLoaded(),
    getYomitanExtension: () => appState.yomitanExt,
    getYomitanExtensionLoadInFlight: () => yomitanLoadInFlight,
    getYomitanSession: () => appState.yomitanSession,
    openYomitanSettingsWindow: ({ yomitanExt, getExistingWindow, setWindow, yomitanSession }) => {
        (0, services_1.openYomitanSettingsWindow)({
            yomitanExt: yomitanExt,
            getExistingWindow: () => getExistingWindow(),
            setWindow: (window) => setWindow(window),
            yomitanSession: yomitanSession ?? appState.yomitanSession,
            onWindowClosed: () => {
                if (appState.yomitanParserWindow) {
                    (0, services_1.clearYomitanParserCachesForWindow)(appState.yomitanParserWindow);
                }
            },
        });
    },
    getExistingWindow: () => appState.yomitanSettingsWindow,
    setWindow: (window) => {
        appState.yomitanSettingsWindow = window;
    },
    logWarn: (message) => logger.warn(message),
    logError: (message, error) => logger.error(message, error),
});
async function updateLastCardFromClipboard() {
    await updateLastCardFromClipboardHandler();
}
async function refreshKnownWordCache() {
    await refreshKnownWordCacheHandler();
}
async function triggerFieldGrouping() {
    await triggerFieldGroupingHandler();
}
async function markLastCardAsAudioCard() {
    await markLastCardAsAudioCardHandler();
}
async function mineSentenceCard() {
    await mineSentenceCardHandler();
}
function handleMineSentenceDigit(count) {
    handleMineSentenceDigitHandler(count);
}
function ensureOverlayWindowsReadyForVisibilityActions() {
    if (!appState.overlayRuntimeInitialized) {
        initializeOverlayRuntime();
        return;
    }
    const mainWindow = overlayManager.getMainWindow();
    if (!mainWindow || mainWindow.isDestroyed()) {
        createMainWindow();
    }
}
function setVisibleOverlayVisible(visible) {
    ensureOverlayWindowsReadyForVisibilityActions();
    if (!visible) {
        cancelPendingLinuxMpvFullscreenOverlayRefreshBurst();
    }
    if (visible) {
        void ensureOverlayMpvSubtitlesHidden();
    }
    setVisibleOverlayVisibleHandler(visible);
    syncOverlayMpvSubtitleSuppression();
}
function toggleVisibleOverlay() {
    ensureOverlayWindowsReadyForVisibilityActions();
    if (overlayManager.getVisibleOverlayVisible()) {
        cancelPendingLinuxMpvFullscreenOverlayRefreshBurst();
    }
    else {
        void ensureOverlayMpvSubtitlesHidden();
    }
    toggleVisibleOverlayHandler();
    syncOverlayMpvSubtitleSuppression();
}
function setOverlayVisible(visible) {
    if (!visible) {
        cancelPendingLinuxMpvFullscreenOverlayRefreshBurst();
    }
    if (visible) {
        void ensureOverlayMpvSubtitlesHidden();
    }
    setOverlayVisibleHandler(visible);
    syncOverlayMpvSubtitleSuppression();
}
function handleOverlayModalClosed(modal) {
    handleOverlayModalClosedHandler(modal);
}
function appendClipboardVideoToQueue() {
    return appendClipboardVideoToQueueHandler();
}
registerIpcRuntimeHandlers();
//# sourceMappingURL=main.js.map
