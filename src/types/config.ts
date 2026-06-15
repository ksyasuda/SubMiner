import type { AnkiConnectConfig } from './anki';
import type {
  AiConfig,
  AiFeatureConfig,
  AnilistCharacterDictionaryCollapsibleSectionsConfig,
  AnilistCharacterDictionaryEvictionPolicy,
  AnilistCharacterDictionaryProfileScope,
  AnilistConfig,
  DiscordPresenceConfig,
  ImmersionTrackingConfig,
  ImmersionTrackingRetentionMode,
  ImmersionTrackingRetentionPreset,
  JellyfinConfig,
  JimakuConfig,
  JimakuLanguagePreference,
  StatsConfig,
  YomitanConfig,
  YoutubeConfig,
  YoutubeSubgenConfig,
} from './integrations';
import type {
  ControllerButtonIndicesConfig,
  ControllerConfig,
  ResolvedControllerProfileConfig,
  ControllerTriggerInputMode,
  Keybinding,
  ResolvedControllerBindingsConfig,
} from './runtime';
import type {
  FrequencyDictionaryMatchMode,
  FrequencyDictionaryMode,
  NPlusOneMatchMode,
  ResolvedSubtitleSidebarConfig,
  SecondarySubConfig,
  SubtitlePosition,
  SubtitleSidebarConfig,
  SubtitleStyleConfig,
} from './subtitle';
import type { NotificationType, OverlayNotificationPosition } from './notification';

export interface WebSocketConfig {
  enabled?: boolean | 'auto';
  port?: number;
}

export interface AnnotationWebSocketConfig {
  enabled?: boolean;
  port?: number;
}

export interface TexthookerConfig {
  launchAtStartup?: boolean;
  openBrowser?: boolean;
}

export type MpvLaunchMode = 'normal' | 'maximized' | 'fullscreen';
export type MpvBackend = 'auto' | 'hyprland' | 'sway' | 'x11' | 'macos' | 'windows';

export interface MpvConfig {
  executablePath?: string;
  launchMode?: MpvLaunchMode;
  profile?: string;
  socketPath?: string;
  backend?: MpvBackend;
  autoStartSubMiner?: boolean;
  pauseUntilOverlayReady?: boolean;
  subminerBinaryPath?: string;
  aniskipEnabled?: boolean;
  aniskipButtonKey?: string;
}

export interface SubsyncConfig {
  alass_path?: string;
  ffsubsync_path?: string;
  ffmpeg_path?: string;
  replace?: boolean;
}

export interface StartupWarmupsConfig {
  lowPowerMode?: boolean;
  mecab?: boolean;
  yomitanExtension?: boolean;
  subtitleDictionaries?: boolean;
  jellyfinRemoteSession?: boolean;
}

export type UpdateNotificationType = NotificationType;
export type UpdateChannel = 'stable' | 'prerelease';

export interface UpdatesConfig {
  enabled?: boolean;
  checkIntervalHours?: number;
  notificationType?: UpdateNotificationType;
  channel?: UpdateChannel;
}

export interface NotificationsConfig {
  overlayPosition?: OverlayNotificationPosition;
}

export type LogRotation = number;

export interface LogFilesConfig {
  app?: boolean;
  launcher?: boolean;
  mpv?: boolean;
}

export interface ShortcutsConfig {
  toggleVisibleOverlayGlobal?: string | null;
  copySubtitle?: string | null;
  copySubtitleMultiple?: string | null;
  updateLastCardFromClipboard?: string | null;
  triggerFieldGrouping?: string | null;
  triggerSubsync?: string | null;
  mineSentence?: string | null;
  mineSentenceMultiple?: string | null;
  multiCopyTimeoutMs?: number;
  toggleSecondarySub?: string | null;
  markAudioCard?: string | null;
  openCharacterDictionaryManager?: string | null;
  openRuntimeOptions?: string | null;
  openJimaku?: string | null;
  openSessionHelp?: string | null;
  openControllerSelect?: string | null;
  openControllerDebug?: string | null;
  toggleSubtitleSidebar?: string | null;
  toggleNotificationHistory?: string | null;
}

export interface Config {
  subtitlePosition?: SubtitlePosition;
  keybindings?: Keybinding[];
  websocket?: WebSocketConfig;
  annotationWebsocket?: AnnotationWebSocketConfig;
  texthooker?: TexthookerConfig;
  mpv?: MpvConfig;
  controller?: ControllerConfig;
  ankiConnect?: AnkiConnectConfig;
  shortcuts?: ShortcutsConfig;
  secondarySub?: SecondarySubConfig;
  subsync?: SubsyncConfig;
  startupWarmups?: StartupWarmupsConfig;
  subtitleStyle?: SubtitleStyleConfig;
  subtitleSidebar?: SubtitleSidebarConfig;
  auto_start_overlay?: boolean;
  uiLanguage?: string;
  jimaku?: JimakuConfig;
  anilist?: AnilistConfig;
  yomitan?: YomitanConfig;
  jellyfin?: JellyfinConfig;
  discordPresence?: DiscordPresenceConfig;
  ai?: AiConfig;
  youtube?: YoutubeConfig;
  youtubeSubgen?: YoutubeSubgenConfig;
  immersionTracking?: ImmersionTrackingConfig;
  stats?: StatsConfig;
  updates?: UpdatesConfig;
  notifications?: NotificationsConfig;
  logging?: {
    level?: 'debug' | 'info' | 'warn' | 'error';
    rotation?: LogRotation;
    files?: LogFilesConfig;
  };
}

export type RawConfig = Config;

export interface ResolvedConfig {
  subtitlePosition: SubtitlePosition;
  keybindings: Keybinding[];
  websocket: Required<WebSocketConfig>;
  annotationWebsocket: Required<AnnotationWebSocketConfig>;
  texthooker: Required<TexthookerConfig>;
  mpv: {
    executablePath: string;
    launchMode: MpvLaunchMode;
    profile: string;
    socketPath: string;
    backend: MpvBackend;
    autoStartSubMiner: boolean;
    pauseUntilOverlayReady: boolean;
    subminerBinaryPath: string;
    aniskipEnabled: boolean;
    aniskipButtonKey: string;
  };
  controller: {
    enabled: boolean;
    preferredGamepadId: string;
    preferredGamepadLabel: string;
    smoothScroll: boolean;
    scrollPixelsPerSecond: number;
    horizontalJumpPixels: number;
    stickDeadzone: number;
    triggerInputMode: ControllerTriggerInputMode;
    triggerDeadzone: number;
    repeatDelayMs: number;
    repeatIntervalMs: number;
    buttonIndices: Required<ControllerButtonIndicesConfig>;
    bindings: Required<ResolvedControllerBindingsConfig>;
    profiles: Record<string, ResolvedControllerProfileConfig>;
  };
  ankiConnect: AnkiConnectConfig & {
    enabled: boolean;
    url: string;
    pollingRate: number;
    proxy: {
      enabled: boolean;
      host: string;
      port: number;
      upstreamUrl: string;
    };
    tags: string[];
    fields: {
      word: string;
      audio: string;
      image: string;
      sentence: string;
      miscInfo: string;
      translation: string;
    };
    ai: AiFeatureConfig & {
      enabled: boolean;
    };
    media: {
      generateAudio: boolean;
      generateImage: boolean;
      imageType: 'static' | 'avif';
      imageFormat: 'jpg' | 'png' | 'webp';
      imageQuality: number;
      imageMaxWidth?: number;
      imageMaxHeight?: number;
      animatedFps: number;
      animatedMaxWidth: number;
      animatedMaxHeight?: number;
      animatedCrf: number;
      syncAnimatedImageToWordAudio: boolean;
      audioPadding: number;
      fallbackDuration: number;
      maxMediaDuration: number;
    };
    knownWords: {
      highlightEnabled: boolean;
      refreshMinutes: number;
      addMinedWordsImmediately: boolean;
      matchMode: NPlusOneMatchMode;
      decks: Record<string, string[]>;
    };
    nPlusOne: {
      enabled: boolean;
      minSentenceWords: number;
    };
    behavior: {
      overwriteAudio: boolean;
      overwriteImage: boolean;
      mediaInsertMode: 'append' | 'prepend';
      highlightWord: boolean;
      notificationType: NotificationType;
      autoUpdateNewCards: boolean;
    };
    metadata: {
      pattern: string;
    };
    isLapis: {
      enabled: boolean;
      sentenceCardModel: string;
    };
    isKiku: {
      enabled: boolean;
      fieldGrouping: 'auto' | 'manual' | 'disabled';
      deleteDuplicateInAuto: boolean;
    };
  };
  shortcuts: Required<ShortcutsConfig>;
  secondarySub: Required<SecondarySubConfig>;
  subsync: Required<SubsyncConfig>;
  startupWarmups: {
    lowPowerMode: boolean;
    mecab: boolean;
    yomitanExtension: boolean;
    subtitleDictionaries: boolean;
    jellyfinRemoteSession: boolean;
  };
  subtitleStyle: Required<Omit<SubtitleStyleConfig, 'secondary' | 'frequencyDictionary'>> & {
    secondary: Required<NonNullable<SubtitleStyleConfig['secondary']>>;
    frequencyDictionary: {
      enabled: boolean;
      sourcePath: string;
      topX: number;
      mode: FrequencyDictionaryMode;
      matchMode: FrequencyDictionaryMatchMode;
      singleColor: string;
      bandedColors: [string, string, string, string, string];
    };
  };
  subtitleSidebar: ResolvedSubtitleSidebarConfig;
  auto_start_overlay: boolean;
  uiLanguage: string;
  jimaku: JimakuConfig & {
    apiBaseUrl: string;
    languagePreference: JimakuLanguagePreference;
    maxEntryResults: number;
  };
  anilist: {
    enabled: boolean;
    accessToken: string;
    characterDictionary: {
      refreshTtlHours: number;
      maxLoaded: number;
      evictionPolicy: AnilistCharacterDictionaryEvictionPolicy;
      profileScope: AnilistCharacterDictionaryProfileScope;
      collapsibleSections: Required<AnilistCharacterDictionaryCollapsibleSectionsConfig>;
    };
  };
  yomitan: {
    externalProfilePath: string;
  };
  jellyfin: {
    enabled: boolean;
    serverUrl: string;
    recentServers: string[];
    username: string;
    defaultLibraryId: string;
    remoteControlEnabled: boolean;
    remoteControlAutoConnect: boolean;
    autoAnnounce: boolean;
    pullPictures: boolean;
    iconCacheDir: string;
    directPlayPreferred: boolean;
    directPlayContainers: string[];
    transcodeVideoCodec: string;
  };
  discordPresence: {
    enabled: boolean;
    presenceStyle: import('./integrations').DiscordPresenceStylePreset;
    updateIntervalMs: number;
    debounceMs: number;
  };
  ai: AiConfig & {
    enabled: boolean;
    apiKey: string;
    apiKeyCommand: string;
    baseUrl: string;
    model: string;
    systemPrompt: string;
    requestTimeoutMs: number;
  };
  youtube: YoutubeConfig & {
    primarySubLanguages: string[];
  };
  youtubeSubgen: YoutubeSubgenConfig & {
    whisperBin: string;
    whisperModel: string;
    whisperVadModel: string;
    whisperThreads: number;
    fixWithAi: boolean;
    ai: AiFeatureConfig;
  };
  immersionTracking: {
    enabled: boolean;
    dbPath?: string;
    batchSize: number;
    flushIntervalMs: number;
    queueCap: number;
    payloadCapBytes: number;
    maintenanceIntervalMs: number;
    retentionMode: ImmersionTrackingRetentionMode;
    retentionPreset: ImmersionTrackingRetentionPreset;
    retention: {
      eventsDays: number;
      telemetryDays: number;
      sessionsDays: number;
      dailyRollupsDays: number;
      monthlyRollupsDays: number;
      vacuumIntervalDays: number;
    };
    lifetimeSummaries: {
      global: boolean;
      anime: boolean;
      media: boolean;
    };
  };
  stats: {
    toggleKey: string;
    markWatchedKey: string;
    serverPort: number;
    autoStartServer: boolean;
    autoOpenBrowser: boolean;
  };
  updates: Required<UpdatesConfig>;
  notifications: Required<NotificationsConfig>;
  logging: {
    level: 'debug' | 'info' | 'warn' | 'error';
    rotation: LogRotation;
    files: Required<LogFilesConfig>;
  };
}

export interface ConfigValidationWarning {
  path: string;
  value: unknown;
  fallback: unknown;
  message: string;
}
