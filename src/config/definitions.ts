import {
  AnkiConnectConfig,
  Config,
  RawConfig,
  ResolvedConfig,
  RuntimeOptionId,
  RuntimeOptionScope,
  RuntimeOptionValue,
  RuntimeOptionValueType,
} from '../types';

export type ConfigValueKind = 'boolean' | 'number' | 'string' | 'enum' | 'array' | 'object';

export interface RuntimeOptionRegistryEntry {
  id: RuntimeOptionId;
  path: string;
  label: string;
  scope: RuntimeOptionScope;
  valueType: RuntimeOptionValueType;
  allowedValues: RuntimeOptionValue[];
  defaultValue: RuntimeOptionValue;
  requiresRestart: boolean;
  formatValueForOsd: (value: RuntimeOptionValue) => string;
  toAnkiPatch: (value: RuntimeOptionValue) => Partial<AnkiConnectConfig>;
}

export interface ConfigOptionRegistryEntry {
  path: string;
  kind: ConfigValueKind;
  defaultValue: unknown;
  description: string;
  enumValues?: readonly string[];
  runtime?: RuntimeOptionRegistryEntry;
}

export interface ConfigTemplateSection {
  title: string;
  description: string[];
  key: keyof ResolvedConfig;
  notes?: string[];
}

export const SPECIAL_COMMANDS = {
  SUBSYNC_TRIGGER: '__subsync-trigger',
  RUNTIME_OPTIONS_OPEN: '__runtime-options-open',
  RUNTIME_OPTION_CYCLE_PREFIX: '__runtime-option-cycle:',
  REPLAY_SUBTITLE: '__replay-subtitle',
  PLAY_NEXT_SUBTITLE: '__play-next-subtitle',
} as const;

export const DEFAULT_KEYBINDINGS: NonNullable<ResolvedConfig['keybindings']> = [
  { key: 'Space', command: ['cycle', 'pause'] },
  { key: 'ArrowRight', command: ['seek', 5] },
  { key: 'ArrowLeft', command: ['seek', -5] },
  { key: 'ArrowUp', command: ['seek', 60] },
  { key: 'ArrowDown', command: ['seek', -60] },
  { key: 'Shift+KeyH', command: ['sub-seek', -1] },
  { key: 'Shift+KeyL', command: ['sub-seek', 1] },
  { key: 'Ctrl+Shift+KeyH', command: [SPECIAL_COMMANDS.REPLAY_SUBTITLE] },
  { key: 'Ctrl+Shift+KeyL', command: [SPECIAL_COMMANDS.PLAY_NEXT_SUBTITLE] },
  { key: 'KeyQ', command: ['quit'] },
  { key: 'Ctrl+KeyW', command: ['quit'] },
];

export const DEFAULT_CONFIG: ResolvedConfig = {
  subtitlePosition: { yPercent: 10 },
  keybindings: [],
  websocket: {
    enabled: 'auto',
    port: 6677,
  },
  logging: {
    level: 'info',
  },
  texthooker: {
    openBrowser: true,
  },
  ankiConnect: {
    enabled: false,
    url: 'http://127.0.0.1:8765',
    pollingRate: 3000,
    tags: ['SubMiner'],
    fields: {
      audio: 'ExpressionAudio',
      image: 'Picture',
      sentence: 'Sentence',
      miscInfo: 'MiscInfo',
      translation: 'SelectionText',
    },
    ai: {
      enabled: false,
      alwaysUseAiTranslation: false,
      apiKey: '',
      model: 'openai/gpt-4o-mini',
      baseUrl: 'https://openrouter.ai/api',
      targetLanguage: 'English',
      systemPrompt:
        'You are a translation engine. Return only the translated text with no explanations.',
    },
    media: {
      generateAudio: true,
      generateImage: true,
      imageType: 'static',
      imageFormat: 'jpg',
      imageQuality: 92,
      imageMaxWidth: undefined,
      imageMaxHeight: undefined,
      animatedFps: 10,
      animatedMaxWidth: 640,
      animatedMaxHeight: undefined,
      animatedCrf: 35,
      audioPadding: 0.5,
      fallbackDuration: 3.0,
      maxMediaDuration: 30,
    },
    behavior: {
      overwriteAudio: true,
      overwriteImage: true,
      mediaInsertMode: 'append',
      highlightWord: true,
      notificationType: 'osd',
      autoUpdateNewCards: true,
    },
    nPlusOne: {
      highlightEnabled: false,
      refreshMinutes: 1440,
      matchMode: 'headword',
      decks: [],
      minSentenceWords: 3,
      nPlusOne: '#c6a0f6',
      knownWord: '#a6da95',
    },
    metadata: {
      pattern: '[SubMiner] %f (%t)',
    },
    isLapis: {
      enabled: false,
      sentenceCardModel: 'Japanese sentences',
    },
    isKiku: {
      enabled: false,
      fieldGrouping: 'disabled',
      deleteDuplicateInAuto: true,
    },
  },
  shortcuts: {
    toggleVisibleOverlayGlobal: 'Alt+Shift+O',
    toggleInvisibleOverlayGlobal: 'Alt+Shift+I',
    copySubtitle: 'CommandOrControl+C',
    copySubtitleMultiple: 'CommandOrControl+Shift+C',
    updateLastCardFromClipboard: 'CommandOrControl+V',
    triggerFieldGrouping: 'CommandOrControl+G',
    triggerSubsync: 'Ctrl+Alt+S',
    mineSentence: 'CommandOrControl+S',
    mineSentenceMultiple: 'CommandOrControl+Shift+S',
    multiCopyTimeoutMs: 3000,
    toggleSecondarySub: 'CommandOrControl+Shift+V',
    markAudioCard: 'CommandOrControl+Shift+A',
    openRuntimeOptions: 'CommandOrControl+Shift+O',
    openJimaku: 'Ctrl+Shift+J',
  },
  secondarySub: {
    secondarySubLanguages: [],
    autoLoadSecondarySub: false,
    defaultMode: 'hover',
  },
  subsync: {
    defaultMode: 'auto',
    alass_path: '',
    ffsubsync_path: '',
    ffmpeg_path: '',
  },
  subtitleStyle: {
    enableJlpt: false,
    fontFamily: 'Noto Sans CJK JP Regular, Noto Sans CJK JP, Arial Unicode MS, Arial, sans-serif',
    fontSize: 35,
    fontColor: '#cad3f5',
    fontWeight: 'normal',
    fontStyle: 'normal',
    backgroundColor: 'rgba(54, 58, 79, 0.5)',
    nPlusOneColor: '#c6a0f6',
    knownWordColor: '#a6da95',
    jlptColors: {
      N1: '#ed8796',
      N2: '#f5a97f',
      N3: '#f9e2af',
      N4: '#a6e3a1',
      N5: '#8aadf4',
    },
    frequencyDictionary: {
      enabled: false,
      sourcePath: '',
      topX: 1000,
      mode: 'single',
      singleColor: '#f5a97f',
      bandedColors: ['#ed8796', '#f5a97f', '#f9e2af', '#a6e3a1', '#8aadf4'],
    },
    secondary: {
      fontSize: 24,
      fontColor: '#ffffff',
      backgroundColor: 'transparent',
      fontWeight: 'normal',
      fontStyle: 'normal',
      fontFamily: 'Noto Sans CJK JP Regular, Noto Sans CJK JP, Arial Unicode MS, Arial, sans-serif',
    },
  },
  auto_start_overlay: false,
  bind_visible_overlay_to_mpv_sub_visibility: true,
  jimaku: {
    apiBaseUrl: 'https://jimaku.cc',
    languagePreference: 'ja',
    maxEntryResults: 10,
  },
  anilist: {
    enabled: false,
    accessToken: '',
  },
  jellyfin: {
    enabled: false,
    serverUrl: '',
    username: '',
    accessToken: '',
    userId: '',
    deviceId: 'subminer',
    clientName: 'SubMiner',
    clientVersion: '0.1.0',
    defaultLibraryId: '',
    remoteControlEnabled: true,
    remoteControlAutoConnect: true,
    autoAnnounce: false,
    remoteControlDeviceName: 'SubMiner',
    pullPictures: false,
    iconCacheDir: '/tmp/subminer-jellyfin-icons',
    directPlayPreferred: true,
    directPlayContainers: ['mkv', 'mp4', 'webm', 'mov', 'flac', 'mp3', 'aac'],
    transcodeVideoCodec: 'h264',
  },
  youtubeSubgen: {
    mode: 'automatic',
    whisperBin: '',
    whisperModel: '',
    primarySubLanguages: ['ja', 'jpn'],
  },
  invisibleOverlay: {
    startupVisibility: 'platform-default',
  },
  immersionTracking: {
    enabled: true,
    dbPath: '',
    batchSize: 25,
    flushIntervalMs: 500,
    queueCap: 1000,
    payloadCapBytes: 256,
    maintenanceIntervalMs: 24 * 60 * 60 * 1000,
    retention: {
      eventsDays: 7,
      telemetryDays: 30,
      dailyRollupsDays: 365,
      monthlyRollupsDays: 5 * 365,
      vacuumIntervalDays: 7,
    },
  },
};

export const DEFAULT_ANKI_CONNECT_CONFIG = DEFAULT_CONFIG.ankiConnect;

export const RUNTIME_OPTION_REGISTRY: RuntimeOptionRegistryEntry[] = [
  {
    id: 'anki.autoUpdateNewCards',
    path: 'ankiConnect.behavior.autoUpdateNewCards',
    label: 'Auto Update New Cards',
    scope: 'ankiConnect',
    valueType: 'boolean',
    allowedValues: [true, false],
    defaultValue: DEFAULT_CONFIG.ankiConnect.behavior.autoUpdateNewCards,
    requiresRestart: false,
    formatValueForOsd: (value) => (value === true ? 'On' : 'Off'),
    toAnkiPatch: (value) => ({
      behavior: { autoUpdateNewCards: value === true },
    }),
  },
  {
    id: 'anki.nPlusOneMatchMode',
    path: 'ankiConnect.nPlusOne.matchMode',
    label: 'N+1 Match Mode',
    scope: 'ankiConnect',
    valueType: 'enum',
    allowedValues: ['headword', 'surface'],
    defaultValue: DEFAULT_CONFIG.ankiConnect.nPlusOne.matchMode,
    requiresRestart: false,
    formatValueForOsd: (value) => String(value),
    toAnkiPatch: (value) => ({
      nPlusOne: {
        matchMode: value === 'headword' || value === 'surface' ? value : 'headword',
      },
    }),
  },
  {
    id: 'anki.kikuFieldGrouping',
    path: 'ankiConnect.isKiku.fieldGrouping',
    label: 'Kiku Field Grouping',
    scope: 'ankiConnect',
    valueType: 'enum',
    allowedValues: ['auto', 'manual', 'disabled'],
    defaultValue: 'disabled',
    requiresRestart: false,
    formatValueForOsd: (value) => String(value),
    toAnkiPatch: (value) => ({
      isKiku: {
        fieldGrouping:
          value === 'auto' || value === 'manual' || value === 'disabled' ? value : 'disabled',
      },
    }),
  },
];

export const CONFIG_OPTION_REGISTRY: ConfigOptionRegistryEntry[] = [
  {
    path: 'logging.level',
    kind: 'enum',
    enumValues: ['debug', 'info', 'warn', 'error'],
    defaultValue: DEFAULT_CONFIG.logging.level,
    description: 'Minimum log level for runtime logging.',
  },
  {
    path: 'websocket.enabled',
    kind: 'enum',
    enumValues: ['auto', 'true', 'false'],
    defaultValue: DEFAULT_CONFIG.websocket.enabled,
    description: 'Built-in subtitle websocket server mode.',
  },
  {
    path: 'websocket.port',
    kind: 'number',
    defaultValue: DEFAULT_CONFIG.websocket.port,
    description: 'Built-in subtitle websocket server port.',
  },
  {
    path: 'subtitleStyle.enableJlpt',
    kind: 'boolean',
    defaultValue: DEFAULT_CONFIG.subtitleStyle.enableJlpt,
    description:
      'Enable JLPT vocabulary level underlines. ' +
      'When disabled, JLPT tagging lookup and underlines are skipped.',
  },
  {
    path: 'subtitleStyle.frequencyDictionary.enabled',
    kind: 'boolean',
    defaultValue: DEFAULT_CONFIG.subtitleStyle.frequencyDictionary.enabled,
    description: 'Enable frequency-dictionary-based highlighting based on token rank.',
  },
  {
    path: 'subtitleStyle.frequencyDictionary.sourcePath',
    kind: 'string',
    defaultValue: DEFAULT_CONFIG.subtitleStyle.frequencyDictionary.sourcePath,
    description:
      'Optional absolute path to a frequency dictionary directory.' +
      ' If empty, built-in discovery search paths are used.',
  },
  {
    path: 'subtitleStyle.frequencyDictionary.topX',
    kind: 'number',
    defaultValue: DEFAULT_CONFIG.subtitleStyle.frequencyDictionary.topX,
    description: 'Only color tokens with frequency rank <= topX (default: 1000).',
  },
  {
    path: 'subtitleStyle.frequencyDictionary.mode',
    kind: 'enum',
    enumValues: ['single', 'banded'],
    defaultValue: DEFAULT_CONFIG.subtitleStyle.frequencyDictionary.mode,
    description:
      'single: use one color for all matching tokens. banded: use color ramp by frequency band.',
  },
  {
    path: 'subtitleStyle.frequencyDictionary.singleColor',
    kind: 'string',
    defaultValue: DEFAULT_CONFIG.subtitleStyle.frequencyDictionary.singleColor,
    description: 'Color used when frequencyDictionary.mode is `single`.',
  },
  {
    path: 'subtitleStyle.frequencyDictionary.bandedColors',
    kind: 'array',
    defaultValue: DEFAULT_CONFIG.subtitleStyle.frequencyDictionary.bandedColors,
    description:
      'Five colors used for rank bands when mode is `banded` (from most common to least within topX).',
  },
  {
    path: 'ankiConnect.enabled',
    kind: 'boolean',
    defaultValue: DEFAULT_CONFIG.ankiConnect.enabled,
    description: 'Enable AnkiConnect integration.',
  },
  {
    path: 'ankiConnect.pollingRate',
    kind: 'number',
    defaultValue: DEFAULT_CONFIG.ankiConnect.pollingRate,
    description: 'Polling interval in milliseconds.',
  },
  {
    path: 'ankiConnect.tags',
    kind: 'array',
    defaultValue: DEFAULT_CONFIG.ankiConnect.tags,
    description:
      'Tags to add to cards mined or updated by SubMiner. Provide an empty array to disable automatic tagging.',
  },
  {
    path: 'ankiConnect.behavior.autoUpdateNewCards',
    kind: 'boolean',
    defaultValue: DEFAULT_CONFIG.ankiConnect.behavior.autoUpdateNewCards,
    description: 'Automatically update newly added cards.',
    runtime: RUNTIME_OPTION_REGISTRY[0],
  },
  {
    path: 'ankiConnect.nPlusOne.matchMode',
    kind: 'enum',
    enumValues: ['headword', 'surface'],
    defaultValue: DEFAULT_CONFIG.ankiConnect.nPlusOne.matchMode,
    description: 'Known-word matching strategy for N+1 highlighting.',
  },
  {
    path: 'ankiConnect.nPlusOne.highlightEnabled',
    kind: 'boolean',
    defaultValue: DEFAULT_CONFIG.ankiConnect.nPlusOne.highlightEnabled,
    description: 'Enable fast local highlighting for words already known in Anki.',
  },
  {
    path: 'ankiConnect.nPlusOne.refreshMinutes',
    kind: 'number',
    defaultValue: DEFAULT_CONFIG.ankiConnect.nPlusOne.refreshMinutes,
    description: 'Minutes between known-word cache refreshes.',
  },
  {
    path: 'ankiConnect.nPlusOne.minSentenceWords',
    kind: 'number',
    defaultValue: DEFAULT_CONFIG.ankiConnect.nPlusOne.minSentenceWords,
    description: 'Minimum sentence word count required for N+1 targeting (default: 3).',
  },
  {
    path: 'ankiConnect.nPlusOne.decks',
    kind: 'array',
    defaultValue: DEFAULT_CONFIG.ankiConnect.nPlusOne.decks,
    description: 'Decks used for N+1 known-word cache scope. Supports one or more deck names.',
  },
  {
    path: 'ankiConnect.nPlusOne.nPlusOne',
    kind: 'string',
    defaultValue: DEFAULT_CONFIG.ankiConnect.nPlusOne.nPlusOne,
    description: 'Color used for the single N+1 target token highlight.',
  },
  {
    path: 'ankiConnect.nPlusOne.knownWord',
    kind: 'string',
    defaultValue: DEFAULT_CONFIG.ankiConnect.nPlusOne.knownWord,
    description: 'Color used for legacy known-word highlights.',
  },
  {
    path: 'ankiConnect.isKiku.fieldGrouping',
    kind: 'enum',
    enumValues: ['auto', 'manual', 'disabled'],
    defaultValue: DEFAULT_CONFIG.ankiConnect.isKiku.fieldGrouping,
    description: 'Kiku duplicate-card field grouping mode.',
    runtime: RUNTIME_OPTION_REGISTRY[1],
  },
  {
    path: 'subsync.defaultMode',
    kind: 'enum',
    enumValues: ['auto', 'manual'],
    defaultValue: DEFAULT_CONFIG.subsync.defaultMode,
    description: 'Subsync default mode.',
  },
  {
    path: 'shortcuts.multiCopyTimeoutMs',
    kind: 'number',
    defaultValue: DEFAULT_CONFIG.shortcuts.multiCopyTimeoutMs,
    description: 'Timeout for multi-copy/mine modes.',
  },
  {
    path: 'bind_visible_overlay_to_mpv_sub_visibility',
    kind: 'boolean',
    defaultValue: DEFAULT_CONFIG.bind_visible_overlay_to_mpv_sub_visibility,
    description: 'Link visible overlay toggles to MPV subtitle visibility (primary and secondary).',
  },
  {
    path: 'jimaku.languagePreference',
    kind: 'enum',
    enumValues: ['ja', 'en', 'none'],
    defaultValue: DEFAULT_CONFIG.jimaku.languagePreference,
    description: 'Preferred language used in Jimaku search.',
  },
  {
    path: 'jimaku.maxEntryResults',
    kind: 'number',
    defaultValue: DEFAULT_CONFIG.jimaku.maxEntryResults,
    description: 'Maximum Jimaku search results returned.',
  },
  {
    path: 'anilist.enabled',
    kind: 'boolean',
    defaultValue: DEFAULT_CONFIG.anilist.enabled,
    description: 'Enable AniList post-watch progress updates.',
  },
  {
    path: 'anilist.accessToken',
    kind: 'string',
    defaultValue: DEFAULT_CONFIG.anilist.accessToken,
    description: 'AniList access token used for post-watch updates.',
  },
  {
    path: 'jellyfin.enabled',
    kind: 'boolean',
    defaultValue: DEFAULT_CONFIG.jellyfin.enabled,
    description: 'Enable optional Jellyfin integration and CLI control commands.',
  },
  {
    path: 'jellyfin.serverUrl',
    kind: 'string',
    defaultValue: DEFAULT_CONFIG.jellyfin.serverUrl,
    description: 'Base Jellyfin server URL (for example: http://localhost:8096).',
  },
  {
    path: 'jellyfin.username',
    kind: 'string',
    defaultValue: DEFAULT_CONFIG.jellyfin.username,
    description: 'Default Jellyfin username used during CLI login.',
  },
  {
    path: 'jellyfin.defaultLibraryId',
    kind: 'string',
    defaultValue: DEFAULT_CONFIG.jellyfin.defaultLibraryId,
    description: 'Optional default Jellyfin library ID for item listing.',
  },
  {
    path: 'jellyfin.remoteControlEnabled',
    kind: 'boolean',
    defaultValue: DEFAULT_CONFIG.jellyfin.remoteControlEnabled,
    description: 'Enable Jellyfin remote cast control mode.',
  },
  {
    path: 'jellyfin.remoteControlAutoConnect',
    kind: 'boolean',
    defaultValue: DEFAULT_CONFIG.jellyfin.remoteControlAutoConnect,
    description: 'Auto-connect to the configured remote control target.',
  },
  {
    path: 'jellyfin.autoAnnounce',
    kind: 'boolean',
    defaultValue: DEFAULT_CONFIG.jellyfin.autoAnnounce,
    description:
      'When enabled, automatically trigger remote announce/visibility check on websocket connect.',
  },
  {
    path: 'jellyfin.remoteControlDeviceName',
    kind: 'string',
    defaultValue: DEFAULT_CONFIG.jellyfin.remoteControlDeviceName,
    description: 'Device name reported for Jellyfin remote control sessions.',
  },
  {
    path: 'jellyfin.pullPictures',
    kind: 'boolean',
    defaultValue: DEFAULT_CONFIG.jellyfin.pullPictures,
    description: 'Enable Jellyfin poster/icon fetching for launcher menus.',
  },
  {
    path: 'jellyfin.iconCacheDir',
    kind: 'string',
    defaultValue: DEFAULT_CONFIG.jellyfin.iconCacheDir,
    description: 'Directory used by launcher for cached Jellyfin poster icons.',
  },
  {
    path: 'jellyfin.directPlayPreferred',
    kind: 'boolean',
    defaultValue: DEFAULT_CONFIG.jellyfin.directPlayPreferred,
    description: 'Try direct play before server-managed transcoding when possible.',
  },
  {
    path: 'jellyfin.directPlayContainers',
    kind: 'array',
    defaultValue: DEFAULT_CONFIG.jellyfin.directPlayContainers,
    description: 'Container allowlist for direct play decisions.',
  },
  {
    path: 'jellyfin.transcodeVideoCodec',
    kind: 'string',
    defaultValue: DEFAULT_CONFIG.jellyfin.transcodeVideoCodec,
    description: 'Preferred transcode video codec when direct play is unavailable.',
  },
  {
    path: 'youtubeSubgen.mode',
    kind: 'enum',
    enumValues: ['automatic', 'preprocess', 'off'],
    defaultValue: DEFAULT_CONFIG.youtubeSubgen.mode,
    description: 'YouTube subtitle generation mode for the launcher script.',
  },
  {
    path: 'youtubeSubgen.whisperBin',
    kind: 'string',
    defaultValue: DEFAULT_CONFIG.youtubeSubgen.whisperBin,
    description: 'Path to whisper.cpp CLI used as fallback transcription engine.',
  },
  {
    path: 'youtubeSubgen.whisperModel',
    kind: 'string',
    defaultValue: DEFAULT_CONFIG.youtubeSubgen.whisperModel,
    description: 'Path to whisper model used for fallback transcription.',
  },
  {
    path: 'youtubeSubgen.primarySubLanguages',
    kind: 'string',
    defaultValue: DEFAULT_CONFIG.youtubeSubgen.primarySubLanguages.join(','),
    description: 'Comma-separated primary subtitle language priority used by the launcher.',
  },
  {
    path: 'immersionTracking.enabled',
    kind: 'boolean',
    defaultValue: DEFAULT_CONFIG.immersionTracking.enabled,
    description: 'Enable immersion tracking for mined subtitle metadata.',
  },
  {
    path: 'immersionTracking.dbPath',
    kind: 'string',
    defaultValue: DEFAULT_CONFIG.immersionTracking.dbPath,
    description:
      'Optional SQLite database path for immersion tracking. Empty value uses the default app data path.',
  },
  {
    path: 'immersionTracking.batchSize',
    kind: 'number',
    defaultValue: DEFAULT_CONFIG.immersionTracking.batchSize,
    description: 'Buffered telemetry/event writes per SQLite transaction.',
  },
  {
    path: 'immersionTracking.flushIntervalMs',
    kind: 'number',
    defaultValue: DEFAULT_CONFIG.immersionTracking.flushIntervalMs,
    description: 'Max delay before queue flush in milliseconds.',
  },
  {
    path: 'immersionTracking.queueCap',
    kind: 'number',
    defaultValue: DEFAULT_CONFIG.immersionTracking.queueCap,
    description: 'In-memory write queue cap before overflow policy applies.',
  },
  {
    path: 'immersionTracking.payloadCapBytes',
    kind: 'number',
    defaultValue: DEFAULT_CONFIG.immersionTracking.payloadCapBytes,
    description: 'Max JSON payload size per event before truncation.',
  },
  {
    path: 'immersionTracking.maintenanceIntervalMs',
    kind: 'number',
    defaultValue: DEFAULT_CONFIG.immersionTracking.maintenanceIntervalMs,
    description: 'Maintenance cadence (prune + rollup + vacuum checks).',
  },
  {
    path: 'immersionTracking.retention.eventsDays',
    kind: 'number',
    defaultValue: DEFAULT_CONFIG.immersionTracking.retention.eventsDays,
    description: 'Raw event retention window in days.',
  },
  {
    path: 'immersionTracking.retention.telemetryDays',
    kind: 'number',
    defaultValue: DEFAULT_CONFIG.immersionTracking.retention.telemetryDays,
    description: 'Telemetry retention window in days.',
  },
  {
    path: 'immersionTracking.retention.dailyRollupsDays',
    kind: 'number',
    defaultValue: DEFAULT_CONFIG.immersionTracking.retention.dailyRollupsDays,
    description: 'Daily rollup retention window in days.',
  },
  {
    path: 'immersionTracking.retention.monthlyRollupsDays',
    kind: 'number',
    defaultValue: DEFAULT_CONFIG.immersionTracking.retention.monthlyRollupsDays,
    description: 'Monthly rollup retention window in days.',
  },
  {
    path: 'immersionTracking.retention.vacuumIntervalDays',
    kind: 'number',
    defaultValue: DEFAULT_CONFIG.immersionTracking.retention.vacuumIntervalDays,
    description: 'Minimum days between VACUUM runs.',
  },
];

export const CONFIG_TEMPLATE_SECTIONS: ConfigTemplateSection[] = [
  {
    title: 'Overlay Auto-Start',
    description: [
      'When overlay connects to mpv, automatically show overlay and hide mpv subtitles.',
    ],
    key: 'auto_start_overlay',
  },
  {
    title: 'Visible Overlay Subtitle Binding',
    description: [
      'Control whether visible overlay toggles also toggle MPV subtitle visibility.',
      'When enabled, visible overlay hides MPV subtitles; when disabled, MPV subtitles are left unchanged.',
    ],
    key: 'bind_visible_overlay_to_mpv_sub_visibility',
  },
  {
    title: 'Texthooker Server',
    description: ['Control whether browser opens automatically for texthooker.'],
    key: 'texthooker',
  },
  {
    title: 'WebSocket Server',
    description: [
      'Built-in WebSocket server broadcasts subtitle text to connected clients.',
      'Auto mode disables built-in server if mpv_websocket is detected.',
    ],
    key: 'websocket',
  },
  {
    title: 'Logging',
    description: ['Controls logging verbosity.', 'Set to debug for full runtime diagnostics.'],
    key: 'logging',
  },
  {
    title: 'AnkiConnect Integration',
    description: ['Automatic Anki updates and media generation options.'],
    notes: [
      'Hot-reload: AI translation settings update live while SubMiner is running.',
      'Most other AnkiConnect settings still require restart.',
    ],
    key: 'ankiConnect',
  },
  {
    title: 'Keyboard Shortcuts',
    description: ['Overlay keyboard shortcuts. Set a shortcut to null to disable.'],
    notes: ['Hot-reload: shortcut changes apply live and update the session help modal on reopen.'],
    key: 'shortcuts',
  },
  {
    title: 'Invisible Overlay',
    description: ['Startup behavior for the invisible interactive subtitle mining layer.'],
    notes: [
      'Invisible subtitle position edit mode: Ctrl/Cmd+Shift+P to toggle, arrow keys to move, Enter or Ctrl/Cmd+S to save, Esc to cancel.',
      'This edit-mode shortcut is fixed and is not currently configurable.',
    ],
    key: 'invisibleOverlay',
  },
  {
    title: 'Keybindings (MPV Commands)',
    description: [
      'Extra keybindings that are merged with built-in defaults.',
      'Set command to null to disable a default keybinding.',
    ],
    notes: [
      'Hot-reload: keybinding changes apply live and update the session help modal on reopen.',
    ],
    key: 'keybindings',
  },
  {
    title: 'Subtitle Appearance',
    description: ['Primary and secondary subtitle styling.'],
    notes: ['Hot-reload: subtitle style changes apply live without restarting SubMiner.'],
    key: 'subtitleStyle',
  },
  {
    title: 'Secondary Subtitles',
    description: [
      'Dual subtitle track options.',
      'Used by subminer YouTube subtitle generation as secondary language preferences.',
    ],
    notes: ['Hot-reload: defaultMode updates live while SubMiner is running.'],
    key: 'secondarySub',
  },
  {
    title: 'Auto Subtitle Sync',
    description: ['Subsync engine and executable paths.'],
    key: 'subsync',
  },
  {
    title: 'Subtitle Position',
    description: ['Initial vertical subtitle position from the bottom.'],
    key: 'subtitlePosition',
  },
  {
    title: 'Jimaku',
    description: ['Jimaku API configuration and defaults.'],
    key: 'jimaku',
  },
  {
    title: 'YouTube Subtitle Generation',
    description: ['Defaults for subminer YouTube subtitle extraction/transcription mode.'],
    key: 'youtubeSubgen',
  },
  {
    title: 'Anilist',
    description: ['Anilist API credentials and update behavior.'],
    key: 'anilist',
  },
  {
    title: 'Jellyfin',
    description: [
      'Optional Jellyfin integration for auth, browsing, and playback launch.',
      'Access token is stored in config and should be treated as a secret.',
    ],
    key: 'jellyfin',
  },
  {
    title: 'Immersion Tracking',
    description: [
      'Enable/disable immersion tracking.',
      'Set dbPath to override the default sqlite database location.',
      'Policy tuning is available for queue, flush, and retention values.',
    ],
    key: 'immersionTracking',
  },
];

export function deepCloneConfig(config: ResolvedConfig): ResolvedConfig {
  return JSON.parse(JSON.stringify(config)) as ResolvedConfig;
}

export function deepMergeRawConfig(base: RawConfig, patch: RawConfig): RawConfig {
  const clone = JSON.parse(JSON.stringify(base)) as Record<string, unknown>;
  const patchObject = patch as Record<string, unknown>;

  const mergeInto = (target: Record<string, unknown>, source: Record<string, unknown>): void => {
    for (const [key, value] of Object.entries(source)) {
      if (
        value !== null &&
        typeof value === 'object' &&
        !Array.isArray(value) &&
        typeof target[key] === 'object' &&
        target[key] !== null &&
        !Array.isArray(target[key])
      ) {
        mergeInto(target[key] as Record<string, unknown>, value as Record<string, unknown>);
      } else {
        target[key] = value;
      }
    }
  };

  mergeInto(clone, patchObject);
  return clone as RawConfig;
}
