import { ConfigTemplateSection } from './shared';

const CORE_TEMPLATE_SECTIONS: ConfigTemplateSection[] = [
  {
    title: 'Visible Overlay Auto-Start',
    description: [
      'Show the visible subtitle overlay automatically after managed mpv playback starts SubMiner.',
      'SubMiner can still auto-start in the background when this is false.',
    ],
    key: 'auto_start_overlay',
  },
  {
    title: 'Texthooker Server',
    description: ['Configure texthooker startup launch and browser opening behavior.'],
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
    title: 'Annotation WebSocket',
    description: [
      'Dedicated annotated subtitle websocket for bundled texthooker and token-aware clients.',
      'Independent from websocket.auto and defaults to port 6678.',
    ],
    key: 'annotationWebsocket',
  },
  {
    title: 'Logging',
    description: ['Controls logging verbosity.', 'Set to debug for full runtime diagnostics.'],
    notes: ['Hot-reload: logging.level and logging.files apply live while SubMiner is running.'],
    key: 'logging',
  },
  {
    title: 'Controller Support',
    description: [
      'Gamepad support for the visible overlay while keyboard-only mode is active.',
      'Use Alt+C to pick a preferred controller and remap actions inline with learn mode.',
      'Trigger input mode can be auto, digital-only, or analog-thresholded depending on the controller.',
      'Override controller.buttonIndices when your pad reports non-standard raw button numbers.',
    ],
    key: 'controller',
  },
  {
    title: 'Startup Warmups',
    description: [
      'Background warmup controls for MeCab, Yomitan, dictionaries, and Jellyfin session.',
      'Disable individual warmups to defer load until first real usage.',
      'lowPowerMode defers all warmups except Yomitan extension.',
    ],
    key: 'startupWarmups',
  },
  {
    title: 'Updates',
    description: [
      'Automatic update check behavior.',
      'Manual checks from the tray or launcher are always allowed.',
    ],
    key: 'updates',
  },
  {
    title: 'Notifications',
    description: ['Overlay notification display behavior.'],
    notes: ['Hot-reload: position changes apply to the next overlay notification.'],
    key: 'notifications',
  },
  {
    title: 'Keyboard Shortcuts',
    description: ['Overlay keyboard shortcuts. Set a shortcut to null to disable.'],
    notes: ['Hot-reload: shortcut changes apply live and update the session help modal on reopen.'],
    key: 'shortcuts',
  },
  {
    title: 'Keybindings (MPV Commands)',
    description: [
      'Default and custom keybindings that are merged with built-in defaults.',
      'Set command to null to disable a default keybinding.',
    ],
    notes: [
      'Hot-reload: keybinding changes apply live and update the session help modal on reopen.',
    ],
    key: 'keybindings',
  },
  {
    title: 'Secondary Subtitles',
    description: [
      'Dual subtitle track options.',
      'Used by managed subtitle loading as secondary language preferences for local and YouTube playback.',
    ],
    notes: ['Hot-reload: defaultMode updates live while SubMiner is running.'],
    key: 'secondarySub',
  },
  {
    title: 'Subtitle Sync',
    description: ['Subsync engine and executable paths.'],
    notes: ['Hot-reload: subsync changes apply to the next subtitle sync run.'],
    key: 'subsync',
  },
  {
    title: 'Subtitle Position',
    description: ['Initial vertical subtitle position from the bottom.'],
    key: 'subtitlePosition',
  },
];

const SUBTITLE_TEMPLATE_SECTIONS: ConfigTemplateSection[] = [
  {
    title: 'Subtitle Appearance',
    description: ['Primary and secondary subtitle styling.'],
    notes: ['Hot-reload: subtitle style changes apply live without restarting SubMiner.'],
    key: 'subtitleStyle',
  },
  {
    title: 'Subtitle Sidebar',
    description: ['Parsed-subtitle sidebar cue list styling, behavior, and toggle key.'],
    notes: ['Hot-reload: subtitle sidebar changes apply live without restarting SubMiner.'],
    key: 'subtitleSidebar',
  },
];

const INTEGRATION_TEMPLATE_SECTIONS: ConfigTemplateSection[] = [
  {
    title: 'Shared AI Provider',
    description: [
      'Canonical OpenAI-compatible provider transport settings shared by Anki and YouTube subtitle fixing.',
    ],
    key: 'ai',
  },
  {
    title: 'AnkiConnect Integration',
    description: ['Automatic Anki updates and media generation options.'],
    notes: [
      'Hot-reload: ankiConnect.ai.enabled, media.normalizeAudio/mirrorMpvVolume, knownWords, nPlusOne, fields.word/audio/image/sentence/miscInfo, behavior.autoUpdateNewCards, isLapis.sentenceCardModel, and isKiku.fieldGrouping update live while SubMiner is running.',
      'Shared AI provider transport settings are read from top-level ai and typically require restart.',
      'Most other AnkiConnect settings still require restart.',
    ],
    key: 'ankiConnect',
  },
  {
    title: 'Jimaku',
    description: ['Jimaku API configuration and defaults.'],
    notes: ['Hot-reload: Jimaku changes apply to the next Jimaku request.'],
    key: 'jimaku',
  },
  {
    title: 'TsukiHime',
    description: [
      'TsukiHime subtitle search configuration (English and Japanese). No API key required.',
      'TsukiHime is the successor to Animetosho; the config key keeps the legacy name.',
    ],
    notes: ['Hot-reload: TsukiHime changes apply to the next TsukiHime request.'],
    key: 'animetosho',
  },
  {
    title: 'YouTube Playback Settings',
    description: [
      'Defaults for managed subtitle language preferences and YouTube subtitle loading.',
    ],
    notes: ['Hot-reload: primarySubLanguages applies to the next YouTube subtitle load.'],
    key: 'youtube',
  },
  {
    title: 'Anilist',
    description: [
      'Anilist API credentials and update behavior.',
      'Includes optional auto-sync for a merged MRU-based character dictionary in bundled Yomitan.',
      'Character dictionaries are keyed by AniList media ID (no season/franchise merge).',
    ],
    key: 'anilist',
  },
  {
    title: 'Yomitan',
    description: [
      'Optional external Yomitan profile integration.',
      'Setting yomitan.externalProfilePath switches SubMiner to read-only external-profile mode.',
      'For GameSentenceMiner on Linux, the default overlay profile is usually ~/.config/gsm_overlay.',
      'In external-profile mode SubMiner will not import, delete, or modify Yomitan dictionaries/settings.',
    ],
    key: 'yomitan',
  },
  {
    title: 'MPV Launcher',
    description: [
      'SubMiner-managed mpv launch and bundled plugin options.',
      'Set mpv.socketPath to the IPC socket used by the launcher, Electron app, and bundled plugin.',
      'autoStartSubMiner starts SubMiner in the background; auto_start_overlay only controls visible overlay display.',
      'Set mpv.launchMode to choose normal, maximized, or fullscreen SubMiner-managed mpv playback.',
      'Set mpv.profile to pass an mpv profile to managed mpv launches; leave it blank to pass none.',
      'Leave mpv.executablePath blank to auto-discover mpv.exe from SUBMINER_MPV_PATH or PATH.',
    ],
    key: 'mpv',
  },
  {
    title: 'Jellyfin',
    description: [
      'Optional Jellyfin integration for auth, browsing, and playback launch.',
      'Access token is stored in local encrypted token storage after login/setup.',
      'jellyfin.accessToken remains an optional explicit override in config.',
    ],
    key: 'jellyfin',
  },
  {
    title: 'Discord Rich Presence',
    description: [
      'Optional Discord Rich Presence activity card updates for current playback/study session.',
      'Uses official SubMiner Discord app assets for polished card visuals.',
    ],
    key: 'discordPresence',
  },
];

const IMMERSION_TEMPLATE_SECTIONS: ConfigTemplateSection[] = [
  {
    title: 'Immersion Tracking',
    description: [
      'Enable/disable immersion tracking.',
      'Set dbPath to override the default sqlite database location.',
      'Policy tuning is available for queue, flush, and retention values.',
    ],
    key: 'immersionTracking',
  },
  {
    title: 'Stats Dashboard',
    description: [
      'Local immersion stats dashboard served on localhost and available as an in-app overlay.',
      'Uses the immersion tracking database for overview, trends, sessions, and vocabulary views.',
    ],
    key: 'stats',
  },
];

export const CONFIG_TEMPLATE_SECTIONS: ConfigTemplateSection[] = [
  ...CORE_TEMPLATE_SECTIONS,
  ...SUBTITLE_TEMPLATE_SECTIONS,
  ...INTEGRATION_TEMPLATE_SECTIONS,
  ...IMMERSION_TEMPLATE_SECTIONS,
];
