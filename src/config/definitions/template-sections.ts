import { ConfigTemplateSection } from './shared';

const CORE_TEMPLATE_SECTIONS: ConfigTemplateSection[] = [
  {
    title: 'Overlay Auto-Start',
    description: [
      'When overlay connects to mpv, automatically show overlay and hide mpv subtitles.',
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
    key: 'logging',
  },
  {
    title: 'Controller Support',
    description: [
      'Gamepad support for the visible overlay while keyboard-only mode is active.',
      'Use the selection modal to save a preferred controller by id for future launches.',
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
    title: 'Keyboard Shortcuts',
    description: ['Overlay keyboard shortcuts. Set a shortcut to null to disable.'],
    notes: ['Hot-reload: shortcut changes apply live and update the session help modal on reopen.'],
    key: 'shortcuts',
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
];

const SUBTITLE_TEMPLATE_SECTIONS: ConfigTemplateSection[] = [
  {
    title: 'Subtitle Appearance',
    description: ['Primary and secondary subtitle styling.'],
    notes: ['Hot-reload: subtitle style changes apply live without restarting SubMiner.'],
    key: 'subtitleStyle',
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
      'Hot-reload: ankiConnect.ai.enabled updates live while SubMiner is running.',
      'Shared AI provider transport settings are read from top-level ai and typically require restart.',
      'Most other AnkiConnect settings still require restart.',
    ],
    key: 'ankiConnect',
  },
  {
    title: 'Jimaku',
    description: ['Jimaku API configuration and defaults.'],
    key: 'jimaku',
  },
  {
    title: 'YouTube Subtitle Generation',
    description: ['Defaults for SubMiner YouTube subtitle generation.'],
    key: 'youtubeSubgen',
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
];

export const CONFIG_TEMPLATE_SECTIONS: ConfigTemplateSection[] = [
  ...CORE_TEMPLATE_SECTIONS,
  ...SUBTITLE_TEMPLATE_SECTIONS,
  ...INTEGRATION_TEMPLATE_SECTIONS,
  ...IMMERSION_TEMPLATE_SECTIONS,
];
