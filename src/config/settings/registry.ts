import type { ResolvedConfig } from '../../types/config';
import type {
  ConfigSettingsCategory,
  ConfigSettingsControl,
  ConfigSettingsField,
  ConfigSettingsRestartBehavior,
} from '../../types/settings';
import { CONFIG_OPTION_REGISTRY, DEFAULT_CONFIG } from '../definitions';
import { i18n } from '../../i18n/index.js';
import {
  getSubtitleCssManagedConfigPaths,
  getSubtitleCssScopeForPath,
} from '../../settings/subtitle-style-css';

type Leaf = {
  path: string;
  value: unknown;
};

export const LEGACY_HIDDEN_CONFIG_PATHS = [
  'ankiConnect.wordField',
  'ankiConnect.audioField',
  'ankiConnect.imageField',
  'ankiConnect.sentenceField',
  'ankiConnect.miscInfoField',
  'ankiConnect.miscInfoPattern',
  'ankiConnect.generateAudio',
  'ankiConnect.generateImage',
  'ankiConnect.imageType',
  'ankiConnect.imageFormat',
  'ankiConnect.imageQuality',
  'ankiConnect.imageMaxWidth',
  'ankiConnect.imageMaxHeight',
  'ankiConnect.animatedFps',
  'ankiConnect.animatedMaxWidth',
  'ankiConnect.animatedMaxHeight',
  'ankiConnect.animatedCrf',
  'ankiConnect.syncAnimatedImageToWordAudio',
  'ankiConnect.audioPadding',
  'ankiConnect.fallbackDuration',
  'ankiConnect.maxMediaDuration',
  'ankiConnect.overwriteAudio',
  'ankiConnect.overwriteImage',
  'ankiConnect.mediaInsertMode',
  'ankiConnect.highlightWord',
  'ankiConnect.notificationType',
  'ankiConnect.autoUpdateNewCards',
  'ankiConnect.nPlusOne.highlightEnabled',
  'ankiConnect.nPlusOne.refreshMinutes',
  'ankiConnect.nPlusOne.matchMode',
  'ankiConnect.nPlusOne.decks',
  'ankiConnect.nPlusOne.knownWord',
  'ankiConnect.nPlusOne.nPlusOne',
  'ankiConnect.knownWords.color',
  'ankiConnect.behavior.nPlusOneHighlightEnabled',
  'ankiConnect.behavior.nPlusOneRefreshMinutes',
  'ankiConnect.behavior.nPlusOneMatchMode',
  'ankiConnect.isLapis.sentenceCardSentenceField',
  'ankiConnect.isLapis.sentenceCardAudioField',
  'ankiConnect.fields.translation',
  'controller.bindings',
  'controller.preferredGamepadId',
  'controller.preferredGamepadLabel',
  'controller.profiles',
  'youtubeSubgen.primarySubLanguages',
  'anilist.characterDictionary.enabled',
  'anilist.characterDictionary.refreshTtlHours',
  'anilist.characterDictionary.evictionPolicy',
  'anilist.characterDictionary.profileScope',
  'jellyfin.accessToken',
  'jellyfin.userId',
  'jellyfin.defaultLibraryId',
  'jellyfin.directPlayContainers',
  'controller.buttonIndices',
  'shortcuts.multiCopyTimeoutMs',
  'subtitleSidebar.toggleKey',
  'jellyfin.recentServers',
] as const;

const EXCLUDED_PREFIXES = [
  'ai',
  'ankiConnect.ai',
  'controller.buttonIndices',
  'youtubeSubgen',
] as const;

const JSON_OBJECT_FIELDS = new Set([
  'keybindings',
  'controller.bindings',
  'controller.profiles',
  'ankiConnect.knownWords.decks',
  'subtitleStyle.css',
  'subtitleStyle.secondary.css',
  'subtitleSidebar.css',
]);

export const SECRET_PATHS = new Set(['ai.apiKey', 'jimaku.apiKey', 'anilist.accessToken']);

const COLOR_SUFFIXES = new Set(['Color', 'color', 'backgroundColor', 'singleColor']);
const SUBTITLE_CSS_MANAGED_CONFIG_PATHS = new Set([
  ...getSubtitleCssManagedConfigPaths('primary'),
  ...getSubtitleCssManagedConfigPaths('secondary'),
  ...getSubtitleCssManagedConfigPaths('sidebar'),
]);

const OPTION_BY_PATH = new Map(CONFIG_OPTION_REGISTRY.map((entry) => [entry.path, entry]));

const CATEGORY_ORDER: ConfigSettingsCategory[] = [
  'appearance',
  'behavior',
  'mining-anki',
  'input',
  'integrations',
  'tracking-app',
  'advanced',
];

const SECTION_ORDER = new Map<string, number>(
  [
    'Annotation Display',
    'Known Words',
    'N+1',
    'Frequency Highlighting',
    'Primary Subtitle Appearance',
    'Secondary Subtitle Appearance',
    'Subtitle Sidebar Appearance',
    'Playback Behavior',
    'Subtitle Behavior',
    'Subtitle Sidebar Behavior',
    'YouTube Playback Settings',
    'mpv Playback',
    'AnkiConnect',
    'Note Fields',
    'Media Capture',
    'Kiku/Lapis Features',
    'Anki AI',
    'AnkiConnect Proxy',
    'Jimaku',
    'Subtitle Sync',
    'MPV Keybindings',
    'Overlay Shortcuts',
    'Controller',
    'Annotation WebSocket',
    'WebSocket server',
    'AniList',
    'Character Dictionary',
    'Discord Rich Presence',
    'Jellyfin',
    'Texthooker',
    'Yomitan',
    'Stats dashboard',
    'Startup warmups',
    'Logging',
    'Updates',
    'Notifications',
    'Immersion tracking',
  ].map((section, index) => [section, index]),
);

const PATH_ORDER = new Map<string, number>(
  [
    'ankiConnect.enabled',
    'ankiConnect.deck',
    'ankiConnect.proxy.enabled',
    'ankiConnect.isLapis.enabled',
    'ankiConnect.isKiku.enabled',
    'subtitleStyle.fontColor',
    'subtitleStyle.backgroundColor',
    'subtitleStyle.hoverTokenColor',
    'subtitleStyle.hoverTokenBackgroundColor',
    'subtitleStyle.css',
    'subtitleStyle.primaryDefaultMode',
    'subtitleStyle.primaryVisibleOnYomitanPopup',
    'subtitleStyle.secondary.fontColor',
    'subtitleStyle.secondary.backgroundColor',
    'subtitleStyle.secondary.css',
    'subtitleSidebar.css',
    'secondarySub.defaultMode',
    'secondarySub.secondarySubLanguages',
    'mpv.autoStartSubMiner',
    'auto_start_overlay',
    'mpv.pauseUntilOverlayReady',
    'mpv.socketPath',
    'mpv.backend',
    'mpv.subminerBinaryPath',
    'mpv.aniskipEnabled',
    'mpv.profile',
    'mpv.launchMode',
    'mpv.executablePath',
    'mpv.aniskipButtonKey',
    'logging.level',
    'logging.rotation',
    'logging.files.app',
    'logging.files.launcher',
    'logging.files.mpv',
  ].map((path, index) => [path, index]),
);

const SUBSECTION_ORDER = new Map<string, number>(
  [
    'Known Words',
    'N+1',
    'JLPT',
    'Frequency Highlighting',
    'Character Names',
    'Mining & Clipboard',
    'Toggle & Visibility',
    'Open Panels',
    'Playback',
    'Default Fold State',
  ].map((subsection, index) => [subsection, index]),
);

const LABEL_OVERRIDES: Record<string, string> = {
  'ankiConnect.knownWords.highlightEnabled': 'Enabled',
  'ankiConnect.nPlusOne.enabled': 'Enabled',
  'ankiConnect.isLapis.enabled': 'Enable Lapis Features',
  'ankiConnect.isKiku.enabled': 'Enable Kiku Features',
  'stats.toggleKey': 'Toggle Stats Overlay',
  'shortcuts.openCharacterDictionaryManager': 'Open Character Dictionary Manager',
  'subtitleSidebar.pauseVideoOnHover': 'Pause Video On Hover - Sidebar',
  'subtitleStyle.autoPauseVideoOnHover': 'Pause Video On Hover - Subtitles',
  'subtitleStyle.autoPauseVideoOnYomitanPopup': 'Pause Video On Yomitan Popup',
  'subtitleStyle.primaryVisibleOnYomitanPopup': 'Keep Primary Visible On Yomitan Popup',
  'subtitleStyle.primaryDefaultMode': 'Primary Subtitle Visibility Mode',
  'subtitleStyle.frequencyDictionary.mode': 'Frequency Mode',
  'subtitleStyle.css': 'CSS Declarations',
  'subtitleStyle.secondary.css': 'CSS Declarations',
  'subtitleSidebar.css': 'CSS Declarations',
  'secondarySub.defaultMode': 'Secondary Subtitle Visibility Mode',
  'subtitlePosition.yPercent': 'Subtitle Position',
  'mpv.executablePath': 'mpv Executable Path',
  'mpv.subminerBinaryPath': 'SubMiner Binary Path',
  'mpv.socketPath': 'mpv IPC Socket Path',
  'mpv.profile': 'mpv Profile',
  'mpv.autoStartSubMiner': 'Auto-start SubMiner',
  'mpv.pauseUntilOverlayReady': 'Pause Until Overlay Ready',
  'mpv.aniskipEnabled': 'Enable AniSkip',
  'mpv.aniskipButtonKey': 'AniSkip Button Key',
  'discordPresence.updateIntervalMs': 'Update Interval (ms)',
};

const DESCRIPTION_OVERRIDES: Record<string, string> = {
  'ankiConnect.pollingRate':
    'Polling interval in milliseconds. Ignored while the local AnkiConnect proxy is enabled because push-based enrichment is used instead.',
  'ankiConnect.isKiku.enabled':
    'Enable Kiku-specific mining behavior. Kiku supersedes Lapis: Lapis features still work, and Kiku adds duplicate handling and field grouping.',
  'ankiConnect.isLapis.enabled':
    'Enable Lapis-specific mining behavior and sentence-card model targeting. When Kiku is enabled, Lapis features still work and Kiku-specific features are added on top.',
  'ankiConnect.isLapis.sentenceCardModel':
    'Anki note type used for Lapis sentence cards. Select from note types reported by AnkiConnect.',
  'subtitleStyle.css':
    'CSS declarations applied to primary subtitles. Includes color, background-color, and all font properties.',
  'subtitleStyle.secondary.css':
    'CSS declarations applied to secondary subtitles. Includes color, background-color, and all font properties.',
  'subtitleSidebar.css':
    'CSS declarations applied to the subtitle sidebar. Includes color, background-color, all font properties, and sidebar CSS variables.',
  'subtitleStyle.primaryVisibleOnYomitanPopup':
    'When primary subtitles are in hover mode, keep the primary subtitle bar visible while a Yomitan popup is open.',
  'websocket.enabled':
    'Built-in subtitle WebSocket server mode. Auto starts the built-in server only when mpv_websocket is not detected; otherwise it defers to the plugin.',
  'discordPresence.updateIntervalMs':
    'Minimum interval between presence payload updates, in milliseconds.',
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function pathStartsWith(path: string, prefix: string): boolean {
  return path === prefix || path.startsWith(`${prefix}.`);
}

function isLegacyHidden(path: string): boolean {
  return (
    LEGACY_HIDDEN_CONFIG_PATHS.some((hiddenPath) => pathStartsWith(path, hiddenPath)) ||
    EXCLUDED_PREFIXES.some((prefix) => pathStartsWith(path, prefix))
  );
}

function flattenConfigLeaves(value: unknown, prefix = ''): Leaf[] {
  if (JSON_OBJECT_FIELDS.has(prefix)) {
    return [{ path: prefix, value }];
  }

  if (Array.isArray(value)) {
    return [{ path: prefix, value }];
  }

  if (isRecord(value)) {
    const entries = Object.entries(value).filter(([, child]) => child !== undefined);
    if (entries.length === 0) {
      return [{ path: prefix, value }];
    }
    return entries.flatMap(([key, child]) =>
      flattenConfigLeaves(child, prefix ? `${prefix}.${key}` : key),
    );
  }

  return prefix ? [{ path: prefix, value }] : [];
}

function humanizePath(path: string): string {
  const override = LABEL_OVERRIDES[path];
  const fallback = override ?? computeHumanizedLabel(path);
  return i18n.t(`optLabel.${path}`, undefined, fallback);
}

function computeHumanizedLabel(path: string): string {
  const key = path.split('.').at(-1) ?? path;
  const spaced = key
    .replace(/_/g, ' ')
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .replace(/\bai\b/i, 'AI')
    .replace(/\bmpv\b/i, 'mpv')
    .replace(/\byomitan\b/i, 'Yomitan')
    .replace(/\bjimaku\b/i, 'Jimaku')
    .replace(/\banilist\b/i, 'AniList')
    .replace(/\banki\b/i, 'Anki');
  return spaced.charAt(0).toUpperCase() + spaced.slice(1);
}

function categoryAndSection(path: string): { category: ConfigSettingsCategory; section: string } {
  if (
    path === 'subtitleStyle.autoPauseVideoOnHover' ||
    path === 'subtitleStyle.autoPauseVideoOnYomitanPopup' ||
    path === 'subtitleSidebar.pauseVideoOnHover'
  ) {
    return { category: 'behavior', section: 'Playback Behavior' };
  }
  if (path === 'subtitleStyle.preserveLineBreaks') {
    return { category: 'behavior', section: 'Subtitle Behavior' };
  }
  if (
    path === 'ankiConnect.knownWords.addMinedWordsImmediately' ||
    path === 'ankiConnect.knownWords.decks' ||
    path === 'ankiConnect.knownWords.matchMode' ||
    path === 'ankiConnect.knownWords.refreshMinutes'
  ) {
    return { category: 'behavior', section: 'Known Words' };
  }
  if (path === 'ankiConnect.nPlusOne.minSentenceWords') {
    return { category: 'behavior', section: 'N+1' };
  }
  if (
    path === 'subtitleStyle.frequencyDictionary.matchMode' ||
    path === 'subtitleStyle.frequencyDictionary.mode' ||
    path === 'subtitleStyle.frequencyDictionary.sourcePath' ||
    path === 'subtitleStyle.frequencyDictionary.topX'
  ) {
    return { category: 'behavior', section: 'Frequency Highlighting' };
  }
  if (
    path.startsWith('ankiConnect.knownWords.') ||
    path.startsWith('ankiConnect.nPlusOne.') ||
    path.startsWith('subtitleStyle.frequencyDictionary.') ||
    path.startsWith('subtitleStyle.jlptColors.') ||
    path === 'subtitleStyle.enableJlpt' ||
    path === 'subtitleStyle.knownWordColor' ||
    path === 'subtitleStyle.nPlusOneColor' ||
    path === 'subtitleStyle.nameMatchEnabled' ||
    path === 'subtitleStyle.nameMatchImagesEnabled' ||
    path === 'subtitleStyle.nameMatchColor'
  ) {
    return { category: 'appearance', section: 'Annotation Display' };
  }
  if (path.startsWith('subtitleStyle.secondary.')) {
    return { category: 'appearance', section: 'Secondary Subtitle Appearance' };
  }
  if (
    path === 'subtitleStyle.primaryDefaultMode' ||
    path === 'subtitleStyle.primaryVisibleOnYomitanPopup'
  ) {
    return { category: 'behavior', section: 'Subtitle Behavior' };
  }
  if (path.startsWith('subtitleStyle.')) {
    return { category: 'appearance', section: 'Primary Subtitle Appearance' };
  }
  if (path.startsWith('subtitleSidebar.')) {
    const sidebarBehaviorPaths = new Set([
      'subtitleSidebar.enabled',
      'subtitleSidebar.autoOpen',
      'subtitleSidebar.autoScroll',
      'subtitleSidebar.layout',
    ]);
    return sidebarBehaviorPaths.has(path)
      ? { category: 'behavior', section: 'Subtitle Sidebar Behavior' }
      : { category: 'appearance', section: 'Subtitle Sidebar Appearance' };
  }
  if (path.startsWith('subtitlePosition.') || path.startsWith('secondarySub.')) {
    return { category: 'behavior', section: 'Subtitle Behavior' };
  }
  if (path.startsWith('ankiConnect.fields.')) {
    return { category: 'mining-anki', section: 'Note Fields' };
  }
  if (path.startsWith('ankiConnect.media.')) {
    return { category: 'mining-anki', section: 'Media Capture' };
  }
  if (path.startsWith('ankiConnect.isKiku.') || path.startsWith('ankiConnect.isLapis.')) {
    return { category: 'mining-anki', section: 'Kiku/Lapis Features' };
  }
  if (path.startsWith('ankiConnect.ai.')) {
    return { category: 'mining-anki', section: 'Anki AI' };
  }
  if (path.startsWith('ankiConnect.proxy.')) {
    return { category: 'mining-anki', section: 'AnkiConnect Proxy' };
  }
  if (path.startsWith('ankiConnect.')) {
    return { category: 'mining-anki', section: 'AnkiConnect' };
  }
  if (
    path === 'auto_start_overlay' ||
    path === 'mpv.autoStartSubMiner' ||
    path === 'mpv.pauseUntilOverlayReady'
  ) {
    return { category: 'behavior', section: 'Playback Behavior' };
  }
  if (path.startsWith('notifications.')) {
    return { category: 'behavior', section: 'Notifications' };
  }
  if (path === 'mpv.aniskipButtonKey') {
    return { category: 'input', section: 'Overlay Shortcuts' };
  }
  if (path.startsWith('mpv.') || path.startsWith('youtube.')) {
    return { category: 'behavior', section: topSection(path) };
  }
  if (path.startsWith('jimaku.')) {
    return { category: 'integrations', section: topSection(path) };
  }
  if (path.startsWith('subsync.')) {
    return { category: 'integrations', section: topSection(path) };
  }
  if (path === 'stats.toggleKey' || path === 'stats.markWatchedKey') {
    return { category: 'input', section: 'Overlay Shortcuts' };
  }
  if (path.startsWith('shortcuts.')) {
    return { category: 'input', section: 'Overlay Shortcuts' };
  }
  if (path === 'keybindings') {
    return { category: 'input', section: 'MPV Keybindings' };
  }
  if (path.startsWith('controller.')) {
    return { category: 'input', section: 'Controller' };
  }
  if (
    path.startsWith('ai.') ||
    path.startsWith('yomitan.') ||
    path.startsWith('jellyfin.') ||
    path.startsWith('discordPresence.') ||
    path.startsWith('websocket.') ||
    path.startsWith('annotationWebsocket.') ||
    path.startsWith('texthooker.')
  ) {
    return { category: 'integrations', section: topSection(path) };
  }
  if (path.startsWith('anilist.characterDictionary.')) {
    return { category: 'integrations', section: 'Character Dictionary' };
  }
  if (path.startsWith('anilist.')) {
    return { category: 'integrations', section: 'AniList' };
  }
  if (
    path.startsWith('immersionTracking.') ||
    path.startsWith('stats.') ||
    path.startsWith('updates.') ||
    path.startsWith('startupWarmups.') ||
    path.startsWith('logging.')
  ) {
    return { category: 'tracking-app', section: topSection(path) };
  }
  return { category: 'advanced', section: 'Advanced' };
}

function topSection(path: string): string {
  const top = path.split('.')[0] ?? path;
  const labels: Record<string, string> = {
    ai: 'Shared AI provider',
    anilist: 'AniList',
    annotationWebsocket: 'Annotation WebSocket',
    discordPresence: 'Discord Rich Presence',
    immersionTracking: 'Immersion tracking',
    jimaku: 'Jimaku',
    jellyfin: 'Jellyfin',
    logging: 'Logging',
    mpv: 'mpv Playback',
    stats: 'Stats dashboard',
    startupWarmups: 'Startup warmups',
    notifications: 'Notifications',
    subsync: 'Subtitle Sync',
    texthooker: 'Texthooker',
    updates: 'Updates',
    websocket: 'WebSocket server',
    yomitan: 'Yomitan',
    youtube: 'YouTube Playback Settings',
    youtubeSubgen: 'YouTube subtitle generation',
    auto_start_overlay: 'Playback Behavior',
  };
  return labels[top] ?? humanizePath(top);
}

function controlForPath(path: string, value: unknown): ConfigSettingsControl {
  if (SECRET_PATHS.has(path)) return 'secret';
  if (getSubtitleCssScopeForPath(path)) return 'css-declarations';
  if (path === 'keybindings') return 'mpv-keybindings';
  if (path === 'ankiConnect.deck') return 'anki-deck';
  if (path === 'ankiConnect.knownWords.decks') return 'known-words-decks';
  if (path === 'ankiConnect.isLapis.sentenceCardModel') return 'anki-note-type';
  if (path.startsWith('ankiConnect.fields.')) return 'anki-field';
  if (path.startsWith('shortcuts.'))
    return path.endsWith('multiCopyTimeoutMs') ? 'number' : 'keyboard-shortcut';
  if (path === 'mpv.aniskipButtonKey') return 'mpv-key';
  if (
    path === 'subtitleSidebar.toggleKey' ||
    path === 'stats.toggleKey' ||
    path === 'stats.markWatchedKey'
  ) {
    return 'key-code';
  }
  if (path.startsWith('subtitleStyle.jlptColors.')) return 'color';
  if (path === 'subtitleStyle.frequencyDictionary.bandedColors') return 'color-list';
  if (OPTION_BY_PATH.get(path)?.enumValues?.length) return 'select';
  if (JSON_OBJECT_FIELDS.has(path)) return 'json';
  if (Array.isArray(value)) return 'string-list';
  if (typeof value === 'boolean') return 'boolean';
  if (typeof value === 'number') return 'number';
  if (typeof value === 'string') {
    const leaf = path.split('.').at(-1) ?? path;
    if ([...COLOR_SUFFIXES].some((suffix) => leaf.endsWith(suffix))) return 'color';
    if (leaf.toLowerCase().includes('prompt')) return 'textarea';
    return 'text';
  }
  return 'json';
}

function subsectionForPath(path: string): string | undefined {
  if (path === 'ankiConnect.knownWords.highlightEnabled') return 'Known Words';
  if (path === 'ankiConnect.nPlusOne.enabled') return 'N+1';
  if (path === 'subtitleStyle.knownWordColor') return 'Known Words';
  if (path === 'subtitleStyle.nPlusOneColor') return 'N+1';
  if (path === 'subtitleStyle.enableJlpt' || path.startsWith('subtitleStyle.jlptColors.')) {
    return 'JLPT';
  }
  if (
    path === 'subtitleStyle.frequencyDictionary.enabled' ||
    path === 'subtitleStyle.frequencyDictionary.singleColor' ||
    path === 'subtitleStyle.frequencyDictionary.bandedColors'
  ) {
    return 'Frequency Highlighting';
  }
  if (
    path === 'subtitleStyle.nameMatchEnabled' ||
    path === 'subtitleStyle.nameMatchImagesEnabled' ||
    path === 'subtitleStyle.nameMatchColor'
  ) {
    return 'Character Names';
  }
  if (path === 'anilist.characterDictionary.collapsibleSections.description') {
    return 'Default Fold State';
  }
  if (path === 'anilist.characterDictionary.collapsibleSections.characterInformation') {
    return 'Default Fold State';
  }
  if (path === 'anilist.characterDictionary.collapsibleSections.voicedBy') {
    return 'Default Fold State';
  }
  if (path === 'stats.toggleKey' || path === 'stats.markWatchedKey') {
    return 'Toggle & Visibility';
  }
  if (path === 'mpv.aniskipButtonKey') {
    return 'Playback';
  }
  if (path.startsWith('shortcuts.')) {
    const leaf = path.split('.').at(-1) ?? '';
    if (
      leaf === 'copySubtitle' ||
      leaf === 'copySubtitleMultiple' ||
      leaf === 'mineSentence' ||
      leaf === 'mineSentenceMultiple' ||
      leaf === 'updateLastCardFromClipboard' ||
      leaf === 'triggerFieldGrouping' ||
      leaf === 'markAudioCard'
    ) {
      return 'Mining & Clipboard';
    }
    if (
      leaf === 'toggleVisibleOverlayGlobal' ||
      leaf === 'toggleSubtitleSidebar' ||
      leaf === 'toggleNotificationHistory' ||
      leaf === 'toggleSecondarySub' ||
      leaf === 'toggleStatsOverlay' ||
      leaf === 'markWatched'
    ) {
      return 'Toggle & Visibility';
    }
    if (
      leaf === 'openCharacterDictionaryManager' ||
      leaf === 'openRuntimeOptions' ||
      leaf === 'openJimaku' ||
      leaf === 'openSessionHelp' ||
      leaf === 'openControllerSelect' ||
      leaf === 'openControllerDebug'
    ) {
      return 'Open Panels';
    }
    if (leaf === 'triggerSubsync') return 'Playback';
    return undefined;
  }
  return undefined;
}

function isFeatureToggle(field: ConfigSettingsField): boolean {
  if (field.control !== 'boolean') return false;
  const leaf = field.configPath.split('.').at(-1) ?? field.configPath;
  return (
    leaf === 'enabled' ||
    leaf.startsWith('enable') ||
    leaf.endsWith('Enabled') ||
    field.label.startsWith('Enable ')
  );
}

function fieldTypeRank(field: ConfigSettingsField): number {
  if (field.configPath === 'subtitleStyle.primaryVisibleOnYomitanPopup') return 2;
  if (field.configPath === 'ankiConnect.deck') return 1;
  if (field.control !== 'boolean') return 2;
  return isFeatureToggle(field) ? 0 : 1;
}

function compareFields(a: ConfigSettingsField, b: ConfigSettingsField): number {
  const category = CATEGORY_ORDER.indexOf(a.category) - CATEGORY_ORDER.indexOf(b.category);
  if (category !== 0) return category;

  const section =
    (SECTION_ORDER.get(a.section) ?? Number.MAX_SAFE_INTEGER) -
    (SECTION_ORDER.get(b.section) ?? Number.MAX_SAFE_INTEGER);
  if (section !== 0) return section;

  const sectionName = a.section.localeCompare(b.section);
  if (sectionName !== 0) return sectionName;

  const aSubOrder =
    a.subsection === undefined
      ? -1
      : (SUBSECTION_ORDER.get(a.subsection) ?? Number.MAX_SAFE_INTEGER);
  const bSubOrder =
    b.subsection === undefined
      ? -1
      : (SUBSECTION_ORDER.get(b.subsection) ?? Number.MAX_SAFE_INTEGER);
  const subsection = aSubOrder - bSubOrder;
  if (subsection !== 0) return subsection;

  const subsectionName = (a.subsection ?? '').localeCompare(b.subsection ?? '');
  if (subsectionName !== 0) return subsectionName;

  const type = fieldTypeRank(a) - fieldTypeRank(b);
  if (type !== 0) return type;

  const pathOrder =
    (PATH_ORDER.get(a.configPath) ?? Number.MAX_SAFE_INTEGER) -
    (PATH_ORDER.get(b.configPath) ?? Number.MAX_SAFE_INTEGER);
  if (pathOrder !== 0) return pathOrder;

  const label = a.label.localeCompare(b.label);
  if (label !== 0) return label;
  return a.configPath.localeCompare(b.configPath);
}

function restartBehaviorForPath(path: string): ConfigSettingsRestartBehavior {
  if (
    path === 'keybindings' ||
    pathStartsWith(path, 'shortcuts') ||
    pathStartsWith(path, 'subtitleStyle') ||
    pathStartsWith(path, 'subtitleSidebar') ||
    path === 'secondarySub.defaultMode' ||
    path === 'ankiConnect.deck' ||
    path === 'ankiConnect.ai.enabled' ||
    path === 'ankiConnect.behavior.autoUpdateNewCards' ||
    path === 'ankiConnect.knownWords.highlightEnabled' ||
    path === 'ankiConnect.knownWords.refreshMinutes' ||
    path === 'ankiConnect.knownWords.addMinedWordsImmediately' ||
    path === 'ankiConnect.knownWords.matchMode' ||
    path === 'ankiConnect.knownWords.decks' ||
    path === 'ankiConnect.nPlusOne.enabled' ||
    path === 'ankiConnect.nPlusOne.minSentenceWords' ||
    path === 'ankiConnect.fields.word' ||
    path === 'ankiConnect.fields.audio' ||
    path === 'ankiConnect.fields.image' ||
    path === 'ankiConnect.fields.sentence' ||
    path === 'ankiConnect.fields.miscInfo' ||
    path === 'ankiConnect.isLapis.sentenceCardModel' ||
    path === 'ankiConnect.isKiku.fieldGrouping' ||
    path === 'mpv.aniskipEnabled' ||
    path === 'mpv.aniskipButtonKey' ||
    path === 'stats.toggleKey' ||
    path === 'stats.markWatchedKey' ||
    path === 'logging.level' ||
    path === 'logging.rotation' ||
    pathStartsWith(path, 'logging.files') ||
    pathStartsWith(path, 'notifications') ||
    path === 'youtube.primarySubLanguages' ||
    pathStartsWith(path, 'jimaku') ||
    pathStartsWith(path, 'subsync')
  ) {
    return 'hot-reload';
  }
  return 'restart';
}

function fieldForLeaf(leaf: Leaf): ConfigSettingsField {
  const option = OPTION_BY_PATH.get(leaf.path);
  const { category, section } = categoryAndSection(leaf.path);
  const description = DESCRIPTION_OVERRIDES[leaf.path] ?? option?.description;
  return {
    id: leaf.path,
    label: option?.path === leaf.path ? humanizePath(leaf.path) : humanizePath(leaf.path),
    description: description ?? `${humanizePath(leaf.path)} setting.`,
    configPath: leaf.path,
    category,
    section,
    ...(subsectionForPath(leaf.path) ? { subsection: subsectionForPath(leaf.path) } : {}),
    control: controlForPath(leaf.path, leaf.value),
    defaultValue: leaf.value,
    ...(option?.settingsEnumValues || option?.enumValues
      ? { enumValues: option.settingsEnumValues ?? option.enumValues }
      : {}),
    restartBehavior: restartBehaviorForPath(leaf.path),
    advanced:
      leaf.path.startsWith('controller.') ||
      leaf.path.startsWith('immersionTracking.retention.') ||
      leaf.path.startsWith('youtubeSubgen.'),
    secret: SECRET_PATHS.has(leaf.path),
    settingsHidden: SUBTITLE_CSS_MANAGED_CONFIG_PATHS.has(leaf.path),
  };
}

export function buildConfigSettingsRegistry(
  defaultConfig: ResolvedConfig = DEFAULT_CONFIG,
): ConfigSettingsField[] {
  const leaves = flattenConfigLeaves(defaultConfig).filter((leaf) => !isLegacyHidden(leaf.path));
  return leaves.map(fieldForLeaf).sort(compareFields);
}

export function getConfigSettingsCoverage(
  defaultConfig: ResolvedConfig,
  fields: ConfigSettingsField[],
): { uncoveredDefaultPaths: string[] } {
  const visibleFields = fields.filter((field) => !field.legacyHidden);
  const uncoveredDefaultPaths = flattenConfigLeaves(defaultConfig)
    .map((leaf) => leaf.path)
    .filter((path) => !isLegacyHidden(path))
    .filter(
      (path) =>
        !visibleFields.some(
          (field) =>
            field.configPath === path ||
            (field.control === 'json' && pathStartsWith(path, field.configPath)),
        ),
    )
    .sort();

  return { uncoveredDefaultPaths };
}

export function getConfigValueAtPath(root: unknown, path: string): unknown {
  let current = root;
  for (const segment of path.split('.')) {
    if (!isRecord(current)) return undefined;
    current = current[segment];
  }
  return current;
}
