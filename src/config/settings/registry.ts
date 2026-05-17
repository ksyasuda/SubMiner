import type { ResolvedConfig } from '../../types/config';
import type {
  ConfigSettingsCategory,
  ConfigSettingsControl,
  ConfigSettingsField,
  ConfigSettingsRestartBehavior,
} from '../../types/settings';
import { CONFIG_OPTION_REGISTRY, DEFAULT_CONFIG } from '../definitions';

type Leaf = {
  path: string;
  value: unknown;
};

export const LEGACY_HIDDEN_CONFIG_PATHS = [
  'ankiConnect.deck',
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
  'controller.bindings',
  'controller.preferredGamepadId',
  'controller.preferredGamepadLabel',
  'controller.profiles',
  'youtubeSubgen.primarySubLanguages',
  'anilist.characterDictionary.refreshTtlHours',
  'anilist.characterDictionary.evictionPolicy',
  'jellyfin.accessToken',
  'jellyfin.userId',
  'jellyfin.clientName',
  'jellyfin.clientVersion',
  'jellyfin.defaultLibraryId',
  'jellyfin.deviceId',
  'controller.buttonIndices',
] as const;

const EXCLUDED_PREFIXES = ['controller.buttonIndices', 'youtubeSubgen'] as const;

const JSON_OBJECT_FIELDS = new Set([
  'keybindings',
  'controller.bindings',
  'controller.profiles',
  'ankiConnect.knownWords.decks',
]);

const SECRET_PATHS = new Set(['ai.apiKey', 'jimaku.apiKey', 'anilist.accessToken']);

const COLOR_SUFFIXES = new Set(['Color', 'color', 'backgroundColor', 'singleColor', 'nPlusOne']);

const OPTION_BY_PATH = new Map(CONFIG_OPTION_REGISTRY.map((entry) => [entry.path, entry]));

const CATEGORY_ORDER: ConfigSettingsCategory[] = [
  'appearance',
  'behavior',
  'mining-anki',
  'playback-sources',
  'input',
  'integrations',
  'tracking-app',
  'advanced',
];

const SECTION_ORDER = new Map<string, number>(
  [
    'Annotation Display',
    'Primary Subtitle Appearance',
    'Secondary Subtitle Appearance',
    'Subtitle Sidebar Appearance',
    'Playback Pause Behavior',
    'Subtitle Behavior',
    'Subtitle Sidebar Behavior',
    'Note Fields',
    'Media Capture',
    'Kiku Features And Lapis Features',
    'Anki AI',
    'AnkiConnect Proxy',
    'AnkiConnect',
    'MPV Keybindings',
    'Overlay Shortcuts',
    'Controller',
    'Character Dictionary',
  ].map((section, index) => [section, index]),
);

const PATH_ORDER = new Map<string, number>(
  [
    'ankiConnect.enabled',
    'ankiConnect.proxy.enabled',
    'ankiConnect.isLapis.enabled',
    'ankiConnect.isKiku.enabled',
  ].map((path, index) => [path, index]),
);

const SUBSECTION_ORDER = new Map<string, number>(
  ['Known Words', 'N+1', 'JLPT', 'Frequency Dictionary', 'Character Names'].map(
    (subsection, index) => [subsection, index],
  ),
);

const LABEL_OVERRIDES: Record<string, string> = {
  'subtitleSidebar.pauseVideoOnHover': 'Pause Video On Hover - Sidebar',
  'subtitleStyle.autoPauseVideoOnHover': 'Pause Video On Hover - Subtitles',
  'subtitleStyle.autoPauseVideoOnYomitanPopup': 'Pause Video On Yomitan Popup',
  'subtitleStyle.primaryDefaultMode': 'Primary Subtitle Visibility Mode',
  'secondarySub.defaultMode': 'Secondary Subtitle Visibility Mode',
  'subtitlePosition.yPercent': 'Subtitle Position',
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
  if (override) {
    return override;
  }
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
    return { category: 'behavior', section: 'Playback Pause Behavior' };
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
    path === 'subtitleStyle.nameMatchColor'
  ) {
    return { category: 'appearance', section: 'Annotation Display' };
  }
  if (path.startsWith('subtitleStyle.secondary.')) {
    return { category: 'appearance', section: 'Secondary Subtitle Appearance' };
  }
  if (path === 'subtitleStyle.primaryDefaultMode') {
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
      'subtitleSidebar.toggleKey',
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
    return { category: 'mining-anki', section: 'Kiku Features And Lapis Features' };
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
    path.startsWith('mpv.') ||
    path.startsWith('youtube.') ||
    path.startsWith('youtubeSubgen.') ||
    path.startsWith('jimaku.') ||
    path.startsWith('subsync.')
  ) {
    return { category: 'playback-sources', section: topSection(path) };
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
    path.startsWith('logging.') ||
    path === 'auto_start_overlay'
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
    mpv: 'mpv launcher',
    stats: 'Stats dashboard',
    startupWarmups: 'Startup warmups',
    subsync: 'Auto subtitle sync',
    texthooker: 'Texthooker',
    updates: 'Updates',
    websocket: 'WebSocket server',
    yomitan: 'Yomitan',
    youtube: 'YouTube playback',
    youtubeSubgen: 'YouTube subtitle generation',
    auto_start_overlay: 'Overlay startup',
  };
  return labels[top] ?? humanizePath(top);
}

function controlForPath(path: string, value: unknown): ConfigSettingsControl {
  if (SECRET_PATHS.has(path)) return 'secret';
  if (path === 'keybindings') return 'mpv-keybindings';
  if (path === 'ankiConnect.knownWords.decks') return 'known-words-decks';
  if (path === 'ankiConnect.isLapis.sentenceCardModel') return 'anki-note-type';
  if (path.startsWith('ankiConnect.fields.')) return 'anki-field';
  if (path.startsWith('shortcuts.'))
    return path.endsWith('multiCopyTimeoutMs') ? 'number' : 'keyboard-shortcut';
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
  if (path.startsWith('ankiConnect.knownWords.')) return 'Known Words';
  if (path.startsWith('ankiConnect.nPlusOne.')) return 'N+1';
  if (path === 'subtitleStyle.knownWordColor') return 'Known Words';
  if (path === 'subtitleStyle.nPlusOneColor') return 'N+1';
  if (path === 'subtitleStyle.enableJlpt' || path.startsWith('subtitleStyle.jlptColors.')) {
    return 'JLPT';
  }
  if (path.startsWith('subtitleStyle.frequencyDictionary.')) return 'Frequency Dictionary';
  if (path === 'subtitleStyle.nameMatchEnabled' || path === 'subtitleStyle.nameMatchColor') {
    return 'Character Names';
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

  const subsection =
    (SUBSECTION_ORDER.get(a.subsection ?? '') ?? Number.MAX_SAFE_INTEGER) -
    (SUBSECTION_ORDER.get(b.subsection ?? '') ?? Number.MAX_SAFE_INTEGER);
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
    pathStartsWith(path, 'ankiConnect.ai')
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
    ...(option?.enumValues ? { enumValues: option.enumValues } : {}),
    restartBehavior: restartBehaviorForPath(leaf.path),
    advanced:
      leaf.path.startsWith('controller.') ||
      leaf.path.startsWith('immersionTracking.retention.') ||
      leaf.path.startsWith('youtubeSubgen.'),
    secret: SECRET_PATHS.has(leaf.path),
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
