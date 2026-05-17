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
  'ankiConnect.behavior.nPlusOneHighlightEnabled',
  'ankiConnect.behavior.nPlusOneRefreshMinutes',
  'ankiConnect.behavior.nPlusOneMatchMode',
  'ankiConnect.isLapis.sentenceCardSentenceField',
  'ankiConnect.isLapis.sentenceCardAudioField',
  'youtubeSubgen.primarySubLanguages',
  'anilist.characterDictionary.refreshTtlHours',
  'anilist.characterDictionary.evictionPolicy',
  'jellyfin.accessToken',
  'jellyfin.userId',
  'controller.buttonIndices',
] as const;

const EXCLUDED_PREFIXES = ['controller.buttonIndices'] as const;

const JSON_OBJECT_FIELDS = new Set([
  'keybindings',
  'controller.bindings',
  'controller.profiles',
  'ankiConnect.knownWords.decks',
]);

const SECRET_PATHS = new Set(['ai.apiKey', 'jimaku.apiKey', 'anilist.accessToken']);

const COLOR_SUFFIXES = new Set([
  'Color',
  'color',
  'backgroundColor',
  'singleColor',
  'knownWordColor',
  'nPlusOne',
]);

const OPTION_BY_PATH = new Map(CONFIG_OPTION_REGISTRY.map((entry) => [entry.path, entry]));

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
    return { category: 'viewing', section: 'Playback pause behavior' };
  }
  if (
    path.startsWith('ankiConnect.knownWords.') ||
    path.startsWith('ankiConnect.nPlusOne.') ||
    path.startsWith('subtitleStyle.frequencyDictionary.') ||
    path.startsWith('subtitleStyle.jlptColors.') ||
    path === 'subtitleStyle.enableJlpt' ||
    path === 'subtitleStyle.nameMatchEnabled' ||
    path === 'subtitleStyle.nameMatchColor'
  ) {
    return { category: 'viewing', section: 'Annotation display' };
  }
  if (path.startsWith('subtitleStyle.secondary.')) {
    return { category: 'viewing', section: 'Secondary subtitle appearance' };
  }
  if (path.startsWith('subtitleStyle.')) {
    return { category: 'viewing', section: 'Primary subtitle appearance' };
  }
  if (path.startsWith('subtitleSidebar.')) {
    return { category: 'viewing', section: 'Subtitle sidebar' };
  }
  if (path.startsWith('subtitlePosition.') || path.startsWith('secondarySub.')) {
    return { category: 'viewing', section: 'Subtitle behavior' };
  }
  if (path.startsWith('ankiConnect.fields.')) {
    return { category: 'mining-anki', section: 'Note fields' };
  }
  if (path.startsWith('ankiConnect.media.')) {
    return { category: 'mining-anki', section: 'Media capture' };
  }
  if (path.startsWith('ankiConnect.isKiku.') || path.startsWith('ankiConnect.isLapis.')) {
    return { category: 'mining-anki', section: 'Kiku and Lapis' };
  }
  if (path.startsWith('ankiConnect.ai.')) {
    return { category: 'mining-anki', section: 'Anki AI' };
  }
  if (path.startsWith('ankiConnect.proxy.')) {
    return { category: 'mining-anki', section: 'AnkiConnect proxy' };
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
    return { category: 'input', section: 'Overlay shortcuts' };
  }
  if (path === 'keybindings') {
    return { category: 'input', section: 'MPV keybindings' };
  }
  if (path.startsWith('controller.')) {
    return { category: 'input', section: 'Controller' };
  }
  if (
    path.startsWith('ai.') ||
    path.startsWith('anilist.') ||
    path.startsWith('yomitan.') ||
    path.startsWith('jellyfin.') ||
    path.startsWith('discordPresence.') ||
    path.startsWith('websocket.') ||
    path.startsWith('annotationWebsocket.') ||
    path.startsWith('texthooker.')
  ) {
    return { category: 'integrations', section: topSection(path) };
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
  return {
    id: leaf.path,
    label: option?.path === leaf.path ? humanizePath(leaf.path) : humanizePath(leaf.path),
    description: option?.description ?? `${humanizePath(leaf.path)} setting.`,
    configPath: leaf.path,
    category,
    section,
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
  return leaves.map(fieldForLeaf).sort((a, b) => {
    const category = a.category.localeCompare(b.category);
    if (category !== 0) return category;
    const section = a.section.localeCompare(b.section);
    if (section !== 0) return section;
    return a.configPath.localeCompare(b.configPath);
  });
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
