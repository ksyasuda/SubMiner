import { RawConfig, ResolvedConfig } from '../types/config';
import { CORE_DEFAULT_CONFIG } from './definitions/defaults-core';
import { IMMERSION_DEFAULT_CONFIG } from './definitions/defaults-immersion';
import { INTEGRATIONS_DEFAULT_CONFIG } from './definitions/defaults-integrations';
import { STATS_DEFAULT_CONFIG } from './definitions/defaults-stats';
import { SUBTITLE_DEFAULT_CONFIG } from './definitions/defaults-subtitle';
import { buildCoreConfigOptionRegistry } from './definitions/options-core';
import { buildImmersionConfigOptionRegistry } from './definitions/options-immersion';
import { buildIntegrationConfigOptionRegistry } from './definitions/options-integrations';
import { buildStatsConfigOptionRegistry } from './definitions/options-stats';
import { buildSubtitleConfigOptionRegistry } from './definitions/options-subtitle';
import { buildRuntimeOptionRegistry } from './definitions/runtime-options';
import { CONFIG_TEMPLATE_SECTIONS } from './definitions/template-sections';

export { DEFAULT_KEYBINDINGS, SPECIAL_COMMANDS } from './definitions/shared';
export type {
  ConfigOptionRegistryEntry,
  ConfigTemplateSection,
  ConfigValueKind,
  RuntimeOptionRegistryEntry,
} from './definitions/shared';

const {
  subtitlePosition,
  keybindings,
  websocket,
  annotationWebsocket,
  logging,
  texthooker,
  controller,
  shortcuts,
  secondarySub,
  youtube,
  subsync,
  startupWarmups,
  updates,
  notifications,
  auto_start_overlay,
} = CORE_DEFAULT_CONFIG;
const {
  ankiConnect,
  anime,
  jimaku,
  tsukihime,
  anilist,
  mpv,
  yomitan,
  jellyfin,
  discordPresence,
  ai,
  youtubeSubgen,
} = INTEGRATIONS_DEFAULT_CONFIG;
const { subtitleStyle, subtitleSidebar } = SUBTITLE_DEFAULT_CONFIG;
const { immersionTracking } = IMMERSION_DEFAULT_CONFIG;
const { stats } = STATS_DEFAULT_CONFIG;

export const DEFAULT_CONFIG: ResolvedConfig = {
  subtitlePosition,
  keybindings,
  websocket,
  annotationWebsocket,
  logging,
  texthooker,
  controller,
  ankiConnect,
  shortcuts,
  secondarySub,
  youtube,
  subsync,
  startupWarmups,
  updates,
  notifications,
  subtitleStyle,
  subtitleSidebar,
  auto_start_overlay,
  anime,
  jimaku,
  tsukihime,
  anilist,
  mpv,
  yomitan,
  jellyfin,
  discordPresence,
  ai,
  youtubeSubgen,
  immersionTracking,
  stats,
};

export const DEFAULT_ANKI_CONNECT_CONFIG = DEFAULT_CONFIG.ankiConnect;

export const RUNTIME_OPTION_REGISTRY = buildRuntimeOptionRegistry(DEFAULT_CONFIG);

export const CONFIG_OPTION_REGISTRY = [
  ...buildCoreConfigOptionRegistry(DEFAULT_CONFIG),
  ...buildSubtitleConfigOptionRegistry(DEFAULT_CONFIG),
  ...buildIntegrationConfigOptionRegistry(DEFAULT_CONFIG, RUNTIME_OPTION_REGISTRY),
  ...buildImmersionConfigOptionRegistry(DEFAULT_CONFIG),
  ...buildStatsConfigOptionRegistry(DEFAULT_CONFIG),
];

export { CONFIG_TEMPLATE_SECTIONS };

export function deepCloneConfig(config: ResolvedConfig): ResolvedConfig {
  return structuredClone(config);
}

export function deepMergeRawConfig(base: RawConfig, patch: RawConfig): RawConfig {
  const clone = structuredClone(base) as Record<string, unknown>;
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
